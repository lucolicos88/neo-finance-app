/**
 * webapp-service.ts
 *
 * Serviço backend para a Web App
 * Fornece dados para o frontend via google.script.run
 */

import { getSheetValues, createSheetIfNotExists, appendRows } from '../shared/sheets-client';
import {
  cacheGet,
  cacheSet,
  cacheRemoveNamespace,
  cacheGetOrLoad,
  CacheNamespace,
  CacheScope,
} from '../shared/cache';
import { combineValidations, validateEnum, validateRequired } from '../shared/validation';
import {
  SHEET_TB_LANCAMENTOS,
  SHEET_TB_EXTRATOS,
  SHEET_REF_FILIAIS,
  SHEET_REF_CANAIS,
  SHEET_REF_CCUSTO,
  SHEET_REF_PLANO_CONTAS,
} from '../config/sheet-mapping';

// ============================================================================
// VIEW RENDERING
// ============================================================================

const SHEET_USUARIOS = 'TB_Usuarios';

const WEBAPP_VIEW_ALLOWLIST = new Set([
  'dashboard',
  'contas-pagar',
  'contas-receber',
  'conciliacao',
  'relatorios',
  'dre',
  'fluxo-caixa',
  'kpis',
  'configuracoes',
]);

function clearReportsCache(): void {
  cacheRemoveNamespace(CacheNamespace.DRE, CacheScope.SCRIPT);
  cacheRemoveNamespace(CacheNamespace.DFC, CacheScope.SCRIPT);
  cacheRemoveNamespace(CacheNamespace.KPI, CacheScope.SCRIPT);
}

function getHeaderIndexMap(sheet: GoogleAppsScript.Spreadsheet.Sheet): Record<string, number> {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  const map: Record<string, number> = {};
  headers.forEach((h, idx) => {
    const key = String(h || '').trim();
    if (!key) return;
    map[key] = idx;
  });
  return map;
}

function findRowByExactValueInColumn(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  columnIndex0: number,
  value: unknown,
  startRow: number = 2
): number | null {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return null;
  const col = columnIndex0 + 1;
  const range = sheet.getRange(startRow, col, lastRow - startRow + 1, 1);
  const found = range.createTextFinder(String(value ?? '')).matchEntireCell(true).findNext();
  return found ? found.getRow() : null;
}

function columnToLetter(col: number): string {
  let temp = col;
  let letter = '';
  while (temp > 0) {
    const modulo = (temp - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    temp = Math.floor((temp - modulo) / 26);
  }
  return letter;
}

function escapeHtml(value: unknown): string {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function isSeedDataEnabled(): boolean {
  return PropertiesService.getScriptProperties().getProperty('ENABLE_SEED_DATA') === 'true';
}

function getRequestingUserEmail(): string {
  return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
}

function sanitizeSheetString(value: unknown): string {
  const MAX_SHEET_CELL_CHARS = 45000;
  let s = String(value ?? '').replace(/\u0000/g, '').trim();
  if (/^[=+\-@]/.test(s)) s = `'${s}`;
  if (s.length > MAX_SHEET_CELL_CHARS) s = s.slice(0, MAX_SHEET_CELL_CHARS - 1) + '…';
  return s;
}

type RequestContext = { __ctx?: boolean; correlationId?: string; view?: string; url?: string } | null;
let activeRequestContext: RequestContext = null;

export function setRequestContext(ctx: RequestContext): void {
  activeRequestContext = ctx && typeof ctx === 'object' ? ctx : null;
}

export function clearRequestContext(): void {
  activeRequestContext = null;
}

function ensureUsuariosSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_USUARIOS);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SHEET_USUARIOS);
  sheet.getRange('A1:H1').setValues([[
    'ID', 'Email', 'Nome', 'Perfil', 'Status', 'Último Acesso', 'Permissões', 'Data Criação'
  ]]);
  sheet.getRange('A1:H1').setFontWeight('bold');
  sheet.setFrozenRows(1);

  const email = getRequestingUserEmail();
  if (email) {
    const id = Utilities.getUuid();
    const now = new Date().toISOString();
    sheet.appendRow([
      id,
      email,
      'Administrador',
      'ADMIN',
      'ATIVO',
      now,
      JSON.stringify(getPermissoesPadrao('ADMIN')),
      now,
    ]);
  }

  return sheet;
}

function getUsuarioByEmail(email: string): Usuario | null {
  if (!email) return null;
  const sheet = ensureUsuariosSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (String(row[1]).toLowerCase() !== email.toLowerCase()) continue;

    return {
      id: String(row[0]),
      email: String(row[1]),
      nome: String(row[2]),
      perfil: String(row[3]) as any,
      status: String(row[4]) as any,
      ultimoAcesso: row[5] ? String(row[5]) : undefined,
      permissoes: row[6] ? JSON.parse(String(row[6])) : getPermissoesPadrao(String(row[3])),
    };
  }

  return null;
}

type PermissionKey = keyof NonNullable<Usuario['permissoes']>;

function requirePermission<T extends { success: boolean; message: string }>(
  permission: PermissionKey,
  action: string
): T | null {
  const email = getRequestingUserEmail();
  const user = getUsuarioByEmail(email);

  if (!user) {
    appendAuditLog('permissionDenied', { permission, action }, false, `Usuário não cadastrado: ${action}`);
    return { success: false, message: `Usuário não cadastrado: ${action}` } as T;
  }
  if (user.status !== 'ATIVO') {
    appendAuditLog('permissionDenied', { permission, action, status: user.status }, false, `Usuário inativo: ${action}`);
    return { success: false, message: `Usuário inativo: ${action}` } as T;
  }
  if (!user.permissoes?.[permission]) {
    appendAuditLog('permissionDenied', { permission, action, perfil: user.perfil }, false, `Sem permissão: ${action}`);
    return { success: false, message: `Sem permissão: ${action}` } as T;
  }

  return null;
}

function enforcePermission(permission: PermissionKey, action: string): void {
  const denied = requirePermission<{ success: boolean; message: string }>(permission, action);
  if (denied) throw new Error(denied.message);
}

const SHEET_AUDIT_LOG = 'TB_AUDIT_LOG';
const MAX_AUDIT_LOG_ROWS = 5000; // exclui header
const AUDIT_LOG_TRIM_BUFFER = 200; // só limpa quando passar do limite + buffer

function appendRowFast(sheet: GoogleAppsScript.Spreadsheet.Sheet, row: unknown[]): void {
  const last = sheet.getLastRow();
  const targetRow = last + 1;
  sheet.getRange(targetRow, 1, 1, row.length).setValues([row as any[]]);
}

function withCorrelationId(payload: unknown): unknown {
  const correlationId = activeRequestContext?.correlationId
    ? String(activeRequestContext.correlationId)
    : '';

  if (!correlationId) return payload ?? null;

  if (payload && typeof payload === 'object' && !Array.isArray(payload)) {
    const p: any = payload as any;
    if (p.correlationId) return payload;
    return { correlationId, ...p };
  }

  return { correlationId, value: payload ?? null };
}

function appendAuditLog(action: string, payload: unknown, success: boolean, message?: string): void {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(5000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_AUDIT_LOG);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_AUDIT_LOG);
      sheet.getRange('A1:F1').setValues([[
        'Timestamp', 'Email', 'Action', 'Success', 'Message', 'Payload'
      ]]);
      sheet.getRange('A1:F1').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const email = getRequestingUserEmail();
    const row = [
      new Date().toISOString(),
      sanitizeSheetString(email),
      sanitizeSheetString(action),
      success ? 'TRUE' : 'FALSE',
      sanitizeSheetString(message || ''),
      sanitizeSheetString(JSON.stringify(withCorrelationId(payload))),
    ];

    appendRowFast(sheet, row);

    const lastRow = sheet.getLastRow();
    const maxWithHeader = MAX_AUDIT_LOG_ROWS + 1;
    if (lastRow > maxWithHeader + AUDIT_LOG_TRIM_BUFFER) {
      const toDelete = lastRow - maxWithHeader;
      sheet.deleteRows(2, toDelete);
    }
  } catch (err) {
    Logger.log(`Erro ao gravar audit log: ${String((err as any)?.message || err)}`);
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

export function logClientError(event: {
  message: string;
  stack?: string;
  view?: string;
  url?: string;
  userAgent?: string;
  correlationId?: string;
}): { success: boolean } {
  try {
    const user = getUsuarioByEmail(getRequestingUserEmail());
    if (!user || user.status !== 'ATIVO') {
      return { success: false };
    }

    const cache = CacheService.getUserCache();
    const minuteKey = `lograte:${new Date().toISOString().slice(0, 16)}`; // YYYY-MM-DDTHH:MM
    const raw = cache.get(minuteKey);
    const next = (raw ? Number(raw) : 0) + 1;
    cache.put(minuteKey, String(next), 120);

    if (next > 20) {
      return { success: true }; // drop silently
    }

    appendAuditLog(
      'clientError',
      {
        message: event?.message,
        stack: event?.stack,
        view: event?.view,
        url: event?.url,
        userAgent: event?.userAgent,
        correlationId: event?.correlationId,
      },
      false,
      event?.message
    );

    return { success: true };
  } catch (error: any) {
    Logger.log(`Erro ao registrar clientError: ${error?.message || String(error)}`);
    return { success: false };
  }
}

export function logServerException(
  endpoint: string,
  context: { correlationId?: string; view?: string; url?: string } | null | undefined,
  error: unknown
): void {
  try {
    const message =
      (error as any)?.message ? String((error as any).message) : String(error);
    const stack = (error as any)?.stack ? String((error as any).stack) : '';
    appendAuditLog(
      'serverException',
      {
        endpoint: String(endpoint || ''),
        correlationId: context?.correlationId ? String(context.correlationId) : '',
        view: context?.view ? String(context.view) : '',
        url: context?.url ? String(context.url) : '',
        message,
        stack,
      },
      false,
      message
    );
  } catch (e: any) {
    Logger.log(`Erro ao registrar serverException: ${e?.message || String(e)}`);
  }
}

export function logEndpointTiming(endpoint: string, durationMs: number): void {
  try {
    const ms = Math.max(0, Number(durationMs) || 0);
    appendAuditLog('slowEndpoint', { endpoint: String(endpoint || ''), durationMs: ms }, true, `${ms}ms`);
  } catch (e: any) {
    Logger.log(`Erro ao registrar slowEndpoint: ${e?.message || String(e)}`);
  }
}

export function runSmokeTests(): {
  ok: boolean;
  ranAt: string;
  steps: Array<{ name: string; ok: boolean; durationMs: number; error?: string }>;
} {
  const requester = getUsuarioByEmail(getRequestingUserEmail());
  if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
    return { ok: false, ranAt: new Date().toISOString(), steps: [{ name: 'auth', ok: false, durationMs: 0, error: 'ADMIN only' }] };
  }

  const steps: Array<{ name: string; ok: boolean; durationMs: number; error?: string }> = [];
  const run = (name: string, fn: () => void) => {
    const startedAt = Date.now();
    try {
      fn();
      steps.push({ name, ok: true, durationMs: Date.now() - startedAt });
    } catch (e: any) {
      steps.push({ name, ok: false, durationMs: Date.now() - startedAt, error: String(e?.message || e) });
    }
  };

  run('getCurrentUserInfo', () => {
    const info = getCurrentUserInfo() as any;
    if (!info || !info.email) throw new Error('missing email');
  });

  run('getReferenceData', () => {
    const data = getReferenceData() as any;
    if (!data) throw new Error('empty');
    if (!Array.isArray(data.filiais)) throw new Error('filiais not array');
  });

  run('views:getViewHtml', () => {
    const views = ['dashboard', 'contas-pagar', 'contas-receber', 'conciliacao', 'dre', 'fluxo-caixa', 'kpis', 'configuracoes'];
    for (const v of views) {
      const html = getViewHtml(v);
      if (!html || typeof html !== 'string') throw new Error(`view ${v} empty`);
      if (html.includes('View inválida')) throw new Error(`view ${v} invalid`);
    }
  });

  run('getDashboardData', () => {
    const d = getDashboardData() as any;
    if (!d) throw new Error('empty');
  });

  run('getContasPagar', () => {
    const d = getContasPagar() as any;
    if (!d) throw new Error('empty');
  });

  run('getContasReceber', () => {
    const d = getContasReceber() as any;
    if (!d) throw new Error('empty');
  });

  run('getConciliacaoData', () => {
    const d = getConciliacaoData() as any;
    if (!d) throw new Error('empty');
  });

  run('reports:DRE/DFC/KPI', () => {
    const now = new Date();
    const mes = now.getMonth() + 1;
    const ano = now.getFullYear();
    getDREMensal(mes, ano);
    getFluxoCaixaMensal(mes, ano);
    getKPIsMensal(mes, ano);
  });

  const ok = steps.every((s) => s.ok);
  const result = { ok, ranAt: new Date().toISOString(), steps };
  appendAuditLog('runSmokeTests', { ok, steps }, ok, ok ? 'OK' : 'FAIL');
  return result;
}

export function getAuditLogEntries(limit: number = 200): Array<{
  timestamp: string;
  email: string;
  action: string;
  success: boolean;
  message: string;
  payload: string;
}> {
  const requester = getUsuarioByEmail(getRequestingUserEmail());
  if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
    return [];
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_AUDIT_LOG);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const safeLimit = Math.min(Math.max(1, Number(limit) || 200), 1000);
  const startRow = Math.max(2, lastRow - safeLimit + 1);
  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, 6).getDisplayValues();

  return values
    .map((r) => ({
      timestamp: String(r[0] || ''),
      email: String(r[1] || ''),
      action: String(r[2] || ''),
      success: String(r[3] || '').toUpperCase() === 'TRUE',
      message: String(r[4] || ''),
      payload: String(r[5] || ''),
    }))
    .reverse();
}

export function getAuditLogEntriesFiltered(filters: {
  limit?: number;
  action?: string;
  email?: string;
  success?: boolean | null;
  correlationId?: string;
  query?: string;
}): Array<{
  timestamp: string;
  email: string;
  action: string;
  success: boolean;
  message: string;
  payload: string;
}> {
  const requester = getUsuarioByEmail(getRequestingUserEmail());
  if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
    return [];
  }

  const safeLimit = Math.min(Math.max(1, Number(filters?.limit) || 200), 1000);
  const actionFilter = String(filters?.action || '').trim().toLowerCase();
  const emailFilter = String(filters?.email || '').trim().toLowerCase();
  const correlationId = String(filters?.correlationId || '').trim();
  const query = String(filters?.query || '').trim().toLowerCase();
  const successFilter =
    typeof filters?.success === 'boolean' ? filters.success : null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_AUDIT_LOG);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // Lê uma janela maior para permitir filtros sem perder resultados
  const scanWindow = Math.min(2000, Math.max(500, safeLimit * 10));
  const startRow = Math.max(2, lastRow - scanWindow + 1);
  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, 6).getDisplayValues();

  const results: Array<{
    timestamp: string;
    email: string;
    action: string;
    success: boolean;
    message: string;
    payload: string;
  }> = [];

  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i] || [];
    const timestamp = String(r[0] || '');
    const email = String(r[1] || '');
    const action = String(r[2] || '');
    const success = String(r[3] || '').toUpperCase() === 'TRUE';
    const message = String(r[4] || '');
    const payload = String(r[5] || '');

    if (successFilter !== null && success !== successFilter) continue;
    if (actionFilter && action.toLowerCase().indexOf(actionFilter) === -1) continue;
    if (emailFilter && email.toLowerCase().indexOf(emailFilter) === -1) continue;
    if (correlationId && payload.indexOf(correlationId) === -1) continue;

    if (query) {
      const hay = `${timestamp} ${email} ${action} ${message} ${payload}`.toLowerCase();
      if (hay.indexOf(query) === -1) continue;
    }

    results.push({ timestamp, email, action, success, message, payload });
    if (results.length >= safeLimit) break;
  }

  return results;
}

export function getAdminDiagnostics(): {
  ok: boolean;
  now: string;
  timezone: string;
  scriptId: string;
  webAppUrl: string;
  flags: Record<string, string>;
} {
  const requester = getUsuarioByEmail(getRequestingUserEmail());
  if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
    return {
      ok: false,
      now: new Date().toISOString(),
      timezone: Session.getScriptTimeZone(),
      scriptId: '',
      webAppUrl: '',
      flags: {},
    };
  }

  const props = PropertiesService.getScriptProperties();
  const flagKeys = ['ENABLE_DEBUG_API', 'ENABLE_SEED_DATA'] as const;
  const flags: Record<string, string> = {};
  flagKeys.forEach((k) => {
    const v = props.getProperty(k);
    flags[k] = v === null ? '' : String(v);
  });

  return {
    ok: true,
    now: new Date().toISOString(),
    timezone: Session.getScriptTimeZone(),
    scriptId: ScriptApp.getScriptId(),
    webAppUrl: ScriptApp.getService().getUrl(),
    flags,
  };
}

export function setAdminFlag(key: string, value: boolean): { success: boolean; message: string } {
  const requester = getUsuarioByEmail(getRequestingUserEmail());
  if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
    return { success: false, message: 'Sem permissão' };
  }

  const k = String(key || '').trim();
  const allow = new Set(['ENABLE_DEBUG_API', 'ENABLE_SEED_DATA']);
  if (!allow.has(k)) return { success: false, message: `Flag inválida: ${k}` };

  PropertiesService.getScriptProperties().setProperty(k, value ? 'true' : 'false');
  appendAuditLog('setAdminFlag', { key: k, value }, true);
  return { success: true, message: `Atualizado: ${k}=${value ? 'true' : 'false'}` };
}

export function clearCaches(): { success: boolean; message: string } {
  const requester = getUsuarioByEmail(getRequestingUserEmail());
  if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
    return { success: false, message: 'Sem permissão' };
  }
  try {
    cacheRemoveNamespace(CacheNamespace.REFERENCE, CacheScope.SCRIPT);
    cacheRemoveNamespace(CacheNamespace.DRE, CacheScope.SCRIPT);
    cacheRemoveNamespace(CacheNamespace.DFC, CacheScope.SCRIPT);
    cacheRemoveNamespace(CacheNamespace.KPI, CacheScope.SCRIPT);
    appendAuditLog('clearCaches', {}, true);
    return { success: true, message: 'Caches limpos' };
  } catch (e: any) {
    appendAuditLog('clearCaches', {}, false, e?.message);
    return { success: false, message: e?.message || String(e) };
  }
}

/**
 * Retorna o HTML de uma view específica
 */
export function getViewHtml(viewName: string): string {
  try {
    const normalized = String(viewName || '').trim();
    if (!WEBAPP_VIEW_ALLOWLIST.has(normalized)) {
      throw new Error('View inv\u00e1lida');
    }
    if (normalized === 'configuracoes') {
      const denied = requirePermission('gerenciarConfig', 'acessar configurações');
      if (denied) {
        return `<div class="empty-state"><div class="empty-state-message">${escapeHtml(denied.message)}</div></div>`;
      }
    }

    if (['dre', 'fluxo-caixa', 'kpis', 'relatorios'].includes(normalized)) {
      const denied = requirePermission('visualizarRelatorios', `acessar ${normalized}`);
      if (denied) {
        return `<div class="empty-state"><div class="empty-state-message">${escapeHtml(denied.message)}</div></div>`;
      }
    }

    return HtmlService.createHtmlOutputFromFile(`frontend/views/${normalized}-view`).getContent();
  } catch (error) {
    return `<div class="empty-state">
      <div class="empty-state-icon">⚠️</div>
      <div class="empty-state-message">Erro ao carregar view</div>
      <div class="empty-state-hint">${escapeHtml(error)}</div>
    </div>`;
  }
}

export function getCurrentUserInfo(): {
  email: string;
  nome: string;
  perfil: string;
  permissoes: NonNullable<Usuario['permissoes']>;
} {
  const email =
    Session.getActiveUser().getEmail() ||
    Session.getEffectiveUser().getEmail() ||
    'usuario@empresa.com';

  const fallbackNome = email.split('@')[0] || email;
  const user = getUsuarioByEmail(email);
  const perfil = user?.perfil || 'USUARIO';
  const permissoes = user?.permissoes || getPermissoesPadrao(perfil);
  return { email, nome: user?.nome || fallbackNome, perfil, permissoes };
}

// ============================================================================
// REFERENCE DATA
// ============================================================================

export function getReferenceData(): {
  filiais: Array<{ codigo: string; nome: string; ativo?: boolean }>;
  canais: Array<{ codigo: string; nome: string; ativo?: boolean }>;
  contas: Array<{ codigo: string; nome: string; tipo?: string; grupoDRE?: string; subgrupoDRE?: string; grupoDFC?: string; variavelFixa?: string; cmaCmv?: string }>;
  centrosCusto: Array<{ codigo: string; nome: string; ativo?: boolean }>;
} {
  const user = getUsuarioByEmail(getRequestingUserEmail());
  if (!user || user.status !== 'ATIVO') {
    return { filiais: [], canais: [], contas: [], centrosCusto: [] };
  }

  const cached = cacheGet<any>(CacheNamespace.REFERENCE, 'all');
  if (cached) {
    return cached;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Filiais (da planilha)
    const sheetFiliais = ss.getSheetByName(SHEET_REF_FILIAIS);
    const filiais = sheetFiliais ? sheetFiliais.getDataRange().getValues().slice(1) : [];

    // Canais (da planilha)
    const sheetCanais = ss.getSheetByName(SHEET_REF_CANAIS);
    const canais = sheetCanais ? sheetCanais.getDataRange().getValues().slice(1) : [];

    // Centros de Custo (da planilha, com fallback para hardcoded)
    const sheetCCusto = ss.getSheetByName(SHEET_REF_CCUSTO);
    let centrosCusto: any[];
    if (sheetCCusto && sheetCCusto.getLastRow() > 1) {
      const ccData = sheetCCusto.getDataRange().getValues().slice(1);
      centrosCusto = ccData.filter((cc: any) => cc[0]).map((cc: any) => ({
        codigo: String(cc[0]),
        nome: String(cc[1]),
        ativo: cc[2] !== false && String(cc[2]).toUpperCase() !== 'FALSE'
      }));
    } else {
      // Fallback hardcoded
      centrosCusto = [
        { codigo: 'ADM', nome: 'Administrativo', ativo: true },
        { codigo: 'COM', nome: 'Comercial', ativo: true },
        { codigo: 'OPS', nome: 'Operacional', ativo: true },
        { codigo: 'FIN', nome: 'Financeiro', ativo: true },
        { codigo: 'TI', nome: 'Tecnologia', ativo: true },
      ];
    }

    // Contas Contábeis (da planilha, com fallback para hardcoded)
    const sheetContas = ss.getSheetByName(SHEET_REF_PLANO_CONTAS);
    let contas: any[];
    if (sheetContas && sheetContas.getLastRow() > 1) {
      const lastRow = sheetContas.getLastRow();
      const lastCol = Math.max(8, sheetContas.getLastColumn());
      const contasData = sheetContas.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
      contas = contasData
        .filter((c: any) => c[0])
        .map((c: any) => ({
          codigo: String(c[0]).trim(),
          nome: String(c[1] || '').trim(),
          tipo: String(c[2] || '').trim(),
          grupoDRE: String(c[3] || '').trim(),
          subgrupoDRE: String(c[4] || '').trim(),
          grupoDFC: String(c[5] || '').trim(),
          variavelFixa: String(c[6] || '').trim(),
          cmaCmv: String(c[7] || '').trim(),
        }));
    } else {
      // Fallback hardcoded
      contas = [
        { codigo: '1.01.001', nome: 'Receita de Serviços', tipo: 'RECEITA', grupoDRE: '', subgrupoDRE: '', grupoDFC: '', variavelFixa: '', cmaCmv: '' },
        { codigo: '1.01.002', nome: 'Receita de Produtos', tipo: 'RECEITA', grupoDRE: '', subgrupoDRE: '', grupoDFC: '', variavelFixa: '', cmaCmv: '' },
        { codigo: '2.01.001', nome: 'Fornecedores', tipo: 'DESPESA', grupoDRE: '', subgrupoDRE: '', grupoDFC: '', variavelFixa: '', cmaCmv: '' },
        { codigo: '2.01.002', nome: 'Salários', tipo: 'DESPESA', grupoDRE: '', subgrupoDRE: '', grupoDFC: '', variavelFixa: '', cmaCmv: '' },
        { codigo: '2.01.003', nome: 'Impostos', tipo: 'DESPESA', grupoDRE: '', subgrupoDRE: '', grupoDFC: '', variavelFixa: '', cmaCmv: '' },
        { codigo: '2.01.004', nome: 'Aluguel', tipo: 'DESPESA', grupoDRE: '', subgrupoDRE: '', grupoDFC: '', variavelFixa: '', cmaCmv: '' },
      ];
    }

    const result = {
      filiais: filiais.filter((f: any) => f[0]).map((f: any) => {
        const ativoIdx = f.length >= 4 ? 3 : 2;
        return {
          codigo: String(f[0]),
          nome: String(f[1]),
          ativo: f[ativoIdx] !== false && String(f[ativoIdx] ?? 'TRUE').toUpperCase() !== 'FALSE',
        };
      }),
      canais: canais.filter((c: any) => c[0]).map((c: any) => ({
        codigo: String(c[0]),
        nome: String(c[1]),
        ativo: c[2] !== false && String(c[2] ?? 'TRUE').toUpperCase() !== 'FALSE',
      })),
      contas: contas,
      centrosCusto: centrosCusto,
    };

    cacheSet(CacheNamespace.REFERENCE, 'all', result, 600);
    return result;
  } catch (error: any) {
    Logger.log(`Erro ao carregar dados de referência: ${error.message}`);
    // Retornar dados vazios em caso de erro
    return {
      filiais: [],
      canais: [],
      contas: [],
      centrosCusto: [],
    };
  }
}

// ============================================================================
// CRUD CONFIGURAÇÕES
// ============================================================================

// Centros de Custo
export function salvarCentroCusto(centroCusto: { codigo: string; nome: string; ativo?: boolean }, editIndex: number): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'salvar centro de custo');
    if (denied) return denied;

    const validation = combineValidations(
      validateRequired(centroCusto?.codigo, 'Código'),
      validateRequired(centroCusto?.nome, 'Nome')
    );
    if (!validation.valid) return { success: false, message: validation.errors.join('; ') };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_REF_CCUSTO);

    // Criar aba se não existir
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_REF_CCUSTO);
      sheet.getRange('A1:C1').setValues([['Código', 'Nome', 'Ativo']]);
      sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#00a8e8').setFontColor('#ffffff');
    }

    // Verificar código duplicado
    if (editIndex < 0) {
      const data = sheet.getDataRange().getValues().slice(1);
      const codigoExiste = data.some((row: any) => String(row[0]).toUpperCase() === centroCusto.codigo.toUpperCase());
      if (codigoExiste) {
        return { success: false, message: 'Código já existe' };
      }
    }

    const ativo = centroCusto.ativo !== false && String(centroCusto.ativo ?? 'TRUE').toUpperCase() !== 'FALSE';
    const codigo = sanitizeSheetString(centroCusto.codigo).toUpperCase();
    const nome = sanitizeSheetString(centroCusto.nome);

    if (editIndex >= 0) {
      // Editar (linha = editIndex + 2)
      sheet.getRange(editIndex + 2, 1, 1, 3).setValues([[codigo, nome, ativo]]);
    } else {
      // Novo
      sheet.appendRow([codigo, nome, ativo]);
    }

    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('salvarCentroCusto', { centroCusto: { codigo, nome, ativo }, editIndex }, true);
    return { success: true, message: 'Centro de custo salvo com sucesso' };
  } catch (error: any) {
    appendAuditLog('salvarCentroCusto', { centroCusto, editIndex }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function excluirCentroCusto(index: number): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'excluir centro de custo');
    if (denied) return denied;
    if (!Number.isFinite(index) || index < 0) return { success: false, message: 'Índice inválido' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_CCUSTO);

    if (!sheet) {
      throw new Error('Aba de centros de custo não encontrada');
    }

    // Deletar linha (index + 2)
    if (index + 2 > sheet.getLastRow()) throw new Error('Índice inválido');
    sheet.deleteRow(index + 2);

    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('excluirCentroCusto', { index }, true);
    return { success: true, message: 'Centro de custo excluído' };
  } catch (error: any) {
    appendAuditLog('excluirCentroCusto', { index }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function toggleCentroCusto(index: number, ativo: boolean): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'alterar centro de custo');
    if (denied) return denied;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_CCUSTO);
    if (!sheet) throw new Error('Aba de centros de custo não encontrada');
    if (index < 0 || index + 2 > sheet.getLastRow()) throw new Error('Índice inválido');
    sheet.getRange(index + 2, 3).setValue(ativo);
    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('toggleCentroCusto', { index, ativo }, true);
    return { success: true, message: `Centro de custo ${ativo ? 'ativado' : 'inativado'}` };
  } catch (error: any) {
    appendAuditLog('toggleCentroCusto', { index, ativo }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

// Plano de Contas
export function salvarContaContabil(conta: { codigo: string; nome: string; tipo: string; grupoDRE?: string; subgrupoDRE?: string; grupoDFC?: string; variavelFixa?: string; cmaCmv?: string }, editIndex: number): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'salvar conta contábil');
    if (denied) return denied;

    const validation = combineValidations(
      validateRequired(conta?.codigo, 'Código'),
      validateRequired(conta?.nome, 'Nome'),
      validateEnum(String(conta?.tipo || ''), ['RECEITA', 'DESPESA'], 'Tipo')
    );
    if (!validation.valid) return { success: false, message: validation.errors.join('; ') };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_REF_PLANO_CONTAS);

    // Criar aba se não existir
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_REF_PLANO_CONTAS);
      sheet.getRange('A1:H1').setValues([['Código', 'Nome', 'Tipo', 'Grupo DRE', 'Subgrupo DRE', 'Grupo DFC', 'Variável/Fixa', 'CMA/CMV']]);
      sheet.getRange('A1:H1').setFontWeight('bold').setBackground('#00a8e8').setFontColor('#ffffff');
    }

    // Verificar código duplicado
    if (editIndex < 0) {
      const data = sheet.getDataRange().getValues().slice(1);
      const codigoExiste = data.some((row: any) => String(row[0]) === conta.codigo);
      if (codigoExiste) {
        return { success: false, message: 'Código já existe' };
      }
    }

    const rowData = [
      sanitizeSheetString(conta.codigo),
      sanitizeSheetString(conta.nome),
      sanitizeSheetString(conta.tipo),
      sanitizeSheetString(conta.grupoDRE || ''),
      sanitizeSheetString(conta.subgrupoDRE || ''),
      sanitizeSheetString(conta.grupoDFC || ''),
      sanitizeSheetString(conta.variavelFixa || ''),
      sanitizeSheetString(conta.cmaCmv || ''),
    ];

    if (editIndex >= 0) {
      // Editar (linha = editIndex + 2)
      sheet.getRange(editIndex + 2, 1, 1, 8).setValues([rowData]);
    } else {
      // Novo
      sheet.appendRow(rowData);
    }

    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('salvarContaContabil', { editIndex, conta: rowData }, true);
    return { success: true, message: 'Conta contábil salva com sucesso' };
  } catch (error: any) {
    appendAuditLog('salvarContaContabil', { editIndex, conta }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function excluirConta(index: number): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'excluir conta contábil');
    if (denied) return denied;
    if (!Number.isFinite(index) || index < 0) return { success: false, message: 'Índice inválido' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_PLANO_CONTAS);

    if (!sheet) {
      throw new Error('Aba de plano de contas não encontrada');
    }

    if (index + 2 > sheet.getLastRow()) throw new Error('Índice inválido');
    sheet.deleteRow(index + 2);

    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('excluirContaContabil', { index }, true);
    return { success: true, message: 'Conta excluída' };
  } catch (error: any) {
    appendAuditLog('excluirContaContabil', { index }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

// Canais
export function salvarCanal(canal: { codigo: string; nome: string; ativo?: boolean }, editIndex: number): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'salvar canal');
    if (denied) return denied;

    const validation = combineValidations(
      validateRequired(canal?.codigo, 'Código'),
      validateRequired(canal?.nome, 'Nome')
    );
    if (!validation.valid) return { success: false, message: validation.errors.join('; ') };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_REF_CANAIS);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_REF_CANAIS);
      sheet.getRange('A1:C1').setValues([['Código', 'Nome', 'Ativo']]);
      sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#00a8e8').setFontColor('#ffffff');
    }

    // Verificar código duplicado
    if (editIndex < 0) {
      const data = sheet.getDataRange().getValues().slice(1);
      const codigoExiste = data.some((row: any) => String(row[0]).toUpperCase() === canal.codigo.toUpperCase());
      if (codigoExiste) {
        return { success: false, message: 'Código já existe' };
      }
    }

    const ativo = canal.ativo !== false && String(canal.ativo ?? 'TRUE').toUpperCase() !== 'FALSE';
    const codigo = sanitizeSheetString(canal.codigo).toUpperCase();
    const nome = sanitizeSheetString(canal.nome);

    if (editIndex >= 0) {
      // Editar (linha = editIndex + 2)
      sheet.getRange(editIndex + 2, 1, 1, 3).setValues([[codigo, nome, ativo]]);
    } else {
      // Novo
      sheet.appendRow([codigo, nome, ativo]);
    }

    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('salvarCanal', { canal: { codigo, nome, ativo }, editIndex }, true);
    return { success: true, message: 'Canal salvo com sucesso' };
  } catch (error: any) {
    appendAuditLog('salvarCanal', { canal, editIndex }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function excluirCanal(index: number): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'excluir canal');
    if (denied) return denied;
    if (!Number.isFinite(index) || index < 0) return { success: false, message: 'Índice inválido' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_CANAIS);

    if (!sheet) {
      throw new Error('Aba de canais não encontrada');
    }

    if (index + 2 > sheet.getLastRow()) throw new Error('Índice inválido');
    sheet.deleteRow(index + 2);

    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('excluirCanal', { index }, true);
    return { success: true, message: 'Canal excluído' };
  } catch (error: any) {
    appendAuditLog('excluirCanal', { index }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function toggleCanal(index: number, ativo: boolean): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'alterar canal');
    if (denied) return denied;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_CANAIS);
    if (!sheet) throw new Error('Aba de canais não encontrada');
    if (index < 0 || index + 2 > sheet.getLastRow()) throw new Error('Índice inválido');
    sheet.getRange(index + 2, 3).setValue(ativo);
    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('toggleCanal', { index, ativo }, true);
    return { success: true, message: `Canal ${ativo ? 'ativado' : 'inativado'}` };
  } catch (error: any) {
    appendAuditLog('toggleCanal', { index, ativo }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

// Filiais
export function salvarFilial(filial: { codigo: string; nome: string; ativo?: boolean }, editIndex: number): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'salvar filial');
    if (denied) return denied;

    const validation = combineValidations(
      validateRequired(filial?.codigo, 'Código'),
      validateRequired(filial?.nome, 'Nome')
    );
    if (!validation.valid) return { success: false, message: validation.errors.join('; ') };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_REF_FILIAIS);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_REF_FILIAIS);
      sheet.getRange('A1:D1').setValues([['Código', 'Nome', 'CNPJ', 'Ativo']]);
      sheet.getRange('A1:D1').setFontWeight('bold').setBackground('#00a8e8').setFontColor('#ffffff');
    }

    // Verificar código duplicado
    if (editIndex < 0) {
      const data = sheet.getDataRange().getValues().slice(1);
      const codigoExiste = data.some((row: any) => String(row[0]).toUpperCase() === filial.codigo.toUpperCase());
      if (codigoExiste) {
        return { success: false, message: 'Código já existe' };
      }
    }

    const ativo = filial.ativo !== false && String(filial.ativo ?? 'TRUE').toUpperCase() !== 'FALSE';
    const codigo = sanitizeSheetString(filial.codigo).toUpperCase();
    const nome = sanitizeSheetString(filial.nome);

    const lastCol = Math.max(4, sheet.getLastColumn());
    const cnpjColIdx = lastCol >= 4 ? 3 : 2; // zero-based for getRange values assembly
    const ativoColIdx = lastCol >= 4 ? 4 : 3; // 1-based positions for getRange

    // Preserve CNPJ existente se houver
    let cnpj = '';
    if (editIndex >= 0 && sheet.getLastRow() >= editIndex + 2) {
      const existing = sheet.getRange(editIndex + 2, 1, 1, lastCol).getValues()[0];
      cnpj = existing[cnpjColIdx - 1] || '';
    }

    const rowData: any[] = [codigo, nome];
    if (ativoColIdx === 4) {
      rowData.push(cnpj || '');
      rowData.push(ativo);
    } else {
      rowData.push(ativo);
    }

    if (editIndex >= 0) {
      sheet.getRange(editIndex + 2, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('salvarFilial', { filial: rowData, editIndex }, true);
    return { success: true, message: 'Filial salva com sucesso' };
  } catch (error: any) {
    appendAuditLog('salvarFilial', { filial, editIndex }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function excluirFilial(index: number): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'excluir filial');
    if (denied) return denied;
    if (!Number.isFinite(index) || index < 0) return { success: false, message: 'Índice inválido' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_FILIAIS);

    if (!sheet) {
      throw new Error('Aba de filiais não encontrada');
    }

    if (index + 2 > sheet.getLastRow()) throw new Error('Índice inválido');
    sheet.deleteRow(index + 2);

    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('excluirFilial', { index }, true);
    return { success: true, message: 'Filial excluída' };
  } catch (error: any) {
    appendAuditLog('excluirFilial', { index }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function toggleFilial(index: number, ativo: boolean): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'alterar filial');
    if (denied) return denied;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_FILIAIS);
    if (!sheet) throw new Error('Aba de filiais não encontrada');
    if (index < 0 || index + 2 > sheet.getLastRow()) throw new Error('Índice inválido');
    const lastCol = sheet.getLastColumn();
    const colAtivo = lastCol >= 4 ? 4 : 3;
    sheet.getRange(index + 2, colAtivo).setValue(ativo);
    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('toggleFilial', { index, ativo }, true);
    return { success: true, message: `Filial ${ativo ? 'ativada' : 'inativada'}` };
  } catch (error: any) {
    appendAuditLog('toggleFilial', { index, ativo }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

// ============================================================================
// DASHBOARD
// ============================================================================

export function getDashboardData() {
  enforcePermission('visualizarRelatorios', 'carregar dashboard');
  const lancamentos = getLancamentosFromSheet();
  const hoje = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

  // Contas a pagar vencidas
  const pagarVencidas = lancamentos.filter(l =>
    l.tipo === 'DESPESA' &&
    (
      String(l.status || '').toUpperCase() === 'VENCIDA' ||
      (String(l.status || '').toUpperCase() === 'PENDENTE' && new Date(l.dataVencimento) < hoje)
    )
  );

  // Contas a pagar próximos 7 dias
  const proximos7Dias = new Date();
  proximos7Dias.setDate(proximos7Dias.getDate() + 7);
  const pagarProximas = lancamentos.filter(l =>
    l.tipo === 'DESPESA' &&
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos7Dias &&
    new Date(l.dataVencimento) >= hoje
  );

  // Contas a receber hoje
  const receberHoje = lancamentos.filter(l =>
    l.tipo === 'RECEITA' &&
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento).toDateString() === hoje.toDateString()
  );

  // Extratos pendentes
  const extratos = getExtratosFromSheet();
  const extratosPendentes = extratos.filter(e => e.statusConciliacao === 'PENDENTE');

  // Últimos lançamentos
  const recentTransactions = lancamentos
    .slice(0, 10)
    .map(l => ({
      id: l.id,
      data: l.dataCompetencia,
      descricao: l.descricao,
      tipo: l.tipo,
      valor: l.valorLiquido,
      status: l.status,
    }));

  // Alertas
  const alerts: Array<{ type: string; title: string; message: string }> = [];

  if (pagarVencidas.length > 0) {
    alerts.push({
      type: 'danger',
      title: 'Contas Vencidas',
      message: `Você tem ${pagarVencidas.length} contas a pagar vencidas no valor de ${formatCurrency(sumValues(pagarVencidas))}`,
    });
  }

  if (pagarProximas.length > 0) {
    alerts.push({
      type: 'warning',
      title: 'Vencimentos Próximos',
      message: `${pagarProximas.length} contas a pagar vencem nos próximos 7 dias`,
    });
  }

  if (extratosPendentes.length > 5) {
    alerts.push({
      type: 'info',
      title: 'Conciliação Pendente',
      message: `${extratosPendentes.length} extratos bancários aguardando conciliação`,
    });
  }

  return {
    pagarVencidas: {
      quantidade: pagarVencidas.length,
      valor: sumValues(pagarVencidas),
    },
    pagarProximas: {
      quantidade: pagarProximas.length,
      valor: sumValues(pagarProximas),
    },
    receberHoje: {
      quantidade: receberHoje.length,
      valor: sumValues(receberHoje),
    },
    conciliacaoPendentes: {
      quantidade: extratosPendentes.length,
      valor: extratosPendentes.reduce((sum, e) => sum + parseFloat(String(e.valor || 0)), 0),
    },
    recentTransactions,
    alerts,
  };
}

// ============================================================================
// CONTAS A PAGAR
// ============================================================================

export function getContasPagar() {
  enforcePermission('visualizarRelatorios', 'listar contas a pagar');
  const lancamentos = getLancamentosFromSheet();
  const hoje = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

  const contasPagar = lancamentos.filter(l => l.tipo === 'DESPESA');

    const isPago = (s: string) => ['PAGO', 'PAGA', 'RECEBIDO', 'RECEBIDA'].includes((s || '').toUpperCase());
    const vencidas = contasPagar.filter(l =>
      (l.status === 'VENCIDA') ||
      (l.status === 'PENDENTE' && new Date(l.dataVencimento) < hoje)
    );

  const proximos7Dias = new Date();
  proximos7Dias.setDate(proximos7Dias.getDate() + 7);
  const vencer7 = contasPagar.filter(l =>
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos7Dias &&
    new Date(l.dataVencimento) >= hoje
  );

  const proximos30Dias = new Date();
  proximos30Dias.setDate(proximos30Dias.getDate() + 30);
  const vencer30 = contasPagar.filter(l =>
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos30Dias &&
    new Date(l.dataVencimento) > proximos7Dias
  );

    const pagas = contasPagar.filter(l =>
      isPago(l.status) &&
      new Date(l.dataPagamento || l.dataCompetencia) >= inicioMes
    );

  return {
    stats: {
      vencidas: { quantidade: vencidas.length, valor: sumValues(vencidas) },
      vencer7: { quantidade: vencer7.length, valor: sumValues(vencer7) },
      vencer30: { quantidade: vencer30.length, valor: sumValues(vencer30) },
      pagas: { quantidade: pagas.length, valor: sumValues(pagas) },
    },
    contas: contasPagar.map(l => ({
      id: l.id,
      vencimento: l.dataVencimento,
      fornecedor: l.descricao.split('-')[0].trim(),
      descricao: l.descricao,
      valor: l.valorLiquido,
      status: l.status,
      filial: l.filial,
    })),
  };
}

export function pagarConta(id: string): { success: boolean; message: string } {
  try {
    const denied = requirePermission('aprovarPagamentos', 'pagar conta');
    if (denied) return denied;
    const validation = validateRequired(id, 'ID');
    if (!validation.valid) return { success: false, message: validation.errors.join('; ') };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lançamentos não encontrada');

    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);
    try {
      const headers = getHeaderIndexMap(sheet);
      const idCol = headers['ID'];
      const statusCol = headers['Status'];
      const dataPagCol = headers['Data Pagamento'];

      if (idCol === undefined || statusCol === undefined || dataPagCol === undefined) {
        throw new Error('Cabeçalhos obrigatórios não encontrados (ID, Status, Data Pagamento)');
      }

      const row = findRowByExactValueInColumn(sheet, idCol, id);
      if (!row) throw new Error('Conta não encontrada');

      const currentStatus = String(sheet.getRange(row, statusCol + 1).getDisplayValue() || '').toUpperCase();
      if (currentStatus !== 'PENDENTE') {
        return { success: false, message: `Conta não está pendente (status: ${currentStatus})` };
      }

      sheet.getRange(row, statusCol + 1).setValue('PAGA');
      sheet.getRange(row, dataPagCol + 1).setValue(new Date());
      appendAuditLog('pagarConta', { id }, true);
      clearReportsCache();
      return { success: true, message: 'Conta paga com sucesso' };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (error: any) {
    appendAuditLog('pagarConta', { id }, false, error?.message);
    return { success: false, message: error.message };
  }
}

export function pagarContasEmLote(ids: string[]): { success: boolean; message: string } {
  try {
    const denied = requirePermission('aprovarPagamentos', 'pagar contas em lote');
    if (denied) return denied;
    if (!Array.isArray(ids) || ids.length === 0) return { success: false, message: 'IDs inválidos' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lançamentos não encontrada');

    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);
    try {
      const headers = getHeaderIndexMap(sheet);
      const idCol = headers['ID'];
      const statusCol = headers['Status'];
      const dataPagCol = headers['Data Pagamento'];
      if (idCol === undefined || statusCol === undefined || dataPagCol === undefined) {
        throw new Error('Cabeçalhos obrigatórios não encontrados (ID, Status, Data Pagamento)');
      }

      const lastRow = sheet.getLastRow();
      const idsColumnValues =
        lastRow > 1
          ? sheet.getRange(2, idCol + 1, lastRow - 1, 1).getDisplayValues()
          : [];

      const idToRow = new Map<string, number>();
      idsColumnValues.forEach((r, idx) => {
        const cell = String(r[0] || '').trim();
        if (!cell) return;
        idToRow.set(cell, idx + 2);
      });

      let count = 0;
      const errors: string[] = [];
      const statusRanges: string[] = [];
      const dateRanges: string[] = [];
      const now = new Date();

      for (const rawId of ids) {
        const wanted = String(rawId || '').trim();
        if (!wanted) continue;
        const row = idToRow.get(wanted);
        if (!row) {
          errors.push(`${wanted}: não encontrada`);
          continue;
        }

        const currentStatus = String(sheet.getRange(row, statusCol + 1).getDisplayValue() || '').toUpperCase();
        if (currentStatus !== 'PENDENTE') {
          errors.push(`${wanted}: status ${currentStatus}`);
          continue;
        }

        statusRanges.push(`${columnToLetter(statusCol + 1)}${row}`);
        dateRanges.push(`${columnToLetter(dataPagCol + 1)}${row}`);
        appendAuditLog('pagarConta', { id: wanted }, true);
        count++;
      }

      if (statusRanges.length > 0) {
        sheet.getRangeList(statusRanges).setValue('PAGA');
        sheet.getRangeList(dateRanges).setValue(now);
        clearReportsCache();
      }

      if (errors.length > 0) {
        appendAuditLog('pagarContasEmLote', { ids, count, errorsCount: errors.length }, true, 'Parcial');
        return { success: true, message: `${count} contas pagas; ${errors.length} falharam` };
      }
      appendAuditLog('pagarContasEmLote', { ids, count }, true);
      return { success: true, message: `${count} contas pagas com sucesso` };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (error: any) {
    appendAuditLog('pagarContasEmLote', { ids }, false, error?.message);
    return { success: false, message: error.message };
  }
}

// ============================================================================
// CONTAS A RECEBER
// ============================================================================

export function getContasReceber() {
  enforcePermission('visualizarRelatorios', 'listar contas a receber');
  const lancamentos = getLancamentosFromSheet();
  const hoje = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

    const isRecebido = (s: string) => ['RECEBIDO', 'RECEBIDA', 'PAGO', 'PAGA'].includes((s || '').toUpperCase());
    const contasReceber = lancamentos.filter(l => l.tipo === 'RECEITA');

  const vencidas = contasReceber.filter(l =>
    l.status === 'VENCIDA' || (l.status === 'PENDENTE' && new Date(l.dataVencimento) < hoje)
  );

  const proximos7Dias = new Date();
  proximos7Dias.setDate(proximos7Dias.getDate() + 7);
  const receber7 = contasReceber.filter(l =>
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos7Dias &&
    new Date(l.dataVencimento) >= hoje
  );

  const proximos30Dias = new Date();
  proximos30Dias.setDate(proximos30Dias.getDate() + 30);
  const receber30 = contasReceber.filter(l =>
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos30Dias &&
    new Date(l.dataVencimento) > proximos7Dias
  );

    const recebidas = contasReceber.filter(l =>
      isRecebido(l.status) &&
      new Date(l.dataPagamento || l.dataCompetencia) >= inicioMes
    );

  return {
    stats: {
      vencidas: { quantidade: vencidas.length, valor: sumValues(vencidas) },
      receber7: { quantidade: receber7.length, valor: sumValues(receber7) },
      receber30: { quantidade: receber30.length, valor: sumValues(receber30) },
      recebidas: { quantidade: recebidas.length, valor: sumValues(recebidas) },
    },
    contas: contasReceber.map(l => ({
      id: l.id,
      vencimento: l.dataVencimento,
      cliente: l.descricao.split('-')[0].trim(),
      descricao: l.descricao,
      valor: l.valorLiquido,
      status: l.status,
      canal: l.canal || 'N/A',
    })),
  };
}

export function receberConta(id: string): { success: boolean; message: string } {
  try {
    const denied = requirePermission('aprovarPagamentos', 'receber conta');
    if (denied) return denied;
    const validation = validateRequired(id, 'ID');
    if (!validation.valid) return { success: false, message: validation.errors.join('; ') };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lançamentos não encontrada');

    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);
    try {
      const headers = getHeaderIndexMap(sheet);
      const idCol = headers['ID'];
      const statusCol = headers['Status'];
      const dataPagCol = headers['Data Pagamento'];

      if (idCol === undefined || statusCol === undefined || dataPagCol === undefined) {
        throw new Error('Cabeçalhos obrigatórios não encontrados (ID, Status, Data Pagamento)');
      }

      const row = findRowByExactValueInColumn(sheet, idCol, id);
      if (!row) throw new Error('Conta não encontrada');

      const currentStatus = String(sheet.getRange(row, statusCol + 1).getDisplayValue() || '').toUpperCase();
      if (currentStatus !== 'PENDENTE') {
        return { success: false, message: `Conta não está pendente (status: ${currentStatus})` };
      }

      sheet.getRange(row, statusCol + 1).setValue('RECEBIDA');
      sheet.getRange(row, dataPagCol + 1).setValue(new Date());
      appendAuditLog('receberConta', { id }, true);
      clearReportsCache();
      return { success: true, message: 'Conta recebida com sucesso' };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (error: any) {
    appendAuditLog('receberConta', { id }, false, error?.message);
    return { success: false, message: error.message };
  }
}

export function receberContasEmLote(ids: string[]): { success: boolean; message: string } {
  try {
    const denied = requirePermission('aprovarPagamentos', 'receber contas em lote');
    if (denied) return denied;
    if (!Array.isArray(ids) || ids.length === 0) return { success: false, message: 'IDs inválidos' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lançamentos não encontrada');

    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);
    try {
      const headers = getHeaderIndexMap(sheet);
      const idCol = headers['ID'];
      const statusCol = headers['Status'];
      const dataPagCol = headers['Data Pagamento'];
      if (idCol === undefined || statusCol === undefined || dataPagCol === undefined) {
        throw new Error('Cabeçalhos obrigatórios não encontrados (ID, Status, Data Pagamento)');
      }

      const lastRow = sheet.getLastRow();
      const idsColumnValues =
        lastRow > 1
          ? sheet.getRange(2, idCol + 1, lastRow - 1, 1).getDisplayValues()
          : [];

      const idToRow = new Map<string, number>();
      idsColumnValues.forEach((r, idx) => {
        const cell = String(r[0] || '').trim();
        if (!cell) return;
        idToRow.set(cell, idx + 2);
      });

      let count = 0;
      const errors: string[] = [];
      const statusRanges: string[] = [];
      const dateRanges: string[] = [];
      const now = new Date();

      for (const rawId of ids) {
        const wanted = String(rawId || '').trim();
        if (!wanted) continue;
        const row = idToRow.get(wanted);
        if (!row) {
          errors.push(`${wanted}: não encontrada`);
          continue;
        }

        const currentStatus = String(sheet.getRange(row, statusCol + 1).getDisplayValue() || '').toUpperCase();
        if (currentStatus !== 'PENDENTE') {
          errors.push(`${wanted}: status ${currentStatus}`);
          continue;
        }

        statusRanges.push(`${columnToLetter(statusCol + 1)}${row}`);
        dateRanges.push(`${columnToLetter(dataPagCol + 1)}${row}`);
        appendAuditLog('receberConta', { id: wanted }, true);
        count++;
      }

      if (statusRanges.length > 0) {
        sheet.getRangeList(statusRanges).setValue('RECEBIDA');
        sheet.getRangeList(dateRanges).setValue(now);
        clearReportsCache();
      }

      if (count === 0) {
        appendAuditLog('receberContasEmLote', { ids, count, errorsCount: errors.length }, false, 'Nenhuma recebida');
        return { success: false, message: errors.length ? errors[0] : 'Nenhuma conta recebida' };
      }

      if (errors.length > 0) {
        appendAuditLog('receberContasEmLote', { ids, count, errorsCount: errors.length }, true, 'Parcial');
        return { success: true, message: `${count} contas recebidas; ${errors.length} falharam` };
      }

      appendAuditLog('receberContasEmLote', { ids, count }, true);
      return { success: true, message: `${count} contas recebidas com sucesso` };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (error: any) {
    appendAuditLog('receberContasEmLote', { ids }, false, error?.message);
    return { success: false, message: error.message };
  }
}

// ============================================================================
// SALVAR LANÇAMENTO
// ============================================================================

export function salvarLancamento(lancamento: any): { success: boolean; message: string; id?: string } {
  try {
    const denied = requirePermission('criarLancamentos', 'salvar lançamento');
    if (denied) return denied;

    const v = combineValidations(
      validateRequired(lancamento?.id, 'ID'),
      validateRequired(lancamento?.dataCompetencia, 'Data competência'),
      validateRequired(lancamento?.dataVencimento, 'Data vencimento'),
      validateEnum(String(lancamento?.tipo || ''), ['RECEITA', 'DESPESA'], 'Tipo'),
      validateRequired(lancamento?.filial, 'Filial'),
      validateRequired(lancamento?.contaContabil, 'Conta contábil'),
      validateRequired(lancamento?.descricao, 'Descrição'),
      validateRequired(lancamento?.status, 'Status'),
      validateEnum(String(lancamento?.status || ''), ['PENDENTE', 'PAGA', 'RECEBIDA', 'VENCIDA', 'CANCELADA'], 'Status')
    );
    if (!v.valid) {
      return { success: false, message: v.errors.join('; ') };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lançamentos não encontrada');

    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);
    try {
      const headers = getHeaderIndexMap(sheet);
      const idCol = headers['ID'];
      if (idCol === undefined) throw new Error('Cabeçalho obrigatório não encontrado (ID)');
      const existingRow = findRowByExactValueInColumn(sheet, idCol, lancamento.id);
      if (existingRow) {
        return { success: false, message: `ID já existe: ${lancamento.id}` };
      }

    // Converter objeto lancamento para array de valores (seguindo a ordem das colunas)
    const valorBruto = Number(lancamento.valorBruto);
    const desconto = Number(lancamento.desconto || 0);
    const juros = Number(lancamento.juros || 0);
    const multa = Number(lancamento.multa || 0);
    const valorLiquido = valorBruto - desconto + juros + multa;

    if (!Number.isFinite(valorBruto) || valorBruto <= 0) {
      return { success: false, message: 'Valor bruto inválido' };
    }
    if (![desconto, juros, multa, valorLiquido].every(Number.isFinite)) {
      return { success: false, message: 'Valores numéricos inválidos' };
    }

    const row = [
      sanitizeSheetString(lancamento.id),                        // ID
      sanitizeSheetString(lancamento.dataCompetencia),           // Data Competência
      sanitizeSheetString(lancamento.dataVencimento),            // Data Vencimento
      sanitizeSheetString(lancamento.dataPagamento || ''),       // Data Pagamento
      sanitizeSheetString(lancamento.tipo),                      // Tipo (RECEITA/DESPESA)
      sanitizeSheetString(lancamento.filial),                    // Filial
      sanitizeSheetString(lancamento.centroCusto || ''),         // Centro de Custo
      sanitizeSheetString(lancamento.contaGerencial || ''),      // Conta Gerencial
      sanitizeSheetString(lancamento.contaContabil),             // Conta Contábil
      sanitizeSheetString(lancamento.grupoReceita || ''),        // Grupo Receita
      sanitizeSheetString(lancamento.canal || ''),               // Canal
      sanitizeSheetString(lancamento.descricao),                 // Descrição
      valorBruto,                                                // Valor Bruto
      desconto,                                                  // Desconto
      juros,                                                     // Juros
      multa,                                                     // Multa
      valorLiquido,                                              // Valor Líquido (recalculado)
      sanitizeSheetString(lancamento.status),                    // Status
      sanitizeSheetString(lancamento.idExtratoBanco || ''),      // ID Extrato Banco
      sanitizeSheetString(lancamento.origem || 'MANUAL'),        // Origem
      sanitizeSheetString(lancamento.observacoes || ''),         // Observações
    ];

      // Adicionar linha à planilha (mais rápido que appendRow)
      const targetRow = sheet.getLastRow() + 1;
      sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);

    appendAuditLog('salvarLancamento', { id: row[0], tipo: row[4], status: row[17] }, true);
    clearReportsCache();
    return {
      success: true,
      message: lancamento.tipo === 'RECEITA' ? 'Conta a receber salva com sucesso' : 'Conta a pagar salva com sucesso',
      id: lancamento.id,
    };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (error: any) {
    appendAuditLog('salvarLancamento', { lancamento }, false, error?.message);
    return {
      success: false,
      message: `Erro ao salvar lançamento: ${error.message}`,
    };
  }
}

// ============================================================================
// CONCILIAÇÃO
// ============================================================================

export function getConciliacaoData() {
  enforcePermission('visualizarRelatorios', 'carregar conciliação');
  const extratos = getExtratosFromSheet();
  const lancamentos = getLancamentosFromSheet();
  const hoje = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

    const extratosPendentes = extratos.filter(e => (e.statusConciliacao || 'PENDENTE').toUpperCase() === 'PENDENTE');
    const lancamentosPendentes = lancamentos.filter(l => !l.idExtratoBanco);

  const conciliadosHoje = extratos.filter(e =>
    (e.statusConciliacao || '').toUpperCase() === 'CONCILIADO' &&
    new Date(e.importadoEm).toDateString() === hoje.toDateString()
  );

    const totalExtratos = extratos.length;
    const totalConciliados = extratos.filter(e => (e.statusConciliacao || '').toUpperCase() === 'CONCILIADO').length;
  const taxaConciliacao = totalExtratos > 0 ? Math.round((totalConciliados / totalExtratos) * 100) : 0;

  // Histórico (últimas 50 conciliações)
  const historico = extratos
    .filter(e => e.statusConciliacao === 'CONCILIADO' && e.idLancamento)
    .slice(0, 50)
    .map(e => ({
      dataConciliacao: e.importadoEm,
      extratoId: e.id,
      lancamentoId: e.idLancamento,
      descricao: e.descricao,
      valor: e.valor,
      banco: e.banco,
      usuario: 'Sistema',
    }));

  return {
    stats: {
      extratosPendentes: extratosPendentes.length,
      extratosValor: extratosPendentes.reduce((sum, e) => sum + parseFloat(String(e.valor || 0)), 0),
      lancamentosPendentes: lancamentosPendentes.length,
      lancamentosValor: sumValues(lancamentosPendentes),
      conciliadosHoje: conciliadosHoje.length,
      conciliadosHojeValor: conciliadosHoje.reduce((sum, e) => sum + parseFloat(String(e.valor || 0)), 0),
      taxaConciliacao,
    },
    extratos: extratosPendentes.map(e => ({
      id: e.id,
      data: e.data,
      descricao: e.descricao,
      valor: e.valor,
      banco: e.banco,
    })),
    lancamentos: lancamentosPendentes.slice(0, 50).map(l => ({
      id: l.id,
      data: l.dataCompetencia,
      descricao: l.descricao,
      valor: l.valorLiquido,
      tipo: l.tipo,
    })),
    historico,
  };
}

export function conciliarItens(extratoId: string, lancamentoId: string): { success: boolean; message: string } {
  try {
    const denied = requirePermission('editarLancamentos', 'conciliar itens');
    if (denied) return denied;

    const v = combineValidations(
      validateRequired(extratoId, 'Extrato ID'),
      validateRequired(lancamentoId, 'Lançamento ID')
    );
    if (!v.valid) return { success: false, message: v.errors.join('; ') };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);
    try {

    // Atualizar extrato
    const sheetExtratos = ss.getSheetByName(SHEET_TB_EXTRATOS);
    if (!sheetExtratos) throw new Error('Aba de extratos não encontrada');
    const extrHeaders = getHeaderIndexMap(sheetExtratos);
    const extrIdCol = extrHeaders['ID'];
    const extrStatusCol = extrHeaders['Status Conciliação'];
    const extrLancCol = extrHeaders['ID Lançamento'];
    if (extrIdCol === undefined || extrStatusCol === undefined || extrLancCol === undefined) {
      throw new Error('Cabeçalhos obrigatórios não encontrados em extratos (ID, Status Conciliação, ID Lançamento)');
    }

    const extratoRow = findRowByExactValueInColumn(sheetExtratos, extrIdCol, extratoId);
    if (!extratoRow) throw new Error('Extrato não encontrado');
    sheetExtratos.getRange(extratoRow, extrStatusCol + 1).setValue('CONCILIADO');
    sheetExtratos.getRange(extratoRow, extrLancCol + 1).setValue(lancamentoId);

    // Atualizar lançamento
    const sheetLanc = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheetLanc) throw new Error('Aba de lançamentos não encontrada');
    const lancHeaders = getHeaderIndexMap(sheetLanc);
    const lancIdCol = lancHeaders['ID'];
    const lancExtratoCol = lancHeaders['ID Extrato Banco'];
    if (lancIdCol === undefined || lancExtratoCol === undefined) {
      throw new Error('Cabeçalhos obrigatórios não encontrados em lançamentos (ID, ID Extrato Banco)');
    }
    const lancRow = findRowByExactValueInColumn(sheetLanc, lancIdCol, lancamentoId);
    if (!lancRow) throw new Error('Lançamento não encontrado');
    sheetLanc.getRange(lancRow, lancExtratoCol + 1).setValue(extratoId);

    appendAuditLog('conciliarItens', { extratoId, lancamentoId }, true);
    clearReportsCache();
    return { success: true, message: 'Conciliação realizada com sucesso' };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (error: any) {
    appendAuditLog('conciliarItens', { extratoId, lancamentoId }, false, error?.message);
    return { success: false, message: error.message };
  }
}

export function conciliarAutomatico(): { success: boolean; conciliados: number; message: string } {
  try {
    const denied = requirePermission('editarLancamentos', 'conciliar automaticamente');
    if (denied) return { success: false, conciliados: 0, message: denied.message };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetExtratos = ss.getSheetByName(SHEET_TB_EXTRATOS);
    const sheetLanc = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheetExtratos) throw new Error('Aba de extratos não encontrada');
    if (!sheetLanc) throw new Error('Aba de lançamentos não encontrada');

    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);
    try {
      const extrHeaders = getHeaderIndexMap(sheetExtratos);
      const extrIdCol = extrHeaders['ID'];
      const extrDateCol = extrHeaders['Data'];
      const extrValueCol = extrHeaders['Valor'];
      const extrStatusCol = extrHeaders['Status Conciliação'];
      const extrLancCol = extrHeaders['ID Lançamento'];
      if (
        extrIdCol === undefined ||
        extrDateCol === undefined ||
        extrValueCol === undefined ||
        extrStatusCol === undefined ||
        extrLancCol === undefined
      ) {
        throw new Error('Cabeçalhos obrigatórios não encontrados em extratos');
      }

      const lancHeaders = getHeaderIndexMap(sheetLanc);
      const lancIdCol = lancHeaders['ID'];
      const lancDateCol = lancHeaders['Data Competência'];
      const lancValueCol = lancHeaders['Valor Líquido'];
      const lancExtratoCol = lancHeaders['ID Extrato Banco'];
      if (
        lancIdCol === undefined ||
        lancDateCol === undefined ||
        lancValueCol === undefined ||
        lancExtratoCol === undefined
      ) {
        throw new Error('Cabeçalhos obrigatórios não encontrados em lançamentos');
      }

      const extrLastRow = sheetExtratos.getLastRow();
      const lancLastRow = sheetLanc.getLastRow();
      if (extrLastRow <= 1 || lancLastRow <= 1) {
        return { success: true, conciliados: 0, message: '0 itens conciliados' };
      }

      const extrValues = sheetExtratos
        .getRange(2, 1, extrLastRow - 1, Math.max(9, extrStatusCol + 1, extrLancCol + 1, extrValueCol + 1, extrDateCol + 1))
        .getValues();

      const lancValues = sheetLanc
        .getRange(2, 1, lancLastRow - 1, Math.max(19, lancExtratoCol + 1, lancValueCol + 1, lancDateCol + 1))
        .getValues();

      const normalizeKey = (n: number): string => {
        const num = Number(n);
        if (!Number.isFinite(num)) return 'NaN';
        return (Math.round(num * 100) / 100).toFixed(2);
      };

      const toTs = (v: any): number => {
        const d = normalizeDateCell(v);
        const ts = Date.parse(d);
        return Number.isFinite(ts) ? ts : new Date(d).getTime();
      };

      type LancCandidate = { row: number; id: string; ts: number; valueKey: string };
      const byValue = new Map<string, LancCandidate[]>();

      for (let i = 0; i < lancValues.length; i++) {
        const rowNum = i + 2;
        const r = lancValues[i];
        const id = String(r[lancIdCol] || '').trim();
        if (!id) continue;
        const already = String(r[lancExtratoCol] || '').trim();
        if (already) continue;
        const ts = toTs(r[lancDateCol]);
        const valueKey = normalizeKey(Number(r[lancValueCol]));
        if (valueKey === 'NaN') continue;
        const arr = byValue.get(valueKey) || [];
        arr.push({ row: rowNum, id, ts, valueKey });
        byValue.set(valueKey, arr);
      }

      // sort candidates by date to speed up selection
      for (const arr of byValue.values()) {
        arr.sort((a, b) => a.ts - b.ts);
      }

      const maxDiffMs = 7 * 24 * 60 * 60 * 1000;
      let conciliados = 0;

      for (let i = 0; i < extrValues.length; i++) {
        const rowNum = i + 2;
        const r = extrValues[i];
        const status = String(r[extrStatusCol] || '').toUpperCase();
        if (status !== 'PENDENTE') continue;

        const extratoId = String(r[extrIdCol] || '').trim();
        if (!extratoId) continue;
        const extrTs = toTs(r[extrDateCol]);
        const valueKey = normalizeKey(Number(r[extrValueCol]));
        if (valueKey === 'NaN') continue;

        const candidates = byValue.get(valueKey);
        if (!candidates || candidates.length === 0) continue;

        let bestIdx = -1;
        let bestDiff = Infinity;
        for (let j = 0; j < candidates.length; j++) {
          const diff = Math.abs(candidates[j].ts - extrTs);
          if (diff <= maxDiffMs && diff < bestDiff) {
            bestDiff = diff;
            bestIdx = j;
            if (diff === 0) break;
          }
          // early exit if candidates are sorted and already too far past
          if (candidates[j].ts - extrTs > maxDiffMs) break;
        }

        if (bestIdx < 0) continue;
        const match = candidates.splice(bestIdx, 1)[0];

        sheetExtratos.getRange(rowNum, extrStatusCol + 1).setValue('CONCILIADO');
        sheetExtratos.getRange(rowNum, extrLancCol + 1).setValue(match.id);
        sheetLanc.getRange(match.row, lancExtratoCol + 1).setValue(extratoId);
        conciliados++;
      }

      appendAuditLog('conciliarAutomatico', { conciliados }, true);
      clearReportsCache();
      return { success: true, conciliados, message: `${conciliados} itens conciliados` };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (error: any) {
    appendAuditLog('conciliarAutomatico', {}, false, error?.message);
    return { success: false, conciliados: 0, message: error.message };
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function normalizeDateCell(value: any): string {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(value).trim();
  const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1];
  return s;
}

function getLancamentosFromSheet(): any[] {
  // garante aba com cabeçalhos
  createSheetIfNotExists(SHEET_TB_LANCAMENTOS, [
    'ID',
    'Data Competência',
    'Data Vencimento',
    'Data Pagamento',
    'Tipo',
    'Filial',
    'Centro Custo',
    'Conta Gerencial',
    'Conta Contábil',
    'Grupo Receita',
    'Canal',
    'Descrição',
    'Valor Bruto',
    'Desconto',
    'Juros',
    'Multa',
    'Valor Líquido',
    'Status',
    'ID Extrato Banco',
    'Origem',
    'Observações'
  ]);

  const data = getSheetValues(SHEET_TB_LANCAMENTOS);
   if (!data || data.length <= 1) {
    if (!isSeedDataEnabled()) return [];
    const seed = [
      ['CP-1001','2025-01-02','2025-01-12','2025-01-11','DESPESA','MATRIZ','OPS','Compra MP','10201','','ONLINE','Compra matéria-prima lote A',1500,0,0,0,1500,'PAGO','EXT-5002','Fornecedor X','Lote inicial'],
      ['CP-1002','2025-01-05','2025-01-20','','DESPESA','MATRIZ','OPS','Frete Compras','10205','','ONLINE','Frete compras fornecedores',400,0,0,0,400,'PENDENTE','','Fornecedor Y','À espera de pagamento'],
      ['CP-1003','2025-01-03','2025-01-30','','DESPESA','MATRIZ','ADM','Honorários','10402','','ONLINE','Honorários contábeis mês jan',900,0,0,0,900,'PENDENTE','','Escritório Z','Contrato mensal'],
      ['CR-2001','2025-01-02','2025-01-02','2025-01-02','RECEITA','MATRIZ','COM','Receita Fórmulas','20101','Receita Varejo','ONLINE','Venda balcão fórmulas',3200,0,0,0,3200,'RECEBIDA','EXT-5001','Venda Direta','Balcão janeiro'],
      ['CR-2002','2025-01-04','2025-01-14','','RECEITA','MATRIZ','COM','Receita Varejo','20102','Receita Varejo','ONLINE','Venda varejo online',2800,0,0,0,2800,'PENDENTE','','E-commerce','Pedido #234'],
      ['CR-2003','2025-01-06','2025-01-21','','RECEITA','FILIAL_RJ','COM','Receita Convênio','20108','Receita Convênio','PARCEIRO','Convenio Varejo',2100,0,0,0,2100,'PENDENTE','','Convênio Varejo','Ref. janeiro'],
    ];
    appendRows(SHEET_TB_LANCAMENTOS, seed);
    return seed.map(r => ({
      id: String(r[0]),
      dataCompetencia: normalizeDateCell(r[1]),
      dataVencimento: normalizeDateCell(r[2]),
      dataPagamento: normalizeDateCell(r[3]),
      tipo: String(r[4]),
      filial: String(r[5]),
      centroCusto: String(r[6]),
      contaGerencial: String(r[7]),
      contaContabil: String(r[8] ?? ''),
      grupoReceita: String(r[9] ?? ''),
      canal: String(r[10] ?? ''),
      descricao: String(r[11] ?? ''),
      valorBruto: parseFloat(String(r[12] || 0)),
      desconto: parseFloat(String(r[13] || 0)),
      juros: parseFloat(String(r[14] || 0)),
      multa: parseFloat(String(r[15] || 0)),
      valorLiquido: parseFloat(String(r[16] || 0)),
      status: String(r[17] || 'PENDENTE'),
      idExtratoBanco: String(r[18] || ''),
      origem: String(r[19] || ''),
      observacoes: String(r[20] || ''),
    }));
  }

  return data.slice(1).map((row: any) => ({
    id: String(row[0]),
    dataCompetencia: normalizeDateCell(row[1]),
    dataVencimento: normalizeDateCell(row[2]),
    dataPagamento: normalizeDateCell(row[3]),
    tipo: String(row[4] || ''),
    filial: String(row[5] || ''),
    centroCusto: String(row[6] || ''),
    contaGerencial: String(row[7] || ''),
    contaContabil: String(row[8] ?? ''),
    grupoReceita: String(row[9] ?? ''),
    canal: String(row[10] ?? ''),
    descricao: String(row[11] ?? ''),
    valorBruto: parseFloat(String(row[12] || 0)),
    desconto: parseFloat(String(row[13] || 0)),
    juros: parseFloat(String(row[14] || 0)),
    multa: parseFloat(String(row[15] || 0)),
    valorLiquido: parseFloat(String(row[16] || (parseFloat(String(row[12] || 0)) - parseFloat(String(row[13] || 0)) + parseFloat(String(row[14] || 0)) + parseFloat(String(row[15] || 0))))),
    status: String(row[17] || 'PENDENTE'),
    idExtratoBanco: String(row[18] || ''),
    origem: String(row[19] || ''),
    observacoes: String(row[20] || ''),
  })).map(l => {
    const tipoNorm = String(l.tipo || '').toUpperCase();
    if (tipoNorm === 'AP') l.tipo = 'DESPESA';
    else if (tipoNorm === 'AR') l.tipo = 'RECEITA';
    return l;
  });
}

function getExtratosFromSheet(): any[] {
  // Garante que a aba existe com cabeçalhos esperados
  createSheetIfNotExists(SHEET_TB_EXTRATOS, [
    'ID',
    'Data',
    'Descrição',
    'Valor',
    'Tipo',
    'Banco',
    'Conta',
    'Status Conciliação',
    'ID Lançamento',
    'Observações',
    'Importado Em',
  ]);

  const data = getSheetValues(SHEET_TB_EXTRATOS);
  if (!data || data.length <= 1) {
    if (!isSeedDataEnabled()) return [];
    const seed = [
      ['EXT-5001','2025-01-02','Recebimento cartão venda balcão',3200,'ENTRADA','BANCO_A','CC_MATRIZ','CONCILIADO','CR-2001','Pedido balcão','2025-01-03'],
      ['EXT-5002','2025-01-11','Pagamento fornecedor matéria-prima',-1500,'SAIDA','BANCO_A','CC_MATRIZ','CONCILIADO','CP-1001','Pagto lote A','2025-01-11'],
      ['EXT-5003','2025-01-15','Taxa bancária jan',-25,'SAIDA','BANCO_A','CC_MATRIZ','PENDENTE','','Tarifa débito','2025-01-15'],
      ['EXT-5004','2025-01-16','Recebimento boleto convênio',2100,'ENTRADA','BANCO_A','CC_MATRIZ','PENDENTE','CR-2003','Convênio varejo','2025-01-16'],
    ];
    appendRows(SHEET_TB_EXTRATOS, seed);
    return seed.map(r => ({
      id: r[0], data: r[1], descricao: r[2], valor: r[3], tipo: r[4], banco: r[5], conta: r[6],
      statusConciliacao: r[7], idLancamento: r[8], observacoes: r[9], importadoEm: r[10],
    }));
  }

  return data.slice(1).map((row: any) => ({
    id: row[0],
    data: normalizeDateCell(row[1]),
    descricao: row[2],
    valor: parseFloat(String(row[3] || 0)),
    tipo: row[4],
    banco: row[5],
    conta: row[6],
    statusConciliacao: row[7] || 'PENDENTE',
    idLancamento: row[8],
    observacoes: row[9],
    importadoEm: normalizeDateCell(row[10]),
  }));
}

function sumValues(items: any[]): number {
  return items.reduce((sum, item) => sum + parseFloat(String(item.valorLiquido || item.valor || 0)), 0);
}

function formatCurrency(value: number): string {
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value);
}

// ============================================================================
// DRE (Demonstração do Resultado do Exercício)
// ============================================================================

function getPlanoContasMap(): Record<string, any> {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_PLANO_CONTAS);
    if (!sheet) return {};

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return {};

    const lastCol = Math.max(8, sheet.getLastColumn());
    const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    const map: Record<string, any> = {};

    for (const r of rows) {
      const codigo = String(r[0] || '').trim();
      if (!codigo) continue;
      map[codigo] = {
        tipo: String(r[2] || '').trim(),
        grupoDRE: String(r[3] || '').trim(),
        subgrupoDRE: String(r[4] || '').trim(),
        grupoDFC: String(r[5] || '').trim(),
        variavelFixa: String(r[6] || '').trim(),
        cmaCmv: String(r[7] || '').trim(),
      };
    }

    return map;
  } catch (_e) {
    return {};
  }
}

export function getDREMensal(mes: number, ano: number, filial?: string): any {
  enforcePermission('visualizarRelatorios', 'carregar DRE');
  const cacheKey = `mensal:${ano}-${mes}:${filial || 'all'}`;
  return cacheGetOrLoad(CacheNamespace.DRE, cacheKey, () => {
  try {
    const lancamentos = getLancamentosFromSheet();
    const planoMap = getPlanoContasMap();

    // Filtrar por período e filial
    const lancamentosMes = lancamentos.filter(l => {
      const data = new Date(l.dataCompetencia);
      const mesLanc = data.getMonth() + 1; // JavaScript months are 0-indexed
      const anoLanc = data.getFullYear();

      const matchPeriodo = mesLanc === mes && anoLanc === ano;
      const matchFilial = !filial || l.filial === filial;

      return matchPeriodo && matchFilial;
    });

    // Separar receitas e despesas
    const receitas = lancamentosMes.filter(l => l.tipo === 'RECEITA');
    const despesas = lancamentosMes.filter(l => l.tipo === 'DESPESA');

    // Calcular valores
    const receitaBruta = sumValues(receitas.map(r => ({ valorLiquido: r.valorBruto })));
    const deducoes = sumValues(receitas.map(r => ({ valorLiquido: r.desconto })));
    const receitaLiquida = receitaBruta - deducoes;

    // Separar custos e despesas operacionais (baseado na conta contábil)
    const isCusto = (d: any) => {
      const codigo = String(d.contaContabil || '').trim();
      const meta = codigo ? planoMap[codigo] : null;
      const cmaCmv = String(meta?.cmaCmv || '').toUpperCase();
      const tipo = String(meta?.tipo || '').toUpperCase();
      const grupoDRE = String(meta?.grupoDRE || '').toUpperCase();
      return (
        cmaCmv === 'CMA' ||
        cmaCmv === 'CMV' ||
        tipo === 'CUSTO' ||
        grupoDRE.includes('CMV') ||
        grupoDRE.includes('CUSTO')
      );
    };

    const isFinanceiro = (d: any) => {
      const codigo = String(d.contaContabil || '').trim();
      const meta = codigo ? planoMap[codigo] : null;
      const grupoDRE = String(meta?.grupoDRE || '').toUpperCase();
      return grupoDRE.includes('FINANCEIRO') || grupoDRE.includes('RESULTADO FINANCEIRO');
    };

    const custos = despesas.filter((d) => isCusto(d));
    const despesasFinanceiras = despesas.filter((d) => !isCusto(d) && isFinanceiro(d));
    const despesasOp = despesas.filter((d) => !isCusto(d) && !isFinanceiro(d));

    const totalCustos = sumValues(custos);
    const margemBruta = receitaLiquida - totalCustos;
    const percMargemBruta = receitaLiquida > 0 ? (margemBruta / receitaLiquida) * 100 : 0;

    // Despesas operacionais por categoria (baseado em centro de custo)
    const despPessoal = despesasOp.filter(d => d.centroCusto === 'ADM' || d.centroCusto === 'OPS');
    const despMarketing = despesasOp.filter(d => d.centroCusto === 'MKT' || d.centroCusto === 'COM');
    const despAdministrativas = despesasOp.filter(d => d.centroCusto === 'FIN' || d.centroCusto === 'TI');

    const totalDespOp = sumValues(despesasOp);
    const ebitda = margemBruta - totalDespOp;
    const percEbitda = receitaLiquida > 0 ? (ebitda / receitaLiquida) * 100 : 0;

    // Resultado Financeiro
    const receitasFinanceiras = receitas.filter((r) => {
      const codigo = String(r.contaContabil || '').trim();
      const meta = codigo ? planoMap[codigo] : null;
      const grupoDRE = String(meta?.grupoDRE || '').toUpperCase();
      return grupoDRE.includes('FINANCEIRO') || grupoDRE.includes('RESULTADO FINANCEIRO');
    });
    const totalResultadoFinanceiro =
      sumValues(receitasFinanceiras) - sumValues(despesasFinanceiras);

    const lucroLiquido = ebitda + totalResultadoFinanceiro;
    const percLucroLiquido = receitaLiquida > 0 ? (lucroLiquido / receitaLiquida) * 100 : 0;

    return {
      periodo: {
        mes,
        ano,
        mesNome: getMesNome(mes),
        filial: filial || 'Consolidado'
      },
      valores: {
        receitaBruta,
        deducoes,
        receitaLiquida,
        custos: totalCustos,
        margemBruta,
        despesasOperacionais: {
          pessoal: sumValues(despPessoal),
          marketing: sumValues(despMarketing),
          administrativas: sumValues(despAdministrativas),
          total: totalDespOp
        },
        ebitda,
        resultadoFinanceiro: totalResultadoFinanceiro,
        lucroLiquido
      },
      percentuais: {
        margemBruta: percMargemBruta,
        ebitda: percEbitda,
        lucroLiquido: percLucroLiquido
      },
      classificacao: {
        margemBruta: classificarIndicador(percMargemBruta, 'margem_bruta'),
        ebitda: classificarIndicador(percEbitda, 'ebitda'),
        lucroLiquido: classificarIndicador(percLucroLiquido, 'lucro_liquido')
      },
      transacoes: {
        totalReceitas: receitas.length,
        totalDespesas: despesas.length
      }
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular DRE: ${error.message}`);
    throw new Error(`Erro ao calcular DRE: ${error.message}`);
  }
  }, 120, CacheScope.SCRIPT);
}

export function getDREComparativo(meses: Array<{ mes: number; ano: number }>, filial?: string): any {
  enforcePermission('visualizarRelatorios', 'carregar DRE comparativo');
  try {
    const dres = meses.map(periodo => getDREMensal(periodo.mes, periodo.ano, filial));

    return {
      periodos: dres.map(d => d.periodo),
      comparativo: dres,
      evolucao: {
        receitaLiquida: calcularEvolucao(dres.map(d => d.valores.receitaLiquida)),
        margemBruta: calcularEvolucao(dres.map(d => d.valores.margemBruta)),
        ebitda: calcularEvolucao(dres.map(d => d.valores.ebitda)),
        lucroLiquido: calcularEvolucao(dres.map(d => d.valores.lucroLiquido))
      }
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular DRE comparativo: ${error.message}`);
    throw new Error(`Erro ao calcular DRE comparativo: ${error.message}`);
  }
}

export function getDREPorFilial(mes: number, ano: number): any {
  enforcePermission('visualizarRelatorios', 'carregar DRE por filial');
  try {
    const lancamentos = getLancamentosFromSheet();

    // Obter lista única de filiais
    const filiais = [...new Set(lancamentos.map(l => l.filial))].filter(f => f);

    // Calcular DRE para cada filial
    const dresPorFilial = filiais.map(filial => ({
      filial,
      dre: getDREMensal(mes, ano, filial)
    }));

    // DRE consolidado
    const dreConsolidado = getDREMensal(mes, ano);

    return {
      consolidado: dreConsolidado,
      porFilial: dresPorFilial,
      filiais
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular DRE por filial: ${error.message}`);
    throw new Error(`Erro ao calcular DRE por filial: ${error.message}`);
  }
}

// Helper functions
function getMesNome(mes: number): string {
  const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
  return meses[mes - 1] || '';
}

function classificarIndicador(percentual: number, tipo: string): string {
  // Benchmarks baseados no PDF do comitê
  const benchmarks: any = {
    margem_bruta: { sensacional: 60, excelente: 50, bom: 40, ruim: 30 },
    ebitda: { sensacional: 25, excelente: 20, bom: 15, ruim: 10 },
    lucro_liquido: { sensacional: 20, excelente: 15, bom: 10, ruim: 5 }
  };

  const bench = benchmarks[tipo] || benchmarks.ebitda;

  if (percentual >= bench.sensacional) return 'Sensacional';
  if (percentual >= bench.excelente) return 'Excelente';
  if (percentual >= bench.bom) return 'Bom';
  if (percentual >= bench.ruim) return 'Ruim';
  return 'Péssimo';
}

function calcularEvolucao(valores: number[]): any {
  if (valores.length < 2) return { percentual: 0, tendencia: 'estavel' };

  const primeiro = valores[0];
  const ultimo = valores[valores.length - 1];

  if (primeiro === 0) return { percentual: 0, tendencia: 'estavel' };

  const percentual = ((ultimo - primeiro) / Math.abs(primeiro)) * 100;
  const tendencia = percentual > 5 ? 'crescimento' : percentual < -5 ? 'queda' : 'estavel';

  return { percentual, tendencia };
}

// ============================================================================
// FLUXO DE CAIXA (DFC)
// ============================================================================

export function getFluxoCaixaMensal(mes: number, ano: number, filial?: string, saldoInicial?: number): any {
  enforcePermission('visualizarRelatorios', 'carregar fluxo de caixa');
  const cacheKey = `mensal:${ano}-${mes}:${filial || 'all'}:${Number(saldoInicial) || 0}`;
  return cacheGetOrLoad(CacheNamespace.DFC, cacheKey, () => {
  try {
    const lancamentos = getLancamentosFromSheet();

    // Filtrar por período e filial
    const lancamentosMes = lancamentos.filter(l => {
      const data = new Date(l.dataCompetencia);
      const mesLanc = data.getMonth() + 1;
      const anoLanc = data.getFullYear();

      const matchPeriodo = mesLanc === mes && anoLanc === ano;
      const matchFilial = !filial || l.filial === filial;

      return matchPeriodo && matchFilial;
    });

    // Separar por tipo e status
      const isPago = (s: string) => ['PAGO', 'PAGA', 'RECEBIDO', 'RECEBIDA'].includes((s || '').toUpperCase());
      const entradas = lancamentosMes.filter(l =>
        l.tipo === 'RECEITA' && isPago(l.status)
      );
      const saidas = lancamentosMes.filter(l =>
        l.tipo === 'DESPESA' && isPago(l.status)
      );

    // Calcular totais
    const totalEntradas = sumValues(entradas);
    const totalSaidas = sumValues(saidas);

    // Saldo inicial: input manual (quando informado) ou 0
    const saldoInicialNum =
      typeof saldoInicial === 'number' && !isNaN(saldoInicial) ? saldoInicial : 0;
    const saldoFinal = saldoInicialNum + totalEntradas - totalSaidas;
    const variacao =
      saldoInicialNum !== 0
        ? ((saldoFinal - saldoInicialNum) / Math.abs(saldoInicialNum)) * 100
        : 0;

    // Agrupar entradas por categoria (conta contábil)
    const entradasPorCategoria = agruparPorCategoria(entradas);
    const saidasPorCategoria = agruparPorCategoria(saidas);

    return {
      periodo: {
        mes,
        ano,
        mesNome: getMesNome(mes),
        filial: filial || 'Consolidado'
      },
      valores: {
        saldoInicial: saldoInicialNum,
        totalEntradas,
        totalSaidas,
        saldoFinal,
        variacao
      },
      transacoes: {
        qtdEntradas: entradas.length,
        qtdSaidas: saidas.length
      },
      entradas: entradasPorCategoria,
      saidas: saidasPorCategoria
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular Fluxo de Caixa: ${error.message}`);
    throw new Error(`Erro ao calcular Fluxo de Caixa: ${error.message}`);
  }
  }, 120, CacheScope.SCRIPT);
}

export function getFluxoCaixaProjecao(meses: number, filial?: string): any {
  enforcePermission('visualizarRelatorios', 'carregar projeção de fluxo de caixa');
  try {
    const hoje = new Date();
    const periodos = [];

    for (let i = 0; i < meses; i++) {
      const data = new Date(hoje.getFullYear(), hoje.getMonth() + i, 1);
      periodos.push({
        mes: data.getMonth() + 1,
        ano: data.getFullYear()
      });
    }

    const fluxos = periodos.map(p => getFluxoCaixaMensal(p.mes, p.ano, filial));

    return {
      periodos: fluxos.map(f => f.periodo),
      fluxos,
      evolucao: {
        entradas: calcularEvolucao(fluxos.map(f => f.valores.totalEntradas)),
        saidas: calcularEvolucao(fluxos.map(f => f.valores.totalSaidas)),
        saldo: calcularEvolucao(fluxos.map(f => f.valores.saldoFinal))
      }
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular projeção de Fluxo de Caixa: ${error.message}`);
    throw new Error(`Erro ao calcular projeção de Fluxo de Caixa: ${error.message}`);
  }
}

// Helper function para agrupar por categoria
function agruparPorCategoria(lancamentos: any[]): any[] {
  const categorias: any = {};

  lancamentos.forEach(l => {
    const categoria = l.categoria || 'Outros';
    if (!categorias[categoria]) {
      categorias[categoria] = {
        categoria,
        valor: 0,
        quantidade: 0
      };
    }
    categorias[categoria].valor += parseFloat(String(l.valorLiquido || l.valor || 0));
    categorias[categoria].quantidade++;
  });

  return Object.values(categorias).sort((a: any, b: any) => b.valor - a.valor);
}

// ============================================================================
// KPIs FINANCEIROS
// ============================================================================

export function getKPIsMensal(mes: number, ano: number, filial?: string): any {
  enforcePermission('visualizarRelatorios', 'carregar KPIs');
  const cacheKey = `mensal:${ano}-${mes}:${filial || 'all'}`;
  return cacheGetOrLoad(CacheNamespace.KPI, cacheKey, () => {
  try {
    const lancamentos = getLancamentosFromSheet();

    // Filtrar lançamentos do mês atual
    const lancamentosMes = lancamentos.filter(l => {
      const data = new Date(l.dataCompetencia);
      const mesLanc = data.getMonth() + 1;
      const anoLanc = data.getFullYear();
      const matchPeriodo = mesLanc === mes && anoLanc === ano;
      const matchFilial = !filial || l.filial === filial;
      return matchPeriodo && matchFilial;
    });

    // Filtrar lançamentos do mês anterior
    const dataAnterior = new Date(ano, mes - 2, 1); // mes - 2 porque JavaScript months são 0-indexed
    const mesAnterior = dataAnterior.getMonth() + 1;
    const anoAnterior = dataAnterior.getFullYear();

    const lancamentosMesAnterior = lancamentos.filter(l => {
      const data = new Date(l.dataCompetencia);
      const mesLanc = data.getMonth() + 1;
      const anoLanc = data.getFullYear();
      const matchPeriodo = mesLanc === mesAnterior && anoLanc === anoAnterior;
      const matchFilial = !filial || l.filial === filial;
      return matchPeriodo && matchFilial;
    });

    // Calcular DRE do mês atual e anterior
    const dreAtual = getDREMensal(mes, ano, filial);
    const dreAnterior = getDREMensal(mesAnterior, anoAnterior, filial);
    const fcAtual = getFluxoCaixaMensal(mes, ano, filial);

    // Separar receitas e despesas
    const receitas = lancamentosMes.filter(l => l.tipo === 'RECEITA');
    const despesas = lancamentosMes.filter(l => l.tipo === 'DESPESA');
    const receitasAnterior = lancamentosMesAnterior.filter(l => l.tipo === 'RECEITA');

    // KPIs de Rentabilidade
    const margemBruta = dreAtual.percentuais.margemBruta;
    const margemEbitda = dreAtual.percentuais.ebitda;
    const margemLiquida = dreAtual.percentuais.lucroLiquido;
    const roi = dreAtual.valores.receitaLiquida > 0
      ? (dreAtual.valores.lucroLiquido / dreAtual.valores.receitaLiquida) * 100
      : 0;

    // KPIs de Liquidez
    const contasReceber = lancamentosMes.filter(l => l.tipo === 'RECEITA' && l.status === 'PENDENTE');
    const contasPagar = lancamentosMes.filter(l => l.tipo === 'DESPESA' && l.status === 'PENDENTE');
    const ativoCirculante = sumValues(contasReceber) + fcAtual.valores.saldoFinal;
    const passivoCirculante = sumValues(contasPagar);
    const liquidezCorrente = passivoCirculante > 0 ? ativoCirculante / passivoCirculante : 0;

    const saldoCaixa = fcAtual.valores.saldoFinal;
    const burnRate = Math.abs(dreAtual.valores.lucroLiquido < 0 ? dreAtual.valores.lucroLiquido : 0);
    const runway = burnRate > 0 ? saldoCaixa / burnRate : 999;

    // KPIs de Crescimento
    const receitaAtual = dreAtual.valores.receitaLiquida;
    const receitaAnteriorVal = dreAnterior.valores.receitaLiquida;
    const crescimentoReceita = receitaAnteriorVal > 0
      ? ((receitaAtual - receitaAnteriorVal) / receitaAnteriorVal) * 100
      : 0;

    const ticketMedio = receitas.length > 0 ? receitaAtual / receitas.length : 0;

    const receitasVencidas = lancamentosMes.filter(l => {
      if (l.tipo !== 'RECEITA' || l.status !== 'PENDENTE') return false;
      const vencimento = new Date(l.dataVencimento);
      return vencimento < new Date();
    });
    const taxaInadimplencia = receitas.length > 0
      ? (receitasVencidas.length / receitas.length) * 100
      : 0;

    // Prazo médio de recebimento
    const receitasRecebidas = lancamentosMes.filter(l =>
      l.tipo === 'RECEITA' && (l.status === 'PAGO' || l.status === 'RECEBIDO')
    );
    let prazoMedioRecebimento = 0;
    if (receitasRecebidas.length > 0) {
      const prazos = receitasRecebidas.map(r => {
        const venc = new Date(r.dataVencimento);
        const pag = new Date(r.dataPagamento || r.dataCompetencia);
        return Math.floor((pag.getTime() - venc.getTime()) / (1000 * 60 * 60 * 24));
      });
      prazoMedioRecebimento = prazos.reduce((a, b) => a + b, 0) / prazos.length;
    }

    // KPIs Operacionais
    const despesasMarketing = despesas.filter(d => d.centroCusto === 'MKT' || d.centroCusto === 'COM');
    const cac = receitas.length > 0 ? sumValues(despesasMarketing) / receitas.length : 0;

    const despOperacionaisPerc = dreAtual.valores.receitaLiquida > 0
      ? (dreAtual.valores.despesasOperacionais.total / dreAtual.valores.receitaLiquida) * 100
      : 0;

    const breakEven = dreAtual.valores.margemBruta > 0
      ? dreAtual.valores.despesasOperacionais.total / (dreAtual.valores.margemBruta / dreAtual.valores.receitaLiquida)
      : 0;

    // Prazo médio de pagamento
    const despesasPagas = lancamentosMes.filter(l =>
      l.tipo === 'DESPESA' && l.status === 'PAGO'
    );
    let prazoMedioPagamento = 0;
    if (despesasPagas.length > 0) {
      const prazos = despesasPagas.map(d => {
        const venc = new Date(d.dataVencimento);
        const pag = new Date(d.dataPagamento || d.dataCompetencia);
        return Math.floor((pag.getTime() - venc.getTime()) / (1000 * 60 * 60 * 24));
      });
      prazoMedioPagamento = prazos.reduce((a, b) => a + b, 0) / prazos.length;
    }

    return {
      periodo: {
        mes,
        ano,
        mesNome: getMesNome(mes),
        filial: filial || 'Consolidado'
      },
      rentabilidade: {
        margemBruta: {
          valor: margemBruta,
          classificacao: classificarIndicador(margemBruta, 'margem_bruta')
        },
        margemEbitda: {
          valor: margemEbitda,
          classificacao: classificarIndicador(margemEbitda, 'ebitda')
        },
        margemLiquida: {
          valor: margemLiquida,
          classificacao: classificarIndicador(margemLiquida, 'lucro_liquido')
        },
        roi: {
          valor: roi,
          descricao: roi > 0 ? 'Positivo' : 'Negativo'
        }
      },
      liquidez: {
        liquidezCorrente: {
          valor: liquidezCorrente,
          classificacao: classificarLiquidez(liquidezCorrente)
        },
        saldoCaixa: {
          valor: saldoCaixa,
          descricao: saldoCaixa > 0 ? 'Saudável' : 'Atenção'
        },
        burnRate: {
          valor: burnRate,
          descricao: `${burnRate > 0 ? 'Queimando' : 'Gerando'} caixa`
        },
        runway: {
          valor: runway,
          descricao: runway < 6 ? 'Crítico' : runway < 12 ? 'Atenção' : 'Saudável'
        }
      },
      crescimento: {
        crescimentoReceita: {
          valor: crescimentoReceita,
          mesAnterior: receitaAnteriorVal
        },
        ticketMedio: {
          valor: ticketMedio,
          qtdTransacoes: receitas.length
        },
        inadimplencia: {
          valor: taxaInadimplencia,
          classificacao: taxaInadimplencia < 5 ? 'Excelente' : taxaInadimplencia < 10 ? 'Bom' : 'Ruim'
        },
        prazoMedioRecebimento: {
          valor: prazoMedioRecebimento,
          descricao: `${prazoMedioRecebimento.toFixed(0)} dias`
        }
      },
      operacional: {
        cac: {
          valor: cac,
          descricao: `Custo por cliente`
        },
        despOperacionaisReceita: {
          valor: despOperacionaisPerc,
          classificacao: despOperacionaisPerc < 30 ? 'Excelente' : despOperacionaisPerc < 50 ? 'Bom' : 'Ruim'
        },
        breakEven: {
          valor: breakEven,
          descricao: `Ponto de equilíbrio`
        },
        prazoMedioPagamento: {
          valor: prazoMedioPagamento,
          descricao: `${prazoMedioPagamento.toFixed(0)} dias`
        }
      },
      comparativo: {
        mesAnterior: {
          receita: receitaAnteriorVal,
          margemBruta: dreAnterior.percentuais.margemBruta,
          margemLiquida: dreAnterior.percentuais.lucroLiquido
        }
      }
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular KPIs: ${error.message}`);
    throw new Error(`Erro ao calcular KPIs: ${error.message}`);
  }
  }, 120, CacheScope.SCRIPT);
}

// Helper function para classificar liquidez
function classificarLiquidez(valor: number): string {
  if (valor >= 2) return 'Excelente';
  if (valor >= 1.5) return 'Bom';
  if (valor >= 1) return 'Aceitável';
  return 'Ruim';
}

// ============================================================================
// USUÁRIOS E PERMISSÕES
// ============================================================================

interface Usuario {
  id?: string;
  email: string;
  nome: string;
  perfil: 'ADMIN' | 'GESTOR' | 'OPERACIONAL' | 'VISUALIZADOR';
  status: 'ATIVO' | 'INATIVO';
  ultimoAcesso?: string;
  permissoes?: {
    criarLancamentos: boolean;
    editarLancamentos: boolean;
    excluirLancamentos: boolean;
    aprovarPagamentos: boolean;
    visualizarRelatorios: boolean;
    gerenciarConfig: boolean;
  };
}

export function getUsuarios(): Usuario[] {
  try {
    const email = getRequestingUserEmail();
    const requester = getUsuarioByEmail(email);
    if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
      return [];
    }

    const sheet = ensureUsuariosSheet();

    const data = sheet.getDataRange().getValues();
    const usuarios: Usuario[] = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // Skip empty rows

      usuarios.push({
        id: String(row[0]),
        email: String(row[1]),
        nome: String(row[2]),
        perfil: String(row[3]) as any,
        status: String(row[4]) as any,
        ultimoAcesso: row[5] ? String(row[5]) : undefined,
        permissoes: row[6] ? JSON.parse(String(row[6])) : getPermissoesPadrao(String(row[3]))
      });
    }

    return usuarios;
  } catch (error: any) {
    Logger.log(`Erro ao buscar usuários: ${error.message}`);
    return [];
  }
}

export function salvarUsuario(usuario: Usuario): { success: boolean; message: string } {
  try {
    const email = getRequestingUserEmail();
    const requester = getUsuarioByEmail(email);
    if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
      return { success: false, message: 'Sem permissão: salvar usuário' };
    }

    const v = combineValidations(
      validateRequired(usuario?.email, 'Email'),
      validateRequired(usuario?.nome, 'Nome'),
      validateEnum(String(usuario?.perfil || ''), ['ADMIN', 'GESTOR', 'OPERACIONAL', 'VISUALIZADOR'], 'Perfil'),
      validateEnum(String(usuario?.status || ''), ['ATIVO', 'INATIVO'], 'Status')
    );
    if (!v.valid) return { success: false, message: v.errors.join('; ') };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureUsuariosSheet();

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    // Procurar usuário existente
    if (usuario.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === usuario.id) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    const permissoes = usuario.permissoes || getPermissoesPadrao(usuario.perfil);
    const rowData = [
      usuario.id || Utilities.getUuid(),
      sanitizeSheetString(usuario.email).toLowerCase(),
      sanitizeSheetString(usuario.nome),
      sanitizeSheetString(usuario.perfil),
      sanitizeSheetString(usuario.status),
      usuario.ultimoAcesso || '',
      JSON.stringify(permissoes),
      rowIndex === -1 ? new Date().toISOString() : data[rowIndex - 1][7]
    ];

    if (rowIndex > 0) {
      // Atualizar existente
      sheet.getRange(rowIndex, 1, 1, 8).setValues([rowData]);
      appendAuditLog('salvarUsuario', { id: rowData[0], email: rowData[1], perfil: rowData[3], status: rowData[4] }, true);
      return { success: true, message: 'Usuário atualizado com sucesso!' };
    } else {
      // Criar novo
      sheet.appendRow(rowData);
      appendAuditLog('salvarUsuario', { id: rowData[0], email: rowData[1], perfil: rowData[3], status: rowData[4] }, true);
      return { success: true, message: 'Usuário criado com sucesso!' };
    }
  } catch (error: any) {
    appendAuditLog('salvarUsuario', { usuario }, false, error?.message);
    Logger.log(`Erro ao salvar usuário: ${error.message}`);
    return { success: false, message: `Erro ao salvar usuário: ${error.message}` };
  }
}

export function excluirUsuario(id: string): { success: boolean; message: string } {
  try {
    const email = getRequestingUserEmail();
    const requester = getUsuarioByEmail(email);
    if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
      return { success: false, message: 'Sem permissão: excluir usuário' };
    }

    const v = validateRequired(id, 'ID');
    if (!v.valid) return { success: false, message: v.errors.join('; ') };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USUARIOS);

    if (!sheet) {
      return { success: false, message: 'Planilha de usuários não encontrada' };
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id || data[i][1] === id) {
        sheet.deleteRow(i + 1);
        appendAuditLog('excluirUsuario', { id }, true);
        return { success: true, message: 'Usuário excluído com sucesso!' };
      }
    }

    return { success: false, message: 'Usuário não encontrado' };
  } catch (error: any) {
    appendAuditLog('excluirUsuario', { id }, false, error?.message);
    Logger.log(`Erro ao excluir usuário: ${error.message}`);
    return { success: false, message: `Erro ao excluir usuário: ${error.message}` };
  }
}

// -----------------------------------------------------------------------------
// SEED: PLANO DE CONTAS (sobrescreve aba REF_PLANO_CONTAS)
// -----------------------------------------------------------------------------
export function seedPlanoContasFromList(): { success: boolean; message: string } {
  const denied = requirePermission('gerenciarConfig', 'seed plano de contas');
  if (denied) return denied;
  const requester = getUsuarioByEmail(getRequestingUserEmail());
  if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
    return { success: false, message: 'Sem permissão: seed plano de contas' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_REF_PLANO_CONTAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_REF_PLANO_CONTAS);
  }

  const header = [['Código', 'Descrição', 'Tipo']];
  const data: Array<[string, string]> = [
    ['1', 'DESPESAS'],
    ['101', 'IMPOSTOS'],
    ['10101', 'ICMS - VENDAS'],
    ['10102', 'SIMPLES'],
    ['10103', 'SETEC'],
    ['10104', 'ISSQN - COMPRAS'],
    ['10105', 'DARF'],
    ['10106', 'TAXAS PREFEITURA'],
    ['10107', 'DARF IRRF FOLHA PGTO'],
    ['10108', 'IPI'],
    ['10109', 'COFINS'],
    ['10110', 'PIS'],
    ['10111', 'IRPJ'],
    ['10112', 'CSLL'],
    ['10113', 'ISSQN - VENDAS'],
    ['10114', 'ICMS - COMPRAS'],
    ['10115', 'INSS RETIDO NF'],
    ['10116', 'IPVA E TAXAS - VEICULOS'],
    ['10199', 'IMPOSTOS EVENTUAIS'],
    ['102', 'FORNECEDORES'],
    ['10201', 'MATERIA-PRIMA'],
    ['10202', 'PRODUTOS DE REVENDA'],
    ['10203', 'EMBALAGENS'],
    ['10204', 'PRODUTOS DE OUTRA FARMACIA LAB'],
    ['10205', 'FRETES COMPRAS'],
    ['10206', 'MEDICAMENTOS HOMEOPATIA E CLIENTES'],
    ['10208', 'FORNECEDOR LIMPEZA E HIGIENE'],
    ['10209', 'ADIANTAMENTO DE FORNECEDOR'],
    ['10210', 'EMPRESTIMO MUTUO CARVALHO'],
    ['10211', 'ADTO DISTRIBUIÇÃO DE LUCRO'],
    ['10213', 'FORNECEDOR CONSUMO INTERNO'],
    ['10299', 'FORNECEDORES EVENTUAIS'],
    ['103', 'DESPESAS COM FUNCIONARIOS'],
    ['10301', 'SALARIOS - OPERACIONAL/LABS/EXPEDIC'],
    ['10302', 'FERIAS'],
    ['10303', 'DECIMO-TERCEIRO'],
    ['10304', 'RESCISOES'],
    ['10305', 'HORAS-EXTRAS'],
    ['10306', 'CONVENIO MEDICO COLABORADORES'],
    ['10307', 'FGTS'],
    ['10308', 'CESTA BASICA'],
    ['10309', 'PCMSO/PPRA/EXAMES'],
    ['10310', 'CURSOS E TREINAMENTOS FUNCIONARIOS'],
    ['10311', 'VALE-TRANSPORTE / CONDUCAO'],
    ['10312', 'INSS FUNCIONARIO'],
    ['10313', 'RECRUTAMENTO / SELECAO'],
    ['10314', 'UNIFORME/MATERIAL TRABALHO'],
    ['10315', 'PREMIACAO METAS'],
    ['10317', 'PREMIACAO ANUAL'],
    ['10318', 'SEGURO FUNCIONARIOS'],
    ['10319', 'CONVENIO ODONTO COLABORADORES'],
    ['10320', 'CONTRIBUICAO SINDICAL FUNCIONARIO'],
    ['10321', 'INSS EMPRESA PATRONAL'],
    ['10322', 'PRESTACAO SERVICO - IN LOCO'],
    ['10323', 'REFEICAO/ALIMENTACAO/IFOOD'],
    ['10324', 'CRACHAS PARA FUNCIONARIOS'],
    ['10325', 'SALARIOS - COMERCIAL'],
    ['10326', 'SALARIOS - VIRTUAL'],
    ['10327', 'SALARIOS - ADM/MKT'],
    ['10328', 'SALARIOS - CPTF'],
    ['10329', 'PREMIACAO METAS - VISITACAO'],
    ['10399', 'FUNCIONARIOS EVENTUAIS'],
    ['104', 'DESPESAS ADMINISTRATIVAS (ESCRIT.)'],
    ['10401', 'GRAFICA/IMPRESSOS'],
    ['10402', 'HONORARIOS CONTADOR'],
    ['10403', 'CONSULTORIA/ASSESSORIA - ADM'],
    ['10404', 'MATERIAL DE PAPELARIA'],
    ['10405', 'ALUGUEL EQUIP - GERAL'],
    ['10406', 'REGULATORIOS - ALVARA/VISA/CRF/POL'],
    ['10407', 'CORRESPONDENCIA CORREIO'],
    ['10408', 'MANUT EQUIPAMENTOS ESCRITORIO'],
    ['10409', 'ADVOGADO'],
    ['10410', 'ESTACIONAMENTO - ALUGUEL/AVULSO'],
    ['10412', 'ALMOXARIFADO'],
    ['10414', 'DEVOLUCAO AO CLIENTE'],
    ['10415', 'CARTORIO'],
    ['10416', 'AGUA GALAO'],
    ['10499', 'ADMINISTRATIVAS EVENTUAIS'],
    ['105', 'DESPESAS COM VEICULOS'],
    ['10501', 'COMBUSTIVEL'],
    ['10502', 'MANUTENCAO VEICULOS'],
    ['10503', 'MULTAS VEICULOS'],
    ['10504', 'SEGURO VEICULOS'],
    ['10599', 'VEICULOS - EVENTUAIS'],
    ['106', 'DESPESAS COM INFORMATICA'],
    ['10601', 'MANUT EQUIPAMENTOS INFORMATICA'],
    ['10602', 'MATERIAL DE INFORMATICA'],
    ['10603', 'CONSULTORIA E ASSESSORIA - INFOR'],
    ['10604', 'PROGRAMA DE INFORMATICA'],
    ['10699', 'INFORMATICA EVENTUAIS'],
    ['107', 'DESPESAS GERAIS'],
    ['10701', 'PROVISOES DIVERSAS - FLUXO CAIXA'],
    ['10702', 'CONTRIBUICOES/DOACOES'],
    ['10704', 'CONFRATERNIZAÇÃO/REUNIÃO'],
    ['10705', 'ASSOCIACOES'],
    ['10706', 'CONTRIBUICAO SINDICAL PATRONAL'],
    ['10707', 'DESPESA DE USO E CONSUMO'],
    ['108', 'DESPESAS DE COMUNICAÇÃO'],
    ['10801', 'TELEFONE FIXO'],
    ['10802', 'MANUTENCAO EQUIPAMENTOS TELEF'],
    ['10803', 'INTERNET'],
    ['10804', 'TELEFONE CELULAR'],
    ['10805', 'MATERIAL DE TELEFONIA'],
    ['10899', 'TELEFONIA EVENTUAIS'],
    ['109', 'DESPESAS FINANCEIRAS'],
    ['10901', 'JUROS'],
    ['10902', 'DESPESAS FINANCEIRAS/BANCARIAS'],
    ['10904', 'CUSTO TAXA CARTAO CREDITO'],
    ['10905', 'PAGAMENTO EMPRESTIMOS'],
    ['10906', 'IOF OPERACOES FINANCEIRAS'],
    ['10999', 'FINANCEIRAS EVENTUAIS'],
    ['110', 'MKT/VISITACAO/COMERCIALIZACAO'],
    ['11001', 'PUBLICIDADE/ANUNCIOS/PUBLICAC - MKT'],
    ['11002', 'CONSULTORIA/ASSESSORIA - MKT'],
    ['11003', 'EVENTOS - VISITACAO'],
    ['11004', 'EVENTOS - MKT'],
    ['11005', 'FRETE VENDAS - COMERCIALIZACAO'],
    ['11006', 'MKT M - VISITACAO'],
    ['11007', 'BRINDES/CORTESIAS - VISITACAO'],
    ['11008', 'BRINDES/CORTESIAS - MKT'],
    ['11009', 'NAO USAR'],
    ['11011', 'CONSULTORIA/ASSESSORIA - VISITACAO'],
    ['11012', 'CORTESIA FORMULAS E VAREJOS'],
    ['11013', 'P&D - PESQUISA E DESENVIMENTO'],
    ['11014', 'MKT PAGO POR FORNECEDOR'],
    ['11015', 'UNIFORMES DE CAMPANHAS - MKT'],
    ['11016', 'REEMBOLSO DESLOCAMENTO - VISITACAO'],
    ['11099', 'DESPESAS EVENTUAIS - MKT/VISITACAO'],
    ['111', 'DESPESAS COM IMÓVEIS'],
    ['11101', 'ALUGUEL COM IMOVEIS'],
    ['11102', 'ÁGUA'],
    ['11103', 'ENERGIA ELETRICA'],
    ['11104', 'MANUT IMOVEIS - MAO DE OBRA'],
    ['11105', 'SEGUROS COM IMOVEIS'],
    ['11106', 'IMPOSTOS E TAXAS'],
    ['11107', 'SEGURANCA COM IMOVEIS'],
    ['11108', 'MANUT AR CONDICIONADO'],
    ['11109', 'MANUT IMOVEIS - MATERIAL'],
    ['11199', 'IMOVEIS EVENTUAIS'],
    ['112', 'DESPESAS OPERACIONAIS (LAB.)'],
    ['11201', 'MANUT EQUIPAMENTOS LABORATORIO'],
    ['11202', 'MATERIAL PARA LABORATORIO'],
    ['11203', 'CONTROLE DE QUALIDADE'],
    ['11204', 'EPI - EQUIPAMENTO PROTECAO INDIV'],
    ['11205', 'PRESTACAO SERVICOS - LAB'],
    ['11206', 'ALUGUEL EQUIP - LABORATORIO'],
    ['11299', 'LABORATORIO EVENTUAIS'],
    ['113', 'DESPESAS DIRETORIA'],
    ['11301', 'PRO LABORE'],
    ['11302', 'DESPESAS/VIAGENS/REFEICOES'],
    ['11303', 'DESPESAS COM VEICULOS'],
    ['11304', 'CURSOS DIRETORIA'],
    ['11305', 'INSS SOBRE PRO LABORE'],
    ['11306', 'DESPESAS PESSOAIS DIRETORIA'],
    ['114', 'INVESTIMENTOS'],
    ['11401', 'AQUISICAO IMOB. - MOVEIS/UTENSILIOS'],
    ['11402', 'AQUISICAO IMOB. - COMERCIAL/MKT'],
    ['11403', 'AQUISICAO IMOB. - INFOR/TELEFONIA'],
    ['11404', 'REFORMA/EXPANSAO'],
    ['11405', 'AQUISICAO IMOB. - LABORATORIO'],
    ['11406', 'ESTOQUE TINTURAS/MATRIZES'],
    ['11407', 'MARCAS / PATENTES / ETC'],
    ['11408', 'APLICACOES/INVESTIMENTOS FINANCEIRO'],
    ['11409', 'AQUISICAO IMOB. - VEICULOS'],
    ['11410', 'AQUISICAO IMOB. - PREDIAL'],
    ['11499', 'IMOBILIZADO EVENTUAIS'],
    ['2', 'RECEITAS'],
    ['201', 'RECEITAS GERAL'],
    ['20101', 'VENDA A VISTA DE FORMULAS'],
    ['20102', 'VENDA A VISTA DE VAREJO'],
    ['20103', 'VENDA A PRAZO DE FORMULAS'],
    ['20104', 'VENDA A PRAZO DE VAREJO'],
    ['20105', 'VENDA CHEQUE PRE-DATADO'],
    ['20106', 'VENDA CARTAO DE CREDITO'],
    ['20107', 'RECEB. CONVENIO FORMULAS'],
    ['20108', 'RECEB. CONVENIO VAREJO'],
    ['20109', 'RECEB. CHEQUE DEVOLVIDO'],
    ['20110', 'ENTRADA PARA ACERTO'],
    ['20111', 'EMPRESTIMO'],
    ['20112', 'VENDA PARA OUTRA FARMACIA'],
    ['20113', 'RECEBIMENTO DE MULTAS'],
    ['20199', 'RECEITAS DIVERSAS'],
  ];

  const rows = data.map(([codigo, descricao]) => {
    const tipo = String(codigo).startsWith('2') ? 'RECEITA' : 'DESPESA';
    return [String(codigo), String(descricao), tipo];
  });

  sheet!.clearContents();
  sheet!.getRange(1, 1, 1, 3).setValues(header);
  sheet!.getRange('A:C').setNumberFormat('@');
  if (rows.length > 0) {
    sheet!.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  return { success: true, message: `Plano de contas atualizado com ${rows.length} linhas` };
}

function getPermissoesPadrao(perfil: string): any {
  const permissoes: any = {
    'ADMIN': {
      criarLancamentos: true,
      editarLancamentos: true,
      excluirLancamentos: true,
      aprovarPagamentos: true,
      visualizarRelatorios: true,
      gerenciarConfig: true
    },
    'GESTOR': {
      criarLancamentos: true,
      editarLancamentos: true,
      excluirLancamentos: false,
      aprovarPagamentos: true,
      visualizarRelatorios: true,
      gerenciarConfig: false
    },
    'OPERACIONAL': {
      criarLancamentos: true,
      editarLancamentos: true,
      excluirLancamentos: false,
      aprovarPagamentos: false,
      visualizarRelatorios: true,
      gerenciarConfig: false
    },
    'VISUALIZADOR': {
      criarLancamentos: false,
      editarLancamentos: false,
      excluirLancamentos: false,
      aprovarPagamentos: false,
      visualizarRelatorios: true,
      gerenciarConfig: false
    }
  };

  return permissoes[perfil] || permissoes['VISUALIZADOR'];
}
