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
  cacheRemove,
  cacheRemoveNamespace,
  cacheGetOrLoad,
  CacheNamespace,
  CacheScope,
} from '../shared/cache';
import { combineValidations, validateEnum, validateRequired } from '../shared/validation';
import {
  SHEET_TB_LANCAMENTOS,
  SHEET_TB_EXTRATOS,
  SHEET_TB_IMPORT_FC,
  SHEET_TB_IMPORT_ITAU,
  SHEET_TB_IMPORT_SIEG,
  SHEET_TB_CAIXAS,
  SHEET_TB_CAIXAS_MOV,
  SHEET_REF_CAIXA_TIPOS,
  SHEET_REF_FILIAIS,
  SHEET_REF_CANAIS,
  SHEET_REF_CCUSTO,
  SHEET_REF_PLANO_CONTAS,
  SHEET_CFG_CONFIG,
  TB_IMPORT_FC_COLS,
  TB_IMPORT_ITAU_COLS,
  TB_IMPORT_SIEG_COLS,
  TB_CAIXAS_COLS,
  TB_CAIXAS_MOV_COLS,
  REF_CAIXA_TIPOS_COLS,
} from '../config/sheet-mapping';

// ============================================================================
// VIEW RENDERING
// ============================================================================

const SHEET_USUARIOS = 'TB_Usuarios';

const WEBAPP_VIEW_ALLOWLIST = new Set([
  'dashboard',
  'contas-pagar',
  'contas-receber',
  'caixas',
  'conciliacao',
  'relatorios',
  'dre',
  'fluxo-caixa',
  'kpis',
  'ajuda',
  'configuracoes',
]);

function clearReportsCache(): void {
  cacheRemoveNamespace(CacheNamespace.DRE, CacheScope.SCRIPT);
  cacheRemoveNamespace(CacheNamespace.DFC, CacheScope.SCRIPT);
  cacheRemoveNamespace(CacheNamespace.KPI, CacheScope.SCRIPT);
  cacheRemoveNamespace(CacheNamespace.DASHBOARD, CacheScope.SCRIPT);
  cacheRemoveNamespace(CacheNamespace.CONCILIACAO, CacheScope.SCRIPT);
  invalidateLancamentosCache();
  invalidateExtratosCache();
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

function ensureRefFiliaisSchema(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const headers = [
    'Código',
    'Nome',
    'CNPJ',
    'Ativo',
    'Filial SIEG Relatorio',
    'Filial SIEG Contabilidade',
  ];
  const lastCol = sheet.getLastColumn();
  if (lastCol < headers.length) {
    sheet.insertColumnsAfter(Math.max(1, lastCol), headers.length - lastCol);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function ensureCfgConfigSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CFG_CONFIG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CFG_CONFIG);
    sheet.getRange('A1:E1').setValues([[
      'Chave', 'Valor', 'Tipo', 'Descricao', 'Ativo'
    ]]);
    sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#00a8e8').setFontColor('#ffffff');
  }
  return sheet;
}

function getConfigValue(key: string): string {
  const sheet = ensureCfgConfigSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '') === key) {
      return String(rows[i][1] || '').trim();
    }
  }
  return '';
}

function setConfigValue(key: string, value: string): void {
  const sheet = ensureCfgConfigSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '') === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value, 'TEXT', '', 'TRUE']);
}

function parseDriveFolderId(input: string): string {
  const raw = String(input || '').trim();
  if (!raw) return '';
  const match = raw.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  const idMatch = raw.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (idMatch) return idMatch[1];
  return raw;
}

function ensureCaixasSheets(): void {
  createSheetIfNotExists(SHEET_TB_CAIXAS, [
    'ID', 'Canal', 'Colaborador', 'Data Fechamento', 'Comunicado Interno',
    'Sistema Valor', 'Reforco', 'Criado Em', 'Atualizado Em',
    'Observacoes Entradas', 'Observacoes Saidas',
  ]);
  createSheetIfNotExists(SHEET_TB_CAIXAS_MOV, [
    'ID', 'Caixa ID', 'Tipo', 'Natureza', 'Valor', 'Data Mov', 'Arquivo URL',
    'Arquivo Nome', 'Criado Em', 'Atualizado Em', 'Observacoes',
  ]);
  createSheetIfNotExists(SHEET_REF_CAIXA_TIPOS, [
    'Tipo', 'Natureza', 'Requer Arquivo', 'Sistema FC', 'Conta Reforco', 'Ativo',
  ]);
  ensureCaixasSchema();
  ensureCaixasMovSchema();
  ensureCaixaTiposSchema();
}

function ensureCaixasSchema(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TB_CAIXAS);
  if (!sheet) return;
  const headers = [
    'ID', 'Canal', 'Colaborador', 'Data Fechamento', 'Comunicado Interno',
    'Sistema Valor', 'Reforco', 'Criado Em', 'Atualizado Em',
    'Observacoes Entradas', 'Observacoes Saidas',
  ];
  const lastCol = sheet.getLastColumn();
  if (lastCol < headers.length) {
    sheet.insertColumnsAfter(Math.max(1, lastCol), headers.length - lastCol);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function ensureCaixasMovSchema(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TB_CAIXAS_MOV);
  if (!sheet) return;
  const headers = [
    'ID', 'Caixa ID', 'Tipo', 'Natureza', 'Valor', 'Data Mov', 'Arquivo URL',
    'Arquivo Nome', 'Criado Em', 'Atualizado Em', 'Observacoes',
  ];
  const lastCol = sheet.getLastColumn();
  if (lastCol < headers.length) {
    sheet.insertColumnsAfter(Math.max(1, lastCol), headers.length - lastCol);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function ensureCaixaTiposSchema(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_REF_CAIXA_TIPOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_REF_CAIXA_TIPOS);
  }
  const headers = ['Tipo', 'Natureza', 'Requer Arquivo', 'Sistema FC', 'Conta Reforco', 'Ativo'];
  const lastCol = sheet.getLastColumn();
  if (lastCol < headers.length) {
    sheet.insertColumnsAfter(Math.max(1, lastCol), headers.length - lastCol);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const defaults: Array<[string, string, string, string, string, string]> = [
    ['Reforco do Caixa', 'ENTRADA', 'FALSE', 'FALSE', 'FALSE', 'TRUE'],
    ['Cartao de Credito', 'ENTRADA', 'TRUE', 'FALSE', 'FALSE', 'TRUE'],
    ['Depositos', 'ENTRADA', 'TRUE', 'FALSE', 'FALSE', 'TRUE'],
    ['Dinheiro', 'ENTRADA', 'FALSE', 'FALSE', 'FALSE', 'TRUE'],
    ['Link', 'ENTRADA', 'TRUE', 'FALSE', 'FALSE', 'TRUE'],
    ['Deposito Caixa', 'ENTRADA', 'FALSE', 'FALSE', 'FALSE', 'TRUE'],
    ['Outras Entradas', 'ENTRADA', 'FALSE', 'FALSE', 'FALSE', 'TRUE'],
    ['Outras Saidas', 'SAIDA', 'FALSE', 'FALSE', 'FALSE', 'TRUE'],
    ['Dinheiro Cofre', 'ENTRADA', 'FALSE', 'FALSE', 'TRUE', 'TRUE'],
    ['Dinheiro Caixa', 'ENTRADA', 'FALSE', 'FALSE', 'TRUE', 'TRUE'],
    ['Moedas', 'ENTRADA', 'FALSE', 'FALSE', 'TRUE', 'TRUE'],
    ['Sistema FC', 'ENTRADA', 'TRUE', 'TRUE', 'FALSE', 'TRUE'],
  ];
  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, defaults.length, headers.length).setValues(defaults);
    return;
  }

  const existingValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  const existing = new Set(existingValues.map((r) => String(r[0] || '').trim().toUpperCase()).filter(Boolean));
  const rowsToAppend = defaults.filter((row) => !existing.has(String(row[0]).trim().toUpperCase()));
  if (rowsToAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
  }
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

function normalizePerfil(raw: unknown): string {
  const value = String(raw ?? '').trim().toUpperCase();
  const allowed = ['ADMIN', 'GESTOR', 'OPERACIONAL', 'CAIXA', 'VISUALIZADOR', 'SEM_ACESSO'];
  return allowed.includes(value) ? value : 'VISUALIZADOR';
}

function safeParsePermissions(raw: unknown): Record<string, boolean> | null {
  if (!raw) return null;
  try {
    const parsed = JSON.parse(String(raw));
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch {
    return null;
  }
}

function normalizePermissoes(
  perfil: string,
  perms: Record<string, boolean> | null | undefined
): NonNullable<Usuario['permissoes']> {
  const defaults = getPermissoesPadrao(perfil);
  if (perms && typeof perms === 'object') {
    return { ...defaults, ...perms };
  }
  return defaults;
}

const USER_CACHE_TTL_SECONDS = 300;
const USER_SHEET_EMAIL_COL = 2; // Coluna B
const USER_SHEET_TOTAL_COLS = 9;

function normalizeEmail(value: string): string {
  return String(value || '').trim().toLowerCase();
}

function getUserCacheKey(email: string): string {
  return `user:${normalizeEmail(email)}`;
}

function readUserFromCache(email: string): Usuario | null {
  const key = getUserCacheKey(email);
  return cacheGet<Usuario>(CacheNamespace.USERS, key, CacheScope.SCRIPT);
}

function writeUserCache(email: string, user: Usuario): void {
  const key = getUserCacheKey(email);
  cacheSet(CacheNamespace.USERS, key, user, USER_CACHE_TTL_SECONDS, CacheScope.SCRIPT);
}

function invalidateUserCache(email: string): void {
  const key = getUserCacheKey(email);
  cacheRemove(CacheNamespace.USERS, key, CacheScope.SCRIPT);
}

function findUserRowByEmail(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  email: string
): number | null {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const target = normalizeEmail(email);
  if (!target) return null;
  const range = sheet.getRange(2, USER_SHEET_EMAIL_COL, lastRow - 1, 1);
  const found = range.createTextFinder(target).matchEntireCell(true).matchCase(false).findNext();
  return found ? found.getRow() : null;
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
  sheet.getRange('A1:I1').setValues([[
    'ID', 'Email', 'Nome', 'Perfil', 'Status', 'Ultimo Acesso', 'Permissoes', 'Data Criacao', 'Canal'
  ]]);
  sheet.getRange('A1:I1').setFontWeight('bold');
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
      '',
    ]);
  }

  return sheet;
}

function ensureUsuariosSchema(): void {
  const sheet = ensureUsuariosSheet();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || '').trim());
  if (!headers.includes('Canal')) {
    sheet.insertColumnAfter(lastCol);
    sheet.getRange(1, lastCol + 1).setValue('Canal');
  }
}

function getUsuarioByEmail(email: string): Usuario | null {
  if (!email) return null;
  const cached = readUserFromCache(email);
  if (cached) return cached;

  ensureUsuariosSchema();
  const sheet = ensureUsuariosSheet();
  const rowIndex = findUserRowByEmail(sheet, email);
  if (!rowIndex) return null;

  const row = sheet.getRange(rowIndex, 1, 1, USER_SHEET_TOTAL_COLS).getValues()[0];
  if (!row || !row[0]) return null;

  const perfil = normalizePerfil(row[3]);
  const permissoes = normalizePermissoes(perfil, safeParsePermissions(row[6]));
  const user: Usuario = {
    id: String(row[0]),
    email: String(row[1]),
    nome: String(row[2]),
    perfil: perfil as any,
    status: String(row[4]) as any,
    canal: row[8] ? String(row[8]) : undefined,
    ultimoAcesso: row[5] ? String(row[5]) : undefined,
    permissoes,
  };
  writeUserCache(email, user);
  return user;

  return null;
}

function updateUserLastAccess(email: string): void {
  if (!email) return;
  const sheet = ensureUsuariosSheet();
  const rowIndex = findUserRowByEmail(sheet, email);
  if (!rowIndex) return;
  sheet.getRange(rowIndex, 6).setValue(new Date().toISOString());
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

function requireAnyPermission<T extends { success: boolean; message: string }>(
  permissions: PermissionKey[],
  action: string
): T | null {
  const email = getRequestingUserEmail();
  const user = getUsuarioByEmail(email);

  if (!user) {
    appendAuditLog('permissionDenied', { permissions, action }, false, `Usuario nao cadastrado: ${action}`);
    return { success: false, message: `Usuario nao cadastrado: ${action}` } as T;
  }
  if (user.status !== 'ATIVO') {
    appendAuditLog('permissionDenied', { permissions, action, status: user.status }, false, `Usuario inativo: ${action}`);
    return { success: false, message: `Usuario inativo: ${action}` } as T;
  }

  const allowed = permissions.some((p) => Boolean(user.permissoes?.[p]));
  if (!allowed) {
    appendAuditLog('permissionDenied', { permissions, action, perfil: user.perfil }, false, `Sem permissao: ${action}`);
    return { success: false, message: `Sem permissao: ${action}` } as T;
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
    const views = [
      'dashboard',
      'contas-pagar',
      'contas-receber',
      'caixas',
      'conciliacao',
      'dre',
      'fluxo-caixa',
      'kpis',
      'ajuda',
      'configuracoes',
    ];
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

export function getAuditLogEntriesPage(params?: {
  page?: number;
  pageSize?: number;
}): {
  items: Array<{
    timestamp: string;
    email: string;
    action: string;
    success: boolean;
    message: string;
    payload: string;
  }>;
  total: number;
  page: number;
  pageSize: number;
} {
  const requester = getUsuarioByEmail(getRequestingUserEmail());
  if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
    return { items: [], total: 0, page: 1, pageSize: 100 };
  }

  const pageSize = Math.max(20, Math.min(200, Number(params?.pageSize) || 100));
  const page = Math.max(1, Number(params?.page) || 1);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_AUDIT_LOG);
  if (!sheet) return { items: [], total: 0, page, pageSize };

  const lastRow = sheet.getLastRow();
  const total = Math.max(0, lastRow - 1);
  if (total === 0) return { items: [], total, page, pageSize };

  const offset = (page - 1) * pageSize;
  const endRow = lastRow - offset;
  if (endRow < 2) return { items: [], total, page, pageSize };

  const startRow = Math.max(2, endRow - pageSize + 1);
  const numRows = endRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, 6).getDisplayValues();

  const items = values
    .map((r) => ({
      timestamp: String(r[0] || ''),
      email: String(r[1] || ''),
      action: String(r[2] || ''),
      success: String(r[3] || '').toUpperCase() === 'TRUE',
      message: String(r[4] || ''),
      payload: String(r[5] || ''),
    }))
    .reverse();

  return { items, total, page, pageSize };
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

export function getAuditLogEntriesFilteredPage(filters: {
  page?: number;
  pageSize?: number;
  action?: string;
  email?: string;
  success?: boolean | null;
  correlationId?: string;
  query?: string;
}): {
  items: Array<{
    timestamp: string;
    email: string;
    action: string;
    success: boolean;
    message: string;
    payload: string;
  }>;
  total: number;
  page: number;
  pageSize: number;
} {
  const requester = getUsuarioByEmail(getRequestingUserEmail());
  if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
    return { items: [], total: 0, page: 1, pageSize: 100 };
  }

  const pageSize = Math.max(20, Math.min(200, Number(filters?.pageSize) || 100));
  const page = Math.max(1, Number(filters?.page) || 1);
  const offset = (page - 1) * pageSize;

  const actionFilter = String(filters?.action || '').trim().toLowerCase();
  const emailFilter = String(filters?.email || '').trim().toLowerCase();
  const correlationId = String(filters?.correlationId || '').trim();
  const query = String(filters?.query || '').trim().toLowerCase();
  const successFilter =
    typeof filters?.success === 'boolean' ? filters.success : null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_AUDIT_LOG);
  if (!sheet) return { items: [], total: 0, page, pageSize };

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { items: [], total: 0, page, pageSize };

  const scanWindow = Math.min(5000, Math.max(1000, pageSize * 20));
  const startRow = Math.max(2, lastRow - scanWindow + 1);
  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, 6).getDisplayValues();

  const items: Array<{
    timestamp: string;
    email: string;
    action: string;
    success: boolean;
    message: string;
    payload: string;
  }> = [];

  let total = 0;
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

    if (total >= offset && items.length < pageSize) {
      items.push({ timestamp, email, action, success, message, payload });
    }
    total += 1;
  }

  return { items, total, page, pageSize };
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
    cacheRemoveNamespace(CacheNamespace.CONCILIACAO, CacheScope.SCRIPT);
    appendAuditLog('clearCaches', {}, true);
    return { success: true, message: 'Caches limpos' };
  } catch (e: any) {
    appendAuditLog('clearCaches', {}, false, e?.message);
    return { success: false, message: e?.message || String(e) };
  }
}

// ============================================================================
// CONFIGURACAO CAIXAS
// ============================================================================

export function getCaixasConfig(): { pastaId: string; pastaUrl?: string; pastaNome?: string } {
  const user = getUsuarioByEmail(getRequestingUserEmail());
  if (!user || user.status !== 'ATIVO') return { pastaId: '' };
  const pastaId = getConfigValue('CAIXAS_PASTA_ID');
  if (!pastaId) return { pastaId: '' };
  try {
    const folder = DriveApp.getFolderById(pastaId);
    return { pastaId, pastaUrl: folder.getUrl(), pastaNome: folder.getName() };
  } catch {
    return { pastaId };
  }
}

export function salvarCaixasConfig(pastaIdOrUrl: string): { success: boolean; message: string } {
  const denied = requirePermission('gerenciarConfig', 'salvar config caixas');
  if (denied) return denied;
  const pastaId = parseDriveFolderId(pastaIdOrUrl);
  if (!pastaId) {
    setConfigValue('CAIXAS_PASTA_ID', '');
    return { success: true, message: 'Pasta de caixas removida' };
  }
  let valid = true;
  try {
    DriveApp.getFolderById(pastaId);
  } catch {
    valid = false;
  }
  setConfigValue('CAIXAS_PASTA_ID', pastaId);
  appendAuditLog('caixas:config', { pastaId }, true);
  return { success: true, message: valid ? 'Pasta de caixas atualizada' : 'Pasta salva, mas nao foi possivel validar acesso' };
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

    if (normalized === 'caixas') {
      const denied = requireAnyPermission<{ success: boolean; message: string }>(
        ['visualizarRelatorios', 'importarArquivos'],
        'acessar caixas'
      );
      if (denied) {
        return `<div class="empty-state"><div class="empty-state-message">${escapeHtml(denied.message)}</div></div>`;
      }
    }

    if (normalized === 'conciliacao') {
      const user = getUsuarioByEmail(getRequestingUserEmail());
      if (user && user.perfil === 'CAIXA') {
        return `<div class="empty-state"><div class="empty-state-message">Sem permissao para acessar conciliacao</div></div>`;
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
  canal?: string;
  permissoes: NonNullable<Usuario['permissoes']>;
} {
  const email =
    Session.getActiveUser().getEmail() ||
    Session.getEffectiveUser().getEmail() ||
    'usuario@empresa.com';

  const fallbackNome = email.split('@')[0] || email;
  const user = getUsuarioByEmail(email);
  if (!user || user.status !== 'ATIVO') {
    const perfil = 'SEM_ACESSO';
    const permissoes = normalizePermissoes(perfil, null);
    return { email, nome: fallbackNome, perfil, permissoes };
  }

  const perfil = normalizePerfil(user.perfil);
  const permissoes = normalizePermissoes(perfil, user.permissoes);
  updateUserLastAccess(email);
  return { email, nome: user.nome || fallbackNome, perfil, canal: user.canal || '', permissoes };
}

// ============================================================================
// REFERENCE DATA
// ============================================================================

export function getReferenceData(): {
  filiais: Array<{ codigo: string; nome: string; cnpj?: string; ativo?: boolean; siegRelatorio?: string; siegContabilidade?: string }>;
  caixaTipos: Array<{ tipo: string; natureza: string; requerArquivo?: boolean; sistemaFc?: boolean; contaReforco?: boolean; ativo?: boolean }>;
  canais: Array<{ codigo: string; nome: string; ativo?: boolean }>;
  contas: Array<{ codigo: string; nome: string; tipo?: string; grupoDRE?: string; subgrupoDRE?: string; grupoDFC?: string; variavelFixa?: string; cmaCmv?: string }>;
  centrosCusto: Array<{ codigo: string; nome: string; ativo?: boolean }>;
} {
  const user = getUsuarioByEmail(getRequestingUserEmail());
  if (!user || user.status !== 'ATIVO') {
    return { filiais: [], caixaTipos: [], canais: [], contas: [], centrosCusto: [] };
  }

  const cached = cacheGet<any>(CacheNamespace.REFERENCE, 'all');
  if (cached) {
    return cached;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Filiais (da planilha)
    const sheetFiliais = ss.getSheetByName(SHEET_REF_FILIAIS);
    if (sheetFiliais) ensureRefFiliaisSchema(sheetFiliais);
    const filiais = sheetFiliais ? sheetFiliais.getDataRange().getValues().slice(1) : [];

    // Canais (da planilha)
    const sheetCanais = ss.getSheetByName(SHEET_REF_CANAIS);
    const canais = sheetCanais ? sheetCanais.getDataRange().getValues().slice(1) : [];

    // Tipos de movimentacao de caixa
    ensureCaixaTiposSchema();
    const sheetCaixaTipos = ss.getSheetByName(SHEET_REF_CAIXA_TIPOS);
    const caixaTipos = sheetCaixaTipos ? sheetCaixaTipos.getDataRange().getValues().slice(1) : [];

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
          cnpj: String(f[2] || ''),
          ativo: f[ativoIdx] !== false && String(f[ativoIdx] ?? 'TRUE').toUpperCase() !== 'FALSE',
          siegRelatorio: String(f[4] || ''),
          siegContabilidade: String(f[5] || ''),
        };
      }),
      caixaTipos: caixaTipos.filter((t: any) => t[0]).map((t: any) => ({
        tipo: String(t[REF_CAIXA_TIPOS_COLS.TIPO]),
        natureza: String(t[REF_CAIXA_TIPOS_COLS.NATUREZA] || 'ENTRADA'),
        requerArquivo: String(t[REF_CAIXA_TIPOS_COLS.REQUER_ARQUIVO] ?? 'FALSE').toUpperCase() === 'TRUE',
        sistemaFc: String(t[REF_CAIXA_TIPOS_COLS.SISTEMA_FC] ?? 'FALSE').toUpperCase() === 'TRUE',
        contaReforco: String(t[REF_CAIXA_TIPOS_COLS.CONTA_REFORCO] ?? 'FALSE').toUpperCase() === 'TRUE',
        ativo: String(t[REF_CAIXA_TIPOS_COLS.ATIVO] ?? 'TRUE').toUpperCase() !== 'FALSE',
      })),
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
      caixaTipos: [],
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

// Tipos de movimentacao do caixa
export function salvarCaixaTipo(tipo: {
  tipo: string;
  natureza: string;
  requerArquivo?: boolean;
  sistemaFc?: boolean;
  contaReforco?: boolean;
  ativo?: boolean;
}, editIndex: number): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'salvar tipo de caixa');
    if (denied) return denied;

    const validation = combineValidations(
      validateRequired(tipo?.tipo, 'Tipo'),
      validateEnum(String(tipo?.natureza || ''), ['ENTRADA', 'SAIDA'], 'Natureza')
    );
    if (!validation.valid) return { success: false, message: validation.errors.join('; ') };

    ensureCaixaTiposSchema();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_CAIXA_TIPOS);
    if (!sheet) return { success: false, message: 'Aba de tipos de caixa nao encontrada' };

    if (editIndex < 0) {
      const data = sheet.getDataRange().getValues().slice(1);
      const existe = data.some((row: any) => String(row[0]).toUpperCase() === String(tipo.tipo).toUpperCase());
      if (existe) return { success: false, message: 'Tipo ja existe' };
    }

    const rowData = [
      sanitizeSheetString(tipo.tipo),
      String(tipo.natureza || 'ENTRADA').toUpperCase(),
      tipo.requerArquivo ? 'TRUE' : 'FALSE',
      tipo.sistemaFc ? 'TRUE' : 'FALSE',
      tipo.contaReforco ? 'TRUE' : 'FALSE',
      tipo.ativo === false ? 'FALSE' : 'TRUE',
    ];

    if (editIndex >= 0) {
      sheet.getRange(editIndex + 2, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('salvarCaixaTipo', { tipo: rowData, editIndex }, true);
    return { success: true, message: 'Tipo salvo com sucesso' };
  } catch (error: any) {
    appendAuditLog('salvarCaixaTipo', { tipo, editIndex }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function excluirCaixaTipo(index: number): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'excluir tipo de caixa');
    if (denied) return denied;
    if (!Number.isFinite(index) || index < 0) return { success: false, message: 'Indice invalido' };

    ensureCaixaTiposSchema();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_CAIXA_TIPOS);
    if (!sheet) return { success: false, message: 'Aba de tipos de caixa nao encontrada' };
    if (index + 2 > sheet.getLastRow()) return { success: false, message: 'Indice invalido' };

    sheet.deleteRow(index + 2);
    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('excluirCaixaTipo', { index }, true);
    return { success: true, message: 'Tipo excluido' };
  } catch (error: any) {
    appendAuditLog('excluirCaixaTipo', { index }, false, error?.message);
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function toggleCaixaTipo(index: number, ativo: boolean): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'ativar/inativar tipo de caixa');
    if (denied) return denied;
    if (!Number.isFinite(index) || index < 0) return { success: false, message: 'Indice invalido' };

    ensureCaixaTiposSchema();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_CAIXA_TIPOS);
    if (!sheet) return { success: false, message: 'Aba de tipos de caixa nao encontrada' };

    if (index + 2 > sheet.getLastRow()) return { success: false, message: 'Indice invalido' };
    sheet.getRange(index + 2, REF_CAIXA_TIPOS_COLS.ATIVO + 1).setValue(ativo ? 'TRUE' : 'FALSE');

    cacheRemoveNamespace(CacheNamespace.REFERENCE);
    clearReportsCache();
    appendAuditLog('toggleCaixaTipo', { index, ativo }, true);
    return { success: true, message: ativo ? 'Tipo ativado' : 'Tipo inativado' };
  } catch (error: any) {
    appendAuditLog('toggleCaixaTipo', { index, ativo }, false, error?.message);
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
export function salvarFilial(
  filial: { codigo: string; nome: string; cnpj?: string; siegRelatorio?: string; siegContabilidade?: string; ativo?: boolean },
  editIndex: number
): { success: boolean; message: string } {
  try {
    const denied = requirePermission('gerenciarConfig', 'salvar filial');
    if (denied) return denied;

    const validation = combineValidations(
      validateRequired(filial?.codigo, 'C?digo'),
      validateRequired(filial?.nome, 'Nome')
    );
    if (!validation.valid) return { success: false, message: validation.errors.join('; ') };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_REF_FILIAIS);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_REF_FILIAIS);
      sheet.getRange('A1:F1').setValues([['C?digo', 'Nome', 'CNPJ', 'Ativo', 'Filial SIEG Relatorio', 'Filial SIEG Contabilidade']]);
      sheet.getRange('A1:F1').setFontWeight('bold').setBackground('#00a8e8').setFontColor('#ffffff');
    } else if (sheet.getLastColumn() < 6) {
      sheet.getRange('A1:F1').setValues([['C?digo', 'Nome', 'CNPJ', 'Ativo', 'Filial SIEG Relatorio', 'Filial SIEG Contabilidade']]);
      sheet.getRange('A1:F1').setFontWeight('bold').setBackground('#00a8e8').setFontColor('#ffffff');
    }

    // Verificar c?digo duplicado
    if (editIndex < 0) {
      const data = sheet.getDataRange().getValues().slice(1);
      const codigoExiste = data.some((row: any) => String(row[0]).toUpperCase() === filial.codigo.toUpperCase());
      if (codigoExiste) {
        return { success: false, message: 'C?digo j? existe' };
      }
    }

    const ativo = filial.ativo !== false && String(filial.ativo ?? 'TRUE').toUpperCase() !== 'FALSE';
    const codigo = sanitizeSheetString(filial.codigo).toUpperCase();
    const nome = sanitizeSheetString(filial.nome);
    const lastCol = Math.max(6, sheet.getLastColumn());
    const existing =
      editIndex >= 0 && sheet.getLastRow() >= editIndex + 2
        ? sheet.getRange(editIndex + 2, 1, 1, lastCol).getValues()[0]
        : [];

    const cnpj =
      typeof filial.cnpj === 'string' ? sanitizeSheetString(filial.cnpj) : (existing[2] || '');
    const siegRelatorio =
      typeof filial.siegRelatorio === 'string' ? sanitizeSheetString(filial.siegRelatorio) : (existing[4] || '');
    const siegContabilidade =
      typeof filial.siegContabilidade === 'string' ? sanitizeSheetString(filial.siegContabilidade) : (existing[5] || '');

    const rowData: any[] = [
      codigo,
      nome,
      cnpj || '',
      ativo,
      siegRelatorio || '',
      siegContabilidade || '',
    ];

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

export function getDashboardData(mes?: number, ano?: number, filial?: string, canal?: string, includeKpis?: boolean) {
  enforcePermission('visualizarRelatorios', 'carregar dashboard');
  const hoje = new Date();
  const targetMes = Number(mes) || hoje.getMonth() + 1;
  const targetAno = Number(ano) || hoje.getFullYear();
  const inicioMes = new Date(targetAno, targetMes - 1, 1);
  const fimMes = new Date(targetAno, targetMes, 1);
  const isCurrent = targetMes === hoje.getMonth() + 1 && targetAno === hoje.getFullYear();
  const referencia = isCurrent ? hoje : new Date(targetAno, targetMes, 0);
  const include = includeKpis !== false;
  const cacheKey = `dash:${targetAno}-${targetMes}:${filial || 'all'}:${canal || 'all'}:${include ? 'full' : 'lite'}`;

  return cacheGetOrLoad(CacheNamespace.DASHBOARD, cacheKey, () => {
    const lancamentosBase = getLancamentosFromSheet();
    const lancamentos = lancamentosBase.filter(l => {
      const matchFilial = !filial || l.filial === filial;
      const matchCanal = !canal || l.canal === canal;
      return matchFilial && matchCanal;
    });

  const toDate = (value: any): Date | null => {
    const norm = normalizeDateInput(value);
    if (!norm) return null;
    const d = new Date(norm);
    return isNaN(d.getTime()) ? null : d;
  };

  const inRange = (value: any, start: Date, end: Date): boolean => {
    const d = toDate(value);
    if (!d) return false;
    return d >= start && d < end;
  };

  const buildSnapshot = (ref: Date) => {
    const next7 = new Date(ref.getTime());
    next7.setDate(next7.getDate() + 7);
    const next30 = new Date(ref.getTime());
    next30.setDate(next30.getDate() + 30);
    const pagarVencidasSnap = lancamentos.filter(l =>
      l.tipo === 'DESPESA' &&
      (
        String(l.status || '').toUpperCase() === 'VENCIDA' ||
        (String(l.status || '').toUpperCase() === 'PENDENTE' && new Date(l.dataVencimento) < ref)
      )
    );
    const pagarProximasSnap = lancamentos.filter(l =>
      l.tipo === 'DESPESA' &&
      l.status === 'PENDENTE' &&
      new Date(l.dataVencimento) <= next7 &&
      new Date(l.dataVencimento) >= ref
    );
    const receberHojeSnap = lancamentos.filter(l =>
      l.tipo === 'RECEITA' &&
      l.status === 'PENDENTE' &&
      new Date(l.dataVencimento).toDateString() === ref.toDateString()
    );
    const pagarProximas30Snap = lancamentos.filter(l =>
      l.tipo === 'DESPESA' &&
      l.status === 'PENDENTE' &&
      new Date(l.dataVencimento) <= next30 &&
      new Date(l.dataVencimento) >= ref
    );
    const receberProximasSnap = lancamentos.filter(l =>
      l.tipo === 'RECEITA' &&
      l.status === 'PENDENTE' &&
      new Date(l.dataVencimento) <= next7 &&
      new Date(l.dataVencimento) >= ref
    );
    const receberAtrasadasSnap = lancamentos.filter(l =>
      l.tipo === 'RECEITA' &&
      l.status === 'PENDENTE' &&
      new Date(l.dataVencimento) < ref
    );
    const extratosPendentesSnap = extratos.filter(e => e.statusConciliacao === 'PENDENTE');
    return {
      pagarVencidas: {
        quantidade: pagarVencidasSnap.length,
        valor: sumValues(pagarVencidasSnap),
      },
      pagarProximas: {
        quantidade: pagarProximasSnap.length,
        valor: sumValues(pagarProximasSnap),
      },
      receberHoje: {
        quantidade: receberHojeSnap.length,
        valor: sumValues(receberHojeSnap),
      },
      pagarProximas30: {
        quantidade: pagarProximas30Snap.length,
        valor: sumValues(pagarProximas30Snap),
      },
      receberProximas: {
        quantidade: receberProximasSnap.length,
        valor: sumValues(receberProximasSnap),
      },
      receberAtrasadas: {
        quantidade: receberAtrasadasSnap.length,
        valor: sumValues(receberAtrasadasSnap),
      },
      conciliacaoPendentes: {
        quantidade: extratosPendentesSnap.length,
        valor: extratosPendentesSnap.reduce((sum, e) => sum + parseFloat(String(e.valor || 0)), 0),
      }
    };
  };

  // Contas a pagar vencidas
  const pagarVencidas = lancamentos.filter(l =>
    l.tipo === 'DESPESA' &&
    (
      String(l.status || '').toUpperCase() === 'VENCIDA' ||
      (String(l.status || '').toUpperCase() === 'PENDENTE' && new Date(l.dataVencimento) < referencia)
    )
  );

  // Contas a pagar proximos 7 dias
  const proximos7Dias = new Date(referencia.getTime());
  proximos7Dias.setDate(proximos7Dias.getDate() + 7);
  const pagarProximas = lancamentos.filter(l =>
    l.tipo === 'DESPESA' &&
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos7Dias &&
    new Date(l.dataVencimento) >= referencia
  );

  const proximos30Dias = new Date(referencia.getTime());
  proximos30Dias.setDate(proximos30Dias.getDate() + 30);
  const pagarProximas30 = lancamentos.filter(l =>
    l.tipo === 'DESPESA' &&
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos30Dias &&
    new Date(l.dataVencimento) >= referencia
  );

  // Contas a receber hoje
  const receberHoje = lancamentos.filter(l =>
    l.tipo === 'RECEITA' &&
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento).toDateString() === referencia.toDateString()
  );

  const receberProximas = lancamentos.filter(l =>
    l.tipo === 'RECEITA' &&
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos7Dias &&
    new Date(l.dataVencimento) >= referencia
  );

  const receberAtrasadas = lancamentos.filter(l =>
    l.tipo === 'RECEITA' &&
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) < referencia
  );

  const receitaMes = lancamentos.filter(l =>
    l.tipo === 'RECEITA' && inRange(l.dataCompetencia || l.dataVencimento, inicioMes, fimMes)
  );
  const despesaMes = lancamentos.filter(l =>
    l.tipo === 'DESPESA' && inRange(l.dataCompetencia || l.dataVencimento, inicioMes, fimMes)
  );

  // Extratos pendentes
  const kpisMes = include ? getKPIsMensal(targetMes, targetAno, filial, canal) : null;

  const extratos = getExtratosFromSheet();
  const extratosPendentes = extratos.filter(e => e.statusConciliacao === 'PENDENTE');
  const extratosConciliados = extratos.filter(e => (e.statusConciliacao || '').toUpperCase() === 'CONCILIADO');
  const conciliacaoTaxa = extratos.length > 0 ? Math.round((extratosConciliados.length / extratos.length) * 100) : 0;
  const referenciaLabel = Utilities.formatDate(referencia, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const prevRef = new Date(targetAno, targetMes - 2, 1);
  const prevEnd = new Date(prevRef.getFullYear(), prevRef.getMonth() + 1, 0);
  const prevSnapshot = buildSnapshot(prevEnd);

  // Ultimos lancamentos (mantido para compatibilidade)
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
  const fluxoMesValor = sumValues(receitaMes) - sumValues(despesaMes);
  const liquidez = include ? Number(kpisMes?.liquidez?.liquidezCorrente?.valor || 0) : 0;
  const margem = include ? Number(kpisMes?.rentabilidade?.margemLiquida?.valor || 0) : 0;
  const inad = include ? Number(kpisMes?.crescimento?.inadimplencia?.valor || 0) : 0;

  if (pagarVencidas.length > 0) {
    alerts.push({
      type: 'danger',
      title: 'Contas Vencidas',
      message: `Voce tem ${pagarVencidas.length} contas a pagar vencidas no valor de ${formatCurrency(sumValues(pagarVencidas))}`,
    });
  }

  if (pagarProximas.length > 0) {
    alerts.push({
      type: 'warning',
      title: 'Vencimentos Proximos',
      message: `${pagarProximas.length} contas a pagar vencem nos proximos 7 dias`,
    });
  }

  if (extratosPendentes.length > 5 || conciliacaoTaxa < 80) {
    alerts.push({
      type: 'warning',
      title: 'Conciliacao Pendente',
      message: `Conciliacao em ${conciliacaoTaxa}% com ${extratosPendentes.length} extratos pendentes`,
    });
  }

  if (fluxoMesValor < 0) {
    alerts.push({
      type: 'warning',
      title: 'Fluxo Mensal Negativo',
      message: `Fluxo do periodo esta negativo em ${formatCurrency(Math.abs(fluxoMesValor))}`,
    });
  }

  if (include && liquidez > 0 && liquidez < 1) {
    alerts.push({
      type: 'warning',
      title: 'Liquidez Abaixo do Ideal',
      message: `Liquidez corrente em ${liquidez.toFixed(2)} indica risco de curto prazo`,
    });
  }

  if (include && margem < 0) {
    alerts.push({
      type: 'info',
      title: 'Margem Negativa',
      message: 'Margem liquida negativa no periodo analisado',
    });
  }

  if (include && inad > 5) {
    alerts.push({
      type: 'info',
      title: 'Inadimplencia Elevada',
      message: `Inadimplencia em ${inad.toFixed(1)}% no periodo`,
    });
  }

  if (receberAtrasadas.length > 0) {
    alerts.push({
      type: 'info',
      title: 'Recebimentos em Atraso',
      message: `${receberAtrasadas.length} recebimentos pendentes em atraso`,
    });
  }

    return {
      periodo: {
        mes: targetMes,
        ano: targetAno,
        mesNome: getMesNome(targetMes),
        filial: filial || 'Consolidado',
        canal: canal || 'Todos',
        referencia: referenciaLabel,
      },
    pagarVencidas: {
      quantidade: pagarVencidas.length,
      valor: sumValues(pagarVencidas),
    },
    pagarProximas: {
      quantidade: pagarProximas.length,
      valor: sumValues(pagarProximas),
    },
    pagarProximas30: {
      quantidade: pagarProximas30.length,
      valor: sumValues(pagarProximas30),
    },
    receberHoje: {
      quantidade: receberHoje.length,
      valor: sumValues(receberHoje),
    },
    receberProximas: {
      quantidade: receberProximas.length,
      valor: sumValues(receberProximas),
    },
    receberAtrasadas: {
      quantidade: receberAtrasadas.length,
      valor: sumValues(receberAtrasadas),
    },
    receitaMes: {
      quantidade: receitaMes.length,
      valor: sumValues(receitaMes),
      ticketMedio: receitaMes.length ? sumValues(receitaMes) / receitaMes.length : 0,
    },
    despesaMes: {
      quantidade: despesaMes.length,
      valor: sumValues(despesaMes),
      ticketMedio: despesaMes.length ? sumValues(despesaMes) / despesaMes.length : 0,
    },
    fluxoMes: {
      valor: fluxoMesValor,
    },
    conciliacaoPendentes: {
      quantidade: extratosPendentes.length,
      valor: extratosPendentes.reduce((sum, e) => sum + parseFloat(String(e.valor || 0)), 0),
    },
    conciliacaoTaxa,
    kpisMes,
    recentTransactions,
    prev: prevSnapshot,
    alerts,
  };
  }, 120, CacheScope.SCRIPT);
}


export function getLancamentoDetalhes(id: string): any {
  enforcePermission('visualizarRelatorios', 'ver lanÇõamento');
  const validation = validateRequired(id, 'ID');
  if (!validation.valid) {
    throw new Error(validation.errors.join('; '));
  }

  const lancamentos = getLancamentosFromSheet();
  const lancamento = lancamentos.find(l => String(l.id) === String(id));
  if (!lancamento) {
    throw new Error('LanÇõamento nÇœo encontrado');
  }
  return lancamento;
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

export function cancelarContasEmLote(ids: string[]): { success: boolean; message: string } {
  try {
    const denied = requirePermission('aprovarPagamentos', 'cancelar contas a pagar em lote');
    if (denied) return denied;
    if (!Array.isArray(ids) || ids.length === 0) return { success: false, message: 'IDs invÇ­lidos' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lanÇõamentos nÇœo encontrada');

    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);
    try {
      const headers = getHeaderIndexMap(sheet);
      const idCol = headers['ID'];
      const statusCol = headers['Status'];
      const tipoCol = headers['Tipo'];
      const dataPagCol = headers['Data Pagamento'];
      if (idCol === undefined || statusCol === undefined || tipoCol === undefined || dataPagCol === undefined) {
        throw new Error('CabeÇõalhos obrigatÇürios nÇœo encontrados (ID, Status, Tipo, Data Pagamento)');
      }

      const lastRow = sheet.getLastRow();
      const numRows = lastRow > 1 ? lastRow - 1 : 0;
      const idsColumnValues =
        numRows > 0 ? sheet.getRange(2, idCol + 1, numRows, 1).getDisplayValues() : [];
      const statusValues =
        numRows > 0 ? sheet.getRange(2, statusCol + 1, numRows, 1).getDisplayValues() : [];
      const tipoValues =
        numRows > 0 ? sheet.getRange(2, tipoCol + 1, numRows, 1).getDisplayValues() : [];

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
      const allowed = new Set(['PENDENTE', 'VENCIDA']);

      for (const rawId of ids) {
        const wanted = String(rawId || '').trim();
        if (!wanted) continue;
        const row = idToRow.get(wanted);
        if (!row) {
          errors.push(`${wanted}: nÇœo encontrada`);
          continue;
        }

        const status = String(statusValues[row - 2]?.[0] || '').toUpperCase();
        const tipo = String(tipoValues[row - 2]?.[0] || '').toUpperCase();
        if (tipo !== 'DESPESA') {
          errors.push(`${wanted}: tipo ${tipo || 'N/A'}`);
          continue;
        }
        if (!allowed.has(status)) {
          errors.push(`${wanted}: status ${status || 'N/A'}`);
          continue;
        }

        statusRanges.push(`${columnToLetter(statusCol + 1)}${row}`);
        dateRanges.push(`${columnToLetter(dataPagCol + 1)}${row}`);
        appendAuditLog('cancelarConta', { id: wanted }, true);
        count++;
      }

      if (statusRanges.length > 0) {
        sheet.getRangeList(statusRanges).setValue('CANCELADA');
        sheet.getRangeList(dateRanges).setValue('');
        clearReportsCache();
      }

      if (count === 0) {
        appendAuditLog('cancelarContasEmLote', { ids, count, errorsCount: errors.length }, false, 'Nenhuma cancelada');
        return { success: false, message: errors.length ? errors[0] : 'Nenhuma conta cancelada' };
      }

      if (errors.length > 0) {
        appendAuditLog('cancelarContasEmLote', { ids, count, errorsCount: errors.length }, true, 'Parcial');
        return { success: true, message: `${count} contas canceladas; ${errors.length} falharam` };
      }

      appendAuditLog('cancelarContasEmLote', { ids, count }, true);
      return { success: true, message: `${count} contas canceladas com sucesso` };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (error: any) {
    appendAuditLog('cancelarContasEmLote', { ids }, false, error?.message);
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

export function cancelarContasReceberEmLote(ids: string[]): { success: boolean; message: string } {
  try {
    const denied = requirePermission('aprovarPagamentos', 'cancelar contas a receber em lote');
    if (denied) return denied;
    if (!Array.isArray(ids) || ids.length === 0) return { success: false, message: 'IDs invÇ­lidos' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lanÇõamentos nÇœo encontrada');

    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);
    try {
      const headers = getHeaderIndexMap(sheet);
      const idCol = headers['ID'];
      const statusCol = headers['Status'];
      const tipoCol = headers['Tipo'];
      const dataPagCol = headers['Data Pagamento'];
      if (idCol === undefined || statusCol === undefined || tipoCol === undefined || dataPagCol === undefined) {
        throw new Error('CabeÇõalhos obrigatÇürios nÇœo encontrados (ID, Status, Tipo, Data Pagamento)');
      }

      const lastRow = sheet.getLastRow();
      const numRows = lastRow > 1 ? lastRow - 1 : 0;
      const idsColumnValues =
        numRows > 0 ? sheet.getRange(2, idCol + 1, numRows, 1).getDisplayValues() : [];
      const statusValues =
        numRows > 0 ? sheet.getRange(2, statusCol + 1, numRows, 1).getDisplayValues() : [];
      const tipoValues =
        numRows > 0 ? sheet.getRange(2, tipoCol + 1, numRows, 1).getDisplayValues() : [];

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
      const allowed = new Set(['PENDENTE', 'VENCIDA']);

      for (const rawId of ids) {
        const wanted = String(rawId || '').trim();
        if (!wanted) continue;
        const row = idToRow.get(wanted);
        if (!row) {
          errors.push(`${wanted}: nÇœo encontrada`);
          continue;
        }

        const status = String(statusValues[row - 2]?.[0] || '').toUpperCase();
        const tipo = String(tipoValues[row - 2]?.[0] || '').toUpperCase();
        if (tipo !== 'RECEITA') {
          errors.push(`${wanted}: tipo ${tipo || 'N/A'}`);
          continue;
        }
        if (!allowed.has(status)) {
          errors.push(`${wanted}: status ${status || 'N/A'}`);
          continue;
        }

        statusRanges.push(`${columnToLetter(statusCol + 1)}${row}`);
        dateRanges.push(`${columnToLetter(dataPagCol + 1)}${row}`);
        appendAuditLog('cancelarContaReceber', { id: wanted }, true);
        count++;
      }

      if (statusRanges.length > 0) {
        sheet.getRangeList(statusRanges).setValue('CANCELADA');
        sheet.getRangeList(dateRanges).setValue('');
        clearReportsCache();
      }

      if (count === 0) {
        appendAuditLog('cancelarContasReceberEmLote', { ids, count, errorsCount: errors.length }, false, 'Nenhuma cancelada');
        return { success: false, message: errors.length ? errors[0] : 'Nenhuma conta cancelada' };
      }

      if (errors.length > 0) {
        appendAuditLog('cancelarContasReceberEmLote', { ids, count, errorsCount: errors.length }, true, 'Parcial');
        return { success: true, message: `${count} contas canceladas; ${errors.length} falharam` };
      }

      appendAuditLog('cancelarContasReceberEmLote', { ids, count }, true);
      return { success: true, message: `${count} contas canceladas com sucesso` };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (error: any) {
    appendAuditLog('cancelarContasReceberEmLote', { ids }, false, error?.message);
    return { success: false, message: error.message };
  }
}

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

    const dataCompetencia = String(lancamento.dataCompetencia || '').trim();
    const dataVencimento = String(lancamento.dataVencimento || '').trim();
    const dataPagamento = String(lancamento.dataPagamento || '').trim();
    const status = String(lancamento.status || '').trim().toUpperCase();

    const isValidDate = (value: string) => Number.isFinite(Date.parse(value));
    if (!isValidDate(dataCompetencia)) {
      return { success: false, message: 'Data competência inválida' };
    }
    if (!isValidDate(dataVencimento)) {
      return { success: false, message: 'Data vencimento inválida' };
    }
    if (dataPagamento && !isValidDate(dataPagamento)) {
      return { success: false, message: 'Data pagamento inválida' };
    }
    if (new Date(dataCompetencia).getTime() > new Date(dataVencimento).getTime()) {
      return { success: false, message: 'Data competência não pode ser maior que data vencimento' };
    }
    if ((status === 'PAGA' || status === 'RECEBIDA') && !dataPagamento) {
      return { success: false, message: 'Informe data pagamento para status pago/recebido' };
    }
    if (dataPagamento && (status === 'PENDENTE' || status === 'VENCIDA' || status === 'CANCELADA')) {
      return { success: false, message: 'Data pagamento não é permitida para status pendente/vencida/cancelada' };
    }

    const row = [
      sanitizeSheetString(lancamento.id),                        // ID
      sanitizeSheetString(dataCompetencia),                      // Data Competência
      sanitizeSheetString(dataVencimento),                       // Data Vencimento
      sanitizeSheetString(dataPagamento),                        // Data Pagamento
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
      sanitizeSheetString((lancamento as any).numeroDocumento || ''), // N Documento
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

export function atualizarLancamento(lancamento: any): { success: boolean; message: string } {
  try {
    const denied = requirePermission('editarLancamentos', 'atualizar lanÇõamento');
    if (denied) return denied;

    const v = combineValidations(
      validateRequired(lancamento?.id, 'ID'),
      validateRequired(lancamento?.dataCompetencia, 'Data competÇ¦ncia'),
      validateRequired(lancamento?.dataVencimento, 'Data vencimento'),
      validateEnum(String(lancamento?.tipo || ''), ['RECEITA', 'DESPESA'], 'Tipo'),
      validateRequired(lancamento?.filial, 'Filial'),
      validateRequired(lancamento?.contaContabil, 'Conta contÇ­bil'),
      validateRequired(lancamento?.descricao, 'DescriÇõÇœo')
    );
    if (!v.valid) {
      return { success: false, message: v.errors.join('; ') };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lanÇõamentos nÇœo encontrada');

    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);
    try {
      const headers = getHeaderIndexMap(sheet);
      const idCol = headers['ID'];
      const statusCol = headers['Status'];
      const tipoCol = headers['Tipo'];
      if (idCol === undefined || statusCol === undefined || tipoCol === undefined) {
        throw new Error('CabeÇõalhos obrigatÇürios nÇœo encontrados (ID, Status, Tipo)');
      }

      const row = findRowByExactValueInColumn(sheet, idCol, lancamento.id);
      if (!row) throw new Error('LanÇõamento nÇœo encontrado');

      const statusAtual = String(sheet.getRange(row, statusCol + 1).getDisplayValue() || '').toUpperCase();
      if (statusAtual !== 'PENDENTE' && statusAtual !== 'VENCIDA') {
        return { success: false, message: `LanÇõamento nÇœo estÇ­ pendente (status: ${statusAtual || 'N/A'})` };
      }

      const tipoAtual = String(sheet.getRange(row, tipoCol + 1).getDisplayValue() || '').toUpperCase();
      const tipoNovo = String(lancamento.tipo || '').toUpperCase();
      if (tipoAtual && tipoAtual !== tipoNovo) {
        return { success: false, message: `Tipo nÇœo pode ser alterado (${tipoAtual})` };
      }

      const valorBruto = Number(lancamento.valorBruto);
      const desconto = Number(lancamento.desconto || 0);
      const juros = Number(lancamento.juros || 0);
      const multa = Number(lancamento.multa || 0);
      const valorLiquido = valorBruto - desconto + juros + multa;

      if (!Number.isFinite(valorBruto) || valorBruto <= 0) {
        return { success: false, message: 'Valor bruto invÇ­lido' };
      }
      if (![desconto, juros, multa, valorLiquido].every(Number.isFinite)) {
        return { success: false, message: 'Valores numÇ¸ricos invÇ­lidos' };
      }

      const dataCompetencia = String(lancamento.dataCompetencia || '').trim();
      const dataVencimento = String(lancamento.dataVencimento || '').trim();
      const isValidDate = (value: string) => Number.isFinite(Date.parse(value));
      if (!isValidDate(dataCompetencia)) {
        return { success: false, message: 'Data competÇ¦ncia invÇ­lida' };
      }
      if (!isValidDate(dataVencimento)) {
        return { success: false, message: 'Data vencimento invÇ­lida' };
      }
      if (new Date(dataCompetencia).getTime() > new Date(dataVencimento).getTime()) {
        return { success: false, message: 'Data competÇ¦ncia nÇœo pode ser maior que data vencimento' };
      }

      const lastCol = sheet.getLastColumn();
      const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

      const setValue = (header: string, value: unknown) => {
        const colIndex = headers[header];
        if (colIndex === undefined) return;
        rowValues[colIndex] = value;
      };

      setValue('Data CompetÇ¦ncia', dataCompetencia);
      setValue('Data Vencimento', dataVencimento);
      setValue('Filial', sanitizeSheetString(lancamento.filial));
      setValue('Centro Custo', sanitizeSheetString(lancamento.centroCusto || ''));
      setValue('Conta Gerencial', sanitizeSheetString(lancamento.contaGerencial || ''));
      setValue('Conta ContÇ­bil', sanitizeSheetString(lancamento.contaContabil));
      setValue('Grupo Receita', sanitizeSheetString(lancamento.grupoReceita || ''));
      setValue('Canal', sanitizeSheetString(lancamento.canal || ''));
      setValue('DescriÇõÇœo', sanitizeSheetString(lancamento.descricao));
      setValue('Valor Bruto', valorBruto);
      setValue('Desconto', desconto);
      setValue('Juros', juros);
      setValue('Multa', multa);
      setValue('Valor LÇðquido', valorLiquido);
      setValue('ObservaÇõÇæes', sanitizeSheetString(lancamento.observacoes || ''));
      setValue('N Documento', sanitizeSheetString((lancamento as any).numeroDocumento || ''));

      sheet.getRange(row, 1, 1, lastCol).setValues([rowValues]);

      appendAuditLog('atualizarLancamento', { id: lancamento.id }, true);
      clearReportsCache();
      return { success: true, message: 'LanÇõamento atualizado com sucesso' };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (error: any) {
    appendAuditLog('atualizarLancamento', { id: lancamento?.id }, false, error?.message);
    return { success: false, message: error.message };
  }
}

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

export function desfazerConciliacao(extratoId: string, lancamentoId: string): { success: boolean; message: string } {
  try {
    const denied = requirePermission('editarLancamentos', 'desfazer conciliaÇõÇœo');
    if (denied) return denied;

    const v = combineValidations(
      validateRequired(extratoId, 'Extrato ID'),
      validateRequired(lancamentoId, 'LanÇõamento ID')
    );
    if (!v.valid) return { success: false, message: v.errors.join('; ') };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);
    try {
      const sheetExtratos = ss.getSheetByName(SHEET_TB_EXTRATOS);
      if (!sheetExtratos) throw new Error('Aba de extratos nÇœo encontrada');
      const extrHeaders = getHeaderIndexMap(sheetExtratos);
      const extrIdCol = extrHeaders['ID'];
      const extrStatusCol = extrHeaders['Status ConciliaÇõÇœo'];
      const extrLancCol = extrHeaders['ID LanÇõamento'];
      if (extrIdCol === undefined || extrStatusCol === undefined || extrLancCol === undefined) {
        throw new Error('CabeÇõalhos obrigatÇürios nÇœo encontrados em extratos (ID, Status ConciliaÇõÇœo, ID LanÇõamento)');
      }

      const extratoRow = findRowByExactValueInColumn(sheetExtratos, extrIdCol, extratoId);
      if (!extratoRow) throw new Error('Extrato nÇœo encontrado');
      const status = String(sheetExtratos.getRange(extratoRow, extrStatusCol + 1).getDisplayValue() || '').toUpperCase();
      if (status !== 'CONCILIADO') {
        return { success: false, message: `Extrato nÇœo estÇ­ conciliado (status: ${status || 'N/A'})` };
      }

      const currentLancId = String(sheetExtratos.getRange(extratoRow, extrLancCol + 1).getDisplayValue() || '').trim();
      if (currentLancId && currentLancId !== String(lancamentoId || '').trim()) {
        return { success: false, message: `Extrato vinculado a outro lanÇõamento: ${currentLancId}` };
      }

      const sheetLanc = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
      if (!sheetLanc) throw new Error('Aba de lanÇõamentos nÇœo encontrada');
      const lancHeaders = getHeaderIndexMap(sheetLanc);
      const lancIdCol = lancHeaders['ID'];
      const lancExtratoCol = lancHeaders['ID Extrato Banco'];
      if (lancIdCol === undefined || lancExtratoCol === undefined) {
        throw new Error('CabeÇõalhos obrigatÇürios nÇœo encontrados em lanÇõamentos (ID, ID Extrato Banco)');
      }

      const lancRow = findRowByExactValueInColumn(sheetLanc, lancIdCol, lancamentoId);
      if (!lancRow) throw new Error('LanÇõamento nÇœo encontrado');
      const currentExtratoId = String(sheetLanc.getRange(lancRow, lancExtratoCol + 1).getDisplayValue() || '').trim();
      if (currentExtratoId && currentExtratoId !== String(extratoId || '').trim()) {
        return { success: false, message: `LanÇõamento vinculado a outro extrato: ${currentExtratoId}` };
      }

      sheetExtratos.getRange(extratoRow, extrStatusCol + 1).setValue('PENDENTE');
      sheetExtratos.getRange(extratoRow, extrLancCol + 1).setValue('');
      sheetLanc.getRange(lancRow, lancExtratoCol + 1).setValue('');

      appendAuditLog('desfazerConciliacao', { extratoId, lancamentoId }, true);
      clearReportsCache();
      return { success: true, message: 'ConciliaÇõÇœo desfeita com sucesso' };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (error: any) {
    appendAuditLog('desfazerConciliacao', { extratoId, lancamentoId }, false, error?.message);
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
// CAIXAS
// ============================================================================

type CaixaRow = {
  id: string;
  canal: string;
  colaborador: string;
  dataFechamento: string;
  comunicadoInterno: string;
  observacoesEntradas: string;
  observacoesSaidas: string;
  sistemaValor: number;
  reforco: number;
  criadoEm: string;
  atualizadoEm: string;
};

type CaixaMovRow = {
  id: string;
  caixaId: string;
  tipo: string;
  natureza: 'ENTRADA' | 'SAIDA';
  valor: number;
  dataMov: string;
  observacoes?: string;
  arquivoUrl?: string;
  arquivoNome?: string;
  criadoEm: string;
  atualizadoEm: string;
};

type CaixaTipoConfig = {
  tipo: string;
  natureza: 'ENTRADA' | 'SAIDA';
  requerArquivo: boolean;
  sistemaFc: boolean;
  contaReforco: boolean;
  ativo: boolean;
};

function getCaixaTiposConfig(): CaixaTipoConfig[] {
  ensureCaixaTiposSchema();
  const values = getSheetValues(SHEET_REF_CAIXA_TIPOS, { skipHeader: true });
  return values
    .filter((r) => r && r[REF_CAIXA_TIPOS_COLS.TIPO])
    .map((r) => ({
      tipo: String(r[REF_CAIXA_TIPOS_COLS.TIPO]),
      natureza: String(r[REF_CAIXA_TIPOS_COLS.NATUREZA] || 'ENTRADA').toUpperCase() as 'ENTRADA' | 'SAIDA',
      requerArquivo: String(r[REF_CAIXA_TIPOS_COLS.REQUER_ARQUIVO] ?? 'FALSE').toUpperCase() === 'TRUE',
      sistemaFc: String(r[REF_CAIXA_TIPOS_COLS.SISTEMA_FC] ?? 'FALSE').toUpperCase() === 'TRUE',
      contaReforco: String(r[REF_CAIXA_TIPOS_COLS.CONTA_REFORCO] ?? 'FALSE').toUpperCase() === 'TRUE',
      ativo: String(r[REF_CAIXA_TIPOS_COLS.ATIVO] ?? 'TRUE').toUpperCase() !== 'FALSE',
    }));
}


function getCaixaTipoByName(tipo: string): CaixaTipoConfig | null {
  const key = String(tipo || '').trim().toUpperCase();
  if (!key) return null;
  const tipos = getCaixaTiposConfig();
  return tipos.find((t) => String(t.tipo).trim().toUpperCase() === key) || null;
}

function getCaixasRows(): CaixaRow[] {
  ensureCaixasSheets();
  const values = getSheetValues(SHEET_TB_CAIXAS, { skipHeader: true });
  return values
    .filter((r) => r && r[TB_CAIXAS_COLS.ID])
    .map((r) => ({
      id: String(r[TB_CAIXAS_COLS.ID] || ''),
      canal: String(r[TB_CAIXAS_COLS.CANAL] || ''),
      colaborador: String(r[TB_CAIXAS_COLS.COLABORADOR] || ''),
      dataFechamento: normalizeDateInput(r[TB_CAIXAS_COLS.DATA_FECHAMENTO]),
      comunicadoInterno: String(r[TB_CAIXAS_COLS.COMUNICADO_INTERNO] || ''),
      observacoesEntradas: String(r[TB_CAIXAS_COLS.OBSERVACOES_ENTRADAS] || ''),
      observacoesSaidas: String(r[TB_CAIXAS_COLS.OBSERVACOES_SAIDAS] || ''),
      sistemaValor: parseMoneyInput(r[TB_CAIXAS_COLS.SISTEMA_VALOR]),
      reforco: parseMoneyInput(r[TB_CAIXAS_COLS.REFORCO]),
      criadoEm: String(r[TB_CAIXAS_COLS.CRIADO_EM] || ''),
      atualizadoEm: String(r[TB_CAIXAS_COLS.ATUALIZADO_EM] || ''),
    }));
}

function getCaixaMovRows(caixaId?: string): CaixaMovRow[] {
  ensureCaixasSheets();
  const values = getSheetValues(SHEET_TB_CAIXAS_MOV, { skipHeader: true });
  const rows = values
    .filter((r) => r && r[TB_CAIXAS_MOV_COLS.ID])
    .map((r) => ({
      id: String(r[TB_CAIXAS_MOV_COLS.ID] || ''),
      caixaId: String(r[TB_CAIXAS_MOV_COLS.CAIXA_ID] || ''),
      tipo: String(r[TB_CAIXAS_MOV_COLS.TIPO] || ''),
      natureza: String(r[TB_CAIXAS_MOV_COLS.NATUREZA] || 'ENTRADA') as 'ENTRADA' | 'SAIDA',
      valor: parseMoneyInput(r[TB_CAIXAS_MOV_COLS.VALOR]),
      dataMov: normalizeDateInput(r[TB_CAIXAS_MOV_COLS.DATA_MOV]),
      observacoes: String(r[TB_CAIXAS_MOV_COLS.OBSERVACOES] || ''),
      arquivoUrl: String(r[TB_CAIXAS_MOV_COLS.ARQUIVO_URL] || ''),
      arquivoNome: String(r[TB_CAIXAS_MOV_COLS.ARQUIVO_NOME] || ''),
      criadoEm: String(r[TB_CAIXAS_MOV_COLS.CRIADO_EM] || ''),
      atualizadoEm: String(r[TB_CAIXAS_MOV_COLS.ATUALIZADO_EM] || ''),
    }));
  return caixaId ? rows.filter((r) => r.caixaId === caixaId) : rows;
}

function computeCaixaResumo(movimentos: CaixaMovRow[]) {
  const tipos = getCaixaTiposConfig();
  const tipoMap = new Map(tipos.map((t) => [String(t.tipo).trim().toUpperCase(), t]));

  let totalEntradas = 0;
  let totalSaidas = 0;
  let sistemaFc = 0;
  let reforco = 0;
  const pendencias: Array<{ id: string; tipo: string }> = [];

  movimentos.forEach((m) => {
    const config = tipoMap.get(String(m.tipo || '').trim().toUpperCase());
    const valor = Number(m.valor || 0);
    const natureza = String(m.natureza || '').toUpperCase();

    if (config?.sistemaFc) {
      sistemaFc += valor;
    } else if (natureza === 'SAIDA') {
      totalSaidas += valor;
    } else {
      totalEntradas += valor;
    }

    if (config?.contaReforco && natureza !== 'SAIDA') {
      reforco += valor;
    }

    if (config?.requerArquivo && !m.arquivoUrl) {
      pendencias.push({ id: m.id, tipo: m.tipo });
    }
  });

  const totalRecebido = totalEntradas - totalSaidas;
  const diferenca = totalRecebido - sistemaFc;

  return { totalEntradas, totalSaidas, totalRecebido, sistemaFc, diferenca, reforco, pendencias };
}

function seedReforcoMovimento(caixaId: string, canal: string, dataFechamento: string): void {
  if (!caixaId || !canal || !dataFechamento) return;
  const prevDate = getPreviousBusinessDate(dataFechamento);
  if (!prevDate || prevDate === dataFechamento) return;

  const caixas = getCaixasRows();
  const prev = caixas.find((c) => String(c.canal) === String(canal) && String(c.dataFechamento) === prevDate);
  if (!prev) return;

  const prevMovs = getCaixaMovRows(prev.id);
  const prevResumo = computeCaixaResumo(prevMovs);
  const valor = Number(prevResumo.reforco || 0);
  if (!valor || Math.abs(valor) < 0.009) return;

  const currentMovs = getCaixaMovRows(caixaId);
  const exists = currentMovs.some((m) => normalizeKey(m.tipo) === normalizeKey('Reforco do Caixa'));
  if (exists) return;

  const movSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TB_CAIXAS_MOV);
  if (!movSheet) return;

  const now = new Date().toISOString();
  movSheet.appendRow([
    Utilities.getUuid(),
    String(caixaId),
    'Reforco do Caixa',
    'ENTRADA',
    valor,
    dataFechamento,
    '',
    '',
    now,
    now,
    ''
  ]);
  appendAuditLog('caixas:movimento:reforco', { caixaId, canal, dataFechamento, valor }, true);
}

export function getCaixasData(): { success: boolean; caixas: CaixaRow[]; movimentos: CaixaMovRow[] } {
  const denied = requireAnyPermission<{ success: boolean; message: string }>(
    ['visualizarRelatorios', 'importarArquivos'],
    'carregar caixas'
  );
  if (denied) return { success: false, caixas: [], movimentos: [] };

  const requester = getUsuarioByEmail(getRequestingUserEmail());
  let caixas = getCaixasRows();
  let movimentos = getCaixaMovRows();
  if (requester && requester.perfil === 'CAIXA' && requester.canal) {
    const canal = String(requester.canal);
    const caixaIds = new Set(caixas.filter((c) => String(c.canal) === canal).map((c) => c.id));
    caixas = caixas.filter((c) => String(c.canal) === canal);
    movimentos = movimentos.filter((m) => caixaIds.has(m.caixaId));
  }
  appendAuditLog('caixas:listar', { total: caixas.length }, true);
  return { success: true, caixas, movimentos };
}

export function getCaixaMovimentos(caixaId: string): { success: boolean; movimentos: CaixaMovRow[] } {
  const denied = requireAnyPermission<{ success: boolean; message: string }>(
    ['visualizarRelatorios', 'importarArquivos'],
    'carregar movimentos de caixa'
  );
  if (denied) return { success: false, movimentos: [] };

  const id = String(caixaId || '').trim();
  if (!id) return { success: true, movimentos: [] };
  const requester = getUsuarioByEmail(getRequestingUserEmail());
  if (requester && requester.perfil === 'CAIXA' && requester.canal) {
    const caixa = getCaixasRows().find((c) => c.id === id);
    if (!caixa || String(caixa.canal) !== String(requester.canal)) {
      return { success: true, movimentos: [] };
    }
  }
  const movimentos = getCaixaMovRows(id);
  appendAuditLog('caixas:movimentos:listar', { id, total: movimentos.length }, true);
  return { success: true, movimentos };
}

export function salvarCaixa(caixa: {
  id?: string;
  canal: string;
  colaborador?: string;
  dataFechamento: string;
  observacoesEntradas?: string;
  observacoesSaidas?: string;
  comunicadoInterno?: string;
  sistemaValor?: number | string;
  reforco?: number | string;
  finalizar?: boolean;
}): { success: boolean; message: string; id?: string } {
  const denied = requirePermission('importarArquivos', 'salvar caixa');
  if (denied) return denied;

  const validation = combineValidations(
    validateRequired(caixa?.canal, 'Canal'),
    validateRequired(caixa?.colaborador, 'Colaborador'),
    validateRequired(caixa?.dataFechamento, 'Data Fechamento')
  );
  if (!validation.valid) return { success: false, message: validation.errors.join('; ') };

  ensureCaixasSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TB_CAIXAS);
  if (!sheet) return { success: false, message: 'Aba de caixas nao encontrada' };

  const now = new Date().toISOString();
  const id = caixa.id ? String(caixa.id) : Utilities.getUuid();
  const dataFechamento = normalizeDateInput(caixa.dataFechamento);
  const dataDate = dataFechamento ? new Date(`${dataFechamento}T00:00:00`) : null;
  if (dataDate && !Number.isNaN(dataDate.getTime()) && dataDate.getDay() === 0) {
    return { success: false, message: 'Fechamento nao permitido aos domingos' };
  }
  const movimentos = getCaixaMovRows(id);
  const resumo = computeCaixaResumo(movimentos);
  const sistemaValor = resumo.sistemaFc;
  const reforco = resumo.reforco;

  const finalizar = caixa.finalizar !== false;
  if (finalizar) {
    if (Math.abs(resumo.diferenca) > 0.009 && !String(caixa.comunicadoInterno || '').trim()) {
      return { success: false, message: 'Informe o Comunicado Interno (CI) quando houver diferenca' };
    }
    if (resumo.pendencias.length > 0) {
      return { success: false, message: 'Ha movimentacoes que exigem comprovante' };
    }
  }

  let rowIndex = findRowByExactValueInColumn(sheet, TB_CAIXAS_COLS.ID, id, 2);
  let criadoEm = now;
  if (rowIndex) {
    const existing = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    criadoEm = String(existing[TB_CAIXAS_COLS.CRIADO_EM] || now);
  }

  let allRows: any[] = [];
  if (sheet.getLastRow() > 1) {
    allRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  }
  const duplicate = allRows.find((r, idx) => {
    const existingId = String(r[TB_CAIXAS_COLS.ID] || '');
    if (existingId === id) return false;
    const existingCanal = String(r[TB_CAIXAS_COLS.CANAL] || '');
    const existingData = normalizeDateInput(r[TB_CAIXAS_COLS.DATA_FECHAMENTO]);
    return existingCanal === String(caixa.canal) && existingData === dataFechamento;
  });
  if (duplicate) {
    return { success: false, message: 'Ja existe um caixa para este canal e data' };
  }

  const rowData = [
    id,
    sanitizeSheetString(caixa.canal || ''),
    sanitizeSheetString(caixa.colaborador || ''),
    dataFechamento,
    sanitizeSheetString(caixa.comunicadoInterno || ''),
    sistemaValor,
    reforco,
    criadoEm,
    now,
    sanitizeSheetString(caixa.observacoesEntradas || ''),
    sanitizeSheetString(caixa.observacoesSaidas || ''),
  ];

  if (rowIndex) {
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  appendAuditLog('caixas:salvar', { id, canal: caixa.canal, finalizar }, true);
  return { success: true, message: finalizar ? 'Caixa fechado com sucesso' : 'Caixa salvo com sucesso', id };
}

export function excluirCaixa(id: string): { success: boolean; message: string } {
  const denied = requirePermission('importarArquivos', 'excluir caixa');
  if (denied) return denied;
  const caixaId = String(id || '').trim();
  if (!caixaId) return { success: false, message: 'ID invalido' };

  ensureCaixasSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TB_CAIXAS);
  const movSheet = ss.getSheetByName(SHEET_TB_CAIXAS_MOV);
  if (!sheet || !movSheet) return { success: false, message: 'Abas de caixas nao encontradas' };

  const rowIndex = findRowByExactValueInColumn(sheet, TB_CAIXAS_COLS.ID, caixaId, 2);
  if (!rowIndex) return { success: false, message: 'Caixa nao encontrado' };

  sheet.deleteRow(rowIndex);

  const movValues = movSheet.getDataRange().getValues();
  for (let i = movValues.length - 1; i >= 1; i--) {
    if (String(movValues[i][TB_CAIXAS_MOV_COLS.CAIXA_ID] || '') === caixaId) {
      movSheet.deleteRow(i + 1);
    }
  }

  appendAuditLog('caixas:excluir', { id: caixaId }, true);
  return { success: true, message: 'Caixa excluido' };
}

export function salvarCaixaMovimento(mov: {
  id?: string;
  caixaId: string;
  tipo: string;
  natureza: string;
  valor: number | string;
  dataMov?: string;
  observacoes?: string;
  arquivoUrl?: string;
  arquivoNome?: string;
}): { success: boolean; message: string; id?: string } {
  const denied = requirePermission('importarArquivos', 'salvar movimento de caixa');
  if (denied) return denied;

  const validation = combineValidations(
    validateRequired(mov?.caixaId, 'Caixa'),
    validateRequired(mov?.tipo, 'Tipo'),
    validateEnum(String(mov?.natureza || ''), ['ENTRADA', 'SAIDA'], 'Natureza')
  );
  if (!validation.valid) return { success: false, message: validation.errors.join('; ') };

  ensureCaixasSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TB_CAIXAS_MOV);
  if (!sheet) return { success: false, message: 'Aba de movimentacoes nao encontrada' };

  const now = new Date().toISOString();
  const id = mov.id ? String(mov.id) : Utilities.getUuid();
  const valor = parseMoneyInput(mov.valor);
  let dataMov = normalizeDateInput(mov.dataMov || '');
  if (!dataMov) {
    const caixaRows = getCaixasRows();
    const caixa = caixaRows.find((c) => c.id === String(mov.caixaId));
    dataMov = caixa ? caixa.dataFechamento : '';
  }

  let rowIndex = findRowByExactValueInColumn(sheet, TB_CAIXAS_MOV_COLS.ID, id, 2);
  let criadoEm = now;
  let arquivoUrl = sanitizeSheetString(mov.arquivoUrl || '');
  let arquivoNome = sanitizeSheetString(mov.arquivoNome || '');
  let observacoes = sanitizeSheetString(mov.observacoes || '');
  if (rowIndex) {
    const existing = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    criadoEm = String(existing[TB_CAIXAS_MOV_COLS.CRIADO_EM] || now);
    if (!arquivoUrl) arquivoUrl = String(existing[TB_CAIXAS_MOV_COLS.ARQUIVO_URL] || '');
    if (!arquivoNome) arquivoNome = String(existing[TB_CAIXAS_MOV_COLS.ARQUIVO_NOME] || '');
    if (!observacoes) observacoes = String(existing[TB_CAIXAS_MOV_COLS.OBSERVACOES] || '');
  }

  const tipoConfig = getCaixaTipoByName(mov.tipo);
  if (tipoConfig?.requerArquivo && !arquivoUrl) {
    return { success: false, message: 'Este tipo requer comprovante' };
  }

  const rowData = [
    id,
    String(mov.caixaId),
    sanitizeSheetString(mov.tipo || ''),
    String(mov.natureza || 'ENTRADA').toUpperCase(),
    valor,
    dataMov,
    arquivoUrl,
    arquivoNome,
    criadoEm,
    now,
    observacoes,
  ];

  if (rowIndex) {
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  appendAuditLog('caixas:movimento', { id, caixaId: mov.caixaId, tipo: mov.tipo }, true);
  return { success: true, message: 'Movimento salvo', id };
}

export function excluirCaixaMovimento(id: string): { success: boolean; message: string } {
  const denied = requirePermission('importarArquivos', 'excluir movimento de caixa');
  if (denied) return denied;
  const movId = String(id || '').trim();
  if (!movId) return { success: false, message: 'ID invalido' };

  ensureCaixasSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TB_CAIXAS_MOV);
  if (!sheet) return { success: false, message: 'Aba de movimentacoes nao encontrada' };

  const rowIndex = findRowByExactValueInColumn(sheet, TB_CAIXAS_MOV_COLS.ID, movId, 2);
  if (!rowIndex) return { success: false, message: 'Movimento nao encontrado' };

  sheet.deleteRow(rowIndex);
  appendAuditLog('caixas:movimento:excluir', { id: movId }, true);
  return { success: true, message: 'Movimento excluido' };
}

export function getCaixaRelatorio(caixaId: string): { success: boolean; report?: any; message?: string } {
  const denied = requireAnyPermission<{ success: boolean; message: string }>(
    ['visualizarRelatorios', 'importarArquivos'],
    'relatorio caixas'
  );
  if (denied) return { success: false, message: denied.message };

  const id = String(caixaId || '').trim();
  if (!id) return { success: false, message: 'Caixa invalido' };

  const requester = getUsuarioByEmail(getRequestingUserEmail());
  const caixas = getCaixasRows();
  const caixa = caixas.find((c) => c.id === id);
  if (!caixa) return { success: false, message: 'Caixa nao encontrado' };
  if (requester && requester.perfil === 'CAIXA' && requester.canal) {
    if (String(caixa.canal) !== String(requester.canal)) {
      return { success: false, message: 'Sem permissao para este canal' };
    }
  }

  const movimentos = getCaixaMovRows(id);
  const porTipo = new Map<string, { tipo: string; entradas: number; saidas: number }>();
  const resumo = computeCaixaResumo(movimentos);

  movimentos.forEach((m) => {
    const key = m.tipo || 'Outros';
    const item = porTipo.get(key) || { tipo: key, entradas: 0, saidas: 0 };
    if (m.natureza === 'SAIDA') {
      item.saidas += m.valor;
    } else {
      item.entradas += m.valor;
    }
    porTipo.set(key, item);
  });

  appendAuditLog('caixas:relatorio', { id }, true);
  return {
    success: true,
    report: {
      caixa,
      totalEntradas: resumo.totalEntradas,
      totalSaidas: resumo.totalSaidas,
      totalRecebido: resumo.totalRecebido,
      diferenca: resumo.diferenca,
      sistemaFc: resumo.sistemaFc,
      reforco: resumo.reforco,
      porTipo: Array.from(porTipo.values()),
      movimentos,
      pendencias: resumo.pendencias,
    },
  };
}

function ensureFolder(parent: GoogleAppsScript.Drive.Folder, name: string): GoogleAppsScript.Drive.Folder {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function sanitizeFileName(value: string): string {
  return String(value || '')
    .trim()
    .replace(/[\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .slice(0, 120) || 'Arquivo';
}

function formatMoneyFile(value: number): string {
  const safe = Number.isFinite(value) ? value : 0;
  return safe.toFixed(2).replace('.', '_');
}

export function uploadCaixaArquivo(payload: {
  base64: string;
  mimeType: string;
  fileName: string;
  canal: string;
  dataFechamento: string;
  tipo: string;
  valor: number | string;
}): { success: boolean; message: string; url?: string } {
  const denied = requirePermission('importarArquivos', 'upload caixa');
  if (denied) return denied;

  const pastaId = getConfigValue('CAIXAS_PASTA_ID');
  if (!pastaId) return { success: false, message: 'Configure a pasta de caixas nas configuracoes' };

  const base64 = String(payload?.base64 || '');
  if (!base64) return { success: false, message: 'Arquivo invalido' };

  const canal = sanitizeFileName(payload?.canal || 'CANAL');
  const dataFechamento = normalizeDateInput(payload?.dataFechamento) || normalizeDateInput(new Date());
  const tipo = sanitizeFileName(payload?.tipo || 'Movimentacao');
  const valor = parseMoneyInput(payload?.valor);

  const ext = (String(payload?.fileName || '').split('.').pop() || 'dat').toLowerCase();
  const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmmss');
  const fileName = `${canal} - ${dataFechamento} - ${tipo} - ${formatMoneyFile(valor)} - ${time}.${ext}`;

  try {
    const root = DriveApp.getFolderById(pastaId);
    const canalFolder = ensureFolder(root, canal);
    const dateFolder = ensureFolder(canalFolder, dataFechamento);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64),
      payload.mimeType || 'application/octet-stream',
      fileName
    );
    const file = dateFolder.createFile(blob);
    appendAuditLog('caixas:upload', { canal, dataFechamento, tipo, valor, fileName }, true);
    return { success: true, message: 'Upload concluido', url: file.getUrl() };
  } catch (error: any) {
    appendAuditLog('caixas:upload', { canal, dataFechamento, tipo, valor }, false, error?.message);
    return { success: false, message: `Erro ao salvar arquivo: ${error?.message || error}` };
  }
}

// ============================================================================
// IMPORTACAO E COMPARATIVO (FC x SIEG x BANCO)
// ============================================================================

function normalizeKey(value: unknown): string {
  let s = String(value || '');
  // Tenta corrigir moji-bake (UTF-8 lido como Latin-1) apenas quando detectado.
  if (/[\u00C3\u00C2\uFFFD]/.test(s)) {
    try {
      const decoded = Utilities.newBlob(s, 'ISO-8859-1').getDataAsString('UTF-8');
      if (decoded && decoded !== s) s = decoded;
    } catch (e) {
      // ignore e segue com o texto original
    }
  }
  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return s.toLowerCase().replace(/[^a-z0-9]/g, '');
}




function normalizeDateInput(value: unknown): string {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(value).trim();
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  const brMatch = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (brMatch) {
    const dd = brMatch[1].padStart(2, '0');
    const mm = brMatch[2].padStart(2, '0');
    const yyyy = brMatch[3];
    return `${yyyy}-${mm}-${dd}`;
  }
  const altMatch = s.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
  if (altMatch) {
    const yyyy = altMatch[1];
    const mm = altMatch[2].padStart(2, '0');
    const dd = altMatch[3].padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }
  return s;
}

function getPreviousBusinessDate(dateStr: string): string {
  if (!dateStr) return '';
  const base = new Date(`${dateStr}T00:00:00`);
  if (Number.isNaN(base.getTime())) return dateStr;
  base.setDate(base.getDate() - 1);
  if (base.getDay() === 0) {
    base.setDate(base.getDate() - 1);
  }
  return Utilities.formatDate(base, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function parseMoneyInput(value: unknown): number {
  if (typeof value === 'number') return value;
  if (value === null || value === undefined) return 0;
  let s = String(value).trim();
  if (!s) return 0;
  s = s.replace(/[^\d,.-]/g, '');
  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function buildMatchKey(dateStr: string, value: number): string {
  const dateKey = normalizeDateInput(dateStr);
  if (!dateKey) return '';
  return `${dateKey}|${Math.abs(value).toFixed(2)}`;
}

function buildImportKey(parts: Array<string | number | null | undefined>): string {
  return parts
    .map((p) => normalizeKey(p))
    .filter((p) => p)
    .join('|');
}

function isItauMovement(lancamento: string): boolean {
  const key = normalizeKey(lancamento);
  if (!key) return false;
  if (key.includes('saldo')) return false;
  return true;
}

function resolveFilialFcFromRelations(
  codigoFilial: string,
  fallback: string,
  relations: Array<{ filialFc: string; filialSiegRelatorio: string; filialSiegContabil: string; ativa?: boolean }>
): string {
  const direct = String(fallback || '').trim();
  if (direct) return direct;
  const key = normalizeKey(codigoFilial);
  if (!key) return '';
  for (const rel of relations) {
    if (rel.ativa === false) continue;
    const relKey1 = normalizeKey(rel.filialSiegRelatorio);
    const relKey2 = normalizeKey(rel.filialSiegContabil);
    if (key === relKey1 || key === relKey2) {
      return String(rel.filialFc || '');
    }
  }
  return '';
}

function getImportFcRows(tipo?: string): any[] {
  createSheetIfNotExists(SHEET_TB_IMPORT_FC, [
    'Data Emissao', 'Num Documento', 'Cod Conta', 'Filial FC', 'Historico', 'Fornecedor',
    'Valor', 'Descricao', 'Data Baixa', 'Flag Baixa', 'Data Vencimento', 'Tipo', 'Importado Em',
  ]);
  const values = getSheetValues(SHEET_TB_IMPORT_FC, { skipHeader: true });
  return values
    .filter((r) => r && (r[0] || r[1]))
    .map((r) => ({
      dataEmissao: normalizeDateInput(r[TB_IMPORT_FC_COLS.DATA_EMISSAO]),
      numDocumento: String(r[TB_IMPORT_FC_COLS.NUM_DOCUMENTO] || ''),
      codConta: String(r[TB_IMPORT_FC_COLS.COD_CONTA] || ''),
      filialFc: String(r[TB_IMPORT_FC_COLS.FILIAL_FC] || ''),
      historico: String(r[TB_IMPORT_FC_COLS.HISTORICO] || ''),
      fornecedor: String(r[TB_IMPORT_FC_COLS.FORNECEDOR] || ''),
      valor: parseMoneyInput(r[TB_IMPORT_FC_COLS.VALOR]),
      descricao: String(r[TB_IMPORT_FC_COLS.DESCRICAO] || ''),
      dataBaixa: normalizeDateInput(r[TB_IMPORT_FC_COLS.DATA_BAIXA]),
      flagBaixa: String(r[TB_IMPORT_FC_COLS.FLAG_BAIXA] || ''),
      dataVencimento: normalizeDateInput(r[TB_IMPORT_FC_COLS.DATA_VENCIMENTO]),
      tipo: String(r[TB_IMPORT_FC_COLS.TIPO] || '').toUpperCase(),
    }))
    .filter((r) => (tipo ? r.tipo === tipo : true));
}

function getImportItauRows(): any[] {
  createSheetIfNotExists(SHEET_TB_IMPORT_ITAU, [
    'Data', 'Lancamento', 'Agencia/Origem', 'Razao Social', 'CPF/CNPJ', 'Valor',
    'Saldo', 'Conta', 'Filial FC', 'Modelo', 'Importado Em',
  ]);
  const values = getSheetValues(SHEET_TB_IMPORT_ITAU, { skipHeader: true });
  return values
    .filter((r) => r && (r[0] || r[1]))
    .map((r) => ({
      data: normalizeDateInput(r[TB_IMPORT_ITAU_COLS.DATA]),
      lancamento: String(r[TB_IMPORT_ITAU_COLS.LANCAMENTO] || ''),
      agenciaOrigem: String(r[TB_IMPORT_ITAU_COLS.AGENCIA_ORIGEM] || ''),
      razaoSocial: String(r[TB_IMPORT_ITAU_COLS.RAZAO_SOCIAL] || ''),
      cpfCnpj: String(r[TB_IMPORT_ITAU_COLS.CPF_CNPJ] || ''),
      valor: parseMoneyInput(r[TB_IMPORT_ITAU_COLS.VALOR]),
      saldo: parseMoneyInput(r[TB_IMPORT_ITAU_COLS.SALDO]),
      conta: String(r[TB_IMPORT_ITAU_COLS.CONTA] || ''),
      filialFc: String(r[TB_IMPORT_ITAU_COLS.FILIAL_FC] || ''),
      modelo: String(r[TB_IMPORT_ITAU_COLS.MODELO] || ''),
    }));
}

function getImportSiegRows(): any[] {
  createSheetIfNotExists(SHEET_TB_IMPORT_SIEG, [
    'Num NFe', 'Valor', 'Data Emissao', 'CNPJ Emit', 'Nome Fant Emit', 'Razao Soc Emit',
    'CNPJ Dest', 'Nome Fant Dest', 'Razao Soc Dest', 'Data Envio Cofre', 'Chave NFe',
    'Tags', 'Codigo Evento', 'Tipo Evento', 'Status', 'Danfe', 'Xml', 'Codigo Filial',
    'Filial FC', 'Importado Em',
  ]);
  const values = getSheetValues(SHEET_TB_IMPORT_SIEG, { skipHeader: true });
  return values
    .filter((r) => r && (r[0] || r[1]))
    .map((r) => ({
      numNfe: String(r[TB_IMPORT_SIEG_COLS.NUM_NFE] || ''),
      valor: parseMoneyInput(r[TB_IMPORT_SIEG_COLS.VALOR]),
      dataEmissao: normalizeDateInput(r[TB_IMPORT_SIEG_COLS.DATA_EMISSAO]),
      cnpjEmit: String(r[TB_IMPORT_SIEG_COLS.CNPJ_EMIT] || ''),
      nomeFantEmit: String(r[TB_IMPORT_SIEG_COLS.NOME_FANT_EMIT] || ''),
      razaoEmit: String(r[TB_IMPORT_SIEG_COLS.RAZAO_EMIT] || ''),
      cnpjDest: String(r[TB_IMPORT_SIEG_COLS.CNPJ_DEST] || ''),
      nomeFantDest: String(r[TB_IMPORT_SIEG_COLS.NOME_FANT_DEST] || ''),
      razaoDest: String(r[TB_IMPORT_SIEG_COLS.RAZAO_DEST] || ''),
      dataEnvioCofre: normalizeDateInput(r[TB_IMPORT_SIEG_COLS.DATA_ENVIO_COFRE]),
      chaveNfe: String(r[TB_IMPORT_SIEG_COLS.CHAVE_NFE] || ''),
      tags: String(r[TB_IMPORT_SIEG_COLS.TAGS] || ''),
      codigoEvento: String(r[TB_IMPORT_SIEG_COLS.CODIGO_EVENTO] || ''),
      tipoEvento: String(r[TB_IMPORT_SIEG_COLS.TIPO_EVENTO] || ''),
      status: String(r[TB_IMPORT_SIEG_COLS.STATUS] || ''),
      danfe: String(r[TB_IMPORT_SIEG_COLS.DANFE] || ''),
      xml: String(r[TB_IMPORT_SIEG_COLS.XML] || ''),
      codigoFilial: String(r[TB_IMPORT_SIEG_COLS.CODIGO_FILIAL_SIEG] || ''),
      filialFc: String(r[TB_IMPORT_SIEG_COLS.FILIAL_FC] || ''),
    }));
}

export function importarFc(
  rows: Array<any>,
  meta?: { tipo?: string; filialFc?: string }
): { success: boolean; message: string; imported?: number } {
  const denied = requirePermission('importarArquivos', 'importar FC');
  if (denied) return denied;

  const tipo = String(meta?.tipo || '').toUpperCase();
  if (!['PAGAR', 'RECEBER'].includes(tipo)) {
    return { success: false, message: 'Tipo invalido para importacao FC' };
  }

  const now = new Date().toISOString();
  const payload = Array.isArray(rows) ? rows : [];
  const existing = getImportFcRows(tipo);
  const existingKeys = new Set(
    existing.map((r) => buildImportKey([r.dataEmissao, r.valor, r.numDocumento, r.filialFc, r.tipo]))
  );

  const values = payload
    .filter((r) => r)
    .map((r) => ({
      dataEmissao: normalizeDateInput(r.dataEmissao),
      numDocumento: sanitizeSheetString(r.numDocumento || ''),
      codConta: sanitizeSheetString(r.codConta || ''),
      filialFc: sanitizeSheetString(r.filialFc || meta?.filialFc || ''),
      historico: sanitizeSheetString(r.historico || ''),
      fornecedor: sanitizeSheetString(r.fornecedor || ''),
      valor: parseMoneyInput(r.valor),
      descricao: sanitizeSheetString(r.descricao || ''),
      dataBaixa: normalizeDateInput(r.dataBaixa),
      flagBaixa: sanitizeSheetString(r.flagBaixa || ''),
      dataVencimento: normalizeDateInput(r.dataVencimento),
      tipo: tipo,
    }))
    .filter((r) => {
      const key = buildImportKey([r.dataEmissao, r.valor, r.numDocumento, r.filialFc, r.tipo]);
      if (!key || existingKeys.has(key)) return false;
      existingKeys.add(key);
      return true;
    })
    .map((r) => ([
      r.dataEmissao,
      r.numDocumento,
      r.codConta,
      r.filialFc,
      r.historico,
      r.fornecedor,
      r.valor,
      r.descricao,
      r.dataBaixa,
      r.flagBaixa,
      r.dataVencimento,
      r.tipo,
      now,
    ]));

  const imported = values.length;
  const total = payload.length;
  const skipped = Math.max(0, total - imported);
  if (!values.length) {
    if (payload.length > 0) {
      return { success: false, message: 'Nenhuma linha importada (todas duplicadas ou invalidas)' };
    }
    return { success: false, message: 'Nenhuma linha valida para importar' };
  }

  appendRows(SHEET_TB_IMPORT_FC, values);
  appendAuditLog('importarFc', { tipo, imported, skipped }, true);
  cacheRemoveNamespace(CacheNamespace.CONCILIACAO, CacheScope.SCRIPT);
  return {
    success: true,
    message: `Importado ${imported} linhas FC${skipped ? ` (ignoradas ${skipped} duplicadas)` : ''}`,
    imported,
  };
}

export function getSheetData(
  sheetIdOrUrl: string,
  sheetName?: string
): { success: boolean; message?: string; values?: string[][]; truncated?: boolean } {
  const denied = requirePermission('importarArquivos', 'ler planilha');
  if (denied) return denied;

  const input = String(sheetIdOrUrl || '').trim();
  if (!input) {
    return { success: false, message: 'Informe o ID ou URL da planilha' };
  }

  let sheetId = input;
  const match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) sheetId = match[1];
  const idMatch = input.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (idMatch) sheetId = idMatch[1];
  if (!/^[a-zA-Z0-9-_]{10,}$/.test(sheetId)) {
    return { success: false, message: 'ID da planilha invalido' };
  }

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
    if (!sheet) {
      return { success: false, message: 'Aba nao encontrada na planilha' };
    }
    const MAX_ROWS = 2000;
    const MAX_COLS = 50;
    const MAX_CELLS = 100000;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) {
      return { success: true, values: [] };
    }
    const rows = Math.min(lastRow, MAX_ROWS);
    const cols = Math.min(lastCol, MAX_COLS);
    const truncated = rows * cols < lastRow * lastCol || rows * cols > MAX_CELLS;
    const safeRows = Math.min(rows, Math.max(1, Math.floor(MAX_CELLS / cols)));
    const values = sheet.getRange(1, 1, safeRows, cols).getDisplayValues();
    if (truncated || safeRows < lastRow) {
      return {
        success: true,
        values,
        truncated: true,
        message: 'Planilha muito grande, leitura limitada para performance',
      };
    }
    return { success: true, values, truncated: false };
  } catch (error: any) {
    return { success: false, message: `Erro ao ler planilha: ${error?.message || error}` };
  }
}



type ParsedContaPaga = {
  dataVencimento: string;
  dataCompetencia: string;
  dataPagamento: string;
  contaContabil: string;
  descricao: string;
  numeroDocumento: string;
  valor: number;
  tipo: 'DESPESA' | 'RECEITA';
  filialOriginal: string;
  filialMapeada: string;
  filialMapeadaOrigem: string;
  rateio: boolean;
  rateioCount: number;
};

function parseContasPagasTxt(content: string, filiais: Array<{ codigo: string; nome: string }> = []): ParsedContaPaga[] {
  const lines = String(content || '').split(/\r?\n/);
  const items: ParsedContaPaga[] = [];
  let filialAtual = '';

  const filialRegex = /^Filial:\s*(\d+)\s*-\s*([^\r]+?)\s{2,}/i;
  const tailRegex = /([A-Z])\s+([A-Z])\s+(\d{2}\/\d{2}\/\d{4})\s*$/;
  let colContaStart = -1;
  let colHistStart = -1;
  let colDocStart = -1;
  let colValorStart = -1;

  for (const rawLine of lines) {
    const line = String(rawLine || '').trimRight();
    if (!line) continue;
    const filialMatch = line.match(filialRegex);
    if (filialMatch) {
      const codigo = filialMatch[1];
      const nome = filialMatch[2].trim();
      filialAtual = `${codigo} - ${nome}`;
      continue;
    }
    if (line.includes('DT.VENC.') && line.includes('VALOR') && line.includes('DT.BAIXA')) {
      colContaStart = line.indexOf('CONTA');
      colHistStart = line.indexOf('HISTORICO');
      colDocStart = line.indexOf('/N.DOCUMENTO');
      colValorStart = line.indexOf('VALOR');
      continue;
    }
    if (!/^\d{2}\/\d{2}\/\d{4}\s+\d{2}\/\d{2}\/\d{4}\s+/.test(line)) continue;
    const tail = line.match(tailRegex);
    if (!tail) continue;

    const tipoFlag = String(tail[1] || '').toUpperCase();
    const dtBaixa = normalizeDateInput(tail[3]);
    const tailIndex = typeof tail.index === 'number' ? tail.index : line.length;

    let valorStr = '';
    if (colValorStart >= 0) {
      valorStr = line.substring(colValorStart, tailIndex).trim();
    } else {
      const valueMatch = line.match(/((?:\d{1,3}(?:\.\d{3})+|\d{1,6}),\d{2})\s+[A-Z]\s+[A-Z]\s+\d{2}\/\d{2}\/\d{4}\s*$/);
      valorStr = valueMatch ? valueMatch[1] : '';
    }
    const valor = parseMoneyInput(valorStr);

    let dtVenc = '';
    let dtOpe = '';
    let conta = '';
    let historico = '';
    let numeroDocumento = '';

    if (colContaStart >= 0 && colHistStart > colContaStart) {
      dtVenc = normalizeDateInput(line.substring(0, 10).trim());
      dtOpe = normalizeDateInput(line.substring(11, 21).trim());
      conta = line.substring(colContaStart, colHistStart).trim();
      if (colDocStart > colHistStart && colValorStart > colDocStart) {
        historico = line.substring(colHistStart, colDocStart).trim();
        numeroDocumento = line.substring(colDocStart, colValorStart).trim().replace(/^\/+/, '').trim();
      }
    } else {
      const left = line.slice(0, tailIndex).trimRight();
      const startMatch = left.match(/^(\d{2}\/\d{2}\/\d{4})\s+(\d{2}\/\d{2}\/\d{4})\s+(\S+)\s+(.*)$/);
      if (!startMatch) continue;
      dtVenc = normalizeDateInput(startMatch[1]);
      dtOpe = normalizeDateInput(startMatch[2]);
      conta = String(startMatch[3] || '').trim();
      const rest = String(startMatch[4] || '').trimRight();
      historico = rest;
      const restParts = rest.split(/\s{2,}/).map(p => p.trim()).filter(Boolean);
      if (restParts.length >= 2) {
        numeroDocumento = restParts.pop() || '';
        historico = restParts.join(' ').trim();
      } else {
        const docMatch = rest.match(/^(.*?)(?:\s+)([0-9A-Z][0-9A-Z.\/-]{3,})$/);
        if (docMatch) {
          historico = docMatch[1].trim();
          numeroDocumento = docMatch[2].trim();
        }
      }
    }

    const tipo = tipoFlag === 'D' ? 'DESPESA' : 'RECEITA';
    if (!dtVenc || !dtOpe || !dtBaixa || !conta || !historico || !Number.isFinite(valor)) continue;

    const mapping = mapFilialTxt(filialAtual || '', filiais);
    const rateio = isRateioFilial(filialAtual) || isRateioFilial(mapping.codigo);
    const rateioTargets = rateio ? getRateioTargets() : [];
    items.push({
      dataVencimento: dtVenc,
      dataCompetencia: dtOpe,
      dataPagamento: dtBaixa,
      contaContabil: conta,
      descricao: historico,
      numeroDocumento,
      valor,
      tipo,
      filialOriginal: filialAtual || 'N/A',
      filialMapeada: mapping.codigo,
      filialMapeadaOrigem: mapping.origem,
      rateio,
      rateioCount: rateioTargets.length,
    });
  }

  return items;
}

function mapFilialTxt(filialTxt: string, filiais: Array<{ codigo: string; nome: string }>): { codigo: string; origem: string } {
  const fallback = filialTxt || 'N/A';
  if (isRateioFilial(filialTxt)) return { codigo: 'RATEIO', origem: 'rateio' };
  if (!filiais || !filiais.length) return { codigo: fallback, origem: 'fallback' };

  const normalize = (v: string) => normalizeKey(v || '');
  const txt = String(filialTxt || '').trim();
  const match = txt.match(/^(\d+)\s*-\s*(.+)$/);
  const codigoTxt = match ? match[1] : '';
  const nomeTxt = match ? match[2] : txt;

  const codeNorm = normalize(codigoTxt);
  const nameNorm = normalize(nomeTxt);

  let best: { codigo: string; score: number; origem: string } | null = null;

  filiais.forEach((f) => {
    const codNorm = normalize(String(f.codigo || ''));
    const nomeNorm = normalize(String(f.nome || ''));
    let score = 0;
    let origem = '';
    if (codeNorm && codNorm === codeNorm) {
      score = 3;
      origem = 'codigo';
    } else if (codeNorm && codNorm.includes(codeNorm)) {
      score = 2;
      origem = 'codigo-parcial';
    }
    if (nameNorm && (nomeNorm.includes(nameNorm) || nameNorm.includes(nomeNorm))) {
      score = Math.max(score, 2);
      origem = origem || 'nome';
    }
    if (score > 0 && (!best || score > best.score)) {
      best = { codigo: String(f.codigo || '').trim() || fallback, score, origem };
    }
  });

  const selected = best as { codigo: string; origem: string } | null;
  if (selected) return { codigo: selected.codigo, origem: selected.origem };
  return { codigo: fallback, origem: 'fallback' };
}

function isRateioFilial(value: string): boolean {
  const key = normalizeKey(value || '');
  return key.includes('rateio');
}

function getFiliaisForMapping(): Array<{ codigo: string; nome: string }> {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFiliais = ss.getSheetByName(SHEET_REF_FILIAIS);
  if (!sheetFiliais) return [];
  ensureRefFiliaisSchema(sheetFiliais);
  const data = sheetFiliais.getDataRange().getValues().slice(1);
  return data
    .filter((row) => row[0])
    .map((row) => ({
      codigo: String(row[0]),
      nome: String(row[1] || ''),
      }));
}

function getRateioTargets(): Array<{ codigo: string; nome: string }> {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFiliais = ss.getSheetByName(SHEET_REF_FILIAIS);
  if (!sheetFiliais) return [];
  ensureRefFiliaisSchema(sheetFiliais);
  const data = sheetFiliais.getDataRange().getValues().slice(1);
  const targets = data
    .filter((row) => row[0])
    .map((row) => ({
      codigo: String(row[0]),
      nome: String(row[1] || ''),
      ativo: row[3] !== false && String(row[3]).toUpperCase() !== 'FALSE'
    }))
    .filter((row) => row.ativo && !isRateioFilial(row.codigo) && !isRateioFilial(row.nome));
  return targets.map(t => ({ codigo: t.codigo, nome: t.nome }));
}

function splitRateio(valor: number, count: number): number[] {
  if (!count || count <= 1) return [Number(valor || 0)];
  const total = Number(valor || 0);
  const base = Math.round((total / count) * 100) / 100;
  const splits = Array(count).fill(base);
  const currentSum = Math.round(splits.reduce((sum, v) => sum + v, 0) * 100) / 100;
  const diff = Math.round((total - currentSum) * 100) / 100;
  splits[count - 1] = Math.round((splits[count - 1] + diff) * 100) / 100;
  return splits;
}

function findHeaderIndexByAliases(headers: string[], aliases: string[]): number {
  const normalized = headers.map(h => normalizeKey(h));
  const aliasSet = new Set(aliases.map(a => normalizeKey(a)));
  return normalized.findIndex(h => aliasSet.has(h));
}

function ensureLancamentosNumeroDocumentoColumn(sheet: GoogleAppsScript.Spreadsheet.Sheet): number {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
  let idx = findHeaderIndexByAliases(headers, [
    'n documento',
    'ndocumento',
    'numero documento',
    'num documento',
    'documento'
  ]);
  if (idx !== -1) return idx;
  sheet.insertColumnsAfter(lastCol, 1);
  sheet.getRange(1, lastCol + 1).setValue('N Documento');
  return lastCol;
}

function isPagoStatus(status: string): boolean {
  return ['PAGO', 'PAGA', 'RECEBIDO', 'RECEBIDA'].includes((status || '').toUpperCase());
}

function buildRateioPercentuais(lancamentosMes: any[]): Array<{ filial: string; percentual: number }> {
  const receitasRecebidas = lancamentosMes.filter(l => l.tipo === 'RECEITA' && isPagoStatus(l.status) && !isRateioFilial(l.filial));
  const porFilial: Record<string, number> = {};
  receitasRecebidas.forEach(r => {
    const filial = String(r.filial || '').trim();
    if (!filial) return;
    porFilial[filial] = (porFilial[filial] || 0) + Number(r.valorLiquido || r.valor || 0);
  });
  let total = Object.values(porFilial).reduce((sum, v) => sum + v, 0);

  const filiaisBase = Object.keys(porFilial);
  if (total <= 0) {
    const fallback = Array.from(new Set(lancamentosMes.map(l => String(l.filial || '').trim())))
      .filter(f => f && !isRateioFilial(f));
    const baseFiliais = fallback.length ? fallback : filiaisBase;
    if (!baseFiliais.length) return [];
    const pct = 1 / baseFiliais.length;
    return baseFiliais.map(f => ({ filial: f, percentual: pct }));
  }

  return filiaisBase.map(f => ({ filial: f, percentual: porFilial[f] / total }));
}

function aplicarRateioDespesas(lancamentosMes: any[]): any[] {
  const rateioDespesas = lancamentosMes.filter(l => l.tipo === 'DESPESA' && isRateioFilial(l.filial));
  if (!rateioDespesas.length) return lancamentosMes;

  const percentuais = buildRateioPercentuais(lancamentosMes);
  if (!percentuais.length) return lancamentosMes;

  const semRateio = lancamentosMes.filter(l => !(l.tipo === 'DESPESA' && isRateioFilial(l.filial)));
  const alocados: any[] = [];

  rateioDespesas.forEach((d) => {
    const total = Number(d.valorLiquido || d.valor || 0);
    if (!Number.isFinite(total) || total === 0) return;
    const valores = percentuais.map(p => total * p.percentual);
    const valoresAjustados = splitRateio(total, percentuais.length);
    percentuais.forEach((p, idx) => {
      const valor = valoresAjustados[idx] ?? valores[idx] ?? 0;
      alocados.push({
        ...d,
        filial: p.filial,
        valorLiquido: valor,
        valorBruto: valor,
        valor: valor,
        descricao: `${d.descricao || ''} | Rateio ${Math.round(p.percentual * 100)}%`,
        observacoes: `${d.observacoes || ''} Rateio origem: ${d.filial || ''}`.trim(),
      });
    });
  });

  return semRateio.concat(alocados);
}

function getLancamentosMesRateados(
  lancamentos: any[],
  mes: number,
  ano: number,
  canal?: string
): any[] {
  const base = lancamentos.filter(l => {
    const data = new Date(l.dataCompetencia);
    const mesLanc = data.getMonth() + 1;
    const anoLanc = data.getFullYear();
    const matchPeriodo = mesLanc === mes && anoLanc === ano;
    const matchCanal = !canal || l.canal === canal;
    return matchPeriodo && matchCanal;
  });
  return aplicarRateioDespesas(base);
}

export function previewContasPagasTxt(content: string): {
  success: boolean;
  message: string;
  total?: number;
  preview?: ParsedContaPaga[];
  unmapped?: number;
} {
  try {
    const denied = requirePermission('importarArquivos', 'preview contas pagas');
    if (denied) return { success: false, message: denied.message };
    if (!content || !String(content).trim()) {
      return { success: false, message: 'Arquivo vazio' };
    }
    const filiais = getFiliaisForMapping();
    const parsed = parseContasPagasTxt(content, filiais);
    if (!parsed.length) {
      return { success: false, message: 'Nenhuma linha valida para importar' };
    }
    const unmapped = parsed.filter(p => p.filialMapeada === p.filialOriginal).length;
    return {
      success: true,
      message: 'Preview gerado',
      total: parsed.length,
      preview: parsed.slice(0, 25),
      unmapped,
    };
  } catch (error: any) {
    return { success: false, message: error.message };
  }
}

export function importarContasPagasTxt(content: string, fileName?: string): { success: boolean; message: string } {
  try {
    const denied = requirePermission('importarArquivos', 'importar contas pagas');
    if (denied) return denied;

    if (!content || !String(content).trim()) {
      return { success: false, message: 'Arquivo vazio' };
    }

    const filiais = getFiliaisForMapping();
    const parsed = parseContasPagasTxt(content, filiais);
    if (!parsed.length) {
      return { success: false, message: 'Nenhuma linha valida para importar' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lan?amentos n?o encontrada');

    ensureLancamentosNumeroDocumentoColumn(sheet);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0] || [];
    const idCol = findHeaderIndexByAliases(headerRow, ['id', 'codigo']);
    const dataPagCol = findHeaderIndexByAliases(headerRow, [
      'data pagamento',
      'dt pagamento',
      'data pgto',
      'dt pgto',
      'data baixa',
      'dt baixa',
      'baixa'
    ]);
    const contaCol = findHeaderIndexByAliases(headerRow, ['conta contabil', 'conta']);
    const valorCol = findHeaderIndexByAliases(headerRow, ['valor liquido', 'valor']);
    const filialCol = findHeaderIndexByAliases(headerRow, ['filial']);
    const descCol = findHeaderIndexByAliases(headerRow, ['descricao', 'historico']);
    const tipoCol = findHeaderIndexByAliases(headerRow, ['tipo']);

    if ([idCol, dataPagCol, contaCol, valorCol, filialCol, descCol, tipoCol].some(v => v == -1)) {
      throw new Error(`Cabecalhos obrigatorios nao encontrados em lancamentos: ${headerRow.join(' | ')}`);
    }

    const lastRow = sheet.getLastRow();
    const numRows = lastRow > 1 ? lastRow - 1 : 0;
    const existingValues = numRows > 0
      ? sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getDisplayValues()
      : [];

    const existingKeys = new Set<string>();
    existingValues.forEach(row => {
      const key = buildImportKey([
        row[dataPagCol ?? 0],
        row[contaCol ?? 0],
        row[valorCol ?? 0],
        row[filialCol ?? 0],
        row[descCol ?? 0],
        row[tipoCol ?? 0],
      ]);
      if (key) existingKeys.add(key);
    });

    const rowsToAppend: any[][] = [];
    let skipped = 0;

    parsed.forEach((item) => {
      if (item.tipo !== 'DESPESA') return;
      const status = item.tipo === 'DESPESA' ? 'PAGA' : 'RECEBIDA';
      const key = buildImportKey([
        item.dataPagamento,
        item.contaContabil,
        Number(item.valor).toFixed(2),
        item.filialMapeada,
        item.descricao,
        item.tipo,
      ]);
      if (existingKeys.has(key)) {
        skipped++;
        return;
      }
      existingKeys.add(key);

      const id = `CP-ERP-${Utilities.getUuid()}`;
      const obs = item.rateio
        ? `Importado de ${fileName || 'TXT'} | Filial origem: ${item.filialOriginal} | Rateio`
        : `Importado de ${fileName || 'TXT'} | Filial origem: ${item.filialOriginal}`;

      const row = [
        sanitizeSheetString(id),
        sanitizeSheetString(item.dataCompetencia),
        sanitizeSheetString(item.dataVencimento),
        sanitizeSheetString(item.dataPagamento),
        sanitizeSheetString(item.tipo),
        sanitizeSheetString(item.filialMapeada),
        '',
        '',
        sanitizeSheetString(item.contaContabil),
        '',
        '',
        sanitizeSheetString(item.descricao),
        item.valor,
        0,
        0,
        0,
        item.valor,
        sanitizeSheetString(status),
        '',
        sanitizeSheetString('ERP_TXT'),
        sanitizeSheetString(obs),
        sanitizeSheetString(item.numeroDocumento || ''),
      ];

      rowsToAppend.push(row);
    });

    if (!rowsToAppend.length) {
      return { success: false, message: 'Nenhuma linha nova para importar' };
    }

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);

    invalidateLancamentosCache();
    appendAuditLog('importarContasPagasTxt', { imported: rowsToAppend.length, skipped }, true);
    return { success: true, message: `${rowsToAppend.length} contas pagas importadas (${skipped} duplicadas)` };
  } catch (error: any) {
    appendAuditLog('importarContasPagasTxt', { fileName }, false, error?.message);
    return { success: false, message: error.message };
  }
}

export function importarItau(
  rows: Array<any>,
  meta?: { modelo?: string; filialFc?: string; conta?: string }
): { success: boolean; message: string; imported?: number } {
  const denied = requirePermission('importarArquivos', 'importar Itau');
  if (denied) return denied;

  const now = new Date().toISOString();
  const payload = Array.isArray(rows) ? rows : [];
  const existing = getImportItauRows();
  const existingKeys = new Set(
    existing.map((r) => buildImportKey([r.data, r.valor, r.lancamento, r.conta, r.cpfCnpj]))
  );

  const mapped = payload
    .filter((r) => r)
    .map((r) => ({
      data: normalizeDateInput(r.data),
      lancamento: sanitizeSheetString(r.lancamento || ''),
      agenciaOrigem: sanitizeSheetString(r.agenciaOrigem || ''),
      razaoSocial: sanitizeSheetString(r.razaoSocial || ''),
      cpfCnpj: sanitizeSheetString(r.cpfCnpj || ''),
      valor: parseMoneyInput(r.valor),
      saldo: parseMoneyInput(r.saldo),
      conta: sanitizeSheetString(r.conta || meta?.conta || ''),
      filialFc: sanitizeSheetString(r.filialFc || meta?.filialFc || ''),
      modelo: sanitizeSheetString(r.modelo || meta?.modelo || ''),
    }));

  const movimentos = mapped.filter((r) => isItauMovement(r.lancamento));
  const skippedNoMov = Math.max(0, mapped.length - movimentos.length);

  const values = movimentos
    .filter((r) => {
      const key = buildImportKey([r.data, r.valor, r.lancamento, r.conta, r.cpfCnpj]);
      if (!key || existingKeys.has(key)) return false;
      existingKeys.add(key);
      return true;
    })
    .map((r) => ([
      r.data,
      r.lancamento,
      r.agenciaOrigem,
      r.razaoSocial,
      r.cpfCnpj,
      r.valor,
      r.saldo,
      r.conta,
      r.filialFc,
      r.modelo,
      now,
    ]));

  const imported = values.length;
  const total = payload.length;
  const skippedDup = Math.max(0, movimentos.length - imported);

  if (!values.length) {
    if (payload.length > 0) {
      return { success: false, message: 'Nenhuma linha importada (todas duplicadas, saldo ou invalidas)' };
    }
    return { success: false, message: 'Nenhuma linha valida para importar' };
  }

  appendRows(SHEET_TB_IMPORT_ITAU, values);
  appendAuditLog('importarItau', { imported, skippedDup, skippedNoMov }, true);
  cacheRemoveNamespace(CacheNamespace.CONCILIACAO, CacheScope.SCRIPT);

  const notes = [] as string[];
  if (skippedDup) notes.push(`ignoradas ${skippedDup} duplicadas`);
  if (skippedNoMov) notes.push(`ignoradas ${skippedNoMov} de saldo`);
  const suffix = notes.length ? ` (${notes.join(', ')})` : '';

  return {
    success: true,
    message: `Importado ${imported} linhas Itau${suffix}`,
    imported,
  };
}

export function importarSieg(
  rows: Array<any>,
  meta?: { filialFc?: string }
): { success: boolean; message: string; imported?: number } {
  const denied = requirePermission('importarArquivos', 'importar SIEG');
  if (denied) return denied;

  const now = new Date().toISOString();
  createSheetIfNotExists(SHEET_REF_FILIAIS, [
    'Código', 'Nome', 'CNPJ', 'Ativo', 'Filial SIEG Relatorio', 'Filial SIEG Contabilidade',
  ]);
  const relations = getSheetValues(SHEET_REF_FILIAIS, { skipHeader: true })
    .map((r) => ({
      filialFc: String(r[0] || ''),
      filialSiegRelatorio: String(r[4] || ''),
      filialSiegContabil: String(r[5] || ''),
      ativa: r[3] !== false && String(r[3] || '').toUpperCase() !== 'FALSE',
    }));

  const payload = Array.isArray(rows) ? rows : [];
  const existing = getImportSiegRows();
  const existingKeys = new Set(
    existing.map((r) => buildImportKey([r.chaveNfe || r.numNfe, r.valor, r.dataEmissao, r.cnpjEmit]))
  );

  const values = payload
    .filter((r) => r)
    .map((r) => {
      const codigoFilial = String(r.codigoFilial || r.codigo_filial || '');
      const filialFc = resolveFilialFcFromRelations(codigoFilial, r.filialFc || meta?.filialFc || '', relations);
      return {
        numNfe: sanitizeSheetString(r.numNfe || r.num_nfe || ''),
        valor: parseMoneyInput(r.valor),
        dataEmissao: normalizeDateInput(r.dataEmissao),
        cnpjEmit: sanitizeSheetString(r.cnpjEmit || ''),
        nomeFantEmit: sanitizeSheetString(r.nomeFantEmit || ''),
        razaoEmit: sanitizeSheetString(r.razaoEmit || ''),
        cnpjDest: sanitizeSheetString(r.cnpjDest || ''),
        nomeFantDest: sanitizeSheetString(r.nomeFantDest || ''),
        razaoDest: sanitizeSheetString(r.razaoDest || ''),
        dataEnvioCofre: normalizeDateInput(r.dataEnvioCofre),
        chaveNfe: sanitizeSheetString(r.chaveNfe || ''),
        tags: sanitizeSheetString(r.tags || ''),
        codigoEvento: sanitizeSheetString(r.codigoEvento || ''),
        tipoEvento: sanitizeSheetString(r.tipoEvento || ''),
        status: sanitizeSheetString(r.status || ''),
        danfe: sanitizeSheetString(r.danfe || ''),
        xml: sanitizeSheetString(r.xml || ''),
        codigoFilial: sanitizeSheetString(codigoFilial),
        filialFc: sanitizeSheetString(filialFc),
      };
    })
    .filter((r) => {
      const key = buildImportKey([r.chaveNfe || r.numNfe, r.valor, r.dataEmissao, r.cnpjEmit]);
      if (!key || existingKeys.has(key)) return false;
      existingKeys.add(key);
      return true;
    })
    .map((r) => ([
      r.numNfe,
      r.valor,
      r.dataEmissao,
      r.cnpjEmit,
      r.nomeFantEmit,
      r.razaoEmit,
      r.cnpjDest,
      r.nomeFantDest,
      r.razaoDest,
      r.dataEnvioCofre,
      r.chaveNfe,
      r.tags,
      r.codigoEvento,
      r.tipoEvento,
      r.status,
      r.danfe,
      r.xml,
      r.codigoFilial,
      r.filialFc,
      now,
    ]));

  const imported = values.length;
  const total = payload.length;
  const skipped = Math.max(0, total - imported);
  if (!values.length) {
    if (payload.length > 0) {
      return { success: false, message: 'Nenhuma linha importada (todas duplicadas ou invalidas)' };
    }
    return { success: false, message: 'Nenhuma linha valida para importar' };
  }

  appendRows(SHEET_TB_IMPORT_SIEG, values);
  appendAuditLog('importarSieg', { imported, skipped }, true);
  cacheRemoveNamespace(CacheNamespace.CONCILIACAO, CacheScope.SCRIPT);
  return {
    success: true,
    message: `Importado ${imported} linhas SIEG${skipped ? ` (ignoradas ${skipped} duplicadas)` : ''}`,
    imported,
  };
}

const COMPARATIVO_CACHE_TTL_SECONDS = 60;

type ComparativoParams = {
  tipo?: string;
  year?: number;
  month?: number;
  page?: number;
  pageSize?: number;
};

export function getComparativoData(
  tipoOrParams: string | ComparativoParams,
  maybeParams?: ComparativoParams
): {
  success: boolean;
  message?: string;
  tipo?: string;
  stats?: { total: number; auto: number; pendente: number; semMatch: number; duplicado: number };
  counts?: { fcTotal: number; itauTotal: number; siegTotal: number };
  items?: any[];
  total?: number;
  page?: number;
  pageSize?: number;
} {
  const denied = requireAnyPermission<{ success: boolean; message: string }>(
    ['visualizarRelatorios', 'importarArquivos'],
    'carregar comparativo'
  );
  if (denied) return denied;

  const params =
    typeof tipoOrParams === 'object' && tipoOrParams !== null ? tipoOrParams : (maybeParams || {});
  const rawTipo = typeof tipoOrParams === 'string' ? tipoOrParams : params?.tipo;
  const normalizedTipo = String(rawTipo || '').toUpperCase();
  if (!['PAGAR', 'RECEBER'].includes(normalizedTipo)) {
    return { success: false, message: 'Tipo invalido para comparativo' };
  }

  try {
    const now = new Date();
    const year = Number(params?.year) || now.getFullYear();
    const month = Number(params?.month) || now.getMonth() + 1;
    const pageSize = Math.max(20, Math.min(200, Number(params?.pageSize) || 100));
    const page = Math.max(1, Number(params?.page) || 1);
    const monthKey = String(month).padStart(2, '0');
    const cacheKey = `comparativo:${normalizedTipo}:${year}-${monthKey}:p${page}:s${pageSize}`;
    const cached = cacheGet<any>(CacheNamespace.CONCILIACAO, cacheKey, CacheScope.SCRIPT);
    if (cached) return cached;

    const isInMonth = (dateValue: any) => {
      const norm = normalizeDateInput(dateValue);
      if (!norm) return false;
      const parts = String(norm).split('-');
      const y = Number(parts[0] || 0);
      const m = Number(parts[1] || 0);
      return y === year && m === month;
    };

    const fcRows = getImportFcRows(normalizedTipo).filter((r) => isInMonth(r.dataEmissao));
    const itauRows = getImportItauRows().filter((r) => isInMonth(r.data));
    const siegRows = normalizedTipo === 'PAGAR' ? getImportSiegRows().filter((r) => isInMonth(r.dataEmissao)) : [];

    const mapByKey = (rows: any[], getKey: (r: any) => string) => {
      const map = new Map<string, any[]>();
      rows.forEach((r) => {
        const key = getKey(r);
        if (!key) return;
        const list = map.get(key) || [];
        list.push(r);
        map.set(key, list);
      });
      return map;
    };

    const itauMap = mapByKey(itauRows, (r) => buildMatchKey(r.data, r.valor));
    const siegMap = mapByKey(siegRows, (r) => buildMatchKey(r.dataEmissao, r.valor));

    const items: any[] = [];
    let auto = 0;
    let pendente = 0;
    let semMatch = 0;
    let duplicado = 0;
    const total = fcRows.length;
    const offset = (page - 1) * pageSize;
    let index = 0;

    fcRows.forEach((fc) => {
      const key = buildMatchKey(fc.dataEmissao, fc.valor);
      const bancoCandidates = key ? (itauMap.get(key) || []) : [];
      const nfeCandidates = normalizedTipo === 'PAGAR' && key ? (siegMap.get(key) || []) : [];

      const banco = bancoCandidates.length === 1 ? bancoCandidates[0] : null;
      const nfe = nfeCandidates.length === 1 ? nfeCandidates[0] : null;

      let status = 'PENDENTE';
      if (!key || (bancoCandidates.length === 0 && nfeCandidates.length === 0)) {
        status = 'SEM_MATCH';
      } else if (bancoCandidates.length > 1 || nfeCandidates.length > 1) {
        status = 'DUPLICADO';
      } else if (normalizedTipo === 'PAGAR') {
        status = banco && nfe ? 'AUTO' : 'PENDENTE';
      } else {
        status = banco ? 'AUTO' : 'PENDENTE';
      }

      if (status === 'AUTO') auto += 1;
      else if (status === 'SEM_MATCH') semMatch += 1;
      else if (status === 'DUPLICADO') duplicado += 1;
      else pendente += 1;

      if (index >= offset && items.length < pageSize) {
        items.push({
          key,
          status,
          fc,
          nfe,
          banco,
          nfeMatches: nfeCandidates.length,
          bancoMatches: bancoCandidates.length,
          nfeCandidates: nfeCandidates.length > 1 ? nfeCandidates.slice(0, 5) : [],
          bancoCandidates: bancoCandidates.length > 1 ? bancoCandidates.slice(0, 5) : [],
          diffs: {
            nfeData: nfe ? normalizeDateInput(nfe.dataEmissao) !== normalizeDateInput(fc.dataEmissao) : false,
            nfeValor: nfe ? Math.abs(parseMoneyInput(nfe.valor) - parseMoneyInput(fc.valor)) > 0.01 : false,
            bancoData: banco ? normalizeDateInput(banco.data) !== normalizeDateInput(fc.dataEmissao) : false,
            bancoValor: banco ? Math.abs(parseMoneyInput(banco.valor) - parseMoneyInput(fc.valor)) > 0.01 : false,
          },
        });
      }

      index += 1;
    });

    const stats = {
      total,
      auto,
      pendente,
      semMatch,
      duplicado,
    };

    const counts = {
      fcTotal: fcRows.length,
      itauTotal: itauRows.length,
      siegTotal: siegRows.length,
    };

    const result = {
      success: true,
      tipo: normalizedTipo,
      stats,
      counts,
      items,
      total,
      page,
      pageSize,
    };
    cacheSet(CacheNamespace.CONCILIACAO, cacheKey, result, COMPARATIVO_CACHE_TTL_SECONDS, CacheScope.SCRIPT);
    return result;
  } catch (error: any) {
    return { success: false, message: `Erro ao montar comparativo: ${error?.message || error}` };
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

const LANCAMENTOS_CACHE_KEY = 'all';
const EXTRATOS_CACHE_KEY = 'all';
const DATA_CACHE_TTL_SECONDS = 30;

function invalidateLancamentosCache(): void {
  cacheRemove(CacheNamespace.LANCAMENTOS, LANCAMENTOS_CACHE_KEY, CacheScope.SCRIPT);
}

function invalidateExtratosCache(): void {
  cacheRemove(CacheNamespace.EXTRATOS, EXTRATOS_CACHE_KEY, CacheScope.SCRIPT);
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
    'Observações',
    'N Documento'
  ]);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
  if (sheet) {
    ensureLancamentosNumeroDocumentoColumn(sheet);
  }


  const cached = cacheGet<any[]>(CacheNamespace.LANCAMENTOS, LANCAMENTOS_CACHE_KEY, CacheScope.SCRIPT);
  if (cached) return cached;

  const data = getSheetValues(SHEET_TB_LANCAMENTOS);
  if (!data || data.length <= 1) {
    cacheSet(CacheNamespace.EXTRATOS, EXTRATOS_CACHE_KEY, [], DATA_CACHE_TTL_SECONDS, CacheScope.SCRIPT);
    return [];
  }

  const parsed = data.slice(1).map((row: any) => ({
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
    numeroDocumento: String(row[21] || ''),
  })).map(l => {
    const tipoNorm = String(l.tipo || '').toUpperCase();
    if (tipoNorm === 'AP') l.tipo = 'DESPESA';
    else if (tipoNorm === 'AR') l.tipo = 'RECEITA';
    return l;
  });
  cacheSet(CacheNamespace.LANCAMENTOS, LANCAMENTOS_CACHE_KEY, parsed, DATA_CACHE_TTL_SECONDS, CacheScope.SCRIPT);
  return parsed;
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

  const cached = cacheGet<any[]>(CacheNamespace.EXTRATOS, EXTRATOS_CACHE_KEY, CacheScope.SCRIPT);
  if (cached) return cached;

  const data = getSheetValues(SHEET_TB_EXTRATOS);
  if (!data || data.length <= 1) {
    if (!isSeedDataEnabled()) {
      cacheSet(CacheNamespace.EXTRATOS, EXTRATOS_CACHE_KEY, [], DATA_CACHE_TTL_SECONDS, CacheScope.SCRIPT);
      return [];
    }
    const seed = [
      ['EXT-5001','2025-01-02','Recebimento cartão venda balcão',3200,'ENTRADA','BANCO_A','CC_MATRIZ','CONCILIADO','CR-2001','Pedido balcão','2025-01-03','' ],
      ['EXT-5002','2025-01-11','Pagamento fornecedor matéria-prima',-1500,'SAIDA','BANCO_A','CC_MATRIZ','CONCILIADO','CP-1001','Pagto lote A','2025-01-11','' ],
      ['EXT-5003','2025-01-15','Taxa bancária jan',-25,'SAIDA','BANCO_A','CC_MATRIZ','PENDENTE','','Tarifa débito','2025-01-15','' ],
      ['EXT-5004','2025-01-16','Recebimento boleto convênio',2100,'ENTRADA','BANCO_A','CC_MATRIZ','PENDENTE','CR-2003','Convênio varejo','2025-01-16','' ],
    ];
    appendRows(SHEET_TB_EXTRATOS, seed);
    const seeded = seed.map(r => ({
      id: r[0], data: r[1], descricao: r[2], valor: r[3], tipo: r[4], banco: r[5], conta: r[6],
      statusConciliacao: r[7], idLancamento: r[8], observacoes: r[9], importadoEm: r[10],
    }));
    cacheSet(CacheNamespace.EXTRATOS, EXTRATOS_CACHE_KEY, seeded, DATA_CACHE_TTL_SECONDS, CacheScope.SCRIPT);
    return seeded;
  }

  const parsed = data.slice(1).map((row: any) => ({
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
  cacheSet(CacheNamespace.EXTRATOS, EXTRATOS_CACHE_KEY, parsed, DATA_CACHE_TTL_SECONDS, CacheScope.SCRIPT);
  return parsed;
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

export function getDREMensal(mes: number, ano: number, filial?: string, canal?: string): any {
  enforcePermission('visualizarRelatorios', 'carregar DRE');
  const cacheKey = `mensal:${ano}-${mes}:${filial || 'all'}:${canal || 'all'}`;
  return cacheGetOrLoad(CacheNamespace.DRE, cacheKey, () => {
  try {
    const lancamentos = getLancamentosFromSheet();
    const planoMap = getPlanoContasMap();

    // Filtrar por per?odo e filial
    const lancamentosMesBase = getLancamentosMesRateados(lancamentos, mes, ano, canal);
    const lancamentosMes = lancamentosMesBase.filter(l => !filial || l.filial === filial);


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
        filial: filial || 'Consolidado',
        canal: canal || 'Todos'
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

export function getFluxoCaixaMensal(mes: number, ano: number, filial?: string, canal?: string, saldoInicial?: number): any {
  enforcePermission('visualizarRelatorios', 'carregar fluxo de caixa');
  const cacheKey = `mensal:${ano}-${mes}:${filial || 'all'}:${canal || 'all'}:${Number(saldoInicial) || 0}`;
  return cacheGetOrLoad(CacheNamespace.DFC, cacheKey, () => {
  try {
    const lancamentos = getLancamentosFromSheet();

    // Filtrar por per?odo e filial
    const lancamentosMesBase = getLancamentosMesRateados(lancamentos, mes, ano, canal);
    const lancamentosMes = lancamentosMesBase.filter(l => !filial || l.filial === filial);


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
        filial: filial || 'Consolidado',
        canal: canal || 'Todos'
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

export function getKPIsMensal(mes: number, ano: number, filial?: string, canal?: string): any {
  enforcePermission('visualizarRelatorios', 'carregar KPIs');
  const cacheKey = `mensal:${ano}-${mes}:${filial || 'all'}:${canal || 'all'}`;
  return cacheGetOrLoad(CacheNamespace.KPI, cacheKey, () => {
  try {
    const lancamentos = getLancamentosFromSheet();

    // Filtrar lançamentos do mês atual
    const lancamentosMesBase = getLancamentosMesRateados(lancamentos, mes, ano, canal);
    const lancamentosMes = lancamentosMesBase.filter(l => !filial || l.filial === filial);

    const receitasRecebidasAll = lancamentosMesBase.filter(l => l.tipo === 'RECEITA' && isPagoStatus(l.status) && !isRateioFilial(l.filial));
    const receitaPorFilialMap: Record<string, number> = {};
    receitasRecebidasAll.forEach(r => {
      const f = String(r.filial || '').trim();
      if (!f) return;
      receitaPorFilialMap[f] = (receitaPorFilialMap[f] || 0) + Number(r.valorLiquido || r.valor || 0);
    });
    const receitaTotalAll = Object.values(receitaPorFilialMap).reduce((sum, v) => sum + v, 0);
    const receitaPorFilial = Object.keys(receitaPorFilialMap).map(f => ({
      filial: f,
      receita: receitaPorFilialMap[f],
      percentual: receitaTotalAll > 0 ? (receitaPorFilialMap[f] / receitaTotalAll) : 0
    })).sort((a, b) => b.receita - a.receita);


    // Filtrar lançamentos do mês anterior
    const dataAnterior = new Date(ano, mes - 2, 1); // mes - 2 porque JavaScript months são 0-indexed
    const mesAnterior = dataAnterior.getMonth() + 1;
    const anoAnterior = dataAnterior.getFullYear();

    const lancamentosMesAnteriorBase = getLancamentosMesRateados(lancamentos, mesAnterior, anoAnterior, canal);
    const lancamentosMesAnterior = lancamentosMesAnteriorBase.filter(l => !filial || l.filial === filial);


    // Calcular DRE do mês atual e anterior
    const dreAtual = getDREMensal(mes, ano, filial, canal);
    const dreAnterior = getDREMensal(mesAnterior, anoAnterior, filial, canal);
    const fcAtual = getFluxoCaixaMensal(mes, ano, filial, canal);
    const fcAnterior = getFluxoCaixaMensal(mesAnterior, anoAnterior, filial, canal);

    const dataAnteriorAnterior = new Date(ano, mes - 3, 1);
    const mesAnteriorAnterior = dataAnteriorAnterior.getMonth() + 1;
    const anoAnteriorAnterior = dataAnteriorAnterior.getFullYear();
    const dreAnteriorAnterior = getDREMensal(mesAnteriorAnterior, anoAnteriorAnterior, filial, canal);

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
    const contasReceberPrev = lancamentosMesAnterior.filter(l => l.tipo === 'RECEITA' && l.status === 'PENDENTE');
    const contasPagarPrev = lancamentosMesAnterior.filter(l => l.tipo === 'DESPESA' && l.status === 'PENDENTE');
    const ativoCirculante = sumValues(contasReceber) + fcAtual.valores.saldoFinal;
    const passivoCirculante = sumValues(contasPagar);
    const liquidezCorrente = passivoCirculante > 0 ? ativoCirculante / passivoCirculante : 0;
    const ativoCirculantePrev = sumValues(contasReceberPrev) + fcAnterior.valores.saldoFinal;
    const passivoCirculantePrev = sumValues(contasPagarPrev);
    const liquidezCorrentePrev = passivoCirculantePrev > 0 ? ativoCirculantePrev / passivoCirculantePrev : 0;

    const saldoCaixa = fcAtual.valores.saldoFinal;
    const burnRate = Math.abs(dreAtual.valores.lucroLiquido < 0 ? dreAtual.valores.lucroLiquido : 0);
    const runway = burnRate > 0 ? saldoCaixa / burnRate : 999;
    const saldoCaixaPrev = fcAnterior.valores.saldoFinal;
    const burnRatePrev = Math.abs(dreAnterior.valores.lucroLiquido < 0 ? dreAnterior.valores.lucroLiquido : 0);
    const runwayPrev = burnRatePrev > 0 ? saldoCaixaPrev / burnRatePrev : 999;

    // KPIs de Crescimento
    const receitaAtual = dreAtual.valores.receitaLiquida;
    const receitaAnteriorVal = dreAnterior.valores.receitaLiquida;
    const crescimentoReceita = receitaAnteriorVal > 0
      ? ((receitaAtual - receitaAnteriorVal) / receitaAnteriorVal) * 100
      : 0;
    const receitaAnteriorAnteriorVal = dreAnteriorAnterior.valores.receitaLiquida;
    const crescimentoReceitaPrev = receitaAnteriorAnteriorVal > 0
      ? ((receitaAnteriorVal - receitaAnteriorAnteriorVal) / receitaAnteriorAnteriorVal) * 100
      : 0;

    const ticketMedio = receitas.length > 0 ? receitaAtual / receitas.length : 0;

    const referenciaAtual = new Date(ano, mes - 1, 1);
    const referenciaAnterior = new Date(anoAnterior, mesAnterior - 1, 1);

    const receitasVencidas = lancamentosMes.filter(l => {
      if (l.tipo !== 'RECEITA' || l.status !== 'PENDENTE') return false;
      const vencimento = new Date(l.dataVencimento);
      return vencimento < new Date();
    });
    const receitasVencidasPrev = lancamentosMesAnterior.filter(l => {
      if (l.tipo !== 'RECEITA' || l.status !== 'PENDENTE') return false;
      const vencimento = new Date(l.dataVencimento);
      return vencimento < referenciaAnterior;
    });
    const taxaInadimplencia = receitas.length > 0
      ? (receitasVencidas.length / receitas.length) * 100
      : 0;
    const taxaInadimplenciaPrev = receitasAnterior.length > 0
      ? (receitasVencidasPrev.length / receitasAnterior.length) * 100
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
    const receitasRecebidasPrev = lancamentosMesAnterior.filter(l =>
      l.tipo === 'RECEITA' && (l.status === 'PAGO' || l.status === 'RECEBIDO')
    );
    let prazoMedioRecebimentoPrev = 0;
    if (receitasRecebidasPrev.length > 0) {
      const prazosPrev = receitasRecebidasPrev.map(r => {
        const venc = new Date(r.dataVencimento);
        const pag = new Date(r.dataPagamento || r.dataCompetencia);
        return Math.floor((pag.getTime() - venc.getTime()) / (1000 * 60 * 60 * 24));
      });
      prazoMedioRecebimentoPrev = prazosPrev.reduce((a, b) => a + b, 0) / prazosPrev.length;
    }

    // KPIs Operacionais
    const despesasMarketing = despesas.filter(d => d.centroCusto === 'MKT' || d.centroCusto === 'COM');
    const despesasMarketingPrev = lancamentosMesAnterior.filter(d =>
      d.tipo === 'DESPESA' && (d.centroCusto === 'MKT' || d.centroCusto === 'COM')
    );
    const cac = receitas.length > 0 ? sumValues(despesasMarketing) / receitas.length : 0;
    const cacPrev = receitasAnterior.length > 0 ? sumValues(despesasMarketingPrev) / receitasAnterior.length : 0;

    const despOperacionaisPerc = dreAtual.valores.receitaLiquida > 0
      ? (dreAtual.valores.despesasOperacionais.total / dreAtual.valores.receitaLiquida) * 100
      : 0;
    const despOperacionaisPercPrev = dreAnterior.valores.receitaLiquida > 0
      ? (dreAnterior.valores.despesasOperacionais.total / dreAnterior.valores.receitaLiquida) * 100
      : 0;

    const breakEven = dreAtual.valores.margemBruta > 0
      ? dreAtual.valores.despesasOperacionais.total / (dreAtual.valores.margemBruta / dreAtual.valores.receitaLiquida)
      : 0;
    const breakEvenPrev = dreAnterior.valores.margemBruta > 0
      ? dreAnterior.valores.despesasOperacionais.total / (dreAnterior.valores.margemBruta / dreAnterior.valores.receitaLiquida)
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
    const despesasPagasPrev = lancamentosMesAnterior.filter(l =>
      l.tipo === 'DESPESA' && l.status === 'PAGO'
    );
    let prazoMedioPagamentoPrev = 0;
    if (despesasPagasPrev.length > 0) {
      const prazosPrev = despesasPagasPrev.map(d => {
        const venc = new Date(d.dataVencimento);
        const pag = new Date(d.dataPagamento || d.dataCompetencia);
        return Math.floor((pag.getTime() - venc.getTime()) / (1000 * 60 * 60 * 24));
      });
      prazoMedioPagamentoPrev = prazosPrev.reduce((a, b) => a + b, 0) / prazosPrev.length;
    }

    return {
      receitaPorFilial: { total: receitaTotalAll, itens: receitaPorFilial },
      periodo: {
        mes,
        ano,
        mesNome: getMesNome(mes),
        filial: filial || 'Consolidado',
        canal: canal || 'Todos'
      },
      rentabilidade: {
        receitaLiquida: {
          valor: dreAtual.valores.receitaLiquida
        },
        custos: {
          valor: dreAtual.valores.custos
        },
        ebitdaValor: {
          valor: dreAtual.valores.ebitda
        },
        lucroLiquidoValor: {
          valor: dreAtual.valores.lucroLiquido
        },
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
        capitalGiro: {
          valor: ativoCirculante - passivoCirculante
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
        despesasOperacionais: {
          valor: dreAtual.valores.despesasOperacionais.total
        },
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
          margemEbitda: dreAnterior.percentuais.ebitda,
          margemLiquida: dreAnterior.percentuais.lucroLiquido,
          roi: dreAnterior.valores.receitaLiquida > 0
            ? (dreAnterior.valores.lucroLiquido / dreAnterior.valores.receitaLiquida) * 100
            : 0,
          receitaLiquida: dreAnterior.valores.receitaLiquida,
          custos: dreAnterior.valores.custos,
          ebitdaValor: dreAnterior.valores.ebitda,
          lucroLiquidoValor: dreAnterior.valores.lucroLiquido,
          capitalGiro: ativoCirculantePrev - passivoCirculantePrev,
          liquidezCorrente: liquidezCorrentePrev,
          saldoCaixa: saldoCaixaPrev,
          burnRate: burnRatePrev,
          runway: runwayPrev,
          crescimentoReceita: crescimentoReceitaPrev,
          ticketMedio: receitasAnterior.length > 0 ? receitaAnteriorVal / receitasAnterior.length : 0,
          inadimplencia: taxaInadimplenciaPrev,
          prazoMedioRecebimento: prazoMedioRecebimentoPrev,
          cac: cacPrev,
          despOperacionaisReceita: despOperacionaisPercPrev,
          despesasOperacionais: dreAnterior.valores.despesasOperacionais.total,
          breakEven: breakEvenPrev,
          prazoMedioPagamento: prazoMedioPagamentoPrev
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
  perfil: 'ADMIN' | 'GESTOR' | 'OPERACIONAL' | 'CAIXA' | 'VISUALIZADOR';
  status: 'ATIVO' | 'INATIVO';
  canal?: string;
  ultimoAcesso?: string;
  permissoes?: {
    criarLancamentos: boolean;
    editarLancamentos: boolean;
    excluirLancamentos: boolean;
    aprovarPagamentos: boolean;
    visualizarRelatorios: boolean;
    gerenciarConfig: boolean;
    importarArquivos: boolean;
  };
}

export function getUsuarios(): Usuario[] {
  try {
    const email = getRequestingUserEmail();
    const requester = getUsuarioByEmail(email);
    if (!requester || requester.status !== 'ATIVO' || requester.perfil !== 'ADMIN') {
      return [];
    }

    ensureUsuariosSchema();
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
        canal: row[8] ? String(row[8]) : undefined,
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
    validateEnum(String(usuario?.perfil || ''), ['ADMIN', 'GESTOR', 'OPERACIONAL', 'CAIXA', 'VISUALIZADOR'], 'Perfil'),
      validateEnum(String(usuario?.status || ''), ['ATIVO', 'INATIVO'], 'Status')
    );
    if (!v.valid) return { success: false, message: v.errors.join('; ') };

    const perfil = normalizePerfil(usuario.perfil);
    if (perfil === 'CAIXA' && !String(usuario.canal || '').trim()) {
      return { success: false, message: 'Canal de venda obrigatório para o perfil Caixa' };
    }
    const permissoes = getPermissoesPadrao(perfil);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureUsuariosSchema();
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

    const rowData = [
      usuario.id || Utilities.getUuid(),
      sanitizeSheetString(usuario.email).toLowerCase(),
      sanitizeSheetString(usuario.nome),
      sanitizeSheetString(perfil),
      sanitizeSheetString(usuario.status),
      usuario.ultimoAcesso || '',
      JSON.stringify(permissoes),
      rowIndex === -1 ? new Date().toISOString() : data[rowIndex - 1][7],
      sanitizeSheetString(usuario.canal || '').toUpperCase()
    ];

    if (rowIndex > 0) {
      const previousEmail = String(data[rowIndex - 1][1] || '');
      // Atualizar existente
      sheet.getRange(rowIndex, 1, 1, 9).setValues([rowData]);
      appendAuditLog('salvarUsuario', { id: rowData[0], email: rowData[1], perfil: rowData[3], status: rowData[4] }, true);
      invalidateUserCache(previousEmail);
      invalidateUserCache(String(rowData[1] || ''));
      return { success: true, message: 'Usuário atualizado com sucesso!' };
    } else {
      // Criar novo
      sheet.appendRow(rowData);
      appendAuditLog('salvarUsuario', { id: rowData[0], email: rowData[1], perfil: rowData[3], status: rowData[4] }, true);
      invalidateUserCache(String(rowData[1] || ''));
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
        const deletedEmail = String(data[i][1] || '');
        sheet.deleteRow(i + 1);
        appendAuditLog('excluirUsuario', { id }, true);
        invalidateUserCache(deletedEmail);
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
    'SEM_ACESSO': {
      criarLancamentos: false,
      editarLancamentos: false,
      excluirLancamentos: false,
      aprovarPagamentos: false,
      visualizarRelatorios: false,
      gerenciarConfig: false,
      importarArquivos: false
    },
    'ADMIN': {
      criarLancamentos: true,
      editarLancamentos: true,
      excluirLancamentos: true,
      aprovarPagamentos: true,
      visualizarRelatorios: true,
      gerenciarConfig: true,
      importarArquivos: true
    },
    'GESTOR': {
      criarLancamentos: true,
      editarLancamentos: true,
      excluirLancamentos: false,
      aprovarPagamentos: true,
      visualizarRelatorios: true,
      gerenciarConfig: false,
      importarArquivos: true
    },
    'OPERACIONAL': {
      criarLancamentos: true,
      editarLancamentos: true,
      excluirLancamentos: false,
      aprovarPagamentos: false,
      visualizarRelatorios: false,
      gerenciarConfig: false,
      importarArquivos: true
    },
    'CAIXA': {
      criarLancamentos: false,
      editarLancamentos: false,
      excluirLancamentos: false,
      aprovarPagamentos: false,
      visualizarRelatorios: false,
      gerenciarConfig: false,
      importarArquivos: true
    },
    'VISUALIZADOR': {
      criarLancamentos: false,
      editarLancamentos: false,
      excluirLancamentos: false,
      aprovarPagamentos: false,
      visualizarRelatorios: true,
      gerenciarConfig: false,
      importarArquivos: false
    }
  };

  return permissoes[perfil] || permissoes['VISUALIZADOR'];
}
