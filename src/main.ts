/**
 * main.ts
 *
 * Entry point principal do Google Apps Script.
 * Centraliza roteamento web (doGet) e registros de menus.
 */

import { include } from './services/ui-service';
import { backupJob, installTriggers } from './services/scheduler-service';
import { exportToPDF } from './services/reporting-service';
import { getCurrentPeriod, Period } from './shared/types';
import { runCompleteSetup } from './setup-sheets';
import { setupAllSampleData, setupBulkSampleData } from './setup-sample-data';
import {
  getViewHtml,
  getReferenceData,
  getDashboardData,
  getContasPagar,
  pagarConta,
  pagarContasEmLote,
  cancelarContasEmLote,
  getContasReceber,
  receberConta,
  receberContasEmLote,
  cancelarContasReceberEmLote,
  salvarLancamento,
  atualizarLancamento,
  getLancamentoDetalhes,
  getConciliacaoData,
  getComparativoData,
  conciliarItens,
  desfazerConciliacao,
  conciliarAutomatico,
  importarFc,
  importarItau,
  importarSieg,
  getSheetData,
  toggleCentroCusto,
  seedPlanoContasFromList,
  salvarCentroCusto,
  excluirCentroCusto,
  salvarContaContabil,
  excluirConta,
  salvarCanal,
  excluirCanal,
  toggleCanal,
  salvarFilial,
  excluirFilial,
  toggleFilial,
  getDREMensal,
  getDREComparativo,
  getDREPorFilial,
  getFluxoCaixaMensal,
  getFluxoCaixaProjecao,
  getKPIsMensal,
  getCaixasConfig,
  salvarCaixasConfig,
  getCaixasData,
  getCaixaMovimentos,
  salvarCaixa,
  excluirCaixa,
  salvarCaixaMovimento,
  excluirCaixaMovimento,
  getCaixaRelatorio,
  uploadCaixaArquivo,
  salvarCaixaTipo,
  excluirCaixaTipo,
  toggleCaixaTipo,
  getUsuarios,
  getCurrentUserInfo,
  logClientError,
  logServerException,
  getAuditLogEntries,
  salvarUsuario,
  excluirUsuario,
  setRequestContext,
  clearRequestContext,
  logEndpointTiming,
  runSmokeTests,
  getAdminDiagnostics,
  setAdminFlag,
  clearCaches,
} from './services/webapp-service';

const DEPLOY_LABEL = 'v125 - datas e logo caixa';

function isDebugApiEnabled(): boolean {
  return PropertiesService.getScriptProperties().getProperty('ENABLE_DEBUG_API') === 'true';
}

/**
 * Função doGet - Entry point para web app
 *
 * IMPORTANTE: Esta função é chamada automaticamente pelo Apps Script
 * quando o usuário acessa a URL do web app
 *
 * Serve a aplicação web standalone (não modal do Sheets)
 */
function doGet(e: any): GoogleAppsScript.HTML.HtmlOutput {
  // Debug endpoint: retorna JSON para facilitar diagnósticos via URL
  // Ex.: /exec?api=dashboard
  try {
    const api = e?.parameter?.api;
    if (api && isDebugApiEnabled()) {
      const respond = (payload: unknown) =>
        ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
          ContentService.MimeType.JSON
        );

      switch (String(api)) {
        case 'dashboard':
          return respond(getDashboardData()) as any;
        case 'contas-pagar':
          return respond(getContasPagar()) as any;
        case 'contas-receber':
          return respond(getContasReceber()) as any;
        case 'conciliacao':
          return respond(getConciliacaoData()) as any;
        default:
          return respond({ error: `api desconhecida: ${api}` }) as any;
      }
    }
  } catch (err: any) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: err?.message || String(err) })
    ).setMimeType(ContentService.MimeType.JSON) as any;
  }

  const template = HtmlService.createTemplateFromFile('frontend/views/app');
  return template.evaluate()
    .setTitle(`Neoformula Finance - ${DEPLOY_LABEL}`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Hook para incluir arquivos parciais nos templates HTML
 *
 * IMPORTANTE: Esta função deve estar no escopo global para ser
 * acessível pelos templates HTML via <?!= include('filename') ?>
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function includeFile(filename: string): string {
  return include(filename);
}

/**
 * Função onOpen - Executada quando a planilha é aberta
 *
 * IMPORTANTE: Adiciona menu customizado no Google Sheets
 */
function onOpen(): void {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Neoformula Finance')
    .addItem('Abrir Dashboard', 'openDashboard')
    .addSeparator()
    .addItem('Novo Lançamento', 'openNovoLancamento')
    .addItem('Conciliação Bancária', 'openConciliacao')
    .addSeparator()
    .addItem('Atualizar Cache', 'refreshCache')
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Relatórios')
        .addItem('DRE', 'openDRE')
        .addItem('Fluxo de Caixa', 'openDFC')
        .addItem('KPIs', 'openKPI')
        .addSeparator()
        .addItem('Exportar PDF (mes atual)', 'exportCurrentReportPdf')
        .addItem('Exportar PDF (YYYY-MM)', 'exportReportPdfForPeriod')
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Administração')
        .addItem('Configurações', 'openConfiguracoes')
        .addItem('Instalar Triggers', 'setupTriggers')
        .addItem('Backup Agora', 'runBackupNow')
        .addSeparator()
        .addItem('⚙️ Setup da Planilha', 'runCompleteSetup')
        .addItem('📝 Criar Dados de Exemplo', 'setupAllSampleData')
        .addItem('📈 Criar Dados de Exemplo (Massa)', 'setupBulkSampleData')
        .addSeparator()
        .addItem('🌐 Abrir Web App', 'openWebApp')
    )
    .addToUi();
}

/**
 * Funções de menu - Abertura de views
 */

function openDashboard(): void {
  const html = HtmlService.createTemplateFromFile('frontend/views/dashboard')
    .evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Financeiro');
}

function openNovoLancamento(): void {
  const html = HtmlService.createTemplateFromFile('frontend/views/lancamentos')
    .evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Lançamentos');
}

function openConciliacao(): void {
  const html = HtmlService.createTemplateFromFile('frontend/views/conciliacao')
    .evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Conciliação Bancária');
}

function openDRE(): void {
  const html = HtmlService.createTemplateFromFile('frontend/views/dre')
    .evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'DRE');
}

function openDFC(): void {
  const html = HtmlService.createTemplateFromFile('frontend/views/dfc')
    .evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Fluxo de Caixa');
}

function openKPI(): void {
  const html = HtmlService.createTemplateFromFile('frontend/views/kpi')
    .evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'KPIs');
}

function openConfiguracoes(): void {
  const html = HtmlService.createTemplateFromFile('frontend/views/configuracoes')
    .evaluate()
    .setWidth(1000)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configurações');
}

/**
 * Funções de administração
 */

function refreshCache(): void {
  try {
    // TODO: Importar e chamar métodos de reload de cache
    // ConfigService.reloadCache();
    // reloadReferenceCache();

    SpreadsheetApp.getUi().alert(
      'Cache atualizado',
      'Cache de configurações e referências foi atualizado com sucesso.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(
      'Erro',
      `Erro ao atualizar cache: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function setupTriggers(): void {
  try {
    installTriggers();

    SpreadsheetApp.getUi().alert(
      'Triggers instalados',
      'Triggers de automação foram instalados com sucesso.\n\n' +
        '- Job diário: 6h\n' +
        '- Fechamento mensal: dia 1 às 8h',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(
      'Erro',
      `Erro ao instalar triggers: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function openWebApp(): void {
  const url = ScriptApp.getService().getUrl();
  const urlJs = JSON.stringify(url);
  const urlHtml = String(url)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
  const html = HtmlService.createHtmlOutput(
    `<html><body>
      <h2>Web App URL</h2>
      <p>Acesse a aplicação web no link abaixo:</p>
      <p><a href="${urlHtml}" target="_blank" rel="noopener noreferrer">${urlHtml}</a></p>
      <p><button onclick="window.open(${urlJs}, '_blank', 'noopener')">Abrir Web App</button></p>
      <br><p><small>Copie este link para acessar a aplicação de qualquer lugar.</small></p>
    </body></html>`
  ).setWidth(600).setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, 'URL da Web App');
}

function exportCurrentReportPdf(): void {
  const period = getCurrentPeriod();
  exportReportPdfForPeriodInternal(period);
}

function exportReportPdfForPeriod(): void {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Exportar PDF',
    'Informe o periodo no formato YYYY-MM (ex: 2025-03)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const input = String(response.getResponseText() || '').trim();
  const period = parsePeriodInput(input);
  if (!period) {
    ui.alert('Periodo invalido', 'Use o formato YYYY-MM, ex: 2025-03.', ui.ButtonSet.OK);
    return;
  }

  exportReportPdfForPeriodInternal(period);
}

function exportReportPdfForPeriodInternal(period: Period): void {
  try {
    const url = exportToPDF(period);
    const urlHtml = String(url)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');

    const html = HtmlService.createHtmlOutput(
      `<html><body>
        <h2>PDF exportado</h2>
        <p><a href="${urlHtml}" target="_blank" rel="noopener noreferrer">${urlHtml}</a></p>
      </body></html>`
    ).setWidth(520).setHeight(180);

    SpreadsheetApp.getUi().showModalDialog(html, 'Exportacao PDF');
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(
      'Erro',
      `Erro ao exportar PDF: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function parsePeriodInput(input: string): Period | null {
  const match = /^(\d{4})-(\d{2})$/.exec(input);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  if (!year || month < 1 || month > 12) return null;
  return { year, month };
}

function runBackupNow(): void {
  try {
    const result = backupJob();
    const folderUrl = String(result.folderUrl)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');

    const html = HtmlService.createHtmlOutput(
      `<html><body>
        <h2>Backup concluído</h2>
        <p>Arquivos gerados: <strong>${result.filesCreated}</strong></p>
        <p>Pasta: <a href="${folderUrl}" target="_blank" rel="noopener noreferrer">${folderUrl}</a></p>
        <p><small>Carimbo: ${result.stamp}</small></p>
      </body></html>`
    ).setWidth(520).setHeight(220);

    SpreadsheetApp.getUi().showModalDialog(html, 'Backup');
  } catch (error: any) {
    SpreadsheetApp.getUi().alert(
      'Erro',
      `Erro ao executar backup: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Exporta funções globais para o escopo do Apps Script
 *
 * IMPORTANTE: gas-webpack-plugin detecta estas declarações 'global'
 * e as move para o escopo global do Apps Script
 */
declare var global: any;

type RequestContext = { __ctx?: boolean; correlationId?: string; view?: string; url?: string } | null;

function wrapApi<T extends (...args: any[]) => any>(name: string, fn: T): (...args: any[]) => any {
  return (...rawArgs: any[]) => {
    const args = rawArgs.slice();
    const last = args.length ? args[args.length - 1] : null;
    const ctx: RequestContext =
      last && typeof last === 'object' && (last as any).__ctx ? (args.pop() as any) : null;

    const startedAt = Date.now();
    try {
      setRequestContext(ctx);
      return (fn as any)(...args);
    } catch (error: any) {
      const correlationId = ctx?.correlationId ? String(ctx.correlationId) : Utilities.getUuid();
      logServerException(name, { correlationId, view: ctx?.view, url: ctx?.url }, error);
      const message = error?.message ? String(error.message) : String(error);
      throw new Error(`${message} (ref: ${correlationId})`);
    } finally {
      const durationMs = Date.now() - startedAt;
      if (durationMs >= 2000) {
        try {
          logEndpointTiming(name, durationMs);
        } catch (_) {}
      }
      clearRequestContext();
    }
  };
}

global.doGet = doGet;
global.include = includeFile;
global.onOpen = onOpen;
global.openDashboard = openDashboard;
global.openNovoLancamento = openNovoLancamento;
global.openConciliacao = openConciliacao;
global.openDRE = openDRE;
global.openDFC = openDFC;
global.openKPI = openKPI;
global.openConfiguracoes = openConfiguracoes;
global.refreshCache = refreshCache;
global.setupTriggers = setupTriggers;
global.runCompleteSetup = runCompleteSetup;
global.setupAllSampleData = setupAllSampleData;
global.setupBulkSampleData = setupBulkSampleData;
global.openWebApp = openWebApp;
global.runBackupNow = runBackupNow;
global.exportCurrentReportPdf = exportCurrentReportPdf;
global.exportReportPdfForPeriod = exportReportPdfForPeriod;

// Web App API Functions
global.getViewHtml = wrapApi('getViewHtml', getViewHtml);
global.getReferenceData = wrapApi('getReferenceData', getReferenceData);
global.getDashboardData = wrapApi('getDashboardData', getDashboardData);
global.getContasPagar = wrapApi('getContasPagar', getContasPagar);
global.pagarConta = wrapApi('pagarConta', pagarConta);
global.pagarContasEmLote = wrapApi('pagarContasEmLote', pagarContasEmLote);
global.cancelarContasEmLote = wrapApi('cancelarContasEmLote', cancelarContasEmLote);
global.getContasReceber = wrapApi('getContasReceber', getContasReceber);
global.receberConta = wrapApi('receberConta', receberConta);
global.receberContasEmLote = wrapApi('receberContasEmLote', receberContasEmLote);
global.cancelarContasReceberEmLote = wrapApi('cancelarContasReceberEmLote', cancelarContasReceberEmLote);
global.salvarLancamento = wrapApi('salvarLancamento', salvarLancamento);
global.atualizarLancamento = wrapApi('atualizarLancamento', atualizarLancamento);
global.getLancamentoDetalhes = wrapApi('getLancamentoDetalhes', getLancamentoDetalhes);
global.getConciliacaoData = wrapApi('getConciliacaoData', getConciliacaoData);
global.getComparativoData = wrapApi('getComparativoData', getComparativoData);
global.conciliarItens = wrapApi('conciliarItens', conciliarItens);
global.desfazerConciliacao = wrapApi('desfazerConciliacao', desfazerConciliacao);
global.conciliarAutomatico = wrapApi('conciliarAutomatico', conciliarAutomatico);
global.importarFc = wrapApi('importarFc', importarFc);
global.importarItau = wrapApi('importarItau', importarItau);
global.importarSieg = wrapApi('importarSieg', importarSieg);
global.getSheetData = wrapApi('getSheetData', getSheetData);
global.toggleCentroCusto = wrapApi('toggleCentroCusto', toggleCentroCusto);
global.seedPlanoContasFromList = wrapApi('seedPlanoContasFromList', seedPlanoContasFromList);

// Configurações CRUD
global.salvarCentroCusto = wrapApi('salvarCentroCusto', salvarCentroCusto);
global.excluirCentroCusto = wrapApi('excluirCentroCusto', excluirCentroCusto);
global.salvarContaContabil = wrapApi('salvarContaContabil', salvarContaContabil);
global.excluirConta = wrapApi('excluirConta', excluirConta);
global.salvarCanal = wrapApi('salvarCanal', salvarCanal);
global.excluirCanal = wrapApi('excluirCanal', excluirCanal);
global.toggleCanal = wrapApi('toggleCanal', toggleCanal);
global.salvarFilial = wrapApi('salvarFilial', salvarFilial);
global.excluirFilial = wrapApi('excluirFilial', excluirFilial);
global.toggleFilial = wrapApi('toggleFilial', toggleFilial);

// DRE Functions
global.getDREMensal = wrapApi('getDREMensal', getDREMensal);
global.getDREComparativo = wrapApi('getDREComparativo', getDREComparativo);
global.getDREPorFilial = wrapApi('getDREPorFilial', getDREPorFilial);

// Fluxo de Caixa Functions
global.getFluxoCaixaMensal = wrapApi('getFluxoCaixaMensal', getFluxoCaixaMensal);
global.getFluxoCaixaProjecao = wrapApi('getFluxoCaixaProjecao', getFluxoCaixaProjecao);

// KPIs Functions
global.getKPIsMensal = wrapApi('getKPIsMensal', getKPIsMensal);

// Caixas Functions
global.getCaixasConfig = wrapApi('getCaixasConfig', getCaixasConfig);
global.salvarCaixasConfig = wrapApi('salvarCaixasConfig', salvarCaixasConfig);
global.getCaixasData = wrapApi('getCaixasData', getCaixasData);
global.getCaixaMovimentos = wrapApi('getCaixaMovimentos', getCaixaMovimentos);
global.salvarCaixa = wrapApi('salvarCaixa', salvarCaixa);
global.excluirCaixa = wrapApi('excluirCaixa', excluirCaixa);
global.salvarCaixaMovimento = wrapApi('salvarCaixaMovimento', salvarCaixaMovimento);
global.excluirCaixaMovimento = wrapApi('excluirCaixaMovimento', excluirCaixaMovimento);
global.getCaixaRelatorio = wrapApi('getCaixaRelatorio', getCaixaRelatorio);
global.uploadCaixaArquivo = wrapApi('uploadCaixaArquivo', uploadCaixaArquivo);
global.salvarCaixaTipo = wrapApi('salvarCaixaTipo', salvarCaixaTipo);
global.excluirCaixaTipo = wrapApi('excluirCaixaTipo', excluirCaixaTipo);
global.toggleCaixaTipo = wrapApi('toggleCaixaTipo', toggleCaixaTipo);

// Usuários Functions
global.getUsuarios = wrapApi('getUsuarios', getUsuarios);
global.getCurrentUserInfo = wrapApi('getCurrentUserInfo', getCurrentUserInfo);
global.logClientError = logClientError;
global.getAuditLogEntries = wrapApi('getAuditLogEntries', getAuditLogEntries);
global.salvarUsuario = wrapApi('salvarUsuario', salvarUsuario);
global.excluirUsuario = wrapApi('excluirUsuario', excluirUsuario);
global.runSmokeTests = wrapApi('runSmokeTests', runSmokeTests);
global.getAdminDiagnostics = wrapApi('getAdminDiagnostics', getAdminDiagnostics);
global.setAdminFlag = wrapApi('setAdminFlag', setAdminFlag);
global.clearCaches = wrapApi('clearCaches', clearCaches);
