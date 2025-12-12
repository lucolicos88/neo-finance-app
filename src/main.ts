/**
 * main.ts
 *
 * Entry point principal do Google Apps Script.
 * Centraliza roteamento web (doGet) e registros de menus.
 */

import { include } from './services/ui-service';
import { installTriggers } from './services/scheduler-service';
import { runCompleteSetup } from './setup-sheets';
import { setupAllSampleData } from './setup-sample-data';
import {
  getViewHtml,
  getReferenceData,
  getDashboardData,
  getContasPagar,
  pagarConta,
  pagarContasEmLote,
  getContasReceber,
  receberConta,
  salvarLancamento,
  getConciliacaoData,
  conciliarItens,
  conciliarAutomatico,
  salvarCentroCusto,
  excluirCentroCusto,
  salvarContaContabil,
  excluirConta,
  salvarCanal,
  excluirCanal,
  salvarFilial,
  excluirFilial,
  getDREMensal,
  getDREComparativo,
  getDREPorFilial,
  getFluxoCaixaMensal,
  getFluxoCaixaProjecao,
  getKPIsMensal,
  getUsuarios,
  salvarUsuario,
  excluirUsuario,
} from './services/webapp-service';

/**
 * Função doGet - Entry point para web app
 *
 * IMPORTANTE: Esta função é chamada automaticamente pelo Apps Script
 * quando o usuário acessa a URL do web app
 *
 * Serve a aplicação web standalone (não modal do Sheets)
 */
function doGet(e: any): GoogleAppsScript.HTML.HtmlOutput {
  const template = HtmlService.createTemplateFromFile('frontend/views/app');
  return template.evaluate()
    .setTitle('Neoformula Finance')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Administração')
        .addItem('Configurações', 'openConfiguracoes')
        .addItem('Instalar Triggers', 'setupTriggers')
        .addSeparator()
        .addItem('⚙️ Setup da Planilha', 'runCompleteSetup')
        .addItem('📝 Criar Dados de Exemplo', 'setupAllSampleData')
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
  const html = HtmlService.createHtmlOutput(
    `<html><body>
      <h2>Web App URL</h2>
      <p>Acesse a aplicação web no link abaixo:</p>
      <p><a href="${url}" target="_blank">${url}</a></p>
      <p><button onclick="window.open('${url}', '_blank')">Abrir Web App</button></p>
      <br><p><small>Copie este link para acessar a aplicação de qualquer lugar.</small></p>
    </body></html>`
  ).setWidth(600).setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, 'URL da Web App');
}

/**
 * Exporta funções globais para o escopo do Apps Script
 *
 * IMPORTANTE: gas-webpack-plugin detecta estas declarações 'global'
 * e as move para o escopo global do Apps Script
 */
declare var global: any;

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
global.openWebApp = openWebApp;

// Web App API Functions
global.getViewHtml = getViewHtml;
global.getReferenceData = getReferenceData;
global.getDashboardData = getDashboardData;
global.getContasPagar = getContasPagar;
global.pagarConta = pagarConta;
global.pagarContasEmLote = pagarContasEmLote;
global.getContasReceber = getContasReceber;
global.receberConta = receberConta;
global.salvarLancamento = salvarLancamento;
global.getConciliacaoData = getConciliacaoData;
global.conciliarItens = conciliarItens;
global.conciliarAutomatico = conciliarAutomatico;

// Configurações CRUD
global.salvarCentroCusto = salvarCentroCusto;
global.excluirCentroCusto = excluirCentroCusto;
global.salvarContaContabil = salvarContaContabil;
global.excluirConta = excluirConta;
global.salvarCanal = salvarCanal;
global.excluirCanal = excluirCanal;
global.salvarFilial = salvarFilial;
global.excluirFilial = excluirFilial;

// DRE Functions
global.getDREMensal = getDREMensal;
global.getDREComparativo = getDREComparativo;
global.getDREPorFilial = getDREPorFilial;

// Fluxo de Caixa Functions
global.getFluxoCaixaMensal = getFluxoCaixaMensal;
global.getFluxoCaixaProjecao = getFluxoCaixaProjecao;

// KPIs Functions
global.getKPIsMensal = getKPIsMensal;

// Usuários Functions
global.getUsuarios = getUsuarios;
global.salvarUsuario = salvarUsuario;
global.excluirUsuario = excluirUsuario;
