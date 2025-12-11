/**
 * webapp-service.ts
 *
 * Camada fina de serviÇõÇœo para o Web App (GAS).
 * - Fornece HTML das views SPA.
 * - ExpÇõe endpoints para o frontend (contas a pagar/receber, conciliaÇõÇœo, dashboard).
 * - Usa mocks em memÇ®ria apenas para demonstraÇõÇœo.
 *
 * Quando os microservices reais estiverem prontos, troque as funÇõÇœes de dados
 * para consumir as planilhas / APIs adequadas mantendo a assinatura.
 */

import { include } from './ui-service';

// -----------------------------------------------------------------------------
// Tipos auxiliares
// -----------------------------------------------------------------------------

interface ApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
}

interface ContaPagar {
  id: string;
  fornecedor: string;
  descricao: string;
  vencimento: string;
  valor: number;
  status: 'PENDENTE' | 'VENCIDA' | 'PAGA';
  filial: string;
}

interface ContaReceber {
  id: string;
  cliente: string;
  descricao: string;
  vencimento: string;
  valor: number;
  status: 'PENDENTE' | 'VENCIDA' | 'RECEBIDA';
  canal: string;
}

interface Extrato {
  id: string;
  data: string;
  descricao: string;
  valor: number;
  banco: string;
}

interface Lancamento {
  id: string;
  data: string;
  descricao: string;
  valor: number;
  tipo: 'RECEITA' | 'DESPESA';
}

// -----------------------------------------------------------------------------
// Helpers
// -----------------------------------------------------------------------------

const ok = <T>(data: T): ApiResponse<T> => ({ success: true, data });
const err = (message: string): ApiResponse<never> => ({ success: false, error: message });

function todayISO(): string {
  return new Date().toISOString().slice(0, 10);
}

function addDays(base: Date, days: number): string {
  const d = new Date(base);
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

function currencySum(items: { valor: number }[]): number {
  return items.reduce((acc, item) => acc + (item.valor || 0), 0);
}

function buildContasPagarMock(): ContaPagar[] {
  const base = new Date();
  return [
    { id: 'CP-101', fornecedor: 'Fornecedor A', descricao: 'Insumos cozinha', vencimento: addDays(base, -3), valor: 1250.5, status: 'VENCIDA', filial: 'SP' },
    { id: 'CP-102', fornecedor: 'Fornecedor B', descricao: 'Energia elÇ¦trica', vencimento: addDays(base, 2), valor: 830.9, status: 'PENDENTE', filial: 'RJ' },
    { id: 'CP-103', fornecedor: 'Fornecedor C', descricao: 'ManutenÇõÇœo', vencimento: addDays(base, 7), valor: 420.0, status: 'PENDENTE', filial: 'SP' },
    { id: 'CP-104', fornecedor: 'Fornecedor D', descricao: 'Aluguel', vencimento: addDays(base, -1), valor: 5600.0, status: 'VENCIDA', filial: 'MG' },
    { id: 'CP-105', fornecedor: 'Fornecedor E', descricao: 'Marketing', vencimento: addDays(base, 1), valor: 2150.0, status: 'PAGA', filial: 'SP' },
  ];
}

function buildContasReceberMock(): ContaReceber[] {
  const base = new Date();
  return [
    { id: 'CR-201', cliente: 'Cliente X', descricao: 'Contrato mensal', vencimento: addDays(base, -1), valor: 3500, status: 'VENCIDA', canal: 'Online' },
    { id: 'CR-202', cliente: 'Cliente Y', descricao: 'Projeto pontual', vencimento: addDays(base, 3), valor: 7200, status: 'PENDENTE', canal: 'Loja' },
    { id: 'CR-203', cliente: 'Cliente Z', descricao: 'Mensalidade', vencimento: addDays(base, 1), valor: 1800, status: 'PENDENTE', canal: 'Online' },
    { id: 'CR-204', cliente: 'Cliente W', descricao: 'ServiÇõÇœ premium', vencimento: addDays(base, -7), valor: 9500, status: 'RECEBIDA', canal: 'Marketplace' },
  ];
}

function buildExtratosMock(): Extrato[] {
  const base = new Date();
  return [
    { id: 'EXT-500', data: addDays(base, -1), descricao: 'TED Cliente X', valor: 3500, banco: 'Banco A' },
    { id: 'EXT-501', data: addDays(base, -2), descricao: 'CartÇœo - Vendas', valor: 1890.75, banco: 'Banco B' },
    { id: 'EXT-502', data: addDays(base, -3), descricao: 'PIX - Fornecedor', valor: -420.0, banco: 'Banco A' },
  ];
}

function buildLancamentosMock(): Lancamento[] {
  const base = new Date();
  return [
    { id: 'LAN-800', data: addDays(base, -1), descricao: 'Recebimento Cliente X', valor: 3500, tipo: 'RECEITA' },
    { id: 'LAN-801', data: addDays(base, -3), descricao: 'Pagamento ManutenÇõÇœo', valor: -420, tipo: 'DESPESA' },
    { id: 'LAN-802', data: addDays(base, -2), descricao: 'Vendas CartÇœo', valor: 1890.75, tipo: 'RECEITA' },
  ];
}

// -----------------------------------------------------------------------------
// View Loader
// -----------------------------------------------------------------------------

export function getViewHtml(view: string): GoogleAppsScript.HTML.HtmlOutput | string {
  try {
    const map: Record<string, string> = {
      dashboard: 'frontend/views/dashboard-view',
      'contas-pagar': 'frontend/views/contas-pagar-view',
      'contas-receber': 'frontend/views/contas-receber-view',
      conciliacao: 'frontend/views/conciliacao-view',
      relatorios: 'frontend/views/relatorios-view',
      configuracoes: 'frontend/views/configuracoes-view',
    };

    const file = map[view] || 'frontend/views/dashboard-view';
    return HtmlService.createTemplateFromFile(file).evaluate().getContent();
  } catch (error: any) {
    console.error('Erro ao carregar view', view, error);
    return `Erro ao carregar view ${view}: ${error.message}`;
  }
}

// -----------------------------------------------------------------------------
// ReferÇõÇœncias
// -----------------------------------------------------------------------------

export function getReferenceData(): ApiResponse<{ filiais: any[]; canais: any[] }> {
  try {
    const filiais = [
      { codigo: 'SP', nome: 'São Paulo' },
      { codigo: 'RJ', nome: 'Rio de Janeiro' },
      { codigo: 'MG', nome: 'Minas Gerais' },
    ];

    const canais = [
      { codigo: 'ONLINE', nome: 'Online' },
      { codigo: 'LOJA', nome: 'Loja FÇ­sica' },
      { codigo: 'MKT', nome: 'Marketplace' },
    ];

    return ok({ filiais, canais });
  } catch (error: any) {
    return err(error.message || 'Erro ao carregar referÇõÇœncias');
  }
}

// -----------------------------------------------------------------------------
// Dashboard
// -----------------------------------------------------------------------------

export function getDashboardData(): ApiResponse<any> {
  try {
    const contasPagar = buildContasPagarMock();
    const contasReceber = buildContasReceberMock();
    const hoje = todayISO();

    const pagarVencidas = contasPagar.filter((c) => c.status === 'VENCIDA');
    const pagarProximas = contasPagar.filter((c) => c.status === 'PENDENTE' && c.vencimento <= addDays(new Date(), 7));
    const receberHoje = contasReceber.filter((c) => c.vencimento === hoje);

    const concPendentes = buildExtratosMock().filter((e) => e.valor !== 0);

    const recentTransactions = [
      ...contasPagar.slice(0, 2).map((c) => ({ id: c.id, data: c.vencimento, descricao: c.descricao, valor: -Math.abs(c.valor), tipo: 'DESPESA', status: c.status })),
      ...contasReceber.slice(0, 2).map((c) => ({ id: c.id, data: c.vencimento, descricao: c.descricao, valor: c.valor, tipo: 'RECEITA', status: c.status })),
    ];

    const alerts = [
      { type: 'warning', title: 'Contas vencidas', message: `${pagarVencidas.length} contas a pagar vencidas` },
      { type: 'info', title: 'Recebimentos', message: `${receberHoje.length} recebimentos previstos hoje` },
    ];

    return ok({
      pagarVencidas: { quantidade: pagarVencidas.length, valor: currencySum(pagarVencidas) },
      pagarProximas: { quantidade: pagarProximas.length, valor: currencySum(pagarProximas) },
      receberHoje: { quantidade: receberHoje.length, valor: currencySum(receberHoje) },
      conciliacaoPendentes: { quantidade: concPendentes.length, valor: currencySum(concPendentes) },
      recentTransactions,
      alerts,
    });
  } catch (error: any) {
    return err(error.message || 'Erro ao carregar dashboard');
  }
}

// -----------------------------------------------------------------------------
// Contas a Pagar
// -----------------------------------------------------------------------------

export function getContasPagar(): ApiResponse<any> {
  try {
    const contas = buildContasPagarMock();
    const stats = buildStats(contas, 'PAGAR');
    return ok({ contas, stats });
  } catch (error: any) {
    return err(error.message || 'Erro ao listar contas a pagar');
  }
}

export function pagarConta(id: string): ApiResponse<{ id: string }> {
  // Mock: apenas confirmaÇõÇœo
  console.log('Conta paga', id);
  return ok({ id });
}

export function pagarContasEmLote(ids: string[]): ApiResponse<{ quantidade: number }> {
  console.log('Pagando em lote', ids);
  return ok({ quantidade: ids?.length || 0 });
}

// -----------------------------------------------------------------------------
// Contas a Receber
// -----------------------------------------------------------------------------

export function getContasReceber(): ApiResponse<any> {
  try {
    const contas = buildContasReceberMock();
    const stats = buildStatsReceber(contas);
    return ok({ contas, stats });
  } catch (error: any) {
    return err(error.message || 'Erro ao listar contas a receber');
  }
}

export function receberConta(id: string): ApiResponse<{ id: string }> {
  console.log('Conta recebida', id);
  return ok({ id });
}

// -----------------------------------------------------------------------------
// ConciliaÇõÇœo
// -----------------------------------------------------------------------------

export function getConciliacaoData(): ApiResponse<any> {
  try {
    const extratos = buildExtratosMock();
    const lancamentos = buildLancamentosMock();

    const stats = {
      extratosPendentes: extratos.length,
      extratosValor: currencySum(extratos),
      lancamentosPendentes: lancamentos.length,
      lancamentosValor: currencySum(lancamentos),
      conciliadosHoje: 0,
      conciliadosHojeValor: 0,
      taxaConciliacao: 0,
    };

    const historico: any[] = []; // Mock vazio

    return ok({ extratos, lancamentos, stats, historico });
  } catch (error: any) {
    return err(error.message || 'Erro ao carregar conciliaÇõÇœo');
  }
}

export function conciliarItens(extratoId: string, lancamentoId: string): ApiResponse<{ extratoId: string; lancamentoId: string }> {
  console.log('Conciliando', extratoId, lancamentoId);
  return ok({ extratoId, lancamentoId });
}

export function conciliarAutomatico(): ApiResponse<{ conciliados: number }> {
  console.log('Conciliando automaticamente');
  return ok({ conciliados: 2 });
}

// -----------------------------------------------------------------------------
// EstatÇ¸sticas auxiliares
// -----------------------------------------------------------------------------

function buildStats(contas: ContaPagar[], tipo: 'PAGAR') {
  const hoje = todayISO();
  const base = new Date();

  return {
    vencidas: summarize(contas.filter((c) => c.status === 'VENCIDA' || c.vencimento < hoje)),
    vencer7: summarize(contas.filter((c) => c.status === 'PENDENTE' && c.vencimento <= addDays(base, 7))),
    vencer30: summarize(contas.filter((c) => c.status === 'PENDENTE' && c.vencimento <= addDays(base, 30))),
    pagas: summarize(contas.filter((c) => c.status === 'PAGA')),
  };
}

function buildStatsReceber(contas: ContaReceber[]) {
  const hoje = todayISO();
  const base = new Date();

  return {
    vencidas: summarize(contas.filter((c) => c.status === 'VENCIDA' || c.vencimento < hoje)),
    receber7: summarize(contas.filter((c) => c.status === 'PENDENTE' && c.vencimento <= addDays(base, 7))),
    receber30: summarize(contas.filter((c) => c.status === 'PENDENTE' && c.vencimento <= addDays(base, 30))),
    recebidas: summarize(contas.filter((c) => c.status === 'RECEBIDA')),
  };
}

function summarize(items: { valor: number }[]) {
  return {
    quantidade: items.length,
    valor: currencySum(items),
  };
}
