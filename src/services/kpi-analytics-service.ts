/**
 * kpi-analytics-service.ts
 *
 * Gerencia cálculo e análise de KPIs financeiros.
 *
 * Responsabilidades:
 * - Calcular KPIs críticos (desconto médio, CMA, CMV, margens, etc.)
 * - Aplicar benchmarks e classificar em faixas
 * - Gerar dados para dashboard
 * - Persistir em TB_KPI_RESUMO e TB_KPI_DETALHE
 */

import { getSheetValues, setSheetValues, createSheetIfNotExists } from '../shared/sheets-client';
import { Sheets } from '../config/sheet-mapping';
import {
  KPILine,
  Period,
  BranchId,
  ChannelId,
  Money,
  DashboardData,
  ReportFilter,
  LedgerEntryStatus,
  LedgerEntryType,
} from '../shared/types';
import { calculatePercentage, sumMoney } from '../shared/money-utils';
import { BenchmarkRange, KPIMetric, getBenchmarkRange, BenchmarkConfig } from '../config/benchmarks';
import { getAllBenchmarks, getAccountByCode } from './reference-data-service';
import { calculateDRE } from './dre-service';
import { listEntries } from './ledger-service';
import { getFirstDayOfPeriod } from '../shared/date-utils';

// ============================================================================
// CÁLCULO DE KPIs
// ============================================================================

/**
 * Estrutura de KPI calculado
 */
export interface CalculatedKPI {
  metric: KPIMetric;
  value: number;
  range: BenchmarkRange;
  unit: string;
}

/**
 * Calcula todos os KPIs para um período
 *
 * @param period - Período a calcular
 * @param branchId - ID da filial (null = consolidado)
 * @param channelId - ID do canal (null = todos)
 * @returns Array de KPIs calculados
 *
 * TODO: Implementar cálculos reais de cada KPI
 * TODO: Otimizar para evitar recalcular DRE múltiplas vezes
 */
export function calculateKPIs(
  period: Period,
  branchId: BranchId | null = null,
  channelId: ChannelId | null = null
): CalculatedKPI[] {
  const kpis: CalculatedKPI[] = [];

  // Calcula DRE do período
  const dre = calculateDRE(period, branchId);

  const benchmarkMap = getBenchmarkMap();
  const withBenchmark = (
    metric: KPIMetric,
    value: number,
    fallbackRange: BenchmarkRange,
    fallbackUnit: string
  ): { range: BenchmarkRange; unit: string } => {
    const benchmark = benchmarkMap.get(metric);
    if (!benchmark) {
      return { range: fallbackRange, unit: fallbackUnit };
    }
    return {
      range: getBenchmarkRange(value, benchmark),
      unit: benchmark.unit,
    };
  };

  // ========================================================================
  // KPI: Desconto Médio
  // ========================================================================
  const descontoMedio = calculateDescontoMedio(period, branchId, channelId);
  const descontoBench = withBenchmark(KPIMetric.DESCONTO_MEDIO, descontoMedio, BenchmarkRange.BOM, '%');
  kpis.push({
    metric: KPIMetric.DESCONTO_MEDIO,
    value: descontoMedio,
    range: descontoBench.range,
    unit: descontoBench.unit,
  });

  // ========================================================================
  // KPI: CMA (Custo de Mercadoria Adquirida)
  // ========================================================================
  const cma = calculateCMA(period, branchId);
  const cmaBench = withBenchmark(KPIMetric.CMA, cma, BenchmarkRange.BOM, 'R$/UNID');
  kpis.push({
    metric: KPIMetric.CMA,
    value: cma,
    range: cmaBench.range,
    unit: cmaBench.unit,
  });

  // ========================================================================
  // KPI: CMV (Custo de Mercadoria Vendida)
  // ========================================================================
  const cmv = calculateCMV(period, branchId);
  const cmvBench = withBenchmark(KPIMetric.CMV, cmv, BenchmarkRange.BOM, 'R$/UNID');
  kpis.push({
    metric: KPIMetric.CMV,
    value: cmv,
    range: cmvBench.range,
    unit: cmvBench.unit,
  });

  // ========================================================================
  // KPI: Margem Bruta
  // ========================================================================
  const margemBruta = calculatePercentage(dre.summary.lucroBruto, dre.summary.receitaLiquida);
  const margemBrutaBench = withBenchmark(KPIMetric.MARGEM_BRUTA, margemBruta, BenchmarkRange.EXCELENTE, '%');
  kpis.push({
    metric: KPIMetric.MARGEM_BRUTA,
    value: margemBruta,
    range: margemBrutaBench.range,
    unit: margemBrutaBench.unit,
  });

  // ========================================================================
  // KPI: EBITDA %
  // ========================================================================
  const ebitdaBench = withBenchmark(KPIMetric.EBITDA_PCT, dre.summary.ebitdaPct, BenchmarkRange.BOM, '%');
  kpis.push({
    metric: KPIMetric.EBITDA_PCT,
    value: dre.summary.ebitdaPct,
    range: ebitdaBench.range,
    unit: ebitdaBench.unit,
  });

  // ========================================================================
  // KPI: Margem Líquida
  // ========================================================================
  const margemLiquidaBench = withBenchmark(KPIMetric.MARGEM_LIQUIDA, dre.summary.margemLiquida, BenchmarkRange.BOM, '%');
  kpis.push({
    metric: KPIMetric.MARGEM_LIQUIDA,
    value: dre.summary.margemLiquida,
    range: margemLiquidaBench.range,
    unit: margemLiquidaBench.unit,
  });

  return kpis;
}

// ============================================================================
// CÁLCULOS ESPECÍFICOS DE KPIs
// ============================================================================

/**
 * Calcula desconto médio percentual
 *
 * TODO: Implementar cálculo real
 */
function calculateDescontoMedio(
  period: Period,
  branchId: BranchId | null,
  channelId: ChannelId | null
): number {
  const entries = listEntries({
    status: LedgerEntryStatus.REALIZADO,
    ...(branchId && { filial: branchId }),
    ...(channelId && { canal: channelId }),
    periodStart: period,
    periodEnd: period,
  });

  let totalBruto = 0;
  let totalDesconto = 0;

  for (const entry of entries) {
    totalBruto += entry.valorBruto;
    totalDesconto += entry.desconto;
  }

  if (totalBruto === 0) return 0;

  return calculatePercentage(totalDesconto, totalBruto);
}

/**
 * Calcula CMA (Custo de Mercadoria Adquirida)
 *
 * TODO: Implementar cálculo real baseado em contas CMA
 */
function calculateCMA(period: Period, branchId: BranchId | null): number {
  return calculateCmaCmv(period, branchId, 'CMA');
}

/**
 * Calcula CMV (Custo de Mercadoria Vendida)
 *
 * TODO: Implementar cálculo real baseado em contas CMV
 */
function calculateCMV(period: Period, branchId: BranchId | null): number {
  return calculateCmaCmv(period, branchId, 'CMV');
}

// ============================================================================
// DASHBOARD
// ============================================================================

/**
 * Gera dados para o dashboard
 *
 * @param period - Período
 * @param branchId - ID da filial (null = consolidado)
 * @returns Dados estruturados para o dashboard
 */
export function getDashboardData(period: Period, branchId: BranchId | null = null): DashboardData {
  const dre = calculateDRE(period, branchId);
  const kpis = calculateKPIs(period, branchId);

  const saldoCaixa = calculateSaldoCaixa(period, branchId);

  const topDespesas = getTopDespesas(period, branchId);

  return {
    period,
    receitaBruta: dre.summary.receitaBruta,
    receitaLiquida: dre.summary.receitaLiquida,
    ebitda: dre.summary.ebitda,
    ebitdaPct: dre.summary.ebitdaPct,
    saldoCaixa,
    kpis: kpis.map((kpi) => ({
      period,
      filial: branchId,
      canal: null,
      metric: kpi.metric,
      value: kpi.value,
      faixa: kpi.range,
    })),
    topDespesas,
  };
}

function calculateSaldoCaixa(period: Period, branchId: BranchId | null): Money {
  const entries = listEntries({
    status: LedgerEntryStatus.REALIZADO,
    ...(branchId && { filial: branchId }),
  });
  const startDate = getFirstDayOfPeriod(period);
  startDate.setHours(0, 0, 0, 0);

  let saldo = 0;
  for (const entry of entries) {
    if (!entry.pagamento) continue;
    const pagamento = new Date(entry.pagamento);
    if (pagamento >= startDate) continue;
    if (entry.tipo === LedgerEntryType.RECEBER) saldo += entry.valorLiquido;
    else if (entry.tipo === LedgerEntryType.PAGAR) saldo -= entry.valorLiquido;
  }
  return saldo;
}

function getTopDespesas(period: Period, branchId: BranchId | null): Array<{ descricao: string; valor: Money }> {
  const entries = listEntries({
    status: LedgerEntryStatus.REALIZADO,
    tipo: LedgerEntryType.PAGAR,
    periodStart: period,
    periodEnd: period,
    ...(branchId && { filial: branchId }),
  });

  const totals = new Map<string, Money>();
  for (const entry of entries) {
    const key = String(entry.descricao || 'Sem descrição');
    totals.set(key, (totals.get(key) || 0) + entry.valorLiquido);
  }

  return Array.from(totals.entries())
    .map(([descricao, valor]) => ({ descricao, valor }))
    .sort((a, b) => b.valor - a.valor)
    .slice(0, 10);
}

function getBenchmarkMap(): Map<KPIMetric, BenchmarkConfig> {
  const benchmarks = getAllBenchmarks();
  const map = new Map<KPIMetric, BenchmarkConfig>();
  benchmarks.forEach((b) => {
    const metric = String(b.metric || '').trim() as KPIMetric;
    if (!metric) return;
    map.set(metric, b);
  });
  return map;
}

function calculateCmaCmv(period: Period, branchId: BranchId | null, kind: 'CMA' | 'CMV'): number {
  const entries = listEntries({
    status: LedgerEntryStatus.REALIZADO,
    tipo: LedgerEntryType.PAGAR,
    periodStart: period,
    periodEnd: period,
    ...(branchId && { filial: branchId }),
  });

  let total = 0;
  let count = 0;
  for (const entry of entries) {
    const accountCode = entry.contaContabil || entry.contaGerencial;
    if (!accountCode) continue;
    const account = getAccountByCode(accountCode);
    const cmaCmv = String(account?.cmaCmv || '').toUpperCase();
    if (cmaCmv !== kind) continue;
    total += entry.valorLiquido;
    count++;
  }

  if (count === 0) return 0;
  return total / count;
}

/**
 * Obtém tendência de um KPI ao longo de vários períodos
 *
 * @param metric - Métrica a analisar
 * @param periods - Períodos a incluir
 * @param filter - Filtros adicionais
 * @returns Array de valores do KPI por período
 *
 * TODO: Implementar leitura de TB_KPI_RESUMO histórico
 */
export function getKPITrend(
  metric: KPIMetric,
  periods: Period[],
  filter: ReportFilter = {}
): Array<{ period: Period; value: number }> {
  const trend: Array<{ period: Period; value: number }> = [];

  for (const period of periods) {
    const kpis = calculateKPIs(period, filter.filial || null, filter.canal || null);
    const kpi = kpis.find((k) => k.metric === metric);

    if (kpi) {
      trend.push({ period, value: kpi.value });
    }
  }

  return trend;
}

// ============================================================================
// PERSISTÊNCIA
// ============================================================================

/**
 * Persiste KPIs na aba TB_KPI_RESUMO
 *
 * TODO: Implementar lógica de merge (atualizar apenas o período específico)
 */
export function persistKPIs(
  period: Period,
  branchId: BranchId | null,
  channelId: ChannelId | null,
  kpis: CalculatedKPI[]
): void {
  const headers = ['Ano', 'Mes', 'Filial', 'Canal', 'Metric', 'Valor', 'Faixa'];
  createSheetIfNotExists(Sheets.TB_KPI_RESUMO, headers);

  const rows = kpis.map((kpi) => [
    period.year,
    period.month,
    branchId || '',
    channelId || '',
    kpi.metric,
    kpi.value,
    kpi.range,
  ]);

  const existing = getSheetValues(Sheets.TB_KPI_RESUMO);
  const filtered = existing.length > 1
    ? existing.slice(1).filter((r) => {
        const samePeriod = r[0] === period.year && r[1] === period.month;
        const sameBranch = String(r[2] || '') === String(branchId || '');
        const sameChannel = String(r[3] || '') === String(channelId || '');
        return !(samePeriod && sameBranch && sameChannel);
      })
    : [];

  const allRows = [headers, ...filtered, ...rows];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Sheets.TB_KPI_RESUMO);
  if (!sheet) return;
  sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);
  const lastRow = sheet.getLastRow();
  if (lastRow > allRows.length) {
    sheet.getRange(allRows.length + 1, 1, lastRow - allRows.length, headers.length).clearContent();
  }
}

/**
 * Persiste KPIs detalhados na aba TB_KPI_DETALHE
 *
 * Para análises mais granulares (por produto, família, etc.)
 *
 * TODO: Definir estrutura de KPI detalhado
 */
export function persistDetailedKPIs(period: Period, data: any[]): void {
  // TODO: Implementar estrutura de KPI detalhado
}
