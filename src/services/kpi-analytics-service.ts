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

import { getSheetValues, setSheetValues } from '../shared/sheets-client';
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
} from '../shared/types';
import { calculatePercentage, sumMoney } from '../shared/money-utils';
import { BenchmarkRange, KPIMetric, getBenchmarkRange } from '../config/benchmarks';
import { getAllBenchmarks } from './reference-data-service';
import { calculateDRE } from './dre-service';
import { listEntries } from './ledger-service';

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

  // TODO: Buscar benchmarks
  // const benchmarks = getAllBenchmarks();

  // ========================================================================
  // KPI: Desconto Médio
  // ========================================================================
  const descontoMedio = calculateDescontoMedio(period, branchId, channelId);
  kpis.push({
    metric: KPIMetric.DESCONTO_MEDIO,
    value: descontoMedio,
    range: BenchmarkRange.BOM, // TODO: Aplicar benchmark real
    unit: '%',
  });

  // ========================================================================
  // KPI: CMA (Custo de Mercadoria Adquirida)
  // ========================================================================
  const cma = calculateCMA(period, branchId);
  kpis.push({
    metric: KPIMetric.CMA,
    value: cma,
    range: BenchmarkRange.BOM,
    unit: 'R$/UNID',
  });

  // ========================================================================
  // KPI: CMV (Custo de Mercadoria Vendida)
  // ========================================================================
  const cmv = calculateCMV(period, branchId);
  kpis.push({
    metric: KPIMetric.CMV,
    value: cmv,
    range: BenchmarkRange.BOM,
    unit: 'R$/UNID',
  });

  // ========================================================================
  // KPI: Margem Bruta
  // ========================================================================
  const margemBruta = calculatePercentage(dre.summary.lucroBruto, dre.summary.receitaLiquida);
  kpis.push({
    metric: KPIMetric.MARGEM_BRUTA,
    value: margemBruta,
    range: BenchmarkRange.EXCELENTE,
    unit: '%',
  });

  // ========================================================================
  // KPI: EBITDA %
  // ========================================================================
  kpis.push({
    metric: KPIMetric.EBITDA_PCT,
    value: dre.summary.ebitdaPct,
    range: BenchmarkRange.BOM,
    unit: '%',
  });

  // ========================================================================
  // KPI: Margem Líquida
  // ========================================================================
  kpis.push({
    metric: KPIMetric.MARGEM_LIQUIDA,
    value: dre.summary.margemLiquida,
    range: BenchmarkRange.BOM,
    unit: '%',
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
    // TODO: Filtrar por período
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
  // TODO: Filtrar lançamentos com cmaCmv = 'CMA'
  // TODO: Dividir por quantidade de unidades adquiridas
  return 0;
}

/**
 * Calcula CMV (Custo de Mercadoria Vendida)
 *
 * TODO: Implementar cálculo real baseado em contas CMV
 */
function calculateCMV(period: Period, branchId: BranchId | null): number {
  // TODO: Filtrar lançamentos com cmaCmv = 'CMV'
  // TODO: Dividir por quantidade de unidades vendidas
  return 0;
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

  // TODO: Calcular saldo de caixa real
  const saldoCaixa = 0;

  // TODO: Buscar top despesas
  const topDespesas: Array<{ descricao: string; valor: Money }> = [];

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
  const rows = kpis.map((kpi) => [
    period.year,
    period.month,
    branchId || '',
    channelId || '',
    kpi.metric,
    kpi.value,
    kpi.range,
  ]);

  // TODO: Usar setSheetValues para atualizar range específico
  // setSheetValues(Sheets.TB_KPI_RESUMO, 'A2', rows);
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
