/**
 * reporting-service.ts
 *
 * Gerencia geração de relatórios formatados para comitê.
 *
 * Responsabilidades:
 * - Gerar relatórios consolidados (faturamento, DRE, DFC, KPIs)
 * - Formatar dados para apresentações
 * - Exportar para PDF/Google Slides
 * - Alimentar abas RPT_COMITE_*
 */

import { setSheetValues, clearRange } from '../shared/sheets-client';
import { Sheets } from '../config/sheet-mapping';
import { Period, BranchId, Money, LedgerEntryStatus, LedgerEntryType } from '../shared/types';
import { formatMoney, formatPercentage } from '../shared/money-utils';
import { formatDate } from '../shared/date-utils';
import { calculateDRE } from './dre-service';
import { calculateKPIs } from './kpi-analytics-service';
import { calculateRealCashflow } from './cashflow-service';
import { listEntries } from './ledger-service';

// ============================================================================
// ESTRUTURAS DE RELATÓRIOS
// ============================================================================

/**
 * Relatório de faturamento para comitê
 */
export interface FaturamentoReport {
  period: Period;
  receitaBrutaTotal: Money;
  receitaBrutaPorFilial: Array<{ filial: BranchId; valor: Money }>;
  receitaBrutaPorCanal: Array<{ canal: string; valor: Money }>;
  comparativoMesAnterior: {
    variacao: Money;
    variacaoPct: number;
  };
}

/**
 * Relatório DRE para comitê
 */
export interface DREReport {
  period: Period;
  resumo: {
    receitaBruta: Money;
    receitaLiquida: Money;
    lucroBruto: Money;
    ebitda: Money;
    ebitdaPct: number;
    lucroLiquido: Money;
    margemLiquida: number;
  };
  porFilial: Array<{
    filial: BranchId;
    receitaLiquida: Money;
    ebitda: Money;
  }>;
}

/**
 * Relatório DFC para comitê
 */
export interface DFCReport {
  period: Period;
  saldoInicial: Money;
  entradasOperacionais: Money;
  saidasOperacionais: Money;
  entradasInvestimento: Money;
  saidasInvestimento: Money;
  entradasFinanciamento: Money;
  saidasFinanciamento: Money;
  saldoFinal: Money;
}

/**
 * Relatório KPIs para comitê
 */
export interface KPIReport {
  period: Period;
  kpis: Array<{
    metric: string;
    value: number;
    faixa: string;
    unit: string;
  }>;
}

// ============================================================================
// GERAÇÃO DE RELATÓRIOS
// ============================================================================

/**
 * Gera relatório de faturamento
 *
 * TODO: Implementar comparativo com mês anterior
 * TODO: Agrupar por canal e filial
 */
export function generateFaturamentoReport(period: Period): FaturamentoReport {
  const entries = listEntries({
    status: LedgerEntryStatus.REALIZADO,
    tipo: LedgerEntryType.RECEBER,
    // TODO: Filtrar por período
  });

  let receitaBrutaTotal = 0;
  const porFilial = new Map<BranchId, Money>();
  const porCanal = new Map<string, Money>();

  for (const entry of entries) {
    receitaBrutaTotal += entry.valorBruto;

    // Agrupa por filial
    const currentFilial = porFilial.get(entry.filial) || 0;
    porFilial.set(entry.filial, currentFilial + entry.valorBruto);

    // Agrupa por canal
    if (entry.canal) {
      const currentCanal = porCanal.get(entry.canal) || 0;
      porCanal.set(entry.canal, currentCanal + entry.valorBruto);
    }
  }

  return {
    period,
    receitaBrutaTotal,
    receitaBrutaPorFilial: Array.from(porFilial.entries()).map(([filial, valor]) => ({
      filial,
      valor,
    })),
    receitaBrutaPorCanal: Array.from(porCanal.entries()).map(([canal, valor]) => ({
      canal,
      valor,
    })),
    comparativoMesAnterior: {
      variacao: 0, // TODO: Calcular
      variacaoPct: 0, // TODO: Calcular
    },
  };
}

/**
 * Gera relatório DRE consolidado
 */
export function generateDREReport(period: Period): DREReport {
  const dreConsolidado = calculateDRE(period, null);

  // TODO: Calcular DRE por filial
  const porFilial: Array<{ filial: BranchId; receitaLiquida: Money; ebitda: Money }> = [];

  return {
    period,
    resumo: {
      receitaBruta: dreConsolidado.summary.receitaBruta,
      receitaLiquida: dreConsolidado.summary.receitaLiquida,
      lucroBruto: dreConsolidado.summary.lucroBruto,
      ebitda: dreConsolidado.summary.ebitda,
      ebitdaPct: dreConsolidado.summary.ebitdaPct,
      lucroLiquido: dreConsolidado.summary.lucroLiquido,
      margemLiquida: dreConsolidado.summary.margemLiquida,
    },
    porFilial,
  };
}

/**
 * Gera relatório DFC
 *
 * TODO: Implementar cálculo real de saldo inicial/final
 * TODO: Agrupar por categoria
 */
export function generateDFCReport(period: Period): DFCReport {
  const cashflowLines = calculateRealCashflow(period);

  let entradasOperacionais = 0;
  let saidasOperacionais = 0;
  let entradasInvestimento = 0;
  let saidasInvestimento = 0;
  let entradasFinanciamento = 0;
  let saidasFinanciamento = 0;

  for (const line of cashflowLines) {
    const valor = line.value;

    if (line.category === 'OPERACIONAL') {
      if (line.type === 'ENTRADA') {
        entradasOperacionais += valor;
      } else {
        saidasOperacionais += valor;
      }
    } else if (line.category === 'INVESTIMENTO') {
      if (line.type === 'ENTRADA') {
        entradasInvestimento += valor;
      } else {
        saidasInvestimento += valor;
      }
    } else if (line.category === 'FINANCIAMENTO') {
      if (line.type === 'ENTRADA') {
        entradasFinanciamento += valor;
      } else {
        saidasFinanciamento += valor;
      }
    }
  }

  // TODO: Calcular saldo inicial e final real
  const saldoInicial = 0;
  const variacao =
    entradasOperacionais -
    saidasOperacionais +
    entradasInvestimento -
    saidasInvestimento +
    entradasFinanciamento -
    saidasFinanciamento;
  const saldoFinal = saldoInicial + variacao;

  return {
    period,
    saldoInicial,
    entradasOperacionais,
    saidasOperacionais,
    entradasInvestimento,
    saidasInvestimento,
    entradasFinanciamento,
    saidasFinanciamento,
    saldoFinal,
  };
}

/**
 * Gera relatório de KPIs
 */
export function generateKPIReport(period: Period): KPIReport {
  const kpis = calculateKPIs(period, null);

  return {
    period,
    kpis: kpis.map((kpi) => ({
      metric: kpi.metric,
      value: kpi.value,
      faixa: kpi.range,
      unit: kpi.unit,
    })),
  };
}

/**
 * Gera relatório completo para comitê (todos os relatórios)
 */
export function generateCommitteeReport(period: Period): {
  faturamento: FaturamentoReport;
  dre: DREReport;
  dfc: DFCReport;
  kpis: KPIReport;
} {
  return {
    faturamento: generateFaturamentoReport(period),
    dre: generateDREReport(period),
    dfc: generateDFCReport(period),
    kpis: generateKPIReport(period),
  };
}

// ============================================================================
// PERSISTÊNCIA EM ABAS RPT_*
// ============================================================================

/**
 * Persiste relatório de faturamento em RPT_COMITE_FATURAMENTO
 */
export function persistFaturamentoReport(report: FaturamentoReport): void {
  // TODO: Formatar dados e escrever na aba
  // clearRange(Sheets.RPT_COMITE_FATURAMENTO, 'A2:Z');

  const rows = [
    ['Receita Bruta Total', formatMoney(report.receitaBrutaTotal)],
    ['', ''],
    ['Por Filial', ''],
    ...report.receitaBrutaPorFilial.map((item) => [item.filial, formatMoney(item.valor)]),
    ['', ''],
    ['Por Canal', ''],
    ...report.receitaBrutaPorCanal.map((item) => [item.canal, formatMoney(item.valor)]),
  ];

  // setSheetValues(Sheets.RPT_COMITE_FATURAMENTO, 'A2', rows);
}

/**
 * Persiste relatório DRE em RPT_COMITE_DRE
 */
export function persistDREReport(report: DREReport): void {
  // TODO: Formatar DRE em layout de comitê
}

/**
 * Persiste relatório DFC em RPT_COMITE_DFC
 */
export function persistDFCReport(report: DFCReport): void {
  // TODO: Formatar DFC em layout de comitê
}

/**
 * Persiste relatório KPIs em RPT_COMITE_KPIS
 */
export function persistKPIReport(report: KPIReport): void {
  // TODO: Formatar KPIs em layout de comitê com cores de faixa
}

// ============================================================================
// EXPORTAÇÃO
// ============================================================================

/**
 * Exporta relatório para PDF
 *
 * TODO: Implementar usando Google Drive API
 */
export function exportToPDF(period: Period): string {
  // TODO: Gerar PDF da planilha ou de template customizado
  throw new Error('exportToPDF não implementado');
}

/**
 * Exporta relatório para Google Slides
 *
 * TODO: Implementar usando Slides API
 */
export function exportToSlides(period: Period): string {
  // TODO: Popular template de apresentação com dados
  throw new Error('exportToSlides não implementado');
}
