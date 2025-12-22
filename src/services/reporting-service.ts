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

import { setSheetValues, clearRange, createSheetIfNotExists } from '../shared/sheets-client';
import { Sheets } from '../config/sheet-mapping';
import { Period, BranchId, Money, LedgerEntryStatus, LedgerEntryType } from '../shared/types';
import { formatMoney, formatPercentage } from '../shared/money-utils';
import { formatDate } from '../shared/date-utils';
import { calculateDRE } from './dre-service';
import { calculateKPIs } from './kpi-analytics-service';
import { calculateRealCashflow } from './cashflow-service';
import { listEntries } from './ledger-service';
import { getActiveBranches } from './reference-data-service';

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
    periodStart: period,
    periodEnd: period,
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

  const previousPeriod = getPreviousPeriod(period);
  const previousEntries = listEntries({
    status: LedgerEntryStatus.REALIZADO,
    tipo: LedgerEntryType.RECEBER,
    periodStart: previousPeriod,
    periodEnd: previousPeriod,
  });
  const receitaAnterior = previousEntries.reduce((sum, e) => sum + e.valorBruto, 0);
  const variacao = receitaBrutaTotal - receitaAnterior;
  const variacaoPct = receitaAnterior ? (variacao / receitaAnterior) * 100 : 0;

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
      variacao,
      variacaoPct,
    },
  };
}

/**
 * Gera relatório DRE consolidado
 */
export function generateDREReport(period: Period): DREReport {
  const dreConsolidado = calculateDRE(period, null);

  const branches = getActiveBranches();
  const porFilial = branches.map((branch) => {
    const dre = calculateDRE(period, branch.id);
    return {
      filial: branch.id,
      receitaLiquida: dre.summary.receitaLiquida,
      ebitda: dre.summary.ebitda,
    };
  });

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

  const saldoInicial = calculateSaldoInicial(period);
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

function getPreviousPeriod(period: Period): Period {
  if (period.month === 1) {
    return { year: period.year - 1, month: 12 };
  }
  return { year: period.year, month: period.month - 1 };
}

function calculateSaldoInicial(period: Period): Money {
  const entries = listEntries({ status: LedgerEntryStatus.REALIZADO });
  const startDate = new Date(period.year, period.month - 1, 1);
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

// ============================================================================
// PERSISTÊNCIA EM ABAS RPT_*
// ============================================================================

/**
 * Persiste relatório de faturamento em RPT_COMITE_FATURAMENTO
 */
export function persistFaturamentoReport(report: FaturamentoReport): void {
  const headers = ['Item', 'Valor'];
  createSheetIfNotExists(Sheets.RPT_COMITE_FATURAMENTO, headers);
  const rows = [
    ['Receita Bruta Total', formatMoney(report.receitaBrutaTotal)],
    ['Variacao Mes Anterior', formatMoney(report.comparativoMesAnterior.variacao)],
    ['Variacao %', formatPercentage(report.comparativoMesAnterior.variacaoPct)],
    ['', ''],
    ['Por Filial', ''],
    ...report.receitaBrutaPorFilial.map((item) => [item.filial, formatMoney(item.valor)]),
    ['', ''],
    ['Por Canal', ''],
    ...report.receitaBrutaPorCanal.map((item) => [item.canal, formatMoney(item.valor)]),
  ];

  clearRange(Sheets.RPT_COMITE_FATURAMENTO, 'A1:Z');
  setSheetValues(Sheets.RPT_COMITE_FATURAMENTO, 'A1', rows);
}

/**
 * Persiste relatório DRE em RPT_COMITE_DRE
 */
export function persistDREReport(report: DREReport): void {
  createSheetIfNotExists(Sheets.RPT_COMITE_DRE, ['Item', 'Valor']);
  const rows: any[][] = [
    ['Resumo DRE', 'Valor'],
    ['Receita Bruta', formatMoney(report.resumo.receitaBruta)],
    ['Receita Liquida', formatMoney(report.resumo.receitaLiquida)],
    ['Lucro Bruto', formatMoney(report.resumo.lucroBruto)],
    ['EBITDA', formatMoney(report.resumo.ebitda)],
    ['EBITDA %', formatPercentage(report.resumo.ebitdaPct)],
    ['Lucro Liquido', formatMoney(report.resumo.lucroLiquido)],
    ['Margem Liquida', formatPercentage(report.resumo.margemLiquida)],
    ['', ''],
    ['Por Filial', ''],
    ['Filial', 'Receita Liquida', 'EBITDA'],
    ...report.porFilial.map((item) => [item.filial, formatMoney(item.receitaLiquida), formatMoney(item.ebitda)]),
  ];

  clearRange(Sheets.RPT_COMITE_DRE, 'A1:Z');
  setSheetValues(Sheets.RPT_COMITE_DRE, 'A1', rows);
}

/**
 * Persiste relatório DFC em RPT_COMITE_DFC
 */
export function persistDFCReport(report: DFCReport): void {
  createSheetIfNotExists(Sheets.RPT_COMITE_DFC, ['Item', 'Valor']);
  const rows: any[][] = [
    ['DFC', 'Valor'],
    ['Saldo Inicial', formatMoney(report.saldoInicial)],
    ['Entradas Operacionais', formatMoney(report.entradasOperacionais)],
    ['Saidas Operacionais', formatMoney(report.saidasOperacionais)],
    ['Entradas Investimento', formatMoney(report.entradasInvestimento)],
    ['Saidas Investimento', formatMoney(report.saidasInvestimento)],
    ['Entradas Financiamento', formatMoney(report.entradasFinanciamento)],
    ['Saidas Financiamento', formatMoney(report.saidasFinanciamento)],
    ['Saldo Final', formatMoney(report.saldoFinal)],
  ];

  clearRange(Sheets.RPT_COMITE_DFC, 'A1:Z');
  setSheetValues(Sheets.RPT_COMITE_DFC, 'A1', rows);
}

/**
 * Persiste relatório KPIs em RPT_COMITE_KPIS
 */
export function persistKPIReport(report: KPIReport): void {
  createSheetIfNotExists(Sheets.RPT_COMITE_KPIS, ['KPI', 'Valor', 'Faixa', 'Unidade']);
  const rows: any[][] = [
    ['KPI', 'Valor', 'Faixa', 'Unidade'],
    ...report.kpis.map((kpi) => [kpi.metric, kpi.value, kpi.faixa, kpi.unit]),
  ];

  clearRange(Sheets.RPT_COMITE_KPIS, 'A1:Z');
  setSheetValues(Sheets.RPT_COMITE_KPIS, 'A1', rows);
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const month = String(period.month).padStart(2, '0');
  const name = `NeoFinance_${period.year}-${month}.pdf`;
  const blob = ss.getBlob().setName(name);
  const folder = getOrCreateFolder_('NeoFinance Exports');
  const file = folder.createFile(blob);
  return file.getUrl();
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

function getOrCreateFolder_(name: string): GoogleAppsScript.Drive.Folder {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(name);
}
