/**
 * cashflow-service.ts
 *
 * Gerencia fluxo de caixa realizado e projetado.
 *
 * Responsabilidades:
 * - Calcular fluxo de caixa realizado (com base em lançamentos pagos)
 * - Calcular fluxo de caixa projetado (com base em lançamentos previstos)
 * - Gerar timeline de contas futuras
 * - Calcular saldo projetado
 */

import { getSheetValues, setSheetValues, clearRange, createSheetIfNotExists } from '../shared/sheets-client';
import { Sheets } from '../config/sheet-mapping';
import {
  CashflowLine,
  CashflowCategory,
  CashflowType,
  Period,
  Money,
  LedgerEntry,
  LedgerEntryStatus,
  LedgerEntryType,
} from '../shared/types';
import { listEntries } from './ledger-service';
import { sumMoney } from '../shared/money-utils';
import { formatDateISO, generatePeriodRange, getFirstDayOfPeriod } from '../shared/date-utils';
import { getAccountByCode } from './reference-data-service';

// ============================================================================
// CÁLCULO DE FLUXO DE CAIXA REALIZADO
// ============================================================================

/**
 * Calcula fluxo de caixa realizado para um período
 *
 * Usa lançamentos com status REALIZADO e data_pagamento no período
 *
 * @param period - Período a calcular
 * @returns Array de linhas de fluxo de caixa
 *
 * TODO: Agrupar por categoria (OPERACIONAL, INVESTIMENTO, FINANCIAMENTO)
 * TODO: Otimizar performance para grandes volumes
 */
export function calculateRealCashflow(period: Period): CashflowLine[] {
  // Busca lançamentos realizados
  const entries = listEntries({ status: LedgerEntryStatus.REALIZADO });

  const lines: CashflowLine[] = [];

  for (const entry of entries) {
    if (!entry.pagamento) continue;

    // Filtra por período
    const entryYear = entry.pagamento.getFullYear();
    const entryMonth = entry.pagamento.getMonth() + 1;

    if (entryYear !== period.year || entryMonth !== period.month) {
      continue;
    }

    // Determina categoria baseado na conta gerencial
    const accountCode = entry.contaContabil || entry.contaGerencial;
    const account = accountCode ? getAccountByCode(accountCode) : null;
    const category: CashflowCategory = account?.grupoDFC || CashflowCategory.OPERACIONAL;

    // Determina tipo (entrada ou saída)
    const type: CashflowType =
      entry.tipo === LedgerEntryType.RECEBER ? CashflowType.ENTRADA : CashflowType.SAIDA;

    const line: CashflowLine = {
      date: entry.pagamento,
      type,
      category,
      description: entry.descricao,
      value: entry.valorLiquido,
      projected: false,
      contaBancaria: undefined, // TODO: Vincular com conta bancária se disponível
    };

    lines.push(line);
  }

  return lines;
}

/**
 * Persiste fluxo de caixa realizado na aba TB_DFC_REAL
 *
 * TODO: Implementar lógica de merge (não sobrescrever tudo)
 */
export function persistRealCashflow(period: Period, lines: CashflowLine[]): void {
  const headers = ['Data', 'Categoria', 'Tipo', 'Descricao', 'Valor', 'Conta'];
  createSheetIfNotExists(Sheets.TB_DFC_REAL, headers);

  const rows = lines.map((line) => [
    formatDateISO(line.date),
    line.category,
    line.type,
    line.description,
    line.value,
    line.contaBancaria || '',
  ]);

  const existing = getSheetValues(Sheets.TB_DFC_REAL);
  const filtered = existing.length > 1
    ? existing.slice(1).filter((r) => {
        const raw = String(r[0] || '');
        const parts = raw.split('-');
        if (parts.length < 2) return true;
        const year = Number(parts[0]);
        const month = Number(parts[1]);
        return !(year === period.year && month === period.month);
      })
    : [];

  const allRows = [headers, ...filtered, ...rows];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Sheets.TB_DFC_REAL);
  if (!sheet) return;
  sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);
  const lastRow = sheet.getLastRow();
  if (lastRow > allRows.length) {
    sheet.getRange(allRows.length + 1, 1, lastRow - allRows.length, headers.length).clearContent();
  }
}

// ============================================================================
// CÁLCULO DE FLUXO DE CAIXA PROJETADO
// ============================================================================

/**
 * Parâmetros para projeção de fluxo de caixa
 */
export interface CashflowProjectionParams {
  horizonMonths: number; // Quantos meses projetar
  includePrevisto: boolean; // Incluir lançamentos previstos
  saldoInicial: Money; // Saldo inicial
}

/**
 * Calcula fluxo de caixa projetado
 *
 * @param startPeriod - Período inicial
 * @param params - Parâmetros de projeção
 * @returns Array de linhas de fluxo de caixa projetado
 *
 * TODO: Implementar projeções baseadas em histórico
 * TODO: Considerar sazonalidade
 * TODO: Permitir ajustes manuais
 */
export function calculateForecastCashflow(
  startPeriod: Period,
  params: CashflowProjectionParams
): CashflowLine[] {
  const lines: CashflowLine[] = [];

  // Busca lançamentos previstos
  const entries = listEntries({ status: LedgerEntryStatus.PREVISTO });

  // Gera períodos de projeção
  const endPeriod: Period = {
    year: startPeriod.year,
    month: startPeriod.month + params.horizonMonths - 1,
  };

  // Ajusta ano se necessário
  while (endPeriod.month > 12) {
    endPeriod.month -= 12;
    endPeriod.year += 1;
  }

  const periods = generatePeriodRange(startPeriod, endPeriod);

  for (const period of periods) {
    // Filtra lançamentos previstos do período
    const periodEntries = entries.filter((entry) => {
      if (!entry.vencimento) return false;

      const entryYear = entry.vencimento.getFullYear();
      const entryMonth = entry.vencimento.getMonth() + 1;

      return entryYear === period.year && entryMonth === period.month;
    });

    for (const entry of periodEntries) {
      const accountCode = entry.contaContabil || entry.contaGerencial;
      const account = accountCode ? getAccountByCode(accountCode) : null;
      const category: CashflowCategory = account?.grupoDFC || CashflowCategory.OPERACIONAL;

      const type: CashflowType =
        entry.tipo === LedgerEntryType.RECEBER ? CashflowType.ENTRADA : CashflowType.SAIDA;

      const line: CashflowLine = {
        date: entry.vencimento!,
        type,
        category,
        description: entry.descricao,
        value: entry.valorLiquido,
        projected: true,
      };

      lines.push(line);
    }
  }

  return lines;
}

/**
 * Persiste fluxo de caixa projetado na aba TB_DFC_PROJ
 *
 * TODO: Implementar estrutura agregada por mês (não dia a dia)
 */
export function persistForecastCashflow(lines: CashflowLine[]): void {
  const headers = ['Ano', 'Mes', 'Categoria', 'Tipo', 'Valor'];
  createSheetIfNotExists(Sheets.TB_DFC_PROJ, headers);

  const aggregated = new Map<string, Money>();
  for (const line of lines) {
    const year = line.date.getFullYear();
    const month = line.date.getMonth() + 1;
    const key = `${year}|${month}|${line.category}|${line.type}`;
    aggregated.set(key, (aggregated.get(key) || 0) + line.value);
  }

  const rows = Array.from(aggregated.entries()).map(([key, value]) => {
    const [year, month, category, type] = key.split('|');
    return [Number(year), Number(month), category, type, value];
  });

  clearRange(Sheets.TB_DFC_PROJ, 'A1:Z');
  setSheetValues(Sheets.TB_DFC_PROJ, 'A1', [headers, ...rows]);
}

// ============================================================================
// TIMELINE E SALDOS
// ============================================================================

/**
 * Gera timeline de contas futuras (a pagar e a receber)
 *
 * @param horizonDays - Quantos dias à frente projetar
 * @returns Array de lançamentos futuros ordenados por vencimento
 */
export function getFutureAccountsTimeline(horizonDays: number = 90): LedgerEntry[] {
  const entries = listEntries({ status: LedgerEntryStatus.PREVISTO });

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const futureDate = new Date(today);
  futureDate.setDate(futureDate.getDate() + horizonDays);

  // Filtra por vencimento
  const futureEntries = entries.filter((entry) => {
    if (!entry.vencimento) return false;

    return entry.vencimento >= today && entry.vencimento <= futureDate;
  });

  // Ordena por vencimento
  futureEntries.sort((a, b) => {
    if (!a.vencimento || !b.vencimento) return 0;
    return a.vencimento.getTime() - b.vencimento.getTime();
  });

  return futureEntries;
}

/**
 * Calcula saldo projetado
 *
 * @param saldoInicial - Saldo inicial
 * @param period - Período a calcular
 * @returns Saldo projetado ao final do período
 *
 * TODO: Considerar saldo por conta bancária
 */
export function calculateProjectedBalance(saldoInicial: Money, period: Period): Money {
  const cashflowLines = calculateForecastCashflow(period, {
    horizonMonths: 1,
    includePrevisto: true,
    saldoInicial,
  });

  let balance = saldoInicial;

  for (const line of cashflowLines) {
    if (line.type === CashflowType.ENTRADA) {
      balance += line.value;
    } else {
      balance -= line.value;
    }
  }

  return balance;
}
