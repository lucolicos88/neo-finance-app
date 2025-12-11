/**
 * dre-service.ts
 *
 * Gerencia DRE (Demonstrativo de Resultados do Exercício) gerencial.
 *
 * Responsabilidades:
 * - Calcular DRE por período, filial, consolidado
 * - Mapear lançamentos para grupos DRE
 * - Validar consistência com lançamentos
 * - Persistir em TB_DRE_MENSAL e TB_DRE_RESUMO
 */

import { getSheetValues, setSheetValues, clearRange } from '../shared/sheets-client';
import { Sheets } from '../config/sheet-mapping';
import {
  DRELine,
  Period,
  BranchId,
  Money,
  LedgerEntry,
  LedgerEntryStatus,
  ReportFilter,
} from '../shared/types';
import { listEntries } from './ledger-service';
import { getAccountByCode } from './reference-data-service';
import { sumMoney, calculatePercentage } from '../shared/money-utils';

// ============================================================================
// GRUPOS DRE
// ============================================================================

/**
 * Estrutura hierárquica de grupos DRE
 */
export const DRE_GROUPS = {
  RECEITA_BRUTA: 'RECEITA_BRUTA',
  DEDUCOES: 'DEDUCOES',
  RECEITA_LIQUIDA: 'RECEITA_LIQUIDA',
  CUSTOS: 'CUSTOS',
  LUCRO_BRUTO: 'LUCRO_BRUTO',
  DESPESAS_OPERACIONAIS: 'DESPESAS_OPERACIONAIS',
  EBITDA: 'EBITDA',
  DEPRECIACAO: 'DEPRECIACAO',
  EBIT: 'EBIT',
  RESULTADO_FINANCEIRO: 'RESULTADO_FINANCEIRO',
  LUCRO_LIQUIDO: 'LUCRO_LIQUIDO',
} as const;

/**
 * Subgrupos de despesas operacionais
 */
export const DRE_SUBGROUPS = {
  PESSOAL: 'PESSOAL',
  MARKETING: 'MARKETING',
  INFRAESTRUTURA: 'INFRAESTRUTURA',
  ADMINISTRATIVA: 'ADMINISTRATIVA',
  TRIBUTARIA: 'TRIBUTARIA',
} as const;

// ============================================================================
// CÁLCULO DE DRE
// ============================================================================

/**
 * Estrutura de DRE calculado
 */
export interface DREStatement {
  period: Period;
  branchId: BranchId | null;
  lines: DRELine[];
  summary: {
    receitaBruta: Money;
    receitaLiquida: Money;
    custos: Money;
    lucroBruto: Money;
    despesasOperacionais: Money;
    ebitda: Money;
    ebitdaPct: number;
    lucroLiquido: Money;
    margemLiquida: number;
  };
}

/**
 * Calcula DRE para um período
 *
 * @param period - Período a calcular
 * @param branchId - ID da filial (null = consolidado)
 * @returns DRE calculado
 *
 * TODO: Implementar cálculos reais com base no plano de contas
 * TODO: Considerar rateios entre filiais
 * TODO: Calcular impostos e deduções
 */
export function calculateDRE(period: Period, branchId: BranchId | null = null): DREStatement {
  // Busca lançamentos realizados do período
  const entries = listEntries({
    status: LedgerEntryStatus.REALIZADO,
    // TODO: Filtrar por período (competência)
    ...(branchId && { filial: branchId }),
  });

  const lines: DRELine[] = [];
  const groupTotals = new Map<string, Money>();

  // Agrupa lançamentos por grupo DRE
  for (const entry of entries) {
    const account = getAccountByCode(entry.contaGerencial);
    if (!account) continue;

    const group = account.grupoDRE;
    const subGroup = account.subgrupoDRE;

    const current = groupTotals.get(group) || 0;
    groupTotals.set(group, current + entry.valorLiquido);

    const line: DRELine = {
      period,
      branchId,
      group,
      subGroup,
      value: entry.valorLiquido,
    };

    lines.push(line);
  }

  // Calcula resumo
  const receitaBruta = groupTotals.get(DRE_GROUPS.RECEITA_BRUTA) || 0;
  const deducoes = groupTotals.get(DRE_GROUPS.DEDUCOES) || 0;
  const receitaLiquida = receitaBruta - deducoes;

  const custos = groupTotals.get(DRE_GROUPS.CUSTOS) || 0;
  const lucroBruto = receitaLiquida - custos;

  const despesasOperacionais = groupTotals.get(DRE_GROUPS.DESPESAS_OPERACIONAIS) || 0;
  const ebitda = lucroBruto - despesasOperacionais;
  const ebitdaPct = calculatePercentage(ebitda, receitaLiquida);

  const depreciacao = groupTotals.get(DRE_GROUPS.DEPRECIACAO) || 0;
  const ebit = ebitda - depreciacao;

  const resultadoFinanceiro = groupTotals.get(DRE_GROUPS.RESULTADO_FINANCEIRO) || 0;
  const lucroLiquido = ebit + resultadoFinanceiro;
  const margemLiquida = calculatePercentage(lucroLiquido, receitaLiquida);

  return {
    period,
    branchId,
    lines,
    summary: {
      receitaBruta,
      receitaLiquida,
      custos,
      lucroBruto,
      despesasOperacionais,
      ebitda,
      ebitdaPct,
      lucroLiquido,
      margemLiquida,
    },
  };
}

/**
 * Calcula DRE para múltiplas filiais (consolidado + individuais)
 *
 * @param period - Período a calcular
 * @param includeConsolidated - Se true, inclui DRE consolidado
 * @returns Array de DREs
 */
export function calculateMultiBranchDRE(
  period: Period,
  includeConsolidated: boolean = true
): DREStatement[] {
  const statements: DREStatement[] = [];

  // DRE consolidado
  if (includeConsolidated) {
    statements.push(calculateDRE(period, null));
  }

  // DRE por filial
  // TODO: Buscar lista de filiais ativas e calcular para cada uma
  // const branches = getActiveBranches();
  // for (const branch of branches) {
  //   statements.push(calculateDRE(period, branch.id));
  // }

  return statements;
}

// ============================================================================
// PERSISTÊNCIA
// ============================================================================

/**
 * Persiste DRE na aba TB_DRE_MENSAL
 *
 * TODO: Implementar lógica de merge (atualizar apenas o período específico)
 */
export function persistDREMensal(statement: DREStatement): void {
  // Agrupa linhas por grupo/subgrupo
  const aggregated = new Map<string, Money>();

  for (const line of statement.lines) {
    const key = `${line.group}|${line.subGroup || ''}`;
    const current = aggregated.get(key) || 0;
    aggregated.set(key, current + line.value);
  }

  // Prepara rows
  const rows: any[][] = [];
  for (const [key, value] of aggregated.entries()) {
    const [group, subGroup] = key.split('|');

    rows.push([
      statement.period.year,
      statement.period.month,
      statement.branchId || '',
      group,
      subGroup || '',
      value,
    ]);
  }

  // TODO: Encontrar range correto para atualizar apenas este período/filial
  // clearRange(Sheets.TB_DRE_MENSAL, 'A2:Z');
  // setSheetValues(Sheets.TB_DRE_MENSAL, 'A2', rows);
}

/**
 * Persiste resumo DRE na aba TB_DRE_RESUMO
 */
export function persistDREResumo(statement: DREStatement): void {
  const rows: any[][] = [
    [statement.period.year, statement.period.month, 'RECEITA_BRUTA', statement.summary.receitaBruta],
    [statement.period.year, statement.period.month, 'RECEITA_LIQUIDA', statement.summary.receitaLiquida],
    [statement.period.year, statement.period.month, 'CUSTOS', statement.summary.custos],
    [statement.period.year, statement.period.month, 'LUCRO_BRUTO', statement.summary.lucroBruto],
    [statement.period.year, statement.period.month, 'DESPESAS_OPERACIONAIS', statement.summary.despesasOperacionais],
    [statement.period.year, statement.period.month, 'EBITDA', statement.summary.ebitda],
    [statement.period.year, statement.period.month, 'EBITDA_PCT', statement.summary.ebitdaPct],
    [statement.period.year, statement.period.month, 'LUCRO_LIQUIDO', statement.summary.lucroLiquido],
    [statement.period.year, statement.period.month, 'MARGEM_LIQUIDA', statement.summary.margemLiquida],
  ];

  // TODO: Atualizar apenas este período
  // setSheetValues(Sheets.TB_DRE_RESUMO, 'A2', rows);
}

// ============================================================================
// VALIDAÇÃO
// ============================================================================

/**
 * Valida DRE contra lançamentos (sanity check)
 *
 * Verifica se totais de receita/despesa batem com lançamentos
 *
 * TODO: Implementar validação completa
 */
export function validateDREAgainstLedger(statement: DREStatement): boolean {
  // TODO: Somar todos os lançamentos do período e comparar com DRE
  return true;
}

// ============================================================================
// HELPERS
// ============================================================================

/**
 * Obtém DRE de um período específico (lê da aba TB_DRE_RESUMO)
 */
export function getDREStatement(filter: ReportFilter): DREStatement | null {
  // TODO: Ler TB_DRE_RESUMO e reconstruir objeto DREStatement
  return null;
}
