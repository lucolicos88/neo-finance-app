/**
 * reconciliation-service.ts
 *
 * Gerencia conciliação entre lançamentos e extratos bancários.
 *
 * Responsabilidades:
 * - Importar extratos bancários
 * - Conciliar automaticamente lançamentos com extratos
 * - Sugerir matches baseado em valor e data
 * - Permitir conciliação manual
 */

import { getSheetValues, appendRows, updateRow } from '../shared/sheets-client';
import { Sheets, TB_EXTRATOS_COLS, TB_LANCAMENTOS_COLS } from '../config/sheet-mapping';
import { BankStatement, LedgerEntry, Money } from '../shared/types';
import { moneyEquals } from '../shared/money-utils';
import { diffDays, formatDateISO, parseDateISO } from '../shared/date-utils';
import { ConfigService } from './config-service';
import { getEntryById, updateEntry } from './ledger-service';

// ============================================================================
// CONVERSÃO ENTRE SHEET E OBJETO
// ============================================================================

/**
 * Converte linha da planilha para BankStatement
 */
function rowToBankStatement(row: any[]): BankStatement {
  return {
    id: row[TB_EXTRATOS_COLS.ID],
    dataMovimento: parseDateISO(row[TB_EXTRATOS_COLS.DATA_MOVIMENTO]) || new Date(),
    contaBancaria: row[TB_EXTRATOS_COLS.CONTA_BANCARIA],
    historico: row[TB_EXTRATOS_COLS.HISTORICO],
    documento: row[TB_EXTRATOS_COLS.DOCUMENTO] || null,
    valor: parseFloat(row[TB_EXTRATOS_COLS.VALOR]) || 0,
    saldoApos: row[TB_EXTRATOS_COLS.SALDO_APOS] ? parseFloat(row[TB_EXTRATOS_COLS.SALDO_APOS]) : null,
    conciliado: row[TB_EXTRATOS_COLS.CONCILIADO] === true || row[TB_EXTRATOS_COLS.CONCILIADO] === 'TRUE',
    idLancamento: row[TB_EXTRATOS_COLS.ID_LANCAMENTO] || null,
  };
}

/**
 * Converte BankStatement para linha da planilha
 */
function bankStatementToRow(statement: BankStatement): any[] {
  const row = new Array(9).fill('');

  row[TB_EXTRATOS_COLS.ID] = statement.id;
  row[TB_EXTRATOS_COLS.DATA_MOVIMENTO] = formatDateISO(statement.dataMovimento);
  row[TB_EXTRATOS_COLS.CONTA_BANCARIA] = statement.contaBancaria;
  row[TB_EXTRATOS_COLS.HISTORICO] = statement.historico;
  row[TB_EXTRATOS_COLS.DOCUMENTO] = statement.documento || '';
  row[TB_EXTRATOS_COLS.VALOR] = statement.valor;
  row[TB_EXTRATOS_COLS.SALDO_APOS] = statement.saldoApos || '';
  row[TB_EXTRATOS_COLS.CONCILIADO] = statement.conciliado;
  row[TB_EXTRATOS_COLS.ID_LANCAMENTO] = statement.idLancamento || '';

  return row;
}

// ============================================================================
// IMPORTAÇÃO DE EXTRATOS
// ============================================================================

/**
 * Gera ID único para extrato (formato: EB2025-000001)
 *
 * TODO: Implementar geração sequencial thread-safe
 */
function generateBankStatementId(): string {
  const year = new Date().getFullYear();
  const random = Math.floor(Math.random() * 999999)
    .toString()
    .padStart(6, '0');
  return `EB${year}-${random}`;
}

/**
 * Importa extrato bancário de um arquivo ou array de dados
 *
 * @param statements - Array de extratos a importar
 * @returns Quantidade de extratos importados
 *
 * TODO: Implementar parsing de OFX, CSV, Excel
 * TODO: Detectar duplicatas antes de importar
 */
export function importBankStatement(statements: Omit<BankStatement, 'id'>[]): number {
  const rows = statements.map((stmt) => {
    const full: BankStatement = {
      ...stmt,
      id: generateBankStatementId(),
      conciliado: false,
      idLancamento: null,
    };
    return bankStatementToRow(full);
  });

  appendRows(Sheets.TB_EXTRATOS, rows);

  return statements.length;
}

/**
 * Lista todos os extratos
 */
export function listBankStatements(onlyConciliado?: boolean): BankStatement[] {
  const values = getSheetValues(Sheets.TB_EXTRATOS, { skipHeader: true });
  const statements: BankStatement[] = [];

  for (const row of values) {
    if (!row || row.length === 0) continue;

    const stmt = rowToBankStatement(row);

    if (onlyConciliado !== undefined && stmt.conciliado !== onlyConciliado) {
      continue;
    }

    statements.push(stmt);
  }

  return statements;
}

/**
 * Busca extrato por ID
 */
export function getBankStatementById(id: string): BankStatement | null {
  const values = getSheetValues(Sheets.TB_EXTRATOS, { skipHeader: true });

  for (const row of values) {
    if (row[TB_EXTRATOS_COLS.ID] === id) {
      return rowToBankStatement(row);
    }
  }

  return null;
}

// ============================================================================
// CONCILIAÇÃO AUTOMÁTICA
// ============================================================================

/**
 * Sugestão de match entre extrato e lançamento
 */
export interface MatchSuggestion {
  bankStatementId: string;
  ledgerEntryId: string;
  confidence: number; // 0-100
  reason: string;
}

/**
 * Sugere matches para um extrato bancário
 *
 * Critérios:
 * - Valor igual (com tolerância)
 * - Data próxima (±7 dias)
 * - Status do lançamento
 *
 * TODO: Implementar algoritmo mais sofisticado (fuzzy matching, ML)
 */
export function suggestMatches(statementId: string): MatchSuggestion[] {
  const statement = getBankStatementById(statementId);
  if (!statement) {
    return [];
  }

  // Busca lançamentos não conciliados
  const allEntries = getSheetValues(Sheets.TB_LANCAMENTOS, { skipHeader: true });
  const tolerance = ConfigService.getToleranciaConciliacao();
  const suggestions: MatchSuggestion[] = [];

  for (const row of allEntries) {
    if (!row || row.length === 0) continue;

    const entryId = row[TB_LANCAMENTOS_COLS.ID];
    const valorLiquido = parseFloat(row[TB_LANCAMENTOS_COLS.VALOR_LIQUIDO]) || 0;
    const dataPagamento = parseDateISO(row[TB_LANCAMENTOS_COLS.DATA_PAGAMENTO]);
    const status = row[TB_LANCAMENTOS_COLS.STATUS];
    const jaTemExtrato = row[TB_LANCAMENTOS_COLS.ID_EXTRATO_BANCO];

    // Ignora se já conciliado
    if (jaTemExtrato) continue;

    // Ignora se cancelado
    if (status === 'CANCELADO') continue;

    // Verifica valor
    if (!moneyEquals(statement.valor, valorLiquido, tolerance)) {
      continue;
    }

    let confidence = 50; // Base: valores iguais

    // Verifica data
    if (dataPagamento) {
      const daysDiff = Math.abs(diffDays(statement.dataMovimento, dataPagamento));

      if (daysDiff === 0) {
        confidence += 40; // Data exata
      } else if (daysDiff <= 3) {
        confidence += 30; // Até 3 dias
      } else if (daysDiff <= 7) {
        confidence += 10; // Até 7 dias
      }
    }

    // Verifica status
    if (status === 'REALIZADO') {
      confidence += 10;
    }

    suggestions.push({
      bankStatementId: statementId,
      ledgerEntryId: entryId,
      confidence,
      reason: `Valor igual, diferença de ${dataPagamento ? diffDays(statement.dataMovimento, dataPagamento) : '?'} dias`,
    });
  }

  // Ordena por confiança decrescente
  suggestions.sort((a, b) => b.confidence - a.confidence);

  return suggestions;
}

/**
 * Concilia automaticamente todos os extratos não conciliados
 *
 * @param minConfidence - Confiança mínima para conciliação automática (0-100)
 * @returns Quantidade de conciliações realizadas
 */
export function autoReconcile(minConfidence: number = 80): number {
  const statements = listBankStatements(false); // não conciliados
  let count = 0;

  for (const stmt of statements) {
    const suggestions = suggestMatches(stmt.id);

    // Pega a melhor sugestão se passar do threshold
    if (suggestions.length > 0 && suggestions[0].confidence >= minConfidence) {
      const best = suggestions[0];

      try {
        reconcile(stmt.id, best.ledgerEntryId);
        count++;
      } catch (error) {
        console.error(`Erro ao conciliar ${stmt.id} com ${best.ledgerEntryId}:`, error);
      }
    }
  }

  return count;
}

// ============================================================================
// CONCILIAÇÃO MANUAL
// ============================================================================

/**
 * Concilia manualmente um extrato com um lançamento
 *
 * @param statementId - ID do extrato
 * @param entryId - ID do lançamento
 */
export function reconcile(statementId: string, entryId: string): void {
  const statement = getBankStatementById(statementId);
  const entry = getEntryById(entryId);

  if (!statement) {
    throw new Error(`Extrato ${statementId} não encontrado`);
  }

  if (!entry) {
    throw new Error(`Lançamento ${entryId} não encontrado`);
  }

  // Atualiza extrato
  statement.conciliado = true;
  statement.idLancamento = entryId;

  // TODO: Atualizar via updateRow no sheets-client
  // Por enquanto, vamos usar a lógica básica

  // Atualiza lançamento
  updateEntry(entryId, {
    idExtratoBanco: statementId,
    pagamento: statement.dataMovimento, // Considera data do extrato como pagamento
  });

  console.log(`Conciliado: ${statementId} <-> ${entryId}`);
}

/**
 * Desconcilia um extrato
 */
export function unreconcile(statementId: string): void {
  const statement = getBankStatementById(statementId);

  if (!statement) {
    throw new Error(`Extrato ${statementId} não encontrado`);
  }

  if (!statement.conciliado) {
    return; // Já desconciliado
  }

  // Limpa lançamento
  if (statement.idLancamento) {
    updateEntry(statement.idLancamento, {
      idExtratoBanco: null,
    });
  }

  // Atualiza extrato
  statement.conciliado = false;
  statement.idLancamento = null;

  // TODO: Atualizar via updateRow no sheets-client

  console.log(`Desconciliado: ${statementId}`);
}
