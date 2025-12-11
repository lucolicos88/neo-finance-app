/**
 * ledger-service.ts
 *
 * Gerencia lançamentos financeiros e contábeis.
 * CRUD de lançamentos na aba TB_LANCAMENTOS.
 *
 * Responsabilidades:
 * - Criar, atualizar, listar e deletar lançamentos
 * - Validar lançamentos antes de gravar
 * - Controlar períodos fechados (lock)
 * - Gerar IDs únicos para lançamentos
 */

import {
  getSheetValues,
  appendRows,
  updateRow,
  findRowByColumnValue,
} from '../shared/sheets-client';
import { Sheets, TB_LANCAMENTOS_COLS } from '../config/sheet-mapping';
import {
  LedgerEntry,
  LedgerFilter,
  LedgerEntryType,
  LedgerEntryStatus,
  LedgerEntryOrigin,
  Period,
} from '../shared/types';
import { validateLedgerEntry, isPeriodLocked } from '../shared/validation';
import { formatDateISO, parseDateISO, dateToPeriod } from '../shared/date-utils';
import { ConfigService } from './config-service';

// ============================================================================
// GERAÇÃO DE IDs
// ============================================================================

/**
 * Gera ID único para lançamento (formato: L2025-000001)
 *
 * TODO: Implementar geração sequencial thread-safe
 * TODO: Considerar uso de Lock Service para evitar colisão
 */
function generateLedgerId(): string {
  const year = new Date().getFullYear();
  const random = Math.floor(Math.random() * 999999)
    .toString()
    .padStart(6, '0');
  return `L${year}-${random}`;
}

// ============================================================================
// CONVERSÃO ENTRE SHEET E OBJETO
// ============================================================================

/**
 * Converte linha da planilha para objeto LedgerEntry
 */
function rowToLedgerEntry(row: any[]): LedgerEntry {
  return {
    id: row[TB_LANCAMENTOS_COLS.ID],
    competencia: parseDateISO(row[TB_LANCAMENTOS_COLS.DATA_COMPETENCIA]) || new Date(),
    vencimento: parseDateISO(row[TB_LANCAMENTOS_COLS.DATA_VENCIMENTO]),
    pagamento: parseDateISO(row[TB_LANCAMENTOS_COLS.DATA_PAGAMENTO]),
    tipo: row[TB_LANCAMENTOS_COLS.TIPO] as LedgerEntryType,
    filial: row[TB_LANCAMENTOS_COLS.FILIAL],
    centroCusto: row[TB_LANCAMENTOS_COLS.CENTRO_CUSTO] || null,
    contaGerencial: row[TB_LANCAMENTOS_COLS.CONTA_GERENCIAL],
    contaContabil: row[TB_LANCAMENTOS_COLS.CONTA_CONTABIL] || null,
    grupoReceita: row[TB_LANCAMENTOS_COLS.GRUPO_RECEITA] || null,
    canal: row[TB_LANCAMENTOS_COLS.CANAL] || null,
    descricao: row[TB_LANCAMENTOS_COLS.DESCRICAO],
    valorBruto: parseFloat(row[TB_LANCAMENTOS_COLS.VALOR_BRUTO]) || 0,
    desconto: parseFloat(row[TB_LANCAMENTOS_COLS.DESCONTO]) || 0,
    juros: parseFloat(row[TB_LANCAMENTOS_COLS.JUROS]) || 0,
    multa: parseFloat(row[TB_LANCAMENTOS_COLS.MULTA]) || 0,
    valorLiquido: parseFloat(row[TB_LANCAMENTOS_COLS.VALOR_LIQUIDO]) || 0,
    status: row[TB_LANCAMENTOS_COLS.STATUS] as LedgerEntryStatus,
    idExtratoBanco: row[TB_LANCAMENTOS_COLS.ID_EXTRATO_BANCO] || null,
    origem: row[TB_LANCAMENTOS_COLS.ORIGEM] as LedgerEntryOrigin,
    observacoes: row[TB_LANCAMENTOS_COLS.OBSERVACOES] || undefined,
  };
}

/**
 * Converte objeto LedgerEntry para linha da planilha
 */
function ledgerEntryToRow(entry: LedgerEntry): any[] {
  const row = new Array(21).fill('');

  row[TB_LANCAMENTOS_COLS.ID] = entry.id;
  row[TB_LANCAMENTOS_COLS.DATA_COMPETENCIA] = formatDateISO(entry.competencia);
  row[TB_LANCAMENTOS_COLS.DATA_VENCIMENTO] = entry.vencimento ? formatDateISO(entry.vencimento) : '';
  row[TB_LANCAMENTOS_COLS.DATA_PAGAMENTO] = entry.pagamento ? formatDateISO(entry.pagamento) : '';
  row[TB_LANCAMENTOS_COLS.TIPO] = entry.tipo;
  row[TB_LANCAMENTOS_COLS.FILIAL] = entry.filial;
  row[TB_LANCAMENTOS_COLS.CENTRO_CUSTO] = entry.centroCusto || '';
  row[TB_LANCAMENTOS_COLS.CONTA_GERENCIAL] = entry.contaGerencial;
  row[TB_LANCAMENTOS_COLS.CONTA_CONTABIL] = entry.contaContabil || '';
  row[TB_LANCAMENTOS_COLS.GRUPO_RECEITA] = entry.grupoReceita || '';
  row[TB_LANCAMENTOS_COLS.CANAL] = entry.canal || '';
  row[TB_LANCAMENTOS_COLS.DESCRICAO] = entry.descricao;
  row[TB_LANCAMENTOS_COLS.VALOR_BRUTO] = entry.valorBruto;
  row[TB_LANCAMENTOS_COLS.DESCONTO] = entry.desconto;
  row[TB_LANCAMENTOS_COLS.JUROS] = entry.juros;
  row[TB_LANCAMENTOS_COLS.MULTA] = entry.multa;
  row[TB_LANCAMENTOS_COLS.VALOR_LIQUIDO] = entry.valorLiquido;
  row[TB_LANCAMENTOS_COLS.STATUS] = entry.status;
  row[TB_LANCAMENTOS_COLS.ID_EXTRATO_BANCO] = entry.idExtratoBanco || '';
  row[TB_LANCAMENTOS_COLS.ORIGEM] = entry.origem;
  row[TB_LANCAMENTOS_COLS.OBSERVACOES] = entry.observacoes || '';

  return row;
}

// ============================================================================
// CRUD OPERATIONS
// ============================================================================

/**
 * Cria um novo lançamento
 *
 * @param entry - Lançamento a criar (id será gerado automaticamente)
 * @returns Lançamento criado com ID
 */
export function createEntry(entry: Omit<LedgerEntry, 'id'>): LedgerEntry {
  // Valida lançamento
  const maxDaysBack = ConfigService.getMaxDiasRetroativo();
  const validation = validateLedgerEntry(entry as LedgerEntry, maxDaysBack);

  if (!validation.valid) {
    throw new Error(`Validação falhou: ${validation.errors.join(', ')}`);
  }

  // Verifica se período está fechado
  const period = dateToPeriod(entry.competencia);
  if (isPeriodLocked(period)) {
    throw new Error(`Período ${period.year}-${period.month} está fechado para lançamentos`);
  }

  // Gera ID
  const newEntry: LedgerEntry = {
    ...entry,
    id: generateLedgerId(),
  } as LedgerEntry;

  // Converte para linha e adiciona na planilha
  const row = ledgerEntryToRow(newEntry);
  appendRows(Sheets.TB_LANCAMENTOS, [row]);

  return newEntry;
}

/**
 * Atualiza um lançamento existente
 *
 * @param id - ID do lançamento
 * @param updates - Campos a atualizar
 */
export function updateEntry(id: string, updates: Partial<LedgerEntry>): void {
  // Busca lançamento atual
  const current = getEntryById(id);
  if (!current) {
    throw new Error(`Lançamento ${id} não encontrado`);
  }

  // Verifica se período está fechado
  const period = dateToPeriod(current.competencia);
  if (isPeriodLocked(period)) {
    throw new Error(`Período ${period.year}-${period.month} está fechado para edição`);
  }

  // Mescla updates
  const updated: LedgerEntry = {
    ...current,
    ...updates,
    id, // Garante que ID não muda
  };

  // Valida
  const maxDaysBack = ConfigService.getMaxDiasRetroativo();
  const validation = validateLedgerEntry(updated, maxDaysBack);

  if (!validation.valid) {
    throw new Error(`Validação falhou: ${validation.errors.join(', ')}`);
  }

  // Encontra linha na planilha
  const rowIndex = findRowByColumnValue(Sheets.TB_LANCAMENTOS, TB_LANCAMENTOS_COLS.ID, id);
  if (!rowIndex) {
    throw new Error(`Linha do lançamento ${id} não encontrada`);
  }

  // Atualiza
  const row = ledgerEntryToRow(updated);
  updateRow(Sheets.TB_LANCAMENTOS, rowIndex, row);
}

/**
 * Busca lançamento por ID
 */
export function getEntryById(id: string): LedgerEntry | null {
  const values = getSheetValues(Sheets.TB_LANCAMENTOS, { skipHeader: true });

  for (const row of values) {
    if (row[TB_LANCAMENTOS_COLS.ID] === id) {
      return rowToLedgerEntry(row);
    }
  }

  return null;
}

/**
 * Lista lançamentos com filtros
 *
 * TODO: Implementar paginação para grandes volumes
 * TODO: Otimizar com índices ou cache
 */
export function listEntries(filter: LedgerFilter = {}): LedgerEntry[] {
  const values = getSheetValues(Sheets.TB_LANCAMENTOS, { skipHeader: true });
  const entries: LedgerEntry[] = [];

  for (const row of values) {
    if (!row || row.length === 0) continue;

    const entry = rowToLedgerEntry(row);

    // Aplica filtros
    if (filter.filial && entry.filial !== filter.filial) continue;
    if (filter.canal && entry.canal !== filter.canal) continue;
    if (filter.centroCusto && entry.centroCusto !== filter.centroCusto) continue;
    if (filter.tipo && entry.tipo !== filter.tipo) continue;
    if (filter.status && entry.status !== filter.status) continue;
    if (filter.contaGerencial && entry.contaGerencial !== filter.contaGerencial) continue;

    // TODO: Filtrar por período (periodStart, periodEnd)

    entries.push(entry);
  }

  return entries;
}

/**
 * Cancela um lançamento (muda status para CANCELADO)
 */
export function cancelEntry(id: string, motivo?: string): void {
  updateEntry(id, {
    status: LedgerEntryStatus.CANCELADO,
    observacoes: motivo,
  });
}

/**
 * Marca lançamento como realizado
 *
 * @param id - ID do lançamento
 * @param dataPagamento - Data de pagamento
 */
export function markAsRealized(id: string, dataPagamento: Date): void {
  updateEntry(id, {
    status: LedgerEntryStatus.REALIZADO,
    pagamento: dataPagamento,
  });
}

// ============================================================================
// CONTROLE DE PERÍODOS
// ============================================================================

/**
 * Bloqueia um período para lançamentos
 *
 * TODO: Implementar tabela de períodos fechados
 */
export function lockPeriod(period: Period): void {
  // TODO: Adicionar período na tabela de períodos fechados
  throw new Error('lockPeriod não implementado');
}

/**
 * Desbloqueia um período
 *
 * TODO: Implementar tabela de períodos fechados
 */
export function unlockPeriod(period: Period): void {
  // TODO: Remover período da tabela de períodos fechados
  throw new Error('unlockPeriod não implementado');
}
