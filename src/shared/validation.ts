/**
 * validation.ts
 *
 * Funções de validação de dados e regras de negócio.
 * Centraliza todas as validações usadas pelos serviços.
 */

import { Money, LedgerEntry, Period } from './types';
import { isFuture, isWithinRetroactiveLimit, diffDays } from './date-utils';
import { isPositive } from './money-utils';
import { getSheetValues } from './sheets-client';
import {
  SHEET_REF_FILIAIS,
  SHEET_REF_CANAIS,
  SHEET_REF_CCUSTO,
  SHEET_REF_PLANO_CONTAS,
} from '../config/sheet-mapping';

/**
 * Resultado de validação
 */
export interface ValidationResult {
  valid: boolean;
  errors: string[];
}

/**
 * Cria resultado de validação bem-sucedida
 */
export function validationSuccess(): ValidationResult {
  return { valid: true, errors: [] };
}

/**
 * Cria resultado de validação com erros
 */
export function validationError(errors: string[]): ValidationResult {
  return { valid: false, errors };
}

/**
 * Valida se um campo obrigatório está preenchido
 */
export function validateRequired(
  value: any,
  fieldName: string
): ValidationResult {
  if (value === null || value === undefined || value === '') {
    return validationError([`${fieldName} é obrigatório`]);
  }
  return validationSuccess();
}

/**
 * Valida se uma data é válida
 */
export function validateDate(
  date: Date | null | undefined,
  fieldName: string
): ValidationResult {
  const errors: string[] = [];

  if (!date) {
    errors.push(`${fieldName} é obrigatório`);
  } else if (!(date instanceof Date) || isNaN(date.getTime())) {
    errors.push(`${fieldName} não é uma data válida`);
  }

  return errors.length > 0 ? validationError(errors) : validationSuccess();
}

/**
 * Valida se uma data não está no futuro
 */
export function validateNotFuture(
  date: Date,
  fieldName: string
): ValidationResult {
  if (isFuture(date)) {
    return validationError([`${fieldName} não pode ser uma data futura`]);
  }
  return validationSuccess();
}

/**
 * Valida se uma data está dentro do limite retroativo
 */
export function validateRetroactiveLimit(
  date: Date,
  maxDaysBack: number,
  fieldName: string
): ValidationResult {
  if (!isWithinRetroactiveLimit(date, maxDaysBack)) {
    return validationError([
      `${fieldName} excede o limite de ${maxDaysBack} dias retroativos`,
    ]);
  }
  return validationSuccess();
}

/**
 * Valida valor monetário
 */
export function validateMoney(
  value: Money,
  fieldName: string,
  mustBePositive: boolean = true
): ValidationResult {
  const errors: string[] = [];

  if (value === null || value === undefined) {
    errors.push(`${fieldName} é obrigatório`);
  } else if (isNaN(value)) {
    errors.push(`${fieldName} deve ser um número válido`);
  } else if (mustBePositive && !isPositive(value)) {
    errors.push(`${fieldName} deve ser maior que zero`);
  }

  return errors.length > 0 ? validationError(errors) : validationSuccess();
}

/**
 * Valida se um valor está em uma lista de opções permitidas
 */
export function validateEnum<T>(
  value: T,
  allowedValues: T[],
  fieldName: string
): ValidationResult {
  if (!allowedValues.includes(value)) {
    return validationError([
      `${fieldName} deve ser um dos seguintes: ${allowedValues.join(', ')}`,
    ]);
  }
  return validationSuccess();
}

/**
 * Valida período (ano e mês)
 */
export function validatePeriod(period: Period): ValidationResult {
  const errors: string[] = [];

  if (!period.year || period.year < 2000 || period.year > 2100) {
    errors.push('Ano inválido');
  }

  if (!period.month || period.month < 1 || period.month > 12) {
    errors.push('Mês deve estar entre 1 e 12');
  }

  return errors.length > 0 ? validationError(errors) : validationSuccess();
}

/**
 * Valida lançamento financeiro completo
 *
 * TODO: Validar referências cruzadas (filial, canal, conta existem)
 */
export function validateLedgerEntry(
  entry: LedgerEntry,
  maxDaysBack: number = 7
): ValidationResult {
  const errors: string[] = [];

  // Data de competência
  const competenciaValidation = validateDate(entry.competencia, 'Data de competência');
  if (!competenciaValidation.valid) {
    errors.push(...competenciaValidation.errors);
  }

  // Data de pagamento não pode ser futura
  if (entry.pagamento) {
    const pagamentoValidation = validateNotFuture(entry.pagamento, 'Data de pagamento');
    if (!pagamentoValidation.valid) {
      errors.push(...pagamentoValidation.errors);
    }
  }

  // Se status REALIZADO, deve ter data de pagamento
  if (entry.status === 'REALIZADO' && !entry.pagamento) {
    errors.push('Lançamento realizado deve ter data de pagamento');
  }

  // Valor bruto deve ser positivo
  const valorBrutoValidation = validateMoney(entry.valorBruto, 'Valor bruto', true);
  if (!valorBrutoValidation.valid) {
    errors.push(...valorBrutoValidation.errors);
  }

  // Validar cálculo de valor líquido
  const expectedLiquido =
    entry.valorBruto - entry.desconto + entry.juros + entry.multa;
  if (Math.abs(entry.valorLiquido - expectedLiquido) > 0.01) {
    errors.push(
      `Valor líquido inconsistente. Esperado: ${expectedLiquido.toFixed(2)}, informado: ${entry.valorLiquido.toFixed(2)}`
    );
  }

  // Campos obrigatórios
  if (!entry.filial) errors.push('Filial é obrigatória');
  if (!entry.contaGerencial) errors.push('Conta gerencial é obrigatória');
  if (!entry.descricao) errors.push('Descrição é obrigatória');

  const refs = getReferenceSets();
  if (entry.filial && refs.filiais.size > 0 && !refs.filiais.has(String(entry.filial))) {
    errors.push(`Filial inválida: ${entry.filial}`);
  }
  if (entry.canal && refs.canais.size > 0 && !refs.canais.has(String(entry.canal))) {
    errors.push(`Canal inválido: ${entry.canal}`);
  }
  if (entry.centroCusto && refs.centrosCusto.size > 0 && !refs.centrosCusto.has(String(entry.centroCusto))) {
    errors.push(`Centro de custo inválido: ${entry.centroCusto}`);
  }
  if (entry.contaContabil && refs.contasContabeis.size > 0 && !refs.contasContabeis.has(String(entry.contaContabil))) {
    errors.push(`Conta contábil inválida: ${entry.contaContabil}`);
  }

  return errors.length > 0 ? validationError(errors) : validationSuccess();
}

/**
 * Valida se um período está fechado (bloqueado para edição)
 *
 * TODO: Implementar lógica de períodos fechados via config ou tabela
 */
export function isPeriodLocked(period: Period): boolean {
  const locked = getLockedPeriods();
  if (locked.size === 0) return false;
  const key = `${period.year}-${String(period.month).padStart(2, '0')}`;
  return locked.has(key);
}

type ReferenceSets = {
  filiais: Set<string>;
  canais: Set<string>;
  centrosCusto: Set<string>;
  contasContabeis: Set<string>;
};

function getReferenceSets(): ReferenceSets {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('refsets:v1');
  if (cached) {
    try {
      const parsed = JSON.parse(cached) as Record<string, string[]>;
      return {
        filiais: new Set(parsed.filiais || []),
        canais: new Set(parsed.canais || []),
        centrosCusto: new Set(parsed.centrosCusto || []),
        contasContabeis: new Set(parsed.contasContabeis || []),
      };
    } catch (_) {
      // ignore cache parse errors
    }
  }

  const filiais = readReferenceSet(SHEET_REF_FILIAIS);
  const canais = readReferenceSet(SHEET_REF_CANAIS);
  const centrosCusto = readReferenceSet(SHEET_REF_CCUSTO);
  const contasContabeis = readReferenceSet(SHEET_REF_PLANO_CONTAS);

  const payload = JSON.stringify({
    filiais: Array.from(filiais),
    canais: Array.from(canais),
    centrosCusto: Array.from(centrosCusto),
    contasContabeis: Array.from(contasContabeis),
  });
  cache.put('refsets:v1', payload, 600);

  return { filiais, canais, centrosCusto, contasContabeis };
}

function readReferenceSet(sheetName: string): Set<string> {
  try {
    const rows = getSheetValues(sheetName, { skipHeader: true });
    const set = new Set<string>();
    for (const row of rows) {
      const value = row && row.length > 0 ? String(row[0] || '').trim() : '';
      if (value) set.add(value);
    }
    return set;
  } catch (_e) {
    return new Set();
  }
}

function getLockedPeriods(): Set<string> {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('locked_periods:v1');
  if (cached) {
    try {
      const parsed = JSON.parse(cached) as string[];
      return new Set(parsed || []);
    } catch (_) {
      // ignore cache parse errors
    }
  }

  const raw = PropertiesService.getScriptProperties().getProperty('LOCKED_PERIODS') || '';
  const items = raw
    .split(',')
    .map((s) => String(s || '').trim())
    .filter(Boolean);
  cache.put('locked_periods:v1', JSON.stringify(items), 600);
  return new Set(items);
}

/**
 * Combina múltiplos resultados de validação
 */
export function combineValidations(
  ...validations: ValidationResult[]
): ValidationResult {
  const allErrors: string[] = [];

  for (const validation of validations) {
    if (!validation.valid) {
      allErrors.push(...validation.errors);
    }
  }

  return allErrors.length > 0
    ? validationError(allErrors)
    : validationSuccess();
}
