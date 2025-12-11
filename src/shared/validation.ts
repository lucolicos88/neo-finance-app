/**
 * validation.ts
 *
 * Funções de validação de dados e regras de negócio.
 * Centraliza todas as validações usadas pelos serviços.
 */

import { Money, LedgerEntry, Period } from './types';
import { isFuture, isWithinRetroactiveLimit, diffDays } from './date-utils';
import { isPositive } from './money-utils';

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

  // TODO: Validar se filial, canal, centroCusto, contaGerencial existem nas REF_*

  return errors.length > 0 ? validationError(errors) : validationSuccess();
}

/**
 * Valida se um período está fechado (bloqueado para edição)
 *
 * TODO: Implementar lógica de períodos fechados via config ou tabela
 */
export function isPeriodLocked(period: Period): boolean {
  // TODO: Consultar tabela de períodos fechados ou config
  return false;
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
