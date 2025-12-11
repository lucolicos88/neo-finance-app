/**
 * money-utils.ts
 *
 * Utilitários para manipulação de valores monetários.
 * Sempre trabalha com BRL, 2 casas decimais.
 */

import { Money } from './types';

/**
 * Formata valor monetário para exibição (R$ 1.234,56)
 */
export function formatMoney(value: Money | null | undefined): string {
  if (value === null || value === undefined) return 'R$ 0,00';

  const formatted = value.toLocaleString('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

  return `R$ ${formatted}`;
}

/**
 * Formata valor monetário sem símbolo (1.234,56)
 */
export function formatMoneyPlain(value: Money | null | undefined): string {
  if (value === null || value === undefined) return '0,00';

  return value.toLocaleString('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

/**
 * Parse string monetária para número
 * Aceita formatos: "1.234,56", "1234.56", "1234,56", "R$ 1.234,56"
 */
export function parseMoney(value: string | number | null | undefined): Money {
  if (value === null || value === undefined || value === '') return 0;

  if (typeof value === 'number') {
    return parseFloat(value.toFixed(2));
  }

  // Remove símbolo de moeda e espaços
  let cleaned = value.replace(/R\$\s?/g, '').trim();

  // Detecta formato: se tem vírgula E ponto, vírgula é decimal
  if (cleaned.includes(',') && cleaned.includes('.')) {
    // Formato: 1.234,56 -> remove pontos e troca vírgula por ponto
    cleaned = cleaned.replace(/\./g, '').replace(',', '.');
  } else if (cleaned.includes(',')) {
    // Formato: 1234,56 -> troca vírgula por ponto
    cleaned = cleaned.replace(',', '.');
  }
  // Se só tem ponto, assume que já está correto: 1234.56

  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parseFloat(parsed.toFixed(2));
}

/**
 * Arredonda para 2 casas decimais
 */
export function roundMoney(value: number): Money {
  return parseFloat(value.toFixed(2));
}

/**
 * Compara dois valores monetários com tolerância
 * Útil para conciliação bancária
 */
export function moneyEquals(
  value1: Money,
  value2: Money,
  tolerance: Money = 0.01
): boolean {
  return Math.abs(value1 - value2) <= tolerance;
}

/**
 * Soma array de valores monetários
 */
export function sumMoney(values: Money[]): Money {
  const sum = values.reduce((acc, val) => acc + val, 0);
  return roundMoney(sum);
}

/**
 * Calcula percentual de um valor sobre outro
 * Retorna percentual com 2 casas decimais (ex: 15.50 para 15,50%)
 */
export function calculatePercentage(part: Money, total: Money): number {
  if (total === 0) return 0;
  return parseFloat(((part / total) * 100).toFixed(2));
}

/**
 * Aplica percentual a um valor
 */
export function applyPercentage(value: Money, percentage: number): Money {
  return roundMoney((value * percentage) / 100);
}

/**
 * Formata percentual para exibição (15,50%)
 */
export function formatPercentage(value: number): string {
  return `${value.toLocaleString('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}%`;
}

/**
 * Calcula desconto
 */
export function calculateDiscount(
  valorBruto: Money,
  valorLiquido: Money
): Money {
  return roundMoney(valorBruto - valorLiquido);
}

/**
 * Calcula valor líquido
 */
export function calculateValorLiquido(
  valorBruto: Money,
  desconto: Money,
  juros: Money,
  multa: Money
): Money {
  return roundMoney(valorBruto - desconto + juros + multa);
}

/**
 * Valida se valor está positivo
 */
export function isPositive(value: Money): boolean {
  return value > 0;
}

/**
 * Valida se valor está negativo
 */
export function isNegative(value: Money): boolean {
  return value < 0;
}

/**
 * Retorna valor absoluto
 */
export function abs(value: Money): Money {
  return Math.abs(value);
}
