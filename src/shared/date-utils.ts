/**
 * date-utils.ts
 *
 * Utilitários para manipulação de datas.
 * Considera timezone configurado (padrão: America/Sao_Paulo).
 */

import { Period } from './types';

/**
 * Timezone padrão da aplicação
 * TODO: Ler de CFG_CONFIG
 */
const DEFAULT_TIMEZONE = 'America/Sao_Paulo';

/**
 * Formata data para string DD/MM/YYYY
 */
export function formatDate(date: Date | null | undefined): string {
  if (!date) return '';

  const d = new Date(date);
  const day = d.getDate().toString().padStart(2, '0');
  const month = (d.getMonth() + 1).toString().padStart(2, '0');
  const year = d.getFullYear();

  return `${day}/${month}/${year}`;
}

/**
 * Formata data para string ISO (YYYY-MM-DD)
 */
export function formatDateISO(date: Date | null | undefined): string {
  if (!date) return '';

  const d = new Date(date);
  const day = d.getDate().toString().padStart(2, '0');
  const month = (d.getMonth() + 1).toString().padStart(2, '0');
  const year = d.getFullYear();

  return `${year}-${month}-${day}`;
}

/**
 * Parse string DD/MM/YYYY para Date
 */
export function parseDate(dateStr: string): Date | null {
  if (!dateStr) return null;

  const parts = dateStr.split('/');
  if (parts.length !== 3) return null;

  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1; // mês base 0
  const year = parseInt(parts[2], 10);

  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;

  return new Date(year, month, day);
}

/**
 * Parse string ISO (YYYY-MM-DD) para Date
 */
export function parseDateISO(dateStr: string): Date | null {
  if (!dateStr) return null;

  const parts = dateStr.split('-');
  if (parts.length !== 3) return null;

  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1; // mês base 0
  const day = parseInt(parts[2], 10);

  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;

  return new Date(year, month, day);
}

/**
 * Adiciona dias a uma data
 */
export function addDays(date: Date, days: number): Date {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

/**
 * Adiciona meses a uma data
 */
export function addMonths(date: Date, months: number): Date {
  const result = new Date(date);
  result.setMonth(result.getMonth() + months);
  return result;
}

/**
 * Retorna o primeiro dia de um período
 */
export function getFirstDayOfPeriod(period: Period): Date {
  return new Date(period.year, period.month - 1, 1);
}

/**
 * Retorna o último dia de um período
 */
export function getLastDayOfPeriod(period: Period): Date {
  return new Date(period.year, period.month, 0);
}

/**
 * Calcula diferença em dias entre duas datas
 */
export function diffDays(date1: Date, date2: Date): number {
  const msPerDay = 24 * 60 * 60 * 1000;
  const utc1 = Date.UTC(date1.getFullYear(), date1.getMonth(), date1.getDate());
  const utc2 = Date.UTC(date2.getFullYear(), date2.getMonth(), date2.getDate());

  return Math.floor((utc2 - utc1) / msPerDay);
}

/**
 * Verifica se uma data está no futuro
 */
export function isFuture(date: Date): boolean {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return date > today;
}

/**
 * Verifica se uma data está no passado
 */
export function isPast(date: Date): boolean {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return date < today;
}

/**
 * Verifica se uma data é hoje
 */
export function isToday(date: Date): boolean {
  const today = new Date();
  return (
    date.getDate() === today.getDate() &&
    date.getMonth() === today.getMonth() &&
    date.getFullYear() === today.getFullYear()
  );
}

/**
 * Retorna data de hoje às 00:00:00
 */
export function getToday(): Date {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return today;
}

/**
 * Converte Date para Period
 */
export function dateToPeriod(date: Date): Period {
  return {
    year: date.getFullYear(),
    month: date.getMonth() + 1,
  };
}

/**
 * Valida se uma data está dentro do limite de dias retroativos permitidos
 *
 * @param date - Data a validar
 * @param maxDaysBack - Máximo de dias retroativos (padrão: 7)
 * @returns true se válido, false caso contrário
 */
export function isWithinRetroactiveLimit(
  date: Date,
  maxDaysBack: number = 7
): boolean {
  const today = getToday();
  const diff = diffDays(date, today);

  return diff >= -maxDaysBack;
}

/**
 * Gera array de períodos entre dois períodos (inclusive)
 */
export function generatePeriodRange(start: Period, end: Period): Period[] {
  const periods: Period[] = [];
  let current = { ...start };

  while (
    current.year < end.year ||
    (current.year === end.year && current.month <= end.month)
  ) {
    periods.push({ ...current });

    // Avança um mês
    current.month++;
    if (current.month > 12) {
      current.month = 1;
      current.year++;
    }
  }

  return periods;
}
