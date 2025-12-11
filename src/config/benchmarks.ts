/**
 * benchmarks.ts
 *
 * Define estrutura e tipos para benchmarks de KPIs.
 * Os valores reais virão da aba CFG_BENCHMARKS.
 */

/**
 * Faixas de performance para KPIs
 */
export enum BenchmarkRange {
  SENSACIONAL = 'SENSACIONAL',
  EXCELENTE = 'EXCELENTE',
  BOM = 'BOM',
  RUIM = 'RUIM',
  PESSIMO = 'PESSIMO',
}

/**
 * Métricas conhecidas do sistema
 */
export enum KPIMetric {
  // Descontos e margens
  DESCONTO_MEDIO = 'DESCONTO_MEDIO',
  CMA = 'CMA',
  CMV = 'CMV',

  // Rentabilidade
  MARGEM_BRUTA = 'MARGEM_BRUTA',
  MARGEM_LIQUIDA = 'MARGEM_LIQUIDA',
  EBITDA_PCT = 'EBITDA_PCT',

  // Eficiência
  DESPESAS_FIXAS_PCT = 'DESPESAS_FIXAS_PCT',
  DESPESAS_VAR_PCT = 'DESPESAS_VAR_PCT',

  // Liquidez
  SALDO_CAIXA = 'SALDO_CAIXA',
  DIAS_CAIXA = 'DIAS_CAIXA',
}

/**
 * Unidades de medida para métricas
 */
export enum MetricUnit {
  PERCENT = '%',
  CURRENCY = 'R$',
  CURRENCY_PER_UNIT = 'R$/UNID',
  DAYS = 'DIAS',
  RATIO = 'RATIO',
}

/**
 * Interface para uma linha da aba CFG_BENCHMARKS
 */
export interface BenchmarkRow {
  metric: string;
  unidade: MetricUnit;
  sensacional_min: number;
  sensacional_max: number;
  excelente_min: number;
  excelente_max: number;
  bom_min: number;
  bom_max: number;
  ruim_min: number;
  ruim_max: number;
  pessimo_min: number;
  pessimo_max: number;
}

/**
 * Configuração de benchmark processada
 */
export interface BenchmarkConfig {
  metric: string;
  unit: MetricUnit;
  ranges: {
    [BenchmarkRange.SENSACIONAL]: { min: number; max: number };
    [BenchmarkRange.EXCELENTE]: { min: number; max: number };
    [BenchmarkRange.BOM]: { min: number; max: number };
    [BenchmarkRange.RUIM]: { min: number; max: number };
    [BenchmarkRange.PESSIMO]: { min: number; max: number };
  };
}

/**
 * Determina a faixa de benchmark para um valor
 */
export function getBenchmarkRange(
  value: number,
  benchmark: BenchmarkConfig
): BenchmarkRange {
  const { ranges } = benchmark;

  if (value >= ranges.SENSACIONAL.min && value <= ranges.SENSACIONAL.max) {
    return BenchmarkRange.SENSACIONAL;
  }
  if (value >= ranges.EXCELENTE.min && value <= ranges.EXCELENTE.max) {
    return BenchmarkRange.EXCELENTE;
  }
  if (value >= ranges.BOM.min && value <= ranges.BOM.max) {
    return BenchmarkRange.BOM;
  }
  if (value >= ranges.RUIM.min && value <= ranges.RUIM.max) {
    return BenchmarkRange.RUIM;
  }

  return BenchmarkRange.PESSIMO;
}

/**
 * Retorna a cor CSS associada a uma faixa de benchmark
 */
export function getBenchmarkColor(range: BenchmarkRange): string {
  switch (range) {
    case BenchmarkRange.SENSACIONAL:
      return 'var(--kpi-sensacional)';
    case BenchmarkRange.EXCELENTE:
      return 'var(--kpi-excelente)';
    case BenchmarkRange.BOM:
      return 'var(--kpi-bom)';
    case BenchmarkRange.RUIM:
      return 'var(--kpi-ruim)';
    case BenchmarkRange.PESSIMO:
      return 'var(--kpi-pessimo)';
    default:
      return 'var(--neo-gray-600)';
  }
}
