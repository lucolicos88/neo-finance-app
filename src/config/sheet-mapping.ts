/**
 * sheet-mapping.ts
 *
 * Centraliza TODOS os nomes de abas do Google Sheets.
 * Nenhum serviço deve hard-codar nomes de abas; tudo deve vir deste módulo.
 *
 * FONTE DA VERDADE: neoformula-finance-app-spec-v1.md seção 4
 */

// ============================================================================
// ABAS DE CONFIGURAÇÃO (prefixo CFG_)
// ============================================================================

export const SHEET_CFG_CONFIG = 'CFG_CONFIG';
export const SHEET_CFG_BENCHMARKS = 'CFG_BENCHMARKS';
export const SHEET_CFG_LABELS = 'CFG_LABELS';
export const SHEET_CFG_THEME = 'CFG_THEME';
export const SHEET_CFG_DFC = 'CFG_DFC';
export const SHEET_CFG_VALIDATION = 'CFG_VALIDATION';

// ============================================================================
// ABAS DE REFERÊNCIA (prefixo REF_)
// ============================================================================

export const SHEET_REF_PLANO_CONTAS = 'REF_PLANO_CONTAS';
export const SHEET_REF_FILIAIS = 'REF_FILIAIS';
export const SHEET_REF_CANAIS = 'REF_CANAIS';
export const SHEET_REF_CCUSTO = 'REF_CCUSTO';
export const SHEET_REF_NATUREZAS = 'REF_NATUREZAS';

// ============================================================================
// ABAS TRANSACIONAIS (prefixo TB_)
// ============================================================================

export const SHEET_TB_LANCAMENTOS = 'TB_LANCAMENTOS';
export const SHEET_TB_EXTRATOS = 'TB_EXTRATOS';
export const SHEET_TB_DRE_MENSAL = 'TB_DRE_MENSAL';
export const SHEET_TB_DRE_RESUMO = 'TB_DRE_RESUMO';
export const SHEET_TB_DFC_REAL = 'TB_DFC_REAL';
export const SHEET_TB_DFC_PROJ = 'TB_DFC_PROJ';
export const SHEET_TB_KPI_RESUMO = 'TB_KPI_RESUMO';
export const SHEET_TB_KPI_DETALHE = 'TB_KPI_DETALHE';

// ============================================================================
// ABAS DE RELATÓRIOS (prefixo RPT_)
// ============================================================================

export const SHEET_RPT_COMITE_FATURAMENTO = 'RPT_COMITE_FATURAMENTO';
export const SHEET_RPT_COMITE_DRE = 'RPT_COMITE_DRE';
export const SHEET_RPT_COMITE_DFC = 'RPT_COMITE_DFC';
export const SHEET_RPT_COMITE_KPIS = 'RPT_COMITE_KPIS';

// ============================================================================
// OBJETO AGREGADOR
// ============================================================================

export const Sheets = {
  // Configuração
  CFG_CONFIG: SHEET_CFG_CONFIG,
  CFG_BENCHMARKS: SHEET_CFG_BENCHMARKS,
  CFG_LABELS: SHEET_CFG_LABELS,
  CFG_THEME: SHEET_CFG_THEME,
  CFG_DFC: SHEET_CFG_DFC,
  CFG_VALIDATION: SHEET_CFG_VALIDATION,

  // Referência
  REF_PLANO_CONTAS: SHEET_REF_PLANO_CONTAS,
  REF_FILIAIS: SHEET_REF_FILIAIS,
  REF_CANAIS: SHEET_REF_CANAIS,
  REF_CCUSTO: SHEET_REF_CCUSTO,
  REF_NATUREZAS: SHEET_REF_NATUREZAS,

  // Transacional
  TB_LANCAMENTOS: SHEET_TB_LANCAMENTOS,
  TB_EXTRATOS: SHEET_TB_EXTRATOS,
  TB_DRE_MENSAL: SHEET_TB_DRE_MENSAL,
  TB_DRE_RESUMO: SHEET_TB_DRE_RESUMO,
  TB_DFC_REAL: SHEET_TB_DFC_REAL,
  TB_DFC_PROJ: SHEET_TB_DFC_PROJ,
  TB_KPI_RESUMO: SHEET_TB_KPI_RESUMO,
  TB_KPI_DETALHE: SHEET_TB_KPI_DETALHE,

  // Relatórios
  RPT_COMITE_FATURAMENTO: SHEET_RPT_COMITE_FATURAMENTO,
  RPT_COMITE_DRE: SHEET_RPT_COMITE_DRE,
  RPT_COMITE_DFC: SHEET_RPT_COMITE_DFC,
  RPT_COMITE_KPIS: SHEET_RPT_COMITE_KPIS,
} as const;

// ============================================================================
// ÍNDICES DE COLUNAS (para facilitar parsing de arrays)
// ============================================================================

/**
 * Índices de colunas da aba CFG_CONFIG
 */
export const CFG_CONFIG_COLS = {
  CHAVE: 0,
  VALOR: 1,
  TIPO: 2,
  DESCRICAO: 3,
  ATIVO: 4,
} as const;

/**
 * Índices de colunas da aba TB_LANCAMENTOS
 */
export const TB_LANCAMENTOS_COLS = {
  ID: 0,
  DATA_COMPETENCIA: 1,
  DATA_VENCIMENTO: 2,
  DATA_PAGAMENTO: 3,
  TIPO: 4,
  FILIAL: 5,
  CENTRO_CUSTO: 6,
  CONTA_GERENCIAL: 7,
  CONTA_CONTABIL: 8,
  GRUPO_RECEITA: 9,
  CANAL: 10,
  DESCRICAO: 11,
  VALOR_BRUTO: 12,
  DESCONTO: 13,
  JUROS: 14,
  MULTA: 15,
  VALOR_LIQUIDO: 16,
  STATUS: 17,
  ID_EXTRATO_BANCO: 18,
  ORIGEM: 19,
  OBSERVACOES: 20,
} as const;

/**
 * Índices de colunas da aba TB_EXTRATOS
 */
export const TB_EXTRATOS_COLS = {
  ID: 0,
  DATA_MOVIMENTO: 1,
  CONTA_BANCARIA: 2,
  HISTORICO: 3,
  DOCUMENTO: 4,
  VALOR: 5,
  SALDO_APOS: 6,
  CONCILIADO: 7,
  ID_LANCAMENTO: 8,
} as const;

/**
 * Índices de colunas da aba REF_PLANO_CONTAS
 */
export const REF_PLANO_CONTAS_COLS = {
  CODIGO: 0,
  DESCRICAO: 1,
  TIPO: 2,
  GRUPO_DRE: 3,
  SUBGRUPO_DRE: 4,
  GRUPO_DFC: 5,
  VARIAVEL_FIXA: 6,
  CMA_CMV: 7,
} as const;
