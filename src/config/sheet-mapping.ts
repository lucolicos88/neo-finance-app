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
export const SHEET_REF_CAIXA_TIPOS = 'REF_CAIXA_TIPOS';

// ============================================================================
// ABAS TRANSACIONAIS (prefixo TB_)
// ============================================================================

export const SHEET_TB_LANCAMENTOS = 'TB_LANCAMENTOS';
export const SHEET_TB_EXTRATOS = 'TB_EXTRATOS';
export const SHEET_TB_IMPORT_FC = 'TB_IMPORT_FC';
export const SHEET_TB_IMPORT_ITAU = 'TB_IMPORT_ITAU';
export const SHEET_TB_IMPORT_SIEG = 'TB_IMPORT_SIEG';
export const SHEET_TB_CAIXAS = 'TB_CAIXAS';
export const SHEET_TB_CAIXAS_MOV = 'TB_CAIXAS_MOV';
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
  REF_CAIXA_TIPOS: SHEET_REF_CAIXA_TIPOS,

  // Transacional
  TB_LANCAMENTOS: SHEET_TB_LANCAMENTOS,
  TB_EXTRATOS: SHEET_TB_EXTRATOS,
  TB_IMPORT_FC: SHEET_TB_IMPORT_FC,
  TB_IMPORT_ITAU: SHEET_TB_IMPORT_ITAU,
  TB_IMPORT_SIEG: SHEET_TB_IMPORT_SIEG,
  TB_CAIXAS: SHEET_TB_CAIXAS,
  TB_CAIXAS_MOV: SHEET_TB_CAIXAS_MOV,
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
  DATA: 1,
  DESCRICAO: 2,
  VALOR: 3,
  TIPO: 4,
  BANCO: 5,
  CONTA: 6,
  STATUS_CONCILIACAO: 7,
  ID_LANCAMENTO: 8,
  OBSERVACOES: 9,
  IMPORTADO_EM: 10,
} as const;

/**
 * Indices de colunas da aba TB_IMPORT_FC
 */
export const TB_IMPORT_FC_COLS = {
  DATA_EMISSAO: 0,
  NUM_DOCUMENTO: 1,
  COD_CONTA: 2,
  FILIAL_FC: 3,
  HISTORICO: 4,
  FORNECEDOR: 5,
  VALOR: 6,
  DESCRICAO: 7,
  DATA_BAIXA: 8,
  FLAG_BAIXA: 9,
  DATA_VENCIMENTO: 10,
  TIPO: 11,
  IMPORTADO_EM: 12,
} as const;

/**
 * Indices de colunas da aba TB_IMPORT_ITAU
 */
export const TB_IMPORT_ITAU_COLS = {
  DATA: 0,
  LANCAMENTO: 1,
  AGENCIA_ORIGEM: 2,
  RAZAO_SOCIAL: 3,
  CPF_CNPJ: 4,
  VALOR: 5,
  SALDO: 6,
  CONTA: 7,
  FILIAL_FC: 8,
  MODELO: 9,
  IMPORTADO_EM: 10,
} as const;

/**
 * Indices de colunas da aba TB_IMPORT_SIEG
 */
export const TB_IMPORT_SIEG_COLS = {
  NUM_NFE: 0,
  VALOR: 1,
  DATA_EMISSAO: 2,
  CNPJ_EMIT: 3,
  NOME_FANT_EMIT: 4,
  RAZAO_EMIT: 5,
  CNPJ_DEST: 6,
  NOME_FANT_DEST: 7,
  RAZAO_DEST: 8,
  DATA_ENVIO_COFRE: 9,
  CHAVE_NFE: 10,
  TAGS: 11,
  CODIGO_EVENTO: 12,
  TIPO_EVENTO: 13,
  STATUS: 14,
  DANFE: 15,
  XML: 16,
  CODIGO_FILIAL_SIEG: 17,
  FILIAL_FC: 18,
  IMPORTADO_EM: 19,
} as const;

/**
 * Indices de colunas da aba TB_CAIXAS
 */
export const TB_CAIXAS_COLS = {
  ID: 0,
  CANAL: 1,
  COLABORADOR: 2,
  DATA_FECHAMENTO: 3,
  OBSERVACOES: 4,
  SISTEMA_VALOR: 5,
  REFORCO: 6,
  CRIADO_EM: 7,
  ATUALIZADO_EM: 8,
} as const;

/**
 * Indices de colunas da aba TB_CAIXAS_MOV
 */
export const TB_CAIXAS_MOV_COLS = {
  ID: 0,
  CAIXA_ID: 1,
  TIPO: 2,
  NATUREZA: 3,
  VALOR: 4,
  DATA_MOV: 5,
  ARQUIVO_URL: 6,
  ARQUIVO_NOME: 7,
  CRIADO_EM: 8,
  ATUALIZADO_EM: 9,
  OBSERVACOES: 10,
} as const;

/**
 * Indices de colunas da aba REF_CAIXA_TIPOS
 */
export const REF_CAIXA_TIPOS_COLS = {
  TIPO: 0,
  NATUREZA: 1,
  REQUER_ARQUIVO: 2,
  SISTEMA_FC: 3,
  CONTA_REFORCO: 4,
  ATIVO: 5,
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
