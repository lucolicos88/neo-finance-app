/**
 * types.ts
 *
 * Tipos e interfaces compartilhados entre todos os serviços.
 * FONTE DA VERDADE: neoformula-finance-app-spec-v1.md seção 3.2
 *
 * IMPORTANTE: Não adicionar campos aleatórios. Apenas estender onde
 * claramente necessário e documentar com comentário.
 */

// ============================================================================
// TIPOS BÁSICOS
// ============================================================================

/**
 * Representa um período mensal (ano + mês)
 */
export interface Period {
  year: number;
  month: number; // 1-12
}

/**
 * IDs de entidades de referência
 */
export type BranchId = string;
export type ChannelId = string;
export type CostCenterId = string;
export type AccountCode = string;

/**
 * Tipo monetário - sempre em moeda local (BRL), 2 casas decimais
 */
export type Money = number;

// ============================================================================
// ENUMS
// ============================================================================

/**
 * Tipos de lançamento financeiro
 */
export enum LedgerEntryType {
  PAGAR = 'PAGAR',
  RECEBER = 'RECEBER',
  TRANSFERENCIA = 'TRANSFERENCIA',
  AJUSTE = 'AJUSTE',
}

/**
 * Status de um lançamento
 */
export enum LedgerEntryStatus {
  PREVISTO = 'PREVISTO',
  REALIZADO = 'REALIZADO',
  CANCELADO = 'CANCELADO',
}

/**
 * Origem de um lançamento
 */
export enum LedgerEntryOrigin {
  MANUAL = 'MANUAL',
  IMPORTADO = 'IMPORTADO',
}

/**
 * Grupos de receita
 */
export enum RevenueGroup {
  SERVICOS = 'SERVICOS',
  REVENDA = 'REVENDA',
}

/**
 * Tipos de conta no plano de contas
 */
export enum AccountType {
  RECEITA = 'RECEITA',
  DESPESA = 'DESPESA',
  CUSTO = 'CUSTO',
}

/**
 * Classificação de despesa/receita
 */
export enum ExpenseClassification {
  VARIAVEL = 'VARIAVEL',
  FIXA = 'FIXA',
}

/**
 * Classificação de custo
 */
export enum CostClassification {
  CMA = 'CMA', // Custo de Mercadoria Adquirida
  CMV = 'CMV', // Custo de Mercadoria Vendida
}

/**
 * Categorias de fluxo de caixa
 */
export enum CashflowCategory {
  OPERACIONAL = 'OPERACIONAL',
  INVESTIMENTO = 'INVESTIMENTO',
  FINANCIAMENTO = 'FINANCIAMENTO',
}

/**
 * Tipo de movimentação de caixa
 */
export enum CashflowType {
  ENTRADA = 'ENTRADA',
  SAIDA = 'SAIDA',
}

// ============================================================================
// INTERFACES DE ENTIDADES
// ============================================================================

/**
 * Lançamento financeiro/contábil
 */
export interface LedgerEntry {
  id: string;
  competencia: Date;
  vencimento: Date | null;
  pagamento: Date | null;
  tipo: LedgerEntryType;
  filial: BranchId;
  centroCusto: CostCenterId | null;
  contaGerencial: AccountCode;
  contaContabil: AccountCode | null;
  grupoReceita: RevenueGroup | null;
  canal: ChannelId | null;
  descricao: string;
  valorBruto: Money;
  desconto: Money;
  juros: Money;
  multa: Money;
  valorLiquido: Money;
  status: LedgerEntryStatus;
  idExtratoBanco: string | null;
  origem: LedgerEntryOrigin;
  observacoes?: string;
}

/**
 * Extrato bancário
 */
export interface BankStatement {
  id: string;
  dataMovimento: Date;
  contaBancaria: string;
  historico: string;
  documento: string | null;
  valor: Money;
  saldoApos: Money | null;
  conciliado: boolean;
  idLancamento: string | null;
}

/**
 * Linha de DRE (Demonstrativo de Resultados)
 */
export interface DRELine {
  period: Period;
  branchId: BranchId | null; // null = consolidado
  group: string; // ex.: RECEITA_BRUTA, IMPOSTOS, DESPESAS_FIXAS
  subGroup: string | null; // ex.: PESSOAL, MARKETING
  value: Money;
}

/**
 * Linha de fluxo de caixa
 */
export interface CashflowLine {
  date: Date;
  type: CashflowType;
  category: CashflowCategory;
  description: string;
  value: Money;
  projected: boolean; // true = projetado, false = realizado
  contaBancaria?: string;
}

/**
 * Linha de KPI
 */
export interface KPILine {
  period: Period;
  filial: BranchId | null;
  canal: ChannelId | null;
  metric: string;
  value: number;
  faixa: string | null; // SENSACIONAL, EXCELENTE, BOM, RUIM, PESSIMO
}

/**
 * Conta do plano de contas gerencial
 */
export interface Account {
  codigo: AccountCode;
  descricao: string;
  tipo: AccountType;
  grupoDRE: string;
  subgrupoDRE: string | null;
  grupoDFC: CashflowCategory | null;
  variavelFixa: ExpenseClassification | null;
  cmaCmv: CostClassification | null;
}

/**
 * Filial
 */
export interface Branch {
  id: BranchId;
  nome: string;
  ativa: boolean;
}

/**
 * Canal de vendas
 */
export interface Channel {
  id: ChannelId;
  nome: string;
  grupo: RevenueGroup | null;
}

/**
 * Centro de custo
 */
export interface CostCenter {
  id: CostCenterId;
  nome: string;
}

/**
 * Natureza de despesa/receita
 */
export interface Nature {
  id: string;
  nome: string;
  grupoDRE: string;
}

// ============================================================================
// DTOs (Data Transfer Objects) para comunicação com frontend
// ============================================================================

/**
 * Filtro para listagem de lançamentos
 */
export interface LedgerFilter {
  periodStart?: Period;
  periodEnd?: Period;
  filial?: BranchId;
  canal?: ChannelId;
  centroCusto?: CostCenterId;
  tipo?: LedgerEntryType;
  status?: LedgerEntryStatus;
  contaGerencial?: AccountCode;
}

/**
 * Filtro para relatórios e KPIs
 */
export interface ReportFilter {
  period?: Period;
  filial?: BranchId;
  canal?: ChannelId;
  consolidado?: boolean;
}

/**
 * Dados do dashboard
 */
export interface DashboardData {
  period: Period;
  receitaBruta: Money;
  receitaLiquida: Money;
  ebitda: Money;
  ebitdaPct: number;
  saldoCaixa: Money;
  kpis: KPILine[];
  topDespesas: Array<{ descricao: string; valor: Money }>;
}

// ============================================================================
// HELPERS DE TIPO
// ============================================================================

/**
 * Converte Period para string no formato YYYY-MM
 */
export function periodToString(period: Period): string {
  const month = period.month.toString().padStart(2, '0');
  return `${period.year}-${month}`;
}

/**
 * Converte string YYYY-MM para Period
 */
export function stringToPeriod(str: string): Period {
  const [year, month] = str.split('-').map(Number);
  return { year, month };
}

/**
 * Compara dois períodos
 * Retorna: -1 se p1 < p2, 0 se p1 === p2, 1 se p1 > p2
 */
export function comparePeriods(p1: Period, p2: Period): number {
  if (p1.year !== p2.year) {
    return p1.year - p2.year;
  }
  return p1.month - p2.month;
}

/**
 * Retorna o período atual
 */
export function getCurrentPeriod(): Period {
  const now = new Date();
  return {
    year: now.getFullYear(),
    month: now.getMonth() + 1,
  };
}
