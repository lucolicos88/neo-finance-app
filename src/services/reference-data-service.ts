/**
 * reference-data-service.ts
 *
 * Gerencia tabelas mestras (dados de referência).
 * Lê das abas REF_* e mantém cache.
 *
 * Responsabilidades:
 * - Carregar plano de contas, filiais, canais, centros de custo, naturezas
 * - Fornecer funções de busca por ID/código
 * - Cachear dados de referência (mudam pouco)
 */

import { getSheetValues } from '../shared/sheets-client';
import { cacheGetOrLoad, CacheNamespace, CacheScope } from '../shared/cache';
import { Sheets, REF_PLANO_CONTAS_COLS } from '../config/sheet-mapping';
import {
  Account,
  Branch,
  Channel,
  CostCenter,
  Nature,
  AccountCode,
  BranchId,
  ChannelId,
  CostCenterId,
  AccountType,
  ExpenseClassification,
  CostClassification,
  CashflowCategory,
  RevenueGroup,
} from '../shared/types';

// ============================================================================
// PLANO DE CONTAS
// ============================================================================

/**
 * Carrega plano de contas da planilha
 */
function loadAccountsFromSheet(): Account[] {
  const values = getSheetValues(Sheets.REF_PLANO_CONTAS, { skipHeader: true });
  const accounts: Account[] = [];

  for (const row of values) {
    if (!row || row.length === 0) continue;

    const account: Account = {
      codigo: row[REF_PLANO_CONTAS_COLS.CODIGO],
      descricao: row[REF_PLANO_CONTAS_COLS.DESCRICAO],
      tipo: row[REF_PLANO_CONTAS_COLS.TIPO] as AccountType,
      grupoDRE: row[REF_PLANO_CONTAS_COLS.GRUPO_DRE],
      subgrupoDRE: row[REF_PLANO_CONTAS_COLS.SUBGRUPO_DRE] || null,
      grupoDFC: row[REF_PLANO_CONTAS_COLS.GRUPO_DFC] as CashflowCategory || null,
      variavelFixa: row[REF_PLANO_CONTAS_COLS.VARIAVEL_FIXA] as ExpenseClassification || null,
      cmaCmv: row[REF_PLANO_CONTAS_COLS.CMA_CMV] as CostClassification || null,
    };

    if (account.codigo && account.descricao) {
      accounts.push(account);
    }
  }

  return accounts;
}

/**
 * Obtém todas as contas do plano de contas (com cache)
 */
export function getAllAccounts(): Account[] {
  return cacheGetOrLoad(
    CacheNamespace.REFERENCE,
    'accounts',
    loadAccountsFromSheet,
    3600, // 1 hora
    CacheScope.SCRIPT
  );
}

/**
 * Busca conta por código
 */
export function getAccountByCode(code: AccountCode): Account | null {
  const accounts = getAllAccounts();
  return accounts.find((a) => a.codigo === code) || null;
}

/**
 * Lista contas por tipo
 */
export function getAccountsByType(type: AccountType): Account[] {
  const accounts = getAllAccounts();
  return accounts.filter((a) => a.tipo === type);
}

/**
 * Lista contas por grupo DRE
 */
export function getAccountsByDREGroup(group: string): Account[] {
  const accounts = getAllAccounts();
  return accounts.filter((a) => a.grupoDRE === group);
}

// ============================================================================
// FILIAIS
// ============================================================================

/**
 * Carrega filiais da planilha
 */
function loadBranchesFromSheet(): Branch[] {
  const values = getSheetValues(Sheets.REF_FILIAIS, { skipHeader: true });
  const branches: Branch[] = [];

  for (const row of values) {
    if (!row || row.length === 0) continue;

    const branch: Branch = {
      id: row[0], // id
      nome: row[1], // nome
      ativa: row[2] !== false && row[2] !== 'FALSE', // ativa
    };

    if (branch.id && branch.nome) {
      branches.push(branch);
    }
  }

  return branches;
}

/**
 * Obtém todas as filiais (com cache)
 */
export function getAllBranches(): Branch[] {
  return cacheGetOrLoad(
    CacheNamespace.REFERENCE,
    'branches',
    loadBranchesFromSheet,
    3600,
    CacheScope.SCRIPT
  );
}

/**
 * Busca filial por ID
 */
export function getBranchById(id: BranchId): Branch | null {
  const branches = getAllBranches();
  return branches.find((b) => b.id === id) || null;
}

/**
 * Lista apenas filiais ativas
 */
export function getActiveBranches(): Branch[] {
  const branches = getAllBranches();
  return branches.filter((b) => b.ativa);
}

// ============================================================================
// CANAIS
// ============================================================================

/**
 * Carrega canais da planilha
 */
function loadChannelsFromSheet(): Channel[] {
  const values = getSheetValues(Sheets.REF_CANAIS, { skipHeader: true });
  const channels: Channel[] = [];

  for (const row of values) {
    if (!row || row.length === 0) continue;

    const channel: Channel = {
      id: row[0], // id
      nome: row[1], // nome
      grupo: row[2] as RevenueGroup || null, // grupo (SERVICOS ou REVENDA)
    };

    if (channel.id && channel.nome) {
      channels.push(channel);
    }
  }

  return channels;
}

/**
 * Obtém todos os canais (com cache)
 */
export function getAllChannels(): Channel[] {
  return cacheGetOrLoad(
    CacheNamespace.REFERENCE,
    'channels',
    loadChannelsFromSheet,
    3600,
    CacheScope.SCRIPT
  );
}

/**
 * Busca canal por ID
 */
export function getChannelById(id: ChannelId): Channel | null {
  const channels = getAllChannels();
  return channels.find((c) => c.id === id) || null;
}

/**
 * Lista canais por grupo
 */
export function getChannelsByGroup(group: RevenueGroup): Channel[] {
  const channels = getAllChannels();
  return channels.filter((c) => c.grupo === group);
}

// ============================================================================
// CENTROS DE CUSTO
// ============================================================================

/**
 * Carrega centros de custo da planilha
 */
function loadCostCentersFromSheet(): CostCenter[] {
  const values = getSheetValues(Sheets.REF_CCUSTO, { skipHeader: true });
  const costCenters: CostCenter[] = [];

  for (const row of values) {
    if (!row || row.length === 0) continue;

    const costCenter: CostCenter = {
      id: row[0], // id
      nome: row[1], // nome
    };

    if (costCenter.id && costCenter.nome) {
      costCenters.push(costCenter);
    }
  }

  return costCenters;
}

/**
 * Obtém todos os centros de custo (com cache)
 */
export function getAllCostCenters(): CostCenter[] {
  return cacheGetOrLoad(
    CacheNamespace.REFERENCE,
    'costCenters',
    loadCostCentersFromSheet,
    3600,
    CacheScope.SCRIPT
  );
}

/**
 * Busca centro de custo por ID
 */
export function getCostCenterById(id: CostCenterId): CostCenter | null {
  const costCenters = getAllCostCenters();
  return costCenters.find((cc) => cc.id === id) || null;
}

// ============================================================================
// NATUREZAS
// ============================================================================

/**
 * Carrega naturezas da planilha
 */
function loadNaturesFromSheet(): Nature[] {
  const values = getSheetValues(Sheets.REF_NATUREZAS, { skipHeader: true });
  const natures: Nature[] = [];

  for (const row of values) {
    if (!row || row.length === 0) continue;

    const nature: Nature = {
      id: row[0], // id
      nome: row[1], // nome
      grupoDRE: row[2], // grupo_dre
    };

    if (nature.id && nature.nome) {
      natures.push(nature);
    }
  }

  return natures;
}

/**
 * Obtém todas as naturezas (com cache)
 */
export function getAllNatures(): Nature[] {
  return cacheGetOrLoad(
    CacheNamespace.REFERENCE,
    'natures',
    loadNaturesFromSheet,
    3600,
    CacheScope.SCRIPT
  );
}

/**
 * Busca natureza por ID
 */
export function getNatureById(id: string): Nature | null {
  const natures = getAllNatures();
  return natures.find((n) => n.id === id) || null;
}

// ============================================================================
// BENCHMARKS
// ============================================================================

/**
 * Carrega benchmarks da planilha CFG_BENCHMARKS
 *
 * TODO: Implementar parsing completo da estrutura de benchmarks
 */
function loadBenchmarksFromSheet(): any[] {
  const values = getSheetValues(Sheets.CFG_BENCHMARKS, { skipHeader: true });
  // TODO: Parsear para BenchmarkConfig[]
  return values;
}

/**
 * Obtém todos os benchmarks (com cache)
 */
export function getAllBenchmarks(): any[] {
  return cacheGetOrLoad(
    CacheNamespace.BENCHMARKS,
    'all',
    loadBenchmarksFromSheet,
    3600,
    CacheScope.SCRIPT
  );
}

// ============================================================================
// RECARREGAR CACHE
// ============================================================================

/**
 * Recarrega todos os caches de dados de referência
 */
export function reloadReferenceCache(): void {
  // Força reload chamando as funções de load diretamente
  loadAccountsFromSheet();
  loadBranchesFromSheet();
  loadChannelsFromSheet();
  loadCostCentersFromSheet();
  loadNaturesFromSheet();
  loadBenchmarksFromSheet();
}
