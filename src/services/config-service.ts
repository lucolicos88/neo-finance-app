/**
 * config-service.ts
 *
 * Gerencia parâmetros globais da aplicação.
 * Lê da aba CFG_CONFIG e mantém cache para performance.
 *
 * Responsabilidades:
 * - Carregar e cachear configurações
 * - Fornecer acesso tipado a parâmetros
 * - Recarregar cache quando necessário
 */

import { getSheetValues } from '../shared/sheets-client';
import { cacheGet, cacheSet, CacheNamespace, CacheScope } from '../shared/cache';
import { Sheets, CFG_CONFIG_COLS } from '../config/sheet-mapping';
import {
  ConfigKey,
  ConfigType,
  ConfigRow,
  parseConfigValue,
  DEFAULT_CONFIG,
} from '../config/config.schema';

/**
 * Mapa de configurações carregadas
 */
type ConfigMap = Map<string, any>;

/**
 * Carrega todas as configurações da aba CFG_CONFIG
 *
 * TODO: Tratar erros de leitura da planilha
 * TODO: Validar schema das linhas
 */
function loadConfigFromSheet(): ConfigMap {
  const values = getSheetValues(Sheets.CFG_CONFIG, { skipHeader: true });
  const configMap = new Map<string, any>();

  for (const row of values) {
    if (!row || row.length === 0) continue;

    const chave = row[CFG_CONFIG_COLS.CHAVE];
    const valor = row[CFG_CONFIG_COLS.VALOR];
    const tipo = row[CFG_CONFIG_COLS.TIPO] as ConfigType;
    const ativo = row[CFG_CONFIG_COLS.ATIVO];

    // Ignora configurações inativas
    if (ativo === false || ativo === 'FALSE') {
      continue;
    }

    if (!chave || !tipo) continue;

    const parsedValue = parseConfigValue(valor, tipo);
    configMap.set(chave, parsedValue);
  }

  return configMap;
}

/**
 * Obtém todas as configurações (com cache)
 */
function getAllConfigs(): ConfigMap {
  // Tenta buscar do cache
  const cached = cacheGet<Record<string, any>>(
    CacheNamespace.CONFIG,
    'all',
    CacheScope.SCRIPT
  );

  if (cached) {
    return new Map(Object.entries(cached));
  }

  // Carrega da planilha
  const configMap = loadConfigFromSheet();

  // Armazena no cache (1 hora)
  const configObj = Object.fromEntries(configMap);
  cacheSet(CacheNamespace.CONFIG, 'all', configObj, 3600, CacheScope.SCRIPT);

  return configMap;
}

/**
 * Obtém valor de uma configuração específica
 *
 * @param key - Chave da configuração
 * @param defaultValue - Valor padrão se não encontrado
 * @returns Valor da configuração
 */
export function getConfig<T = any>(key: string, defaultValue?: T): T {
  const configs = getAllConfigs();

  if (configs.has(key)) {
    return configs.get(key) as T;
  }

  // Tenta buscar dos defaults
  if (DEFAULT_CONFIG[key] !== undefined) {
    return DEFAULT_CONFIG[key] as T;
  }

  // Retorna valor padrão fornecido
  if (defaultValue !== undefined) {
    return defaultValue;
  }

  throw new Error(`Configuração não encontrada: ${key}`);
}

/**
 * Obtém múltiplas configurações de uma vez
 */
export function getConfigs(keys: string[]): Record<string, any> {
  const configs = getAllConfigs();
  const result: Record<string, any> = {};

  for (const key of keys) {
    if (configs.has(key)) {
      result[key] = configs.get(key);
    } else if (DEFAULT_CONFIG[key] !== undefined) {
      result[key] = DEFAULT_CONFIG[key];
    }
  }

  return result;
}

/**
 * Recarrega cache de configurações
 * Útil após atualização manual da planilha
 */
export function reloadConfigCache(): void {
  const configMap = loadConfigFromSheet();
  const configObj = Object.fromEntries(configMap);
  cacheSet(CacheNamespace.CONFIG, 'all', configObj, 3600, CacheScope.SCRIPT);
}

/**
 * Atualiza uma configuração na planilha
 *
 * TODO: Implementar atualização via sheets-client
 * TODO: Invalidar cache após atualização
 */
export function updateConfig(key: string, value: any): void {
  throw new Error('updateConfig não implementado ainda');
  // TODO:
  // 1. Encontrar linha na aba CFG_CONFIG com a chave
  // 2. Atualizar valor
  // 3. Invalidar cache
}

/**
 * Helpers para configurações específicas comuns
 */
export const ConfigService = {
  getMasterSpreadsheetId(): string {
    return getConfig(ConfigKey.MASTER_SPREADSHEET_ID);
  },

  getReportsSpreadsheetId(): string {
    return getConfig(ConfigKey.REPORTS_SPREADSHEET_ID);
  },

  getTimezone(): string {
    return getConfig(ConfigKey.TIMEZONE, 'America/Sao_Paulo');
  },

  getMaxDiasRetroativo(): number {
    return getConfig(ConfigKey.MAX_DIAS_RETROATIVO, 7);
  },

  getToleranciaConciliacao(): number {
    return getConfig(ConfigKey.TOLERANCIA_CONCILIACAO, 0.01);
  },

  getCacheTTL(): number {
    return getConfig(ConfigKey.CACHE_TTL_MINUTES, 60);
  },

  isAutoReconciliationEnabled(): boolean {
    return getConfig(ConfigKey.FEATURE_AUTO_RECONCILIATION, true);
  },

  isDFCProjectionEnabled(): boolean {
    return getConfig(ConfigKey.FEATURE_DFC_PROJECTION, true);
  },

  reloadCache: reloadConfigCache,
};
