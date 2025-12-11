/**
 * config.schema.ts
 *
 * Define esquemas de validação e tipos para configurações da aplicação.
 * Mapeia chaves esperadas na aba CFG_CONFIG.
 */

/**
 * Tipos permitidos para valores de configuração
 */
export enum ConfigType {
  STRING = 'STRING',
  NUMBER = 'NUMBER',
  BOOLEAN = 'BOOLEAN',
}

/**
 * Chaves de configuração conhecidas pelo sistema
 */
export enum ConfigKey {
  // IDs de planilhas
  MASTER_SPREADSHEET_ID = 'MASTER_SPREADSHEET_ID',
  REPORTS_SPREADSHEET_ID = 'REPORTS_SPREADSHEET_ID',

  // Configurações de sistema
  APP_NAME = 'APP_NAME',
  APP_VERSION = 'APP_VERSION',
  TIMEZONE = 'TIMEZONE',

  // Limites e validações
  MAX_DIAS_RETROATIVO = 'MAX_DIAS_RETROATIVO',
  TOLERANCIA_CONCILIACAO = 'TOLERANCIA_CONCILIACAO',

  // Cache
  CACHE_TTL_MINUTES = 'CACHE_TTL_MINUTES',

  // Features toggles
  FEATURE_AUTO_RECONCILIATION = 'FEATURE_AUTO_RECONCILIATION',
  FEATURE_DFC_PROJECTION = 'FEATURE_DFC_PROJECTION',
}

/**
 * Interface para uma linha da aba CFG_CONFIG
 */
export interface ConfigRow {
  chave: string;
  valor: string;
  tipo: ConfigType;
  descricao?: string;
  ativo?: boolean;
}

/**
 * Valores padrão para configurações críticas
 */
export const DEFAULT_CONFIG: Record<string, any> = {
  [ConfigKey.APP_NAME]: 'Neoformula Finance App',
  [ConfigKey.APP_VERSION]: '1.0.0',
  [ConfigKey.TIMEZONE]: 'America/Sao_Paulo',
  [ConfigKey.MAX_DIAS_RETROATIVO]: 7,
  [ConfigKey.TOLERANCIA_CONCILIACAO]: 0.01, // R$ 0,01
  [ConfigKey.CACHE_TTL_MINUTES]: 60,
  [ConfigKey.FEATURE_AUTO_RECONCILIATION]: true,
  [ConfigKey.FEATURE_DFC_PROJECTION]: true,
};

/**
 * Parse um valor de string para o tipo correto
 */
export function parseConfigValue(value: string, type: ConfigType): any {
  switch (type) {
    case ConfigType.STRING:
      return value;
    case ConfigType.NUMBER:
      return parseFloat(value);
    case ConfigType.BOOLEAN:
      return value.toLowerCase() === 'true' || value === '1';
    default:
      return value;
  }
}
