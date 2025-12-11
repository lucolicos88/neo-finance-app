/**
 * cache.ts
 *
 * Wrapper para CacheService do Google Apps Script.
 * Centraliza lógica de cache para dados de referência e configurações.
 *
 * IMPORTANTE:
 * - CacheService tem limite de 100KB por entrada
 * - Cache Script: persiste entre execuções do mesmo script
 * - Cache User: persiste por usuário
 * - TTL máximo: 6 horas (21600 segundos)
 */

const DEFAULT_TTL_SECONDS = 3600; // 1 hora

/**
 * Enum para tipos de cache disponíveis
 */
export enum CacheScope {
  SCRIPT = 'SCRIPT', // Compartilhado entre todos os usuários
  USER = 'USER', // Específico por usuário
}

/**
 * Obtém a instância correta do CacheService
 */
function getCacheInstance(scope: CacheScope): GoogleAppsScript.Cache.Cache {
  switch (scope) {
    case CacheScope.SCRIPT:
      return CacheService.getScriptCache();
    case CacheScope.USER:
      return CacheService.getUserCache();
    default:
      return CacheService.getScriptCache();
  }
}

/**
 * Gera uma chave de cache com namespace
 */
function buildCacheKey(namespace: string, key: string): string {
  return `${namespace}:${key}`;
}

/**
 * Obtém um valor do cache
 *
 * @param namespace - Namespace do cache (ex: 'config', 'reference')
 * @param key - Chave do valor
 * @param scope - Escopo do cache
 * @returns Valor do cache ou null se não encontrado/expirado
 */
export function cacheGet<T = any>(
  namespace: string,
  key: string,
  scope: CacheScope = CacheScope.SCRIPT
): T | null {
  try {
    const cache = getCacheInstance(scope);
    const cacheKey = buildCacheKey(namespace, key);
    const cached = cache.get(cacheKey);

    if (!cached) {
      return null;
    }

    // Parse JSON
    return JSON.parse(cached) as T;
  } catch (error) {
    console.error(`Erro ao ler cache [${namespace}:${key}]:`, error);
    return null;
  }
}

/**
 * Armazena um valor no cache
 *
 * @param namespace - Namespace do cache
 * @param key - Chave do valor
 * @param value - Valor a armazenar (será serializado como JSON)
 * @param ttlSeconds - Tempo de vida em segundos (padrão: 1 hora)
 * @param scope - Escopo do cache
 */
export function cacheSet(
  namespace: string,
  key: string,
  value: any,
  ttlSeconds: number = DEFAULT_TTL_SECONDS,
  scope: CacheScope = CacheScope.SCRIPT
): void {
  try {
    const cache = getCacheInstance(scope);
    const cacheKey = buildCacheKey(namespace, key);
    const serialized = JSON.stringify(value);

    // Verifica limite de 100KB
    if (serialized.length > 100000) {
      console.warn(
        `Valor muito grande para cache [${namespace}:${key}]: ${serialized.length} bytes`
      );
      return;
    }

    // Limita TTL a 6 horas (máximo do CacheService)
    const ttl = Math.min(ttlSeconds, 21600);

    cache.put(cacheKey, serialized, ttl);
  } catch (error) {
    console.error(`Erro ao escrever cache [${namespace}:${key}]:`, error);
  }
}

/**
 * Remove um valor do cache
 */
export function cacheRemove(
  namespace: string,
  key: string,
  scope: CacheScope = CacheScope.SCRIPT
): void {
  try {
    const cache = getCacheInstance(scope);
    const cacheKey = buildCacheKey(namespace, key);
    cache.remove(cacheKey);
  } catch (error) {
    console.error(`Erro ao remover cache [${namespace}:${key}]:`, error);
  }
}

/**
 * Remove todos os valores de um namespace
 *
 * ATENÇÃO: CacheService não tem método removeAll por namespace,
 * então precisamos rastrear as chaves manualmente
 *
 * TODO: Implementar tracking de chaves por namespace
 */
export function cacheRemoveNamespace(
  namespace: string,
  scope: CacheScope = CacheScope.SCRIPT
): void {
  try {
    const cache = getCacheInstance(scope);
    cache.removeAll([namespace]);
  } catch (error) {
    console.error(`Erro ao limpar namespace ${namespace}:`, error);
  }
}

/**
 * Obtém ou calcula um valor (cache-aside pattern)
 *
 * @param namespace - Namespace do cache
 * @param key - Chave do valor
 * @param loader - Função que carrega o valor se não estiver em cache
 * @param ttlSeconds - TTL do cache
 * @param scope - Escopo do cache
 * @returns Valor do cache ou resultado do loader
 */
export function cacheGetOrLoad<T>(
  namespace: string,
  key: string,
  loader: () => T,
  ttlSeconds: number = DEFAULT_TTL_SECONDS,
  scope: CacheScope = CacheScope.SCRIPT
): T {
  // Tenta buscar do cache
  const cached = cacheGet<T>(namespace, key, scope);

  if (cached !== null) {
    return cached;
  }

  // Não encontrou: executa loader
  const value = loader();

  // Armazena no cache
  cacheSet(namespace, key, value, ttlSeconds, scope);

  return value;
}

/**
 * Namespaces padrão usados no sistema
 */
export const CacheNamespace = {
  CONFIG: 'config',
  REFERENCE: 'reference',
  BENCHMARKS: 'benchmarks',
  LABELS: 'labels',
  THEME: 'theme',
  DRE: 'dre',
  DFC: 'dfc',
  KPI: 'kpi',
} as const;
