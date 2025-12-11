/**
 * scheduler-service.ts
 *
 * Gerencia triggers e jobs automatizados.
 *
 * Responsabilidades:
 * - Criar e gerenciar triggers
 * - Jobs diários (atualizações de cache, conciliação automática)
 * - Jobs mensais (fechamento de período, cálculo de DRE/DFC/KPIs)
 * - Orquestração de processos batch
 * - Monitoramento de limites de execução
 */

import { getCurrentPeriod, Period } from '../shared/types';
import { ConfigService } from './config-service';
import { reloadReferenceCache } from './reference-data-service';
import { autoReconcile } from './reconciliation-service';
import { calculateDRE, persistDREMensal, persistDREResumo } from './dre-service';
import { calculateRealCashflow, persistRealCashflow } from './cashflow-service';
import { calculateKPIs, persistKPIs } from './kpi-analytics-service';
import { generateCommitteeReport } from './reporting-service';

// ============================================================================
// CONFIGURAÇÃO DE TRIGGERS
// ============================================================================

/**
 * Instala todos os triggers necessários
 *
 * IMPORTANTE: Executar manualmente uma vez após deploy
 *
 * TODO: Considerar uso de Properties Service para rastrear triggers instalados
 */
export function installTriggers(): void {
  // Remove triggers antigos para evitar duplicação
  removeAllTriggers();

  // Job diário às 6h
  ScriptApp.newTrigger('dailyJob')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  // Job mensal no dia 1 às 8h
  ScriptApp.newTrigger('monthlyClosing')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  console.log('Triggers instalados com sucesso');
}

/**
 * Remove todos os triggers do projeto
 */
export function removeAllTriggers(): void {
  const triggers = ScriptApp.getProjectTriggers();

  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }

  console.log(`${triggers.length} triggers removidos`);
}

/**
 * Lista triggers ativos
 */
export function listActiveTriggers(): Array<{
  handlerFunction: string;
  eventType: string;
  uniqueId: string;
}> {
  const triggers = ScriptApp.getProjectTriggers();

  return triggers.map((trigger) => ({
    handlerFunction: trigger.getHandlerFunction(),
    eventType: trigger.getEventType().toString(),
    uniqueId: trigger.getUniqueId(),
  }));
}

// ============================================================================
// JOB DIÁRIO
// ============================================================================

/**
 * Job executado diariamente
 *
 * Tarefas:
 * - Recarregar cache de configurações e referências
 * - Conciliação automática de extratos
 * - Atualizar KPIs do dia anterior
 * - Verificar limites de quota
 *
 * IMPORTANTE:
 * - Tempo máximo de execução: ~6 minutos
 * - Dividir em sub-jobs se necessário
 */
export function dailyJob(): void {
  const startTime = new Date().getTime();
  console.log('=== Iniciando job diário ===');

  try {
    // ========================================================================
    // 1. Recarregar cache
    // ========================================================================
    console.log('[1/4] Recarregando cache...');
    ConfigService.reloadCache();
    reloadReferenceCache();

    // ========================================================================
    // 2. Conciliação automática
    // ========================================================================
    if (ConfigService.isAutoReconciliationEnabled()) {
      console.log('[2/4] Executando conciliação automática...');
      const reconciled = autoReconcile(80); // Min 80% de confiança
      console.log(`  → ${reconciled} conciliações realizadas`);
    }

    // ========================================================================
    // 3. Atualizar KPIs do dia anterior (se necessário)
    // ========================================================================
    console.log('[3/4] Atualizando KPIs...');
    // TODO: Implementar atualização incremental de KPIs

    // ========================================================================
    // 4. Verificar limites
    // ========================================================================
    console.log('[4/4] Verificando limites...');
    checkLimits();

    const duration = (new Date().getTime() - startTime) / 1000;
    console.log(`=== Job diário concluído em ${duration}s ===`);
  } catch (error) {
    console.error('Erro no job diário:', error);
    // TODO: Enviar notificação de erro
  }
}

// ============================================================================
// JOB MENSAL (FECHAMENTO)
// ============================================================================

/**
 * Job de fechamento mensal
 *
 * Executa no primeiro dia do mês para processar o mês anterior
 *
 * Tarefas:
 * - Calcular DRE do mês anterior
 * - Calcular DFC do mês anterior
 * - Calcular KPIs do mês anterior
 * - Gerar relatórios para comitê
 * - Persistir em abas TB_* e RPT_*
 *
 * IMPORTANTE:
 * - Pode demorar vários minutos
 * - Considerar dividir em sub-jobs se ultrapassar 6min
 */
export function monthlyClosing(): void {
  const startTime = new Date().getTime();
  console.log('=== Iniciando fechamento mensal ===');

  try {
    // Calcula período do mês anterior
    const current = getCurrentPeriod();
    const previousPeriod: Period = {
      year: current.month === 1 ? current.year - 1 : current.year,
      month: current.month === 1 ? 12 : current.month - 1,
    };

    console.log(`Processando período: ${previousPeriod.year}-${previousPeriod.month}`);

    // ========================================================================
    // 1. Calcular e persistir DRE
    // ========================================================================
    console.log('[1/5] Calculando DRE...');
    const dre = calculateDRE(previousPeriod, null);
    persistDREMensal(dre);
    persistDREResumo(dre);
    console.log(`  → DRE calculado: EBITDA = R$ ${dre.summary.ebitda.toFixed(2)}`);

    // TODO: Calcular DRE por filial
    checkExecutionTime(startTime, 'DRE');

    // ========================================================================
    // 2. Calcular e persistir DFC
    // ========================================================================
    console.log('[2/5] Calculando DFC...');
    const cashflow = calculateRealCashflow(previousPeriod);
    persistRealCashflow(previousPeriod, cashflow);
    console.log(`  → ${cashflow.length} movimentações de caixa`);

    checkExecutionTime(startTime, 'DFC');

    // ========================================================================
    // 3. Calcular e persistir KPIs
    // ========================================================================
    console.log('[3/5] Calculando KPIs...');
    const kpis = calculateKPIs(previousPeriod, null);
    persistKPIs(previousPeriod, null, null, kpis);
    console.log(`  → ${kpis.length} KPIs calculados`);

    checkExecutionTime(startTime, 'KPIs');

    // ========================================================================
    // 4. Gerar relatórios de comitê
    // ========================================================================
    console.log('[4/5] Gerando relatórios de comitê...');
    const committeeReport = generateCommitteeReport(previousPeriod);
    // TODO: Persistir relatórios em RPT_*

    checkExecutionTime(startTime, 'Relatórios');

    // ========================================================================
    // 5. Notificar conclusão
    // ========================================================================
    console.log('[5/5] Enviando notificações...');
    // TODO: Enviar e-mail ou notificação de conclusão

    const duration = (new Date().getTime() - startTime) / 1000;
    console.log(`=== Fechamento mensal concluído em ${duration}s ===`);
  } catch (error) {
    console.error('Erro no fechamento mensal:', error);
    // TODO: Enviar notificação de erro crítico
  }
}

// ============================================================================
// MONITORAMENTO E LIMITES
// ============================================================================

/**
 * Verifica tempo de execução e para se próximo do limite
 *
 * Apps Script tem limite de ~6 minutos por execução
 */
function checkExecutionTime(startTime: number, taskName: string): void {
  const elapsed = (new Date().getTime() - startTime) / 1000;
  const MAX_EXECUTION_TIME = 330; // 5min 30s (margem de segurança)

  console.log(`  → Tempo decorrido: ${elapsed}s`);

  if (elapsed > MAX_EXECUTION_TIME) {
    throw new Error(
      `Limite de tempo excedido após tarefa "${taskName}". Considere dividir em sub-jobs.`
    );
  }
}

/**
 * Verifica limites de quota do Apps Script
 *
 * TODO: Implementar verificação de quotas via API
 */
export function checkLimits(): void {
  // Apps Script quotas principais:
  // - Email: 100/dia (gratuito), ilimitado (Workspace)
  // - URL Fetch: 20.000/dia
  // - Script runtime: 6 min/execução, 90 min/dia (gratuito)

  // TODO: Registrar uso e alertar se próximo dos limites
  console.log('Verificação de limites não implementada');
}

/**
 * Executa job de backup de dados críticos
 *
 * TODO: Implementar backup para Google Drive
 */
export function backupJob(): void {
  console.log('Job de backup não implementado');
  // TODO: Exportar TB_* para arquivos CSV no Drive
}
