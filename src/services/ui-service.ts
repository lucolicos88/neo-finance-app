/**
 * ui-service.ts
 *
 * Gerencia interface de usuário (web app).
 *
 * Responsabilidades:
 * - Servir views HTML
 * - Processar formulários
 * - Validação client-side
 * - Roteamento web (doGet)
 */

/**
 * Rotas disponíveis na aplicação
 */
export enum AppRoute {
  INDEX = '',
  DASHBOARD = 'dashboard',
  LANCAMENTOS = 'lancamentos',
  CONCILIACAO = 'conciliacao',
  DRE = 'dre',
  DFC = 'dfc',
  KPI = 'kpi',
  CONFIGURACOES = 'configuracoes',
}

/**
 * Mapeia rota para arquivo HTML
 */
const ROUTE_TO_FILE: Record<string, string> = {
  [AppRoute.INDEX]: 'index',
  [AppRoute.DASHBOARD]: 'dashboard',
  [AppRoute.LANCAMENTOS]: 'lancamentos',
  [AppRoute.CONCILIACAO]: 'conciliacao',
  [AppRoute.DRE]: 'dre',
  [AppRoute.DFC]: 'dfc',
  [AppRoute.KPI]: 'kpi',
  [AppRoute.CONFIGURACOES]: 'configuracoes',
};

/**
 * Renderiza uma view HTML
 *
 * @param viewName - Nome do arquivo HTML (sem extensão)
 * @returns HtmlOutput do Apps Script
 */
export function renderView(viewName: string): GoogleAppsScript.HTML.HtmlOutput {
  try {
    const template = HtmlService.createTemplateFromFile(`frontend/views/${viewName}`);

    // TODO: Injetar dados no template se necessário
    // template.data = { ... };

    const output = template.evaluate();
    output.setTitle('Neoformula Finance App');

    // TODO: Configurar sandbox mode e X-Frame-Options conforme necessário
    // output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
    // output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return output;
  } catch (error) {
    console.error(`Erro ao renderizar view ${viewName}:`, error);
    throw error;
  }
}

/**
 * Handler principal para doGet (roteamento)
 *
 * Este método deve ser chamado pelo main.ts
 *
 * @param e - Event object do Apps Script
 * @returns HtmlOutput
 */
export function handleDoGet(e: any): GoogleAppsScript.HTML.HtmlOutput {
  const route = (e.parameter.page || '') as string;

  // Valida rota
  const fileName = ROUTE_TO_FILE[route];

  if (!fileName) {
    // Rota não encontrada, redireciona para dashboard
    return renderView('dashboard');
  }

  return renderView(fileName);
}

/**
 * Include helper para carregar arquivos parciais (components, styles, scripts)
 *
 * Este método é chamado pelo template HTML via <?!= include('filename') ?>
 */
export function include(filename: string): string {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    console.error(`Erro ao incluir arquivo ${filename}:`, error);
    return `<!-- Erro ao carregar ${filename} -->`;
  }
}

// ============================================================================
// HANDLERS DE FORMULÁRIOS
// ============================================================================

/**
 * Processa formulário de novo lançamento
 *
 * Este método é chamado via google.script.run do frontend
 *
 * TODO: Implementar validação completa
 */
export function handleNovoLancamento(formData: any): { success: boolean; message: string; id?: string } {
  try {
    // TODO: Validar formData
    // TODO: Criar lançamento via ledger-service
    // const entry = createEntry(formData);

    return {
      success: true,
      message: 'Lançamento criado com sucesso',
      // id: entry.id,
    };
  } catch (error: any) {
    return {
      success: false,
      message: error.message || 'Erro ao criar lançamento',
    };
  }
}

/**
 * Processa formulário de conciliação
 */
export function handleConciliacao(statementId: string, entryId: string): { success: boolean; message: string } {
  try {
    // TODO: Chamar reconciliation-service.reconcile
    // reconcile(statementId, entryId);

    return {
      success: true,
      message: 'Conciliação realizada com sucesso',
    };
  } catch (error: any) {
    return {
      success: false,
      message: error.message || 'Erro ao conciliar',
    };
  }
}

/**
 * Processa atualização de configurações
 */
export function handleAtualizarConfig(configs: Record<string, any>): { success: boolean; message: string } {
  try {
    // TODO: Atualizar configs via config-service
    // for (const [key, value] of Object.entries(configs)) {
    //   updateConfig(key, value);
    // }

    return {
      success: true,
      message: 'Configurações atualizadas com sucesso',
    };
  } catch (error: any) {
    return {
      success: false,
      message: error.message || 'Erro ao atualizar configurações',
    };
  }
}

// ============================================================================
// HELPERS PARA FRONTEND
// ============================================================================

/**
 * Retorna dados para popular dropdowns/selects
 */
export function getFormOptions(): {
  filiais: Array<{ id: string; nome: string }>;
  canais: Array<{ id: string; nome: string }>;
  centrosCusto: Array<{ id: string; nome: string }>;
  contas: Array<{ codigo: string; descricao: string }>;
} {
  // TODO: Buscar de reference-data-service
  // const filiais = getAllBranches();
  // const canais = getAllChannels();
  // const centrosCusto = getAllCostCenters();
  // const contas = getAllAccounts();

  return {
    filiais: [],
    canais: [],
    centrosCusto: [],
    contas: [],
  };
}

/**
 * Valida formulário server-side
 *
 * TODO: Implementar validações específicas por tipo de formulário
 */
export function validateForm(formType: string, formData: any): { valid: boolean; errors: string[] } {
  const errors: string[] = [];

  // TODO: Implementar validações

  return {
    valid: errors.length === 0,
    errors,
  };
}
