/**
 * setup-sheets.ts
 *
 * Script para criar toda a estrutura de abas e dados iniciais da planilha
 */

import {
  SHEET_CFG_CONFIG,
  SHEET_CFG_BENCHMARKS,
  SHEET_CFG_LABELS,
  SHEET_CFG_THEME,
  SHEET_CFG_DFC,
  SHEET_CFG_VALIDATION,
  SHEET_REF_PLANO_CONTAS,
  SHEET_REF_FILIAIS,
  SHEET_REF_CANAIS,
  SHEET_REF_CCUSTO,
  SHEET_REF_NATUREZAS,
  SHEET_TB_LANCAMENTOS,
  SHEET_TB_EXTRATOS,
  SHEET_TB_DRE_MENSAL,
  SHEET_TB_DRE_RESUMO,
  SHEET_TB_DFC_REAL,
  SHEET_TB_DFC_PROJ,
  SHEET_TB_KPI_RESUMO,
  SHEET_TB_KPI_DETALHE,
  SHEET_RPT_COMITE_FATURAMENTO,
  SHEET_RPT_COMITE_DRE,
  SHEET_RPT_COMITE_DFC,
  SHEET_RPT_COMITE_KPIS,
} from './config/sheet-mapping';

/**
 * Cria todas as abas necessárias
 */
function setupAllSheets(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetsToCreate = [
    // Configuração
    SHEET_CFG_CONFIG,
    SHEET_CFG_BENCHMARKS,
    SHEET_CFG_LABELS,
    SHEET_CFG_THEME,
    SHEET_CFG_DFC,
    SHEET_CFG_VALIDATION,

    // Referência
    SHEET_REF_PLANO_CONTAS,
    SHEET_REF_FILIAIS,
    SHEET_REF_CANAIS,
    SHEET_REF_CCUSTO,
    SHEET_REF_NATUREZAS,

    // Transacional
    SHEET_TB_LANCAMENTOS,
    SHEET_TB_EXTRATOS,
    SHEET_TB_DRE_MENSAL,
    SHEET_TB_DRE_RESUMO,
    SHEET_TB_DFC_REAL,
    SHEET_TB_DFC_PROJ,
    SHEET_TB_KPI_RESUMO,
    SHEET_TB_KPI_DETALHE,

    // Relatórios
    SHEET_RPT_COMITE_FATURAMENTO,
    SHEET_RPT_COMITE_DRE,
    SHEET_RPT_COMITE_DFC,
    SHEET_RPT_COMITE_KPIS,
  ];

  const existingSheets = ss.getSheets().map(s => s.getName());
  let created = 0;
  let skipped = 0;

  sheetsToCreate.forEach(sheetName => {
    if (!existingSheets.includes(sheetName)) {
      ss.insertSheet(sheetName);
      created++;
    } else {
      skipped++;
    }
  });

  SpreadsheetApp.getUi().alert(
    'Setup Completo',
    `Abas criadas: ${created}\nAbas já existentes: ${skipped}\n\nTotal: ${sheetsToCreate.length} abas`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Popula dados iniciais nas abas de configuração
 */
function setupInitialData(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // CFG_CONFIG - Configurações gerais
  const cfgConfig = ss.getSheetByName(SHEET_CFG_CONFIG);
  if (cfgConfig) {
    cfgConfig.clear();
    cfgConfig.getRange('A1:E1').setValues([[
      'Chave', 'Valor', 'Tipo', 'Descrição', 'Ativo'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');

    cfgConfig.getRange('A2:E10').setValues([
      ['EMPRESA_NOME', 'Neoformula', 'TEXT', 'Nome da empresa', 'TRUE'],
      ['MOEDA_PADRAO', 'BRL', 'TEXT', 'Moeda padrão', 'TRUE'],
      ['TIMEZONE', 'America/Sao_Paulo', 'TEXT', 'Fuso horário', 'TRUE'],
      ['DRE_FORMATO', 'GERENCIAL', 'TEXT', 'Formato da DRE', 'TRUE'],
      ['CACHE_TTL_MINUTOS', '60', 'NUMBER', 'Tempo de cache em minutos', 'TRUE'],
      ['PERMITIR_LANCAMENTO_FUTURO', 'FALSE', 'BOOLEAN', 'Permitir lançamentos futuros', 'TRUE'],
      ['DIAS_AVISO_VENCIMENTO', '3', 'NUMBER', 'Dias de aviso antes do vencimento', 'TRUE'],
      ['EMAIL_NOTIFICACOES', 'financeiro@neoformula.com', 'TEXT', 'Email para notificações', 'TRUE'],
      ['APROVACAO_NECESSARIA', 'TRUE', 'BOOLEAN', 'Lançamentos precisam aprovação', 'TRUE'],
    ]);

    cfgConfig.autoResizeColumns(1, 5);
  }

  // REF_FILIAIS - Filiais
  const refFiliais = ss.getSheetByName(SHEET_REF_FILIAIS);
  if (refFiliais) {
    refFiliais.clear();
    refFiliais.getRange('A1:D1').setValues([[
      'Código', 'Nome', 'CNPJ', 'Ativa'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');

    refFiliais.getRange('A2:D5').setValues([
      ['MATRIZ', 'Matriz São Paulo', '00.000.000/0001-00', 'TRUE'],
      ['FILIAL_RJ', 'Filial Rio de Janeiro', '00.000.000/0002-00', 'TRUE'],
      ['FILIAL_BH', 'Filial Belo Horizonte', '00.000.000/0003-00', 'TRUE'],
      ['FILIAL_DF', 'Filial Brasília', '00.000.000/0004-00', 'TRUE'],
    ]);

    refFiliais.autoResizeColumns(1, 4);
  }

  // REF_PLANO_CONTAS - Plano de contas simplificado
  const refPlanoContas = ss.getSheetByName(SHEET_REF_PLANO_CONTAS);
  if (refPlanoContas) {
    refPlanoContas.clear();
    refPlanoContas.getRange('A1:H1').setValues([[
      'Código', 'Descrição', 'Tipo', 'Grupo DRE', 'Subgrupo DRE', 'Grupo DFC', 'Variável/Fixa', 'CMA/CMV'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');

    refPlanoContas.getRange('A2:H20').setValues([
      // Receitas
      ['1.01.001', 'Receita de Serviços', 'RECEITA', 'RECEITA_BRUTA', 'Serviços', 'OPERACIONAL', 'VARIAVEL', ''],
      ['1.01.002', 'Receita de Produtos', 'RECEITA', 'RECEITA_BRUTA', 'Produtos', 'OPERACIONAL', 'VARIAVEL', ''],
      ['1.02.001', 'Descontos Concedidos', 'RECEITA', 'DEDUCOES', 'Descontos', 'OPERACIONAL', 'VARIAVEL', ''],
      ['1.02.002', 'Impostos sobre Vendas', 'RECEITA', 'DEDUCOES', 'Impostos', 'OPERACIONAL', 'VARIAVEL', ''],

      // Custos
      ['2.01.001', 'Custo de Pessoal Direto', 'CUSTO', 'CMV_CSP', 'Pessoal', 'OPERACIONAL', 'VARIAVEL', 'CMV'],
      ['2.01.002', 'Materiais Diretos', 'CUSTO', 'CMV_CSP', 'Materiais', 'OPERACIONAL', 'VARIAVEL', 'CMV'],
      ['2.01.003', 'Serviços de Terceiros', 'CUSTO', 'CMV_CSP', 'Terceiros', 'OPERACIONAL', 'VARIAVEL', 'CMV'],

      // Despesas Operacionais
      ['3.01.001', 'Salários Administrativos', 'DESPESA', 'DESPESAS_OPERACIONAIS', 'Pessoal', 'OPERACIONAL', 'FIXA', ''],
      ['3.01.002', 'Encargos Sociais', 'DESPESA', 'DESPESAS_OPERACIONAIS', 'Pessoal', 'OPERACIONAL', 'FIXA', ''],
      ['3.02.001', 'Aluguel', 'DESPESA', 'DESPESAS_OPERACIONAIS', 'Ocupação', 'OPERACIONAL', 'FIXA', ''],
      ['3.02.002', 'Energia Elétrica', 'DESPESA', 'DESPESAS_OPERACIONAIS', 'Ocupação', 'OPERACIONAL', 'VARIAVEL', ''],
      ['3.03.001', 'Marketing e Publicidade', 'DESPESA', 'DESPESAS_OPERACIONAIS', 'Comercial', 'OPERACIONAL', 'VARIAVEL', ''],
      ['3.03.002', 'Comissões de Vendas', 'DESPESA', 'DESPESAS_OPERACIONAIS', 'Comercial', 'OPERACIONAL', 'VARIAVEL', ''],

      // Despesas Financeiras
      ['4.01.001', 'Juros de Empréstimos', 'DESPESA', 'RESULTADO_FINANCEIRO', 'Despesas Financeiras', 'FINANCEIRO', 'VARIAVEL', ''],
      ['4.01.002', 'Tarifas Bancárias', 'DESPESA', 'RESULTADO_FINANCEIRO', 'Despesas Financeiras', 'FINANCEIRO', 'FIXA', ''],

      // Receitas Financeiras
      ['4.02.001', 'Juros Recebidos', 'RECEITA', 'RESULTADO_FINANCEIRO', 'Receitas Financeiras', 'FINANCEIRO', 'VARIAVEL', ''],
      ['4.02.002', 'Descontos Obtidos', 'RECEITA', 'RESULTADO_FINANCEIRO', 'Receitas Financeiras', 'FINANCEIRO', 'VARIAVEL', ''],

      // Outras
      ['5.01.001', 'Impostos s/ Lucro', 'DESPESA', 'IMPOSTOS', 'Impostos', 'OPERACIONAL', 'VARIAVEL', ''],
      ['6.01.001', 'Resultado Não Operacional', 'RECEITA', 'OUTROS', 'Outros', 'OUTROS', 'VARIAVEL', ''],
    ]);

    refPlanoContas.autoResizeColumns(1, 8);
  }

  // REF_CCUSTO - Centros de custo
  const refCcusto = ss.getSheetByName(SHEET_REF_CCUSTO);
  if (refCcusto) {
    refCcusto.clear();
    refCcusto.getRange('A1:C1').setValues([[
      'Código', 'Descrição', 'Ativo'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');

    refCcusto.getRange('A2:C10').setValues([
      ['ADM', 'Administrativo', 'TRUE'],
      ['COM', 'Comercial', 'TRUE'],
      ['FIN', 'Financeiro', 'TRUE'],
      ['MKT', 'Marketing', 'TRUE'],
      ['OPS', 'Operações', 'TRUE'],
      ['TI', 'Tecnologia da Informação', 'TRUE'],
      ['RH', 'Recursos Humanos', 'TRUE'],
      ['JUR', 'Jurídico', 'TRUE'],
      ['LOG', 'Logística', 'TRUE'],
    ]);

    refCcusto.autoResizeColumns(1, 3);
  }

  // REF_CANAIS - Canais de venda
  const refCanais = ss.getSheetByName(SHEET_REF_CANAIS);
  if (refCanais) {
    refCanais.clear();
    refCanais.getRange('A1:C1').setValues([[
      'Código', 'Descrição', 'Ativo'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');

    refCanais.getRange('A2:C6').setValues([
      ['DIRETO', 'Venda Direta', 'TRUE'],
      ['ONLINE', 'E-commerce', 'TRUE'],
      ['PARCEIRO', 'Parceiros/Revendas', 'TRUE'],
      ['MARKETPLACE', 'Marketplaces', 'TRUE'],
      ['LICITACAO', 'Licitações', 'TRUE'],
    ]);

    refCanais.autoResizeColumns(1, 3);
  }

  // TB_LANCAMENTOS - Cabeçalho
  const tbLancamentos = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
  if (tbLancamentos) {
    tbLancamentos.clear();
    tbLancamentos.getRange('A1:U1').setValues([[
      'ID', 'Data Competência', 'Data Vencimento', 'Data Pagamento',
      'Tipo', 'Filial', 'Centro Custo', 'Conta Gerencial', 'Conta Contábil',
      'Grupo Receita', 'Canal', 'Descrição', 'Valor Bruto', 'Desconto',
      'Juros', 'Multa', 'Valor Líquido', 'Status', 'ID Extrato Banco',
      'Origem', 'Observações'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');

    tbLancamentos.autoResizeColumns(1, 21);
  }

  // TB_EXTRATOS - Extratos bancários
  const tbExtratos = ss.getSheetByName(SHEET_TB_EXTRATOS);
  if (tbExtratos) {
    tbExtratos.clear();
    tbExtratos.getRange('A1:K1').setValues([[
      'ID', 'Data', 'Descrição', 'Valor', 'Tipo', 'Banco',
      'Conta', 'Status Conciliação', 'ID Lançamento', 'Observações', 'Importado Em'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    tbExtratos.autoResizeColumns(1, 11);
  }

  // TB_DRE_MENSAL - DRE mensal
  const tbDreMensal = ss.getSheetByName(SHEET_TB_DRE_MENSAL);
  if (tbDreMensal) {
    tbDreMensal.clear();
    tbDreMensal.getRange('A1:H1').setValues([[
      'Mês/Ano', 'Grupo DRE', 'Subgrupo', 'Conta', 'Filial', 'Centro Custo', 'Valor', 'Atualizado Em'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    tbDreMensal.autoResizeColumns(1, 8);
  }

  // TB_DRE_RESUMO - DRE resumo
  const tbDreResumo = ss.getSheetByName(SHEET_TB_DRE_RESUMO);
  if (tbDreResumo) {
    tbDreResumo.clear();
    tbDreResumo.getRange('A1:F1').setValues([[
      'Mês/Ano', 'Grupo DRE', 'Valor Real', 'Valor Orçado', 'Variação', 'Variação %'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    tbDreResumo.autoResizeColumns(1, 6);
  }

  // TB_DFC_REAL - DFC realizado
  const tbDfcReal = ss.getSheetByName(SHEET_TB_DFC_REAL);
  if (tbDfcReal) {
    tbDfcReal.clear();
    tbDfcReal.getRange('A1:G1').setValues([[
      'Data', 'Grupo DFC', 'Descrição', 'Valor', 'Saldo Acumulado', 'Filial', 'ID Lançamento'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    tbDfcReal.autoResizeColumns(1, 7);
  }

  // TB_DFC_PROJ - DFC projetado
  const tbDfcProj = ss.getSheetByName(SHEET_TB_DFC_PROJ);
  if (tbDfcProj) {
    tbDfcProj.clear();
    tbDfcProj.getRange('A1:G1').setValues([[
      'Data Prevista', 'Grupo DFC', 'Descrição', 'Valor Previsto', 'Saldo Projetado', 'Filial', 'Status'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    tbDfcProj.autoResizeColumns(1, 7);
  }

  // TB_KPI_RESUMO - KPIs resumo
  const tbKpiResumo = ss.getSheetByName(SHEET_TB_KPI_RESUMO);
  if (tbKpiResumo) {
    tbKpiResumo.clear();
    tbKpiResumo.getRange('A1:H1').setValues([[
      'Mês/Ano', 'KPI', 'Valor', 'Meta', 'Variação', 'Status', 'Tendência', 'Atualizado Em'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    tbKpiResumo.autoResizeColumns(1, 8);
  }

  // TB_KPI_DETALHE - KPIs detalhados
  const tbKpiDetalhe = ss.getSheetByName(SHEET_TB_KPI_DETALHE);
  if (tbKpiDetalhe) {
    tbKpiDetalhe.clear();
    tbKpiDetalhe.getRange('A1:J1').setValues([[
      'Data', 'KPI', 'Dimensão', 'Valor Dimensão', 'Valor KPI', 'Meta', 'Filial', 'Canal', 'Centro Custo', 'Observações'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    tbKpiDetalhe.autoResizeColumns(1, 10);
  }

  // CFG_BENCHMARKS - Benchmarks e metas
  const cfgBenchmarks = ss.getSheetByName(SHEET_CFG_BENCHMARKS);
  if (cfgBenchmarks) {
    cfgBenchmarks.clear();
    cfgBenchmarks.getRange('A1:F1').setValues([[
      'Métrica', 'Valor Meta', 'Benchmark Mercado', 'Período', 'Unidade', 'Ativo'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    cfgBenchmarks.autoResizeColumns(1, 6);
  }

  // CFG_LABELS - Labels personalizadas
  const cfgLabels = ss.getSheetByName(SHEET_CFG_LABELS);
  if (cfgLabels) {
    cfgLabels.clear();
    cfgLabels.getRange('A1:D1').setValues([[
      'Chave', 'Label PT-BR', 'Label EN', 'Categoria'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    cfgLabels.autoResizeColumns(1, 4);
  }

  // CFG_THEME - Configurações de tema
  const cfgTheme = ss.getSheetByName(SHEET_CFG_THEME);
  if (cfgTheme) {
    cfgTheme.clear();
    cfgTheme.getRange('A1:D1').setValues([[
      'Elemento', 'Propriedade', 'Valor', 'Descrição'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    cfgTheme.autoResizeColumns(1, 4);
  }

  // CFG_DFC - Configurações DFC
  const cfgDfc = ss.getSheetByName(SHEET_CFG_DFC);
  if (cfgDfc) {
    cfgDfc.clear();
    cfgDfc.getRange('A1:D1').setValues([[
      'Grupo DFC', 'Ordem', 'Descrição', 'Tipo Fluxo'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    cfgDfc.autoResizeColumns(1, 4);
  }

  // CFG_VALIDATION - Regras de validação
  const cfgValidation = ss.getSheetByName(SHEET_CFG_VALIDATION);
  if (cfgValidation) {
    cfgValidation.clear();
    cfgValidation.getRange('A1:E1').setValues([[
      'Regra', 'Campo', 'Tipo Validação', 'Parâmetros', 'Mensagem Erro'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    cfgValidation.autoResizeColumns(1, 5);
  }

  // REF_NATUREZAS - Naturezas financeiras
  const refNaturezas = ss.getSheetByName(SHEET_REF_NATUREZAS);
  if (refNaturezas) {
    refNaturezas.clear();
    refNaturezas.getRange('A1:D1').setValues([[
      'Código', 'Descrição', 'Tipo', 'Ativa'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    refNaturezas.autoResizeColumns(1, 4);
  }

  // RPT_COMITE_FATURAMENTO - Relatório de faturamento
  const rptFaturamento = ss.getSheetByName(SHEET_RPT_COMITE_FATURAMENTO);
  if (rptFaturamento) {
    rptFaturamento.clear();
    rptFaturamento.getRange('A1:H1').setValues([[
      'Período', 'Filial', 'Canal', 'Faturamento Bruto', 'Deduções', 'Faturamento Líquido', 'Meta', 'Atingimento %'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    rptFaturamento.autoResizeColumns(1, 8);
  }

  // RPT_COMITE_DRE - Relatório DRE
  const rptDre = ss.getSheetByName(SHEET_RPT_COMITE_DRE);
  if (rptDre) {
    rptDre.clear();
    rptDre.getRange('A1:F1').setValues([[
      'Período', 'Linha DRE', 'Valor Real', 'Valor Orçado', 'Variação', 'Variação %'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    rptDre.autoResizeColumns(1, 6);
  }

  // RPT_COMITE_DFC - Relatório DFC
  const rptDfc = ss.getSheetByName(SHEET_RPT_COMITE_DFC);
  if (rptDfc) {
    rptDfc.clear();
    rptDfc.getRange('A1:E1').setValues([[
      'Período', 'Grupo DFC', 'Valor Realizado', 'Valor Projetado', 'Saldo Final'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    rptDfc.autoResizeColumns(1, 5);
  }

  // RPT_COMITE_KPIS - Relatório KPIs
  const rptKpis = ss.getSheetByName(SHEET_RPT_COMITE_KPIS);
  if (rptKpis) {
    rptKpis.clear();
    rptKpis.getRange('A1:G1').setValues([[
      'Período', 'KPI', 'Valor', 'Meta', 'Variação', 'Status', 'Observações'
    ]]).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    rptKpis.autoResizeColumns(1, 7);
  }

  SpreadsheetApp.getUi().alert(
    'Dados Iniciais Criados',
    'Estrutura da planilha configurada com sucesso!\n\n' +
    '✓ Configurações iniciais\n' +
    '✓ Filiais de exemplo\n' +
    '✓ Plano de contas básico\n' +
    '✓ Centros de custo\n' +
    '✓ Canais de venda\n' +
    '✓ Estrutura de lançamentos\n' +
    '✓ Todas as tabelas transacionais\n' +
    '✓ Relatórios do comitê',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Função para executar o setup completo
 */
function runCompleteSetup(): void {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup da Planilha',
    'Este processo irá:\n\n' +
    '1. Criar todas as abas necessárias\n' +
    '2. Configurar estrutura inicial\n' +
    '3. Adicionar dados de exemplo\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    setupAllSheets();
    setupInitialData();

    ui.alert(
      'Setup Completo!',
      'A planilha está pronta para uso!\n\n' +
      'Próximo passo: teste o Dashboard no menu\n' +
      'Neoformula Finance → Abrir Dashboard',
      ui.ButtonSet.OK
    );
  }
}

// Exporta as funções
export { setupAllSheets, setupInitialData, runCompleteSetup };
