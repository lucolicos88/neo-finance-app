/**
 * setup-sample-data.ts
 *
 * Popula dados de exemplo para teste da aplicação
 */

import {
  SHEET_TB_LANCAMENTOS,
  SHEET_TB_EXTRATOS,
} from './config/sheet-mapping';

/**
 * Popula dados de exemplo em TB_LANCAMENTOS
 */
export function populateSampleLancamentos(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Erro', 'Aba TB_LANCAMENTOS não encontrada', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const hoje = new Date();
  const mes = hoje.getMonth();
  const ano = hoje.getFullYear();

  const sampleData = [
    // Contas a Pagar - Vencidas
    [
      'CP001',
      new Date(ano, mes - 1, 15),
      new Date(ano, mes - 1, 20),
      '',
      'DESPESA',
      'MATRIZ',
      'ADM',
      'Aluguel',
      '3.02.001',
      '',
      '',
      'Aluguel Escritório - Janeiro',
      5000,
      0,
      0,
      0,
      5000,
      'VENCIDA',
      '',
      'MANUAL',
      'Pagamento em atraso',
    ],
    [
      'CP002',
      new Date(ano, mes - 1, 10),
      new Date(ano, mes - 1, 15),
      '',
      'DESPESA',
      'MATRIZ',
      'TI',
      'Serviços',
      '3.01.001',
      '',
      '',
      'Software - Licença Microsoft',
      3500,
      0,
      0,
      0,
      3500,
      'VENCIDA',
      '',
      'MANUAL',
      '',
    ],

    // Contas a Pagar - Próximos 7 dias
    [
      'CP003',
      new Date(ano, mes, 15),
      addDays(hoje, 3),
      '',
      'DESPESA',
      'MATRIZ',
      'COM',
      'Marketing',
      '3.03.001',
      '',
      '',
      'Google Ads - Campanha Fevereiro',
      2500,
      0,
      0,
      0,
      2500,
      'PENDENTE',
      '',
      'MANUAL',
      '',
    ],
    [
      'CP004',
      new Date(ano, mes, 10),
      addDays(hoje, 5),
      '',
      'DESPESA',
      'FILIAL_RJ',
      'ADM',
      'Salários',
      '3.01.001',
      '',
      '',
      'Salários Administrativos',
      15000,
      0,
      0,
      0,
      15000,
      'PENDENTE',
      '',
      'MANUAL',
      '',
    ],

    // Contas a Pagar - Pagas
    [
      'CP005',
      new Date(ano, mes, 1),
      new Date(ano, mes, 5),
      new Date(ano, mes, 5),
      'DESPESA',
      'MATRIZ',
      'ADM',
      'Energia',
      '3.02.002',
      '',
      '',
      'Conta de Luz - Matriz',
      800,
      0,
      0,
      0,
      800,
      'PAGA',
      '',
      'MANUAL',
      '',
    ],

    // Contas a Receber - Vencidas
    [
      'CR001',
      new Date(ano, mes - 1, 20),
      new Date(ano, mes - 1, 25),
      '',
      'RECEITA',
      'MATRIZ',
      'COM',
      'Receita Serviços',
      '1.01.001',
      'Serviços',
      'DIRETO',
      'Cliente ABC - Consultoria Janeiro',
      10000,
      0,
      0,
      0,
      10000,
      'VENCIDA',
      '',
      'MANUAL',
      'Cliente em atraso',
    ],

    // Contas a Receber - Hoje
    [
      'CR002',
      new Date(ano, mes, 15),
      hoje,
      '',
      'RECEITA',
      'MATRIZ',
      'COM',
      'Receita Serviços',
      '1.01.001',
      'Serviços',
      'ONLINE',
      'Cliente XYZ - Sistema',
      8500,
      0,
      0,
      0,
      8500,
      'PENDENTE',
      '',
      'MANUAL',
      '',
    ],

    // Contas a Receber - Futuras
    [
      'CR003',
      new Date(ano, mes, 10),
      addDays(hoje, 5),
      '',
      'RECEITA',
      'FILIAL_RJ',
      'COM',
      'Receita Produtos',
      '1.01.002',
      'Produtos',
      'PARCEIRO',
      'Venda Produtos - Empresa DEF',
      12000,
      0,
      0,
      0,
      12000,
      'PENDENTE',
      '',
      'MANUAL',
      '',
    ],
    [
      'CR004',
      new Date(ano, mes, 20),
      addDays(hoje, 10),
      '',
      'RECEITA',
      'MATRIZ',
      'COM',
      'Receita Serviços',
      '1.01.001',
      'Serviços',
      'LICITACAO',
      'Licitação - Prefeitura',
      25000,
      0,
      0,
      0,
      25000,
      'PENDENTE',
      '',
      'MANUAL',
      '',
    ],

    // Contas a Receber - Recebidas
    [
      'CR005',
      new Date(ano, mes, 1),
      new Date(ano, mes, 5),
      new Date(ano, mes, 5),
      'RECEITA',
      'MATRIZ',
      'COM',
      'Receita Serviços',
      '1.01.001',
      'Serviços',
      'DIRETO',
      'Cliente GHI - Mensalidade',
      7500,
      0,
      0,
      0,
      7500,
      'RECEBIDA',
      '',
      'MANUAL',
      '',
    ],
  ];

  // Adicionar dados após o cabeçalho
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  SpreadsheetApp.getUi().alert(
    'Dados de Exemplo Criados',
    `${sampleData.length} lançamentos de exemplo foram adicionados à planilha.\n\n` +
      '✓ Contas a Pagar (vencidas, pendentes, pagas)\n' +
      '✓ Contas a Receber (vencidas, pendentes, recebidas)\n\n' +
      'Você pode agora acessar a Web App para visualizar os dados.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Popula dados de exemplo em TB_EXTRATOS
 */
export function populateSampleExtratos(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TB_EXTRATOS);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Erro', 'Aba TB_EXTRATOS não encontrada', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const hoje = new Date();
  const mes = hoje.getMonth();
  const ano = hoje.getFullYear();

  const sampleData = [
    [
      'EXT001',
      new Date(ano, mes, 5),
      'PGTO FORNECEDOR ABC',
      -800,
      'DEBITO',
      'BRADESCO',
      '1234-5',
      'PENDENTE',
      '',
      '',
      new Date(),
    ],
    [
      'EXT002',
      new Date(ano, mes, 5),
      'RECEBIMENTO CLIENTE XYZ',
      7500,
      'CREDITO',
      'BRADESCO',
      '1234-5',
      'PENDENTE',
      '',
      '',
      new Date(),
    ],
    [
      'EXT003',
      addDays(hoje, -3),
      'TED FORNECEDOR DEF',
      -3500,
      'DEBITO',
      'ITAU',
      '5678-9',
      'PENDENTE',
      '',
      '',
      new Date(),
    ],
  ];

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  SpreadsheetApp.getUi().alert(
    'Extratos de Exemplo Criados',
    `${sampleData.length} extratos bancários foram adicionados.\n\n` +
      'Estes extratos podem ser conciliados com os lançamentos.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Função principal para popular todos os dados de exemplo
 */
export function setupAllSampleData(): void {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Criar Dados de Exemplo',
    'Deseja criar dados de exemplo para testar a aplicação?\n\n' +
      'Isso irá adicionar:\n' +
      '- 11 lançamentos (receitas e despesas)\n' +
      '- 3 extratos bancários\n\n' +
      'Continuar?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    populateSampleLancamentos();
    populateSampleExtratos();
  }
}

function addDays(date: Date, days: number): Date {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}
