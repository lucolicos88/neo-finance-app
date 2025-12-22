/**
 * setup-sample-data.ts
 *
 * Popula dados de exemplo para teste da aplicação
 */

import {
  SHEET_TB_LANCAMENTOS,
  SHEET_TB_EXTRATOS,
  SHEET_REF_FILIAIS,
  SHEET_REF_CANAIS,
  SHEET_REF_CCUSTO,
  SHEET_REF_PLANO_CONTAS,
} from './config/sheet-mapping';
import { getSheetValues } from './shared/sheets-client';

type SeedRefs = {
  filiais: string[];
  canais: string[];
  centros: string[];
  contasReceita: string[];
  contasDespesa: string[];
};

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

function getSeedReferences(): SeedRefs {
  const filiais = getSheetValues(SHEET_REF_FILIAIS, { skipHeader: true })
    .map((r) => String(r[0] || '').trim())
    .filter(Boolean);
  const canais = getSheetValues(SHEET_REF_CANAIS, { skipHeader: true })
    .map((r) => String(r[0] || '').trim())
    .filter(Boolean);
  const centros = getSheetValues(SHEET_REF_CCUSTO, { skipHeader: true })
    .map((r) => String(r[0] || '').trim())
    .filter(Boolean);

  const contasReceita: string[] = [];
  const contasDespesa: string[] = [];
  const contas = getSheetValues(SHEET_REF_PLANO_CONTAS, { skipHeader: true });
  for (const row of contas) {
    const codigo = String(row[0] || '').trim();
    const tipo = String(row[2] || '').trim().toUpperCase();
    if (!codigo) continue;
    if (tipo === 'RECEITA') contasReceita.push(codigo);
    if (tipo === 'DESPESA' || tipo === 'CUSTO') contasDespesa.push(codigo);
  }

  return {
    filiais: filiais.length ? filiais : ['MATRIZ', 'FILIAL_RJ', 'FILIAL_SP'],
    canais: canais.length ? canais : ['DIRETO', 'ONLINE', 'PARCEIRO'],
    centros: centros.length ? centros : ['ADM', 'COM', 'FIN', 'MKT', 'OPS', 'TI'],
    contasReceita: contasReceita.length ? contasReceita : ['1.01.001', '1.01.002'],
    contasDespesa: contasDespesa.length ? contasDespesa : ['3.01.001', '3.02.001', '3.02.002', '3.03.001'],
  };
}

function randomFrom<T>(items: T[]): T {
  return items[Math.floor(Math.random() * items.length)];
}

function randomNumber(min: number, max: number): number {
  return Math.round((min + Math.random() * (max - min)) * 100) / 100;
}

function randomDateBetween(start: Date, end: Date): Date {
  const ts = start.getTime() + Math.random() * (end.getTime() - start.getTime());
  return new Date(ts);
}

export function setupBulkSampleData(): void {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Criar Dados de Exemplo (Massa)',
    'Deseja gerar muitos dados ficticios para teste?\n' +
      'Isso adiciona lancamentos e extratos sem apagar os atuais.\n\n' +
      'Continuar?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const lancPrompt = ui.prompt('Quantidade de lancamentos', '300', ui.ButtonSet.OK_CANCEL);
  if (lancPrompt.getSelectedButton() !== ui.Button.OK) return;
  const extrPrompt = ui.prompt('Quantidade de extratos', '150', ui.ButtonSet.OK_CANCEL);
  if (extrPrompt.getSelectedButton() !== ui.Button.OK) return;

  const numLanc = Math.max(50, Math.min(2000, Number(lancPrompt.getResponseText()) || 300));
  const numExtr = Math.max(20, Math.min(2000, Number(extrPrompt.getResponseText()) || 150));

  populateBulkSampleData(numLanc, numExtr);
}

function populateBulkSampleData(numLanc: number, numExtr: number): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lancSheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
  const extrSheet = ss.getSheetByName(SHEET_TB_EXTRATOS);

  if (!lancSheet) {
    SpreadsheetApp.getUi().alert('Erro', 'Aba TB_LANCAMENTOS nao encontrada', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  if (!extrSheet) {
    SpreadsheetApp.getUi().alert('Erro', 'Aba TB_EXTRATOS nao encontrada', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const refs = getSeedReferences();
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth() - 4, 1);
  const end = new Date(now.getFullYear(), now.getMonth() + 2, 28);

  const paidCandidates: Array<{ id: string; data: Date; valor: number; tipo: string; descricao: string }> = [];
  const rows: any[][] = [];

  for (let i = 0; i < numLanc; i++) {
    const isReceita = Math.random() < 0.5;
    const tipo = isReceita ? 'RECEITA' : 'DESPESA';
    const idPrefix = isReceita ? 'CR' : 'CP';
    const id = `${idPrefix}-${Utilities.getUuid().slice(0, 8).toUpperCase()}-${i + 1}`;
    const dataCompetencia = randomDateBetween(start, end);
    let dataVencimento = addDays(dataCompetencia, Math.floor(Math.random() * 40) + 3);

    const statusPool = isReceita
      ? ['PENDENTE', 'VENCIDA', 'RECEBIDA', 'CANCELADA']
      : ['PENDENTE', 'VENCIDA', 'PAGA', 'CANCELADA'];
    const status = randomFrom(statusPool);

    if (status === 'VENCIDA') {
      dataVencimento = addDays(now, -Math.floor(Math.random() * 20) - 1);
    }

    let dataPagamento = '';
    if (status === 'PAGA' || status === 'RECEBIDA') {
      const pg = addDays(dataVencimento, Math.floor(Math.random() * 5) - 2);
      dataPagamento = pg > now ? now.toISOString() : pg.toISOString();
    }

    const valorBruto = randomNumber(200, 25000);
    const desconto = Math.random() < 0.2 ? randomNumber(10, 300) : 0;
    const juros = status === 'VENCIDA' ? randomNumber(5, 100) : 0;
    const multa = status === 'VENCIDA' ? randomNumber(5, 80) : 0;
    const valorLiquido = valorBruto - desconto + juros + multa;

    const filial = randomFrom(refs.filiais);
    const centro = randomFrom(refs.centros);
    const canal = isReceita ? randomFrom(refs.canais) : '';
    const contaContabil = isReceita ? randomFrom(refs.contasReceita) : randomFrom(refs.contasDespesa);
    const contaGerencial = isReceita ? 'Receita' : 'Despesa';
    const grupoReceita = isReceita ? (Math.random() < 0.5 ? 'Servicos' : 'Produtos') : '';
    const descricao = `${tipo} seed ${i + 1} ${filial}`;

    rows.push([
      id,
      dataCompetencia,
      dataVencimento,
      dataPagamento ? new Date(dataPagamento) : '',
      tipo,
      filial,
      centro,
      contaGerencial,
      contaContabil,
      grupoReceita,
      canal,
      descricao,
      valorBruto,
      desconto,
      juros,
      multa,
      valorLiquido,
      status,
      '',
      'SEED',
      'Gerado automaticamente',
    ]);

    if (status === 'PAGA' || status === 'RECEBIDA') {
      const pgDate = dataPagamento ? new Date(dataPagamento) : dataVencimento;
      paidCandidates.push({ id, data: pgDate, valor: valorLiquido, tipo, descricao });
    }
  }

  const startRow = lancSheet.getLastRow() + 1;
  lancSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

  const bancos = ['BRADESCO', 'ITAU', 'SANTANDER', 'BB', 'CAIXA'];
  const contas = ['1234-5', '5678-9', '9999-0'];
  const extrRows: any[][] = [];

  const conciliarCount = Math.min(Math.floor(numExtr * 0.5), paidCandidates.length);
  for (let i = 0; i < conciliarCount; i++) {
    const item = paidCandidates[i];
    const credito = item.tipo === 'RECEITA';
    extrRows.push([
      `EXT-${Utilities.getUuid().slice(0, 8).toUpperCase()}-${i + 1}`,
      item.data,
      `Conciliado ${item.descricao}`,
      credito ? item.valor : -Math.abs(item.valor),
      credito ? 'CREDITO' : 'DEBITO',
      randomFrom(bancos),
      randomFrom(contas),
      'CONCILIADO',
      item.id,
      '',
      new Date(),
    ]);
  }

  for (let i = conciliarCount; i < numExtr; i++) {
    const credito = Math.random() < 0.5;
    const valor = randomNumber(50, 15000);
    const data = randomDateBetween(start, end);
    extrRows.push([
      `EXT-${Utilities.getUuid().slice(0, 8).toUpperCase()}-${i + 1}`,
      data,
      `Movimento seed ${i + 1}`,
      credito ? valor : -Math.abs(valor),
      credito ? 'CREDITO' : 'DEBITO',
      randomFrom(bancos),
      randomFrom(contas),
      'PENDENTE',
      '',
      '',
      new Date(),
    ]);
  }

  const extrStartRow = extrSheet.getLastRow() + 1;
  extrSheet.getRange(extrStartRow, 1, extrRows.length, extrRows[0].length).setValues(extrRows);

  SpreadsheetApp.getUi().alert(
    'Dados de Exemplo Criados',
    `${rows.length} lancamentos e ${extrRows.length} extratos foram adicionados.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
