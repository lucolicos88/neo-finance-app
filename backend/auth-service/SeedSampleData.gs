/**
 * Gera dados fictícios para:
 * - tb_contas_pagar
 * - tb_contas_receber
 * - tb_movimentos_caixa
 *
 * Rode uma vez a função seedSampleData() no editor do Apps Script.
 */

function seedSampleData() {
  var ssId = '1e-u2qTehu-iT4P68wP8nOQgcWlKU32VxAHsDzqP7Vnc';

  var contasPagar = [
    {
      id: 1,
      fornecedor: 'Fornecedor Alfa',
      descricao: 'Compra de materiais',
      categoria: 'Operacional',
      centro_custo: 'ADM',
      plano_contas: '5.1.1',
      filial: 'bosque',
      data_lancamento: new Date(),
      data_vencimento: addDays(new Date(), 5),
      data_pagamento: '',
      valor_bruto: 1500,
      valor_pago: '',
      status: 'aberto',
      origem: 'manual',
      observacao: 'Lote inicial'
    },
    {
      id: 2,
      fornecedor: 'Fornecedor Beta',
      descricao: 'Serviços de TI',
      categoria: 'Servicos',
      centro_custo: 'TI',
      plano_contas: '5.2.3',
      filial: 'bosque',
      data_lancamento: new Date(),
      data_vencimento: addDays(new Date(), 10),
      data_pagamento: addDays(new Date(), 2),
      valor_bruto: 900,
      valor_pago: 900,
      status: 'pago',
      origem: 'manual',
      observacao: ''
    }
  ];

  var contasReceber = [
    {
      id: 1,
      cliente: 'Cliente A',
      origem: 'Venda',
      filial: 'bosque',
      grupo_receita: 'Consultoria',
      data_emissao: new Date(),
      data_vencimento: addDays(new Date(), 7),
      data_recebimento: '',
      valor_bruto: 3500,
      descontos: 0,
      impostos: 350,
      valor_liquido: 3150,
      status: 'aberto',
      forma_pagamento: 'Boleto'
    },
    {
      id: 2,
      cliente: 'Cliente B',
      origem: 'Venda',
      filial: 'bosque',
      grupo_receita: 'Serviços',
      data_emissao: addDays(new Date(), -5),
      data_vencimento: addDays(new Date(), -1),
      data_recebimento: addDays(new Date(), -1),
      valor_bruto: 2000,
      descontos: 0,
      impostos: 200,
      valor_liquido: 1800,
      status: 'recebido',
      forma_pagamento: 'PIX'
    }
  ];

  var movimentosCaixa = [
    {
      id: 1,
      data: addDays(new Date(), -2),
      tipo: 'entrada',
      categoria: 'Recebimento',
      subcategoria: 'Venda',
      descricao: 'Recebimento Cliente B',
      valor: 1800,
      servico_origem: 'ar',
      id_origem: 2,
      conta_bancaria: 'Conta Corrente',
      filial: 'bosque'
    },
    {
      id: 2,
      data: addDays(new Date(), -1),
      tipo: 'saida',
      categoria: 'Pagamento',
      subcategoria: 'Fornecedor',
      descricao: 'Pagamento Fornecedor Beta',
      valor: 900,
      servico_origem: 'ap',
      id_origem: 2,
      conta_bancaria: 'Conta Corrente',
      filial: 'bosque'
    }
  ];

  sheetCache.clearCache();
  appendObjects(ssId, 'tb_contas_pagar', contasPagar);
  appendObjects(ssId, 'tb_contas_receber', contasReceber);
  appendObjects(ssId, 'tb_movimentos_caixa', movimentosCaixa);
}

function appendObjects(spreadsheetId, sheetName, objects) {
  if (!objects || !objects.length) return;
  var sheet = getSheetByName(spreadsheetId, sheetName);
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rows = objects.map(function (obj) {
    return header.map(function (h) { return obj[h] || ''; });
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, header.length).setValues(rows);
}

function addDays(date, days) {
  var d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}
