/**
 * CashflowService
 * Planilha: tb_movimentos_caixa
 * Colunas: id, data, tipo, categoria, subcategoria, descricao, valor, servico_origem,
 * id_origem, conta_bancaria, filial
 */
var CASHFLOW_SPREADSHEET_ID = '1e-u2qTehu-iT4P68wP8nOQgcWlKU32VxAHsDzqP7Vnc';
var CASHFLOW_SHEET_NAME = 'tb_movimentos_caixa';

/**
 * Converte listas de AP/AR em movimentos de caixa (não grava).
 * apList: [{id, data_pagamento, valor_pago, status, ...}]
 * arList: [{id, data_recebimento, valor_liquido, status, ...}]
 */
function gerarMovimentosDeAp(apList, arList) {
  var movimentos = [];
  (apList || []).forEach(function (ap) {
    if (String(ap.status).toLowerCase() === 'pago') {
      movimentos.push({
        id_origem: ap.id,
        servico_origem: 'ap',
        data: ap.data_pagamento,
        tipo: 'saida',
        valor: Number(ap.valor_pago) || 0,
        descricao: ap.descricao || ap.fornecedor || 'AP',
        filial: ap.filial || ''
      });
    }
  });
  (arList || []).forEach(function (ar) {
    if (String(ar.status).toLowerCase() === 'recebido') {
      movimentos.push({
        id_origem: ar.id,
        servico_origem: 'ar',
        data: ar.data_recebimento,
        tipo: 'entrada',
        valor: Number(ar.valor_liquido) || 0,
        descricao: ar.cliente || 'AR',
        filial: ar.filial || ''
      });
    }
  });
  return movimentos;
}

/**
 * Calcula fluxo realizado com base na planilha tb_movimentos_caixa.
 * @param {{dataInicio: Date|string, dataFim: Date|string}} periodo
 * @param {string} filial
 * @param {number} saldoInicial
 * @return {Object} { dias:[{data, entradas, saidas, saldo}], saldoFinal, totalEntradas, totalSaidas }
 */
function calcularFluxoRealizado(periodo, filial, saldoInicial) {
  saldoInicial = Number(saldoInicial) || 0;
  var dataIni = periodo && periodo.dataInicio ? new Date(periodo.dataInicio) : null;
  var dataFim = periodo && periodo.dataFim ? new Date(periodo.dataFim) : null;
  var rows = readTable(CASHFLOW_SPREADSHEET_ID, CASHFLOW_SHEET_NAME);

  var filtrados = rows.filter(function (row) {
    if (filial && row.filial && String(row.filial).toLowerCase() !== String(filial).toLowerCase()) return false;
    var d = row.data instanceof Date ? row.data : new Date(row.data);
    if (dataIni && d < dataIni) return false;
    if (dataFim && d > dataFim) return false;
    return true;
  });

  var porDia = {};
  filtrados.forEach(function (row) {
    var d = row.data instanceof Date ? row.data : new Date(row.data);
    var key = formatDateKey(d);
    if (!porDia[key]) porDia[key] = { data: d, entradas: 0, saidas: 0 };
    if (String(row.tipo).toLowerCase() === 'entrada') {
      porDia[key].entradas += Number(row.valor) || 0;
    } else {
      porDia[key].saidas += Number(row.valor) || 0;
    }
  });

  var diasOrdenados = Object.keys(porDia).sort().map(function (k) { return porDia[k]; });
  var saldo = saldoInicial;
  diasOrdenados.forEach(function (dia) {
    saldo += dia.entradas - dia.saidas;
    dia.saldo = saldo;
  });

  var totalEntradas = diasOrdenados.reduce(function (acc, d) { return acc + d.entradas; }, 0);
  var totalSaidas = diasOrdenados.reduce(function (acc, d) { return acc + d.saidas; }, 0);

  return {
    dias: diasOrdenados,
    saldoFinal: saldo,
    totalEntradas: totalEntradas,
    totalSaidas: totalSaidas
  };
}

/**
 * Fluxo projetado baseado em títulos futuros de AP/AR já carregados.
 * @param {{saldoInicial:number, apAbertos:Object[], arAbertos:Object[], agruparPor:string}} parametros
 * @return {Object} { dias:[...], saldoFinal }
 */
function calcularFluxoProjetado(parametros) {
  parametros = parametros || {};
  var saldo = Number(parametros.saldoInicial) || 0;
  var ap = parametros.apAbertos || [];
  var ar = parametros.arAbertos || [];
  var agruparPor = parametros.agruparPor || 'dia'; // dia ou mes

  var movimentos = [];
  ap.forEach(function (item) {
    movimentos.push({
      data: item.data_vencimento,
      tipo: 'saida',
      valor: Number(item.valor_bruto) || 0
    });
  });
  ar.forEach(function (item) {
    movimentos.push({
      data: item.data_vencimento,
      tipo: 'entrada',
      valor: Number(item.valor_liquido || item.valor_bruto) || 0
    });
  });

  var bucket = {};
  movimentos.forEach(function (mov) {
    var d = mov.data instanceof Date ? mov.data : new Date(mov.data);
    var key = (agruparPor === 'mes') ? (d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2)) : formatDateKey(d);
    if (!bucket[key]) bucket[key] = { data: d, entradasPrevistas: 0, saidasPrevistas: 0 };
    if (mov.tipo === 'entrada') bucket[key].entradasPrevistas += mov.valor;
    else bucket[key].saidasPrevistas += mov.valor;
  });

  var dias = Object.keys(bucket).sort().map(function (k) { return bucket[k]; });
  dias.forEach(function (dia) {
    saldo += dia.entradasPrevistas - dia.saidasPrevistas;
    dia.saldoFinal = saldo;
  });

  return {
    dias: dias,
    saldoFinal: saldo
  };
}

function formatDateKey(d) {
  var day = ('0' + d.getDate()).slice(-2);
  var month = ('0' + (d.getMonth() + 1)).slice(-2);
  var year = d.getFullYear();
  return year + '-' + month + '-' + day;
}
