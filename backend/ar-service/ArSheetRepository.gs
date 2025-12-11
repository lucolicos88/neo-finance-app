/**
 * Repositório de Contas a Receber em planilha tb_contas_receber.
 * Colunas: id, cliente, origem, filial, grupo_receita, data_emissao, data_vencimento,
 * data_recebimento, valor_bruto, descontos, impostos, valor_liquido, status, forma_pagamento
 */
var AR_SPREADSHEET_ID = '1e-u2qTehu-iT4P68wP8nOQgcWlKU32VxAHsDzqP7Vnc';
var AR_SHEET_NAME = 'tb_contas_receber';

function listAr(filtros) {
  filtros = filtros || {};
  var statusFiltro = filtros.status ? String(filtros.status).toLowerCase() : null;
  var filialFiltro = filtros.filial ? String(filtros.filial).toLowerCase() : null;
  var dataInicio = filtros.dataInicio ? new Date(filtros.dataInicio) : null;
  var dataFim = filtros.dataFim ? new Date(filtros.dataFim) : null;

  var rows = readTable(AR_SPREADSHEET_ID, AR_SHEET_NAME);
  return rows.filter(function (row) {
    if (statusFiltro && row.status && String(row.status).toLowerCase() !== statusFiltro) return false;
    if (filialFiltro && row.filial && String(row.filial).toLowerCase() !== filialFiltro) return false;
    if (dataInicio || dataFim) {
      var venc = row.data_vencimento instanceof Date ? row.data_vencimento : new Date(row.data_vencimento);
      if (dataInicio && venc < dataInicio) return false;
      if (dataFim && venc > dataFim) return false;
    }
    return true;
  });
}

function createAr(arObject) {
  var sheet = getSheetByName(AR_SPREADSHEET_ID, AR_SHEET_NAME);
  var data = readTable(AR_SPREADSHEET_ID, AR_SHEET_NAME);
  var nextId = data.reduce(function (max, row) {
    var val = Number(row.id);
    return isNaN(val) ? max : Math.max(max, val);
  }, 0) + 1;

  var record = Object.assign({}, arObject, {
    id: nextId,
    status: arObject.status || 'aberto'
  });
  sheetCache.clearCache();
  writeRow(AR_SPREADSHEET_ID, AR_SHEET_NAME, record);
  return record;
}

function updateAr(id, arObject) {
  var sheet = getSheetByName(AR_SPREADSHEET_ID, AR_SHEET_NAME);
  var values = sheet.getDataRange().getValues();
  if (!values.length) return false;
  var header = values[0];
  var idCol = header.indexOf('id');
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) {
      for (var c = 0; c < header.length; c++) {
        var key = String(header[c]).trim();
        if (key && arObject.hasOwnProperty(key)) {
          values[i][c] = arObject[key];
        }
      }
      sheet.getRange(1, 1, values.length, header.length).setValues(values);
      sheetCache.clearCache();
      return true;
    }
  }
  return false;
}

function receberAr(id, data_recebimento, valor_recebido, descontos, impostos) {
  var payload = {
    data_recebimento: data_recebimento,
    valor_liquido: valor_recebido,
    descontos: descontos,
    impostos: impostos,
    status: 'recebido'
  };
  return updateAr(id, payload);
}
