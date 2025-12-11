/**
 * Repositório de Contas a Pagar em planilha tb_contas_pagar.
 * Colunas: id, fornecedor, descricao, categoria, centro_custo, plano_contas, filial,
 * data_lancamento, data_vencimento, data_pagamento, valor_bruto, valor_pago,
 * status, origem, observacao
 */
var AP_SPREADSHEET_ID = '1e-u2qTehu-iT4P68wP8nOQgcWlKU32VxAHsDzqP7Vnc';
var AP_SHEET_NAME = 'tb_contas_pagar';

/**
 * Lista contas a pagar aplicando filtros simples (status, filial, intervalo de vencimento).
 * @param {Object} filtros
 * @return {Object[]}
 */
function listAp(filtros) {
  filtros = filtros || {};
  var dataInicio = filtros.dataInicio ? new Date(filtros.dataInicio) : null;
  var dataFim = filtros.dataFim ? new Date(filtros.dataFim) : null;
  var statusFiltro = filtros.status ? String(filtros.status).toLowerCase() : null;
  var filialFiltro = filtros.filial ? String(filtros.filial).toLowerCase() : null;

  var rows = readTable(AP_SPREADSHEET_ID, AP_SHEET_NAME);
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

/**
 * Cria um novo lançamento de AP gerando id sequencial.
 * @param {Object} apObject
 * @return {Object} registro criado
 */
function createAp(apObject) {
  var sheet = getSheetByName(AP_SPREADSHEET_ID, AP_SHEET_NAME);
  var data = readTable(AP_SPREADSHEET_ID, AP_SHEET_NAME);
  var nextId = data.reduce(function (max, row) {
    var val = Number(row.id);
    return isNaN(val) ? max : Math.max(max, val);
  }, 0) + 1;

  var record = Object.assign({}, apObject, {
    id: nextId,
    status: apObject.status || 'aberto'
  });
  sheetCache.clearCache();
  writeRow(AP_SPREADSHEET_ID, AP_SHEET_NAME, record);
  return record;
}

/**
 * Atualiza um lançamento pelo id.
 * @param {string|number} id
 * @param {Object} apObject
 * @return {boolean} true se atualizou
 */
function updateAp(id, apObject) {
  var sheet = getSheetByName(AP_SPREADSHEET_ID, AP_SHEET_NAME);
  var values = sheet.getDataRange().getValues();
  if (!values.length) return false;
  var header = values[0];
  var idCol = header.indexOf('id');
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) {
      for (var c = 0; c < header.length; c++) {
        var key = String(header[c]).trim();
        if (key && apObject.hasOwnProperty(key)) {
          values[i][c] = apObject[key];
        }
      }
      sheet.getRange(1, 1, values.length, header.length).setValues(values);
      sheetCache.clearCache();
      return true;
    }
  }
  return false;
}

/**
 * Marca pagamento de AP.
 * @param {string|number} id
 * @param {Date|string} data_pagamento
 * @param {number} valor_pago
 * @return {boolean}
 */
function pagarAp(id, data_pagamento, valor_pago) {
  var payload = {
    data_pagamento: data_pagamento,
    valor_pago: valor_pago,
    status: 'pago'
  };
  return updateAp(id, payload);
}
