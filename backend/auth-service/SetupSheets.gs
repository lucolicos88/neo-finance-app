/**
 * Script utilitário para criar as abas básicas no spreadsheet de configuração.
 * Abas: tb_usuarios, tb_filiais, tb_plano_contas, tb_benchmarks.
 * Rode manualmente no editor do Apps Script (função setupBaseSheets).
 */

var BASE_SPREADSHEET_ID = '1e-u2qTehu-iT4P68wP8nOQgcWlKU32VxAHsDzqP7Vnc';

function setupBaseSheets() {
  var ss = SpreadsheetApp.openById(BASE_SPREADSHEET_ID);

  ensureSheetWithHeader(ss, 'tb_usuarios', [
    'id',
    'nome',
    'email',
    'senha_hash',
    'papel',          // admin, socio, financeiro, gerente_filial
    'filial_padrao',
    'ativo'
  ]);

  ensureSheetWithHeader(ss, 'tb_filiais', [
    'id',
    'nome',
    'cnpj',
    'cidade',
    'uf',
    'ativa'
  ]);

  ensureSheetWithHeader(ss, 'tb_plano_contas', [
    'id',
    'codigo',
    'descricao',
    'tipo',       // receita, despesa, imposto, cpv
    'natureza',   // variavel, fixa
    'ativo'
  ]);

  ensureSheetWithHeader(ss, 'tb_benchmarks', [
    'id',
    'tipo',         // cpv, cma, despesas_variaveis, despesas_fixas, rejeicao, descontos
    'faixa_label',
    'faixa_min',
    'faixa_max',
    'cor'
  ]);
}

/**
 * Cria a aba se não existir e garante cabeçalho se estiver vazia.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @param {string[]} headers
 */
function ensureSheetWithHeader(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}
