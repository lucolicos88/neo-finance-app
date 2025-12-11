/**
 * ConfigService - leitura de planilhas de configuração.
 * Abas esperadas:
 * - tb_filiais: id, nome, cnpj, cidade, uf, ativa
 * - tb_plano_contas: id, codigo, descricao, tipo (receita, despesa, imposto, cpv), natureza (variavel, fixa), ativo
 * - tb_benchmarks: id, tipo, faixa_label, faixa_min, faixa_max, cor
 */
var CONFIG_SPREADSHEET_ID = '1e-u2qTehu-iT4P68wP8nOQgcWlKU32VxAHsDzqP7Vnc';

function listFiliais() {
  return readTable(CONFIG_SPREADSHEET_ID, 'tb_filiais');
}

function listPlanoContas() {
  return readTable(CONFIG_SPREADSHEET_ID, 'tb_plano_contas');
}

/**
 * Lista benchmarks filtrando por tipo (ex.: cpv, cma, despesas_variaveis, despesas_fixas, rejeicao, descontos).
 * @param {string} tipo
 * @return {Object[]}
 */
function listBenchmarks(tipo) {
  var all = readTable(CONFIG_SPREADSHEET_ID, 'tb_benchmarks');
  if (!tipo) return all;
  var target = String(tipo).toLowerCase();
  return all.filter(function (row) {
    return row.tipo && String(row.tipo).toLowerCase() === target;
  });
}
