/**
 * Testes simples usando QUnitGS (ou similar).
 * Certifique-se de ter o framework carregado no projeto.
 */

function test_calcularDreMensal() {
  var receitas = [{ filial: 'A', mes: '2025-01', valor: 100000 }];
  var despesas = [
    { filial: 'A', mes: '2025-01', tipo: 'imposto', natureza: 'variavel', valor: 10000 },
    { filial: 'A', mes: '2025-01', tipo: 'cpv', natureza: 'variavel', valor: 30000 },
    { filial: 'A', mes: '2025-01', tipo: 'despesa', natureza: 'variavel', valor: 10000 },
    { filial: 'A', mes: '2025-01', tipo: 'despesa', natureza: 'fixa', valor: 15000 }
  ];
  var dre = calcularDreMensal(receitas, despesas, { filtrarFilial: 'A', filtrarMes: '2025-01' });
  QUnit.assert_equal(100000, dre.receita_bruta, 'Receita bruta');
  QUnit.assert_equal(90000, dre.receita_liquida, 'Receita líquida');
  QUnit.assert_equal(50000, dre.margem_contribuicao, 'Margem de contribuição');
  QUnit.assert_equal(35000, dre.ebitda, 'EBITDA');
  QUnit.assert_equal(35000, dre.lucro_liquido, 'Lucro líquido');
}

function test_classificarIndicadores() {
  var dre = {
    receita_liquida: 100,
    cpv_total: 30,
    despesas_variaveis: 20,
    despesas_fixas: 10
  };
  var benchmarks = {
    cpv: [{ label: 'OK', min: 0, max: 0.4, cor: 'green' }],
    desp_var: [{ label: 'OK', min: 0, max: 0.3, cor: 'blue' }],
    desp_fixa: [{ label: 'OK', min: 0, max: 0.2, cor: 'orange' }]
  };
  var cls = classificarIndicadores(dre, benchmarks);
  QUnit.assert_equal('OK', cls.classificacao_cpv.label, 'CPV faixa');
  QUnit.assert_equal('OK', cls.classificacao_desp_var.label, 'Desp var faixa');
  QUnit.assert_equal('OK', cls.classificacao_desp_fixa.label, 'Desp fixa faixa');
}
