/**
 * DREService - cálculos gerenciais.
 */

/**
 * dadosReceitas: [{filial, mes, valor}]
 * dadosDespesas: [{filial, mes, tipo, natureza, valor}] tipo: receita, despesa, imposto, cpv
 * parametros: { filtrarFilial, filtrarMes }
 */
function calcularDreMensal(dadosReceitas, dadosDespesas, parametros) {
  parametros = parametros || {};
  var filialFiltro = parametros.filtrarFilial || null;
  var mesFiltro = parametros.filtrarMes || null;

  var receitasFiltradas = (dadosReceitas || []).filter(function (r) {
    if (filialFiltro && r.filial !== filialFiltro) return false;
    if (mesFiltro && r.mes !== mesFiltro) return false;
    return true;
  });
  var despesasFiltradas = (dadosDespesas || []).filter(function (d) {
    if (filialFiltro && d.filial !== filialFiltro) return false;
    if (mesFiltro && d.mes !== mesFiltro) return false;
    return true;
  });

  var receita_bruta = sum(receitasFiltradas, 'valor');
  var impostos_venda = sumByTipo(despesasFiltradas, 'imposto');
  var receita_liquida = receita_bruta - impostos_venda;

  var cpv_total = sumByTipo(despesasFiltradas, 'cpv');
  var desp_var = sumByNatureza(despesasFiltradas, 'variavel');
  var desp_fixa = sumByNatureza(despesasFiltradas, 'fixa');

  var margem_contribuicao = receita_liquida - cpv_total - desp_var;
  var ebitda = margem_contribuicao - desp_fixa;
  var lucro_liquido = ebitda; // sem depreciação/IR aqui

  return {
    receita_bruta: receita_bruta,
    impostos_venda: impostos_venda,
    receita_liquida: receita_liquida,
    cpv_total: cpv_total,
    margem_contribuicao: margem_contribuicao,
    despesas_variaveis: desp_var,
    despesas_fixas: desp_fixa,
    ebitda: ebitda,
    lucro_liquido: lucro_liquido
  };
}

function sum(arr, field) {
  return (arr || []).reduce(function (acc, item) {
    return acc + (Number(item[field]) || 0);
  }, 0);
}

function sumByTipo(arr, tipo) {
  return (arr || []).reduce(function (acc, item) {
    if (String(item.tipo).toLowerCase() === tipo) {
      acc += Number(item.valor) || 0;
    }
    return acc;
  }, 0);
}

function sumByNatureza(arr, natureza) {
  return (arr || []).reduce(function (acc, item) {
    if (String(item.natureza).toLowerCase() === natureza) {
      acc += Number(item.valor) || 0;
    }
    return acc;
  }, 0);
}

/**
 * Classifica indicadores comparando com benchmarks.
 * benchmarks esperado: { cpv: [{label,min,max,cor}], desp_var:[...], desp_fixa:[...] }
 */
function classificarIndicadores(dre, benchmarks) {
  var receita = dre.receita_liquida || 0;
  var cpvPerc = receita ? dre.cpv_total / receita : 0;
  var despVarPerc = receita ? dre.despesas_variaveis / receita : 0;
  var despFixaPerc = receita ? dre.despesas_fixas / receita : 0;

  return {
    classificacao_cpv: encontraFaixa(cpvPerc, benchmarks && benchmarks.cpv),
    classificacao_desp_var: encontraFaixa(despVarPerc, benchmarks && benchmarks.desp_var),
    classificacao_desp_fixa: encontraFaixa(despFixaPerc, benchmarks && benchmarks.desp_fixa)
  };
}

function encontraFaixa(valorPerc, faixas) {
  if (!faixas || !faixas.length) return null;
  for (var i = 0; i < faixas.length; i++) {
    var f = faixas[i];
    if (valorPerc >= f.min && valorPerc <= f.max) {
      return { label: f.label, cor: f.cor, valor: valorPerc };
    }
  }
  return { label: 'Sem faixa', cor: '#999', valor: valorPerc };
}
