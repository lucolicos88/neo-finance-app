/**
 * ReportsService - consolidação de dashboard e relatório PDF.
 */

function getDashboardGeral(params) {
  // Placeholders: assumimos funções auxiliares em DREService e CashflowService.
  var dreResumo = (typeof obterResumoDre === 'function') ? obterResumoDre(params) : { faturamentoMes: 0, lucroMes: 0, cpvPercentual: 0 };
  var caixaResumo = (typeof obterSaldoCaixa === 'function') ? obterSaldoCaixa(params) : { saldoCaixa: 0 };
  var series = (typeof obterSeriesMensais === 'function') ? obterSeriesMensais(params) : { faturamentoPorMes: [], lucroPorMes: [] };
  var ranking = (typeof obterRankingFiliais === 'function') ? obterRankingFiliais(params) : [];

  return {
    cards: {
      faturamentoMes: dreResumo.faturamentoMes,
      lucroMes: dreResumo.lucroMes,
      saldoCaixa: caixaResumo.saldoCaixa,
      cpvPercentual: dreResumo.cpvPercentual
    },
    series: {
      faturamentoPorMes: series.faturamentoPorMes,
      lucroPorMes: series.lucroPorMes
    },
    rankingFiliais: ranking
  };
}

/**
 * Gera PDF para comitê de sócios.
 * params: { periodoLabel, dreResumo, fluxoResumo, rankingFiliais }
 */
function gerarRelatorioComiteSociosPdf(params) {
  var doc = DocumentApp.create('Relatorio Comitê - ' + (params.periodoLabel || 'Periodo'));
  var body = doc.getBody();

  body.appendParagraph('Relatório Comitê de Sócios').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(params.periodoLabel || '').setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph('Resumo DRE').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  var dre = params.dreResumo || {};
  body.appendParagraph('Receita Líquida: ' + (dre.receita_liquida || 0));
  body.appendParagraph('CPV: ' + (dre.cpv_total || 0));
  body.appendParagraph('EBITDA: ' + (dre.ebitda || 0));
  body.appendParagraph('Lucro Líquido: ' + (dre.lucro_liquido || 0));

  body.appendParagraph('Resumo Fluxo de Caixa').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  var fluxo = params.fluxoResumo || {};
  body.appendParagraph('Entradas: ' + (fluxo.entradas || 0));
  body.appendParagraph('Saídas: ' + (fluxo.saidas || 0));
  body.appendParagraph('Saldo Final: ' + (fluxo.saldoFinal || 0));

  body.appendParagraph('Ranking de Filiais').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  var tableData = [['Filial', 'Faturamento', 'Margem']];
  (params.rankingFiliais || []).forEach(function (r) {
    tableData.push([r.filial || '', r.faturamento || 0, r.margem || 0]);
  });
  body.appendTable(tableData);

  doc.saveAndClose();
  var pdf = doc.getAs('application/pdf');
  var file = DriveApp.createFile(pdf);
  return { url: file.getUrl(), fileId: file.getId() };
}
