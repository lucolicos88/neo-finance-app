/**
 * webapp-service.ts
 *
 * Serviço backend para a Web App
 * Fornece dados para o frontend via google.script.run
 */

import { getSheetValues } from '../shared/sheets-client';
import {
  SHEET_TB_LANCAMENTOS,
  SHEET_TB_EXTRATOS,
  SHEET_REF_FILIAIS,
  SHEET_REF_CANAIS,
} from '../config/sheet-mapping';

// ============================================================================
// VIEW RENDERING
// ============================================================================

/**
 * Retorna o HTML de uma view específica
 */
export function getViewHtml(viewName: string): string {
  try {
    return HtmlService.createHtmlOutputFromFile(`frontend/views/${viewName}-view`).getContent();
  } catch (error) {
    return `<div class="empty-state">
      <div class="empty-state-icon">⚠️</div>
      <div class="empty-state-message">Erro ao carregar view</div>
      <div class="empty-state-hint">${error}</div>
    </div>`;
  }
}

// ============================================================================
// REFERENCE DATA
// ============================================================================

export function getReferenceData(): {
  filiais: Array<{ codigo: string; nome: string }>;
  canais: Array<{ codigo: string; nome: string }>;
} {
  const filiais = getSheetValues(SHEET_REF_FILIAIS).slice(1); // Skip header
  const canais = getSheetValues(SHEET_REF_CANAIS).slice(1); // Skip header

  return {
    filiais: filiais.map((f: any) => ({ codigo: f[0], nome: f[1] })),
    canais: canais.map((c: any) => ({ codigo: c[0], nome: c[1] })),
  };
}

// ============================================================================
// DASHBOARD
// ============================================================================

export function getDashboardData() {
  const lancamentos = getLancamentosFromSheet();
  const hoje = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

  // Contas a pagar vencidas
  const pagarVencidas = lancamentos.filter(l =>
    l.tipo === 'DESPESA' &&
    l.status === 'VENCIDA' &&
    new Date(l.dataVencimento) < hoje
  );

  // Contas a pagar próximos 7 dias
  const proximos7Dias = new Date();
  proximos7Dias.setDate(proximos7Dias.getDate() + 7);
  const pagarProximas = lancamentos.filter(l =>
    l.tipo === 'DESPESA' &&
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos7Dias &&
    new Date(l.dataVencimento) >= hoje
  );

  // Contas a receber hoje
  const receberHoje = lancamentos.filter(l =>
    l.tipo === 'RECEITA' &&
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento).toDateString() === hoje.toDateString()
  );

  // Extratos pendentes
  const extratos = getExtratosFromSheet();
  const extratosPendentes = extratos.filter(e => e.statusConciliacao === 'PENDENTE');

  // Últimos lançamentos
  const recentTransactions = lancamentos
    .slice(0, 10)
    .map(l => ({
      id: l.id,
      data: l.dataCompetencia,
      descricao: l.descricao,
      tipo: l.tipo,
      valor: l.valorLiquido,
      status: l.status,
    }));

  // Alertas
  const alerts: Array<{ type: string; title: string; message: string }> = [];

  if (pagarVencidas.length > 0) {
    alerts.push({
      type: 'danger',
      title: 'Contas Vencidas',
      message: `Você tem ${pagarVencidas.length} contas a pagar vencidas no valor de ${formatCurrency(sumValues(pagarVencidas))}`,
    });
  }

  if (pagarProximas.length > 0) {
    alerts.push({
      type: 'warning',
      title: 'Vencimentos Próximos',
      message: `${pagarProximas.length} contas a pagar vencem nos próximos 7 dias`,
    });
  }

  if (extratosPendentes.length > 5) {
    alerts.push({
      type: 'info',
      title: 'Conciliação Pendente',
      message: `${extratosPendentes.length} extratos bancários aguardando conciliação`,
    });
  }

  return {
    pagarVencidas: {
      quantidade: pagarVencidas.length,
      valor: sumValues(pagarVencidas),
    },
    pagarProximas: {
      quantidade: pagarProximas.length,
      valor: sumValues(pagarProximas),
    },
    receberHoje: {
      quantidade: receberHoje.length,
      valor: sumValues(receberHoje),
    },
    conciliacaoPendentes: {
      quantidade: extratosPendentes.length,
      valor: extratosPendentes.reduce((sum, e) => sum + parseFloat(String(e.valor || 0)), 0),
    },
    recentTransactions,
    alerts,
  };
}

// ============================================================================
// CONTAS A PAGAR
// ============================================================================

export function getContasPagar() {
  const lancamentos = getLancamentosFromSheet();
  const hoje = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

  const contasPagar = lancamentos.filter(l => l.tipo === 'DESPESA');

  const vencidas = contasPagar.filter(l =>
    l.status === 'VENCIDA' || (l.status === 'PENDENTE' && new Date(l.dataVencimento) < hoje)
  );

  const proximos7Dias = new Date();
  proximos7Dias.setDate(proximos7Dias.getDate() + 7);
  const vencer7 = contasPagar.filter(l =>
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos7Dias &&
    new Date(l.dataVencimento) >= hoje
  );

  const proximos30Dias = new Date();
  proximos30Dias.setDate(proximos30Dias.getDate() + 30);
  const vencer30 = contasPagar.filter(l =>
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos30Dias &&
    new Date(l.dataVencimento) > proximos7Dias
  );

  const pagas = contasPagar.filter(l =>
    l.status === 'PAGA' &&
    new Date(l.dataPagamento || l.dataCompetencia) >= inicioMes
  );

  return {
    stats: {
      vencidas: { quantidade: vencidas.length, valor: sumValues(vencidas) },
      vencer7: { quantidade: vencer7.length, valor: sumValues(vencer7) },
      vencer30: { quantidade: vencer30.length, valor: sumValues(vencer30) },
      pagas: { quantidade: pagas.length, valor: sumValues(pagas) },
    },
    contas: contasPagar.map(l => ({
      id: l.id,
      vencimento: l.dataVencimento,
      fornecedor: l.descricao.split('-')[0].trim(),
      descricao: l.descricao,
      valor: l.valorLiquido,
      status: l.status,
      filial: l.filial,
    })),
  };
}

export function pagarConta(id: string): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lançamentos não encontrada');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const statusCol = headers.indexOf('Status');
    const dataPagCol = headers.indexOf('Data Pagamento');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === id) {
        sheet.getRange(i + 1, statusCol + 1).setValue('PAGA');
        sheet.getRange(i + 1, dataPagCol + 1).setValue(new Date());
        return { success: true, message: 'Conta paga com sucesso' };
      }
    }

    throw new Error('Conta não encontrada');
  } catch (error: any) {
    return { success: false, message: error.message };
  }
}

export function pagarContasEmLote(ids: string[]): { success: boolean; message: string } {
  try {
    let count = 0;
    for (const id of ids) {
      const result = pagarConta(id);
      if (result.success) count++;
    }
    return { success: true, message: `${count} contas pagas com sucesso` };
  } catch (error: any) {
    return { success: false, message: error.message };
  }
}

// ============================================================================
// CONTAS A RECEBER
// ============================================================================

export function getContasReceber() {
  const lancamentos = getLancamentosFromSheet();
  const hoje = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

  const contasReceber = lancamentos.filter(l => l.tipo === 'RECEITA');

  const vencidas = contasReceber.filter(l =>
    l.status === 'VENCIDA' || (l.status === 'PENDENTE' && new Date(l.dataVencimento) < hoje)
  );

  const proximos7Dias = new Date();
  proximos7Dias.setDate(proximos7Dias.getDate() + 7);
  const receber7 = contasReceber.filter(l =>
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos7Dias &&
    new Date(l.dataVencimento) >= hoje
  );

  const proximos30Dias = new Date();
  proximos30Dias.setDate(proximos30Dias.getDate() + 30);
  const receber30 = contasReceber.filter(l =>
    l.status === 'PENDENTE' &&
    new Date(l.dataVencimento) <= proximos30Dias &&
    new Date(l.dataVencimento) > proximos7Dias
  );

  const recebidas = contasReceber.filter(l =>
    l.status === 'RECEBIDA' &&
    new Date(l.dataPagamento || l.dataCompetencia) >= inicioMes
  );

  return {
    stats: {
      vencidas: { quantidade: vencidas.length, valor: sumValues(vencidas) },
      receber7: { quantidade: receber7.length, valor: sumValues(receber7) },
      receber30: { quantidade: receber30.length, valor: sumValues(receber30) },
      recebidas: { quantidade: recebidas.length, valor: sumValues(recebidas) },
    },
    contas: contasReceber.map(l => ({
      id: l.id,
      vencimento: l.dataVencimento,
      cliente: l.descricao.split('-')[0].trim(),
      descricao: l.descricao,
      valor: l.valorLiquido,
      status: l.status,
      canal: l.canal || 'N/A',
    })),
  };
}

export function receberConta(id: string): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lançamentos não encontrada');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const statusCol = headers.indexOf('Status');
    const dataPagCol = headers.indexOf('Data Pagamento');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === id) {
        sheet.getRange(i + 1, statusCol + 1).setValue('RECEBIDA');
        sheet.getRange(i + 1, dataPagCol + 1).setValue(new Date());
        return { success: true, message: 'Conta recebida com sucesso' };
      }
    }

    throw new Error('Conta não encontrada');
  } catch (error: any) {
    return { success: false, message: error.message };
  }
}

// ============================================================================
// CONCILIAÇÃO
// ============================================================================

export function getConciliacaoData() {
  const extratos = getExtratosFromSheet();
  const lancamentos = getLancamentosFromSheet();
  const hoje = new Date();
  const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

  const extratosPendentes = extratos.filter(e => e.statusConciliacao === 'PENDENTE');
  const lancamentosPendentes = lancamentos.filter(l => !l.idExtratoBanco);

  const conciliadosHoje = extratos.filter(e =>
    e.statusConciliacao === 'CONCILIADO' &&
    new Date(e.importadoEm).toDateString() === hoje.toDateString()
  );

  const totalExtratos = extratos.length;
  const totalConciliados = extratos.filter(e => e.statusConciliacao === 'CONCILIADO').length;
  const taxaConciliacao = totalExtratos > 0 ? Math.round((totalConciliados / totalExtratos) * 100) : 0;

  // Histórico (últimas 50 conciliações)
  const historico = extratos
    .filter(e => e.statusConciliacao === 'CONCILIADO' && e.idLancamento)
    .slice(0, 50)
    .map(e => ({
      dataConciliacao: e.importadoEm,
      extratoId: e.id,
      lancamentoId: e.idLancamento,
      descricao: e.descricao,
      valor: e.valor,
      banco: e.banco,
      usuario: 'Sistema',
    }));

  return {
    stats: {
      extratosPendentes: extratosPendentes.length,
      extratosValor: extratosPendentes.reduce((sum, e) => sum + parseFloat(String(e.valor || 0)), 0),
      lancamentosPendentes: lancamentosPendentes.length,
      lancamentosValor: sumValues(lancamentosPendentes),
      conciliadosHoje: conciliadosHoje.length,
      conciliadosHojeValor: conciliadosHoje.reduce((sum, e) => sum + parseFloat(String(e.valor || 0)), 0),
      taxaConciliacao,
    },
    extratos: extratosPendentes.map(e => ({
      id: e.id,
      data: e.data,
      descricao: e.descricao,
      valor: e.valor,
      banco: e.banco,
    })),
    lancamentos: lancamentosPendentes.slice(0, 50).map(l => ({
      id: l.id,
      data: l.dataCompetencia,
      descricao: l.descricao,
      valor: l.valorLiquido,
      tipo: l.tipo,
    })),
    historico,
  };
}

export function conciliarItens(extratoId: string, lancamentoId: string): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Atualizar extrato
    const sheetExtratos = ss.getSheetByName(SHEET_TB_EXTRATOS);
    if (sheetExtratos) {
      const dataExtratos = sheetExtratos.getDataRange().getValues();
      const headersExtratos = dataExtratos[0];
      const idColE = headersExtratos.indexOf('ID');
      const statusColE = headersExtratos.indexOf('Status Conciliação');
      const lancColE = headersExtratos.indexOf('ID Lançamento');

      for (let i = 1; i < dataExtratos.length; i++) {
        if (dataExtratos[i][idColE] === extratoId) {
          sheetExtratos.getRange(i + 1, statusColE + 1).setValue('CONCILIADO');
          sheetExtratos.getRange(i + 1, lancColE + 1).setValue(lancamentoId);
          break;
        }
      }
    }

    // Atualizar lançamento
    const sheetLanc = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (sheetLanc) {
      const dataLanc = sheetLanc.getDataRange().getValues();
      const headersLanc = dataLanc[0];
      const idColL = headersLanc.indexOf('ID');
      const extratoColL = headersLanc.indexOf('ID Extrato Banco');

      for (let i = 1; i < dataLanc.length; i++) {
        if (dataLanc[i][idColL] === lancamentoId) {
          sheetLanc.getRange(i + 1, extratoColL + 1).setValue(extratoId);
          break;
        }
      }
    }

    return { success: true, message: 'Conciliação realizada com sucesso' };
  } catch (error: any) {
    return { success: false, message: error.message };
  }
}

export function conciliarAutomatico(): { success: boolean; conciliados: number; message: string } {
  try {
    const extratos = getExtratosFromSheet().filter(e => e.statusConciliacao === 'PENDENTE');
    const lancamentos = getLancamentosFromSheet().filter(l => !l.idExtratoBanco);

    let conciliados = 0;

    for (const extrato of extratos) {
      // Tentar encontrar lançamento com valor e data próximos
      const match = lancamentos.find(l =>
        Math.abs(parseFloat(String(l.valorLiquido)) - parseFloat(String(extrato.valor))) < 0.01 &&
        Math.abs(new Date(l.dataCompetencia).getTime() - new Date(extrato.data).getTime()) < 7 * 24 * 60 * 60 * 1000
      );

      if (match) {
        conciliarItens(extrato.id, match.id);
        conciliados++;
      }
    }

    return { success: true, conciliados, message: `${conciliados} itens conciliados` };
  } catch (error: any) {
    return { success: false, conciliados: 0, message: error.message };
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function getLancamentosFromSheet(): any[] {
  const data = getSheetValues(SHEET_TB_LANCAMENTOS).slice(1); // Skip header
  if (data.length === 0) return [];

  return data.map((row: any) => ({
    id: row[0],
    dataCompetencia: row[1],
    dataVencimento: row[2],
    dataPagamento: row[3],
    tipo: row[4],
    filial: row[5],
    centroCusto: row[6],
    contaGerencial: row[7],
    contaContabil: row[8],
    grupoReceita: row[9],
    canal: row[10],
    descricao: row[11],
    valorBruto: parseFloat(String(row[12] || 0)),
    desconto: parseFloat(String(row[13] || 0)),
    juros: parseFloat(String(row[14] || 0)),
    multa: parseFloat(String(row[15] || 0)),
    valorLiquido: parseFloat(String(row[16] || 0)),
    status: row[17],
    idExtratoBanco: row[18],
    origem: row[19],
    observacoes: row[20],
  }));
}

function getExtratosFromSheet(): any[] {
  const data = getSheetValues(SHEET_TB_EXTRATOS).slice(1); // Skip header
  if (data.length === 0) return [];

  return data.map((row: any) => ({
    id: row[0],
    data: row[1],
    descricao: row[2],
    valor: parseFloat(String(row[3] || 0)),
    tipo: row[4],
    banco: row[5],
    conta: row[6],
    statusConciliacao: row[7],
    idLancamento: row[8],
    observacoes: row[9],
    importadoEm: row[10],
  }));
}

function sumValues(items: any[]): number {
  return items.reduce((sum, item) => sum + parseFloat(String(item.valorLiquido || item.valor || 0)), 0);
}

function formatCurrency(value: number): string {
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value);
}
