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
  SHEET_REF_CCUSTO,
  SHEET_REF_PLANO_CONTAS,
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
  contas: Array<{ codigo: string; nome: string; tipo?: string }>;
  centrosCusto: Array<{ codigo: string; nome: string }>;
} {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Filiais (da planilha)
    const sheetFiliais = ss.getSheetByName(SHEET_REF_FILIAIS);
    const filiais = sheetFiliais ? sheetFiliais.getDataRange().getValues().slice(1) : [];

    // Canais (da planilha)
    const sheetCanais = ss.getSheetByName(SHEET_REF_CANAIS);
    const canais = sheetCanais ? sheetCanais.getDataRange().getValues().slice(1) : [];

    // Centros de Custo (da planilha, com fallback para hardcoded)
    const sheetCCusto = ss.getSheetByName(SHEET_REF_CCUSTO);
    let centrosCusto: any[];
    if (sheetCCusto && sheetCCusto.getLastRow() > 1) {
      const ccData = sheetCCusto.getDataRange().getValues().slice(1);
      centrosCusto = ccData.filter((cc: any) => cc[0]).map((cc: any) => ({
        codigo: String(cc[0]),
        nome: String(cc[1])
      }));
    } else {
      // Fallback hardcoded
      centrosCusto = [
        { codigo: 'ADM', nome: 'Administrativo' },
        { codigo: 'COM', nome: 'Comercial' },
        { codigo: 'OPS', nome: 'Operacional' },
        { codigo: 'FIN', nome: 'Financeiro' },
        { codigo: 'TI', nome: 'Tecnologia' },
      ];
    }

    // Contas Contábeis (da planilha, com fallback para hardcoded)
    const sheetContas = ss.getSheetByName(SHEET_REF_PLANO_CONTAS);
    let contas: any[];
    if (sheetContas && sheetContas.getLastRow() > 1) {
      const contasData = sheetContas.getDataRange().getValues().slice(1);
      contas = contasData.filter((c: any) => c[0]).map((c: any) => ({
        codigo: String(c[0]),
        nome: String(c[1]),
        tipo: String(c[2] || '')
      }));
    } else {
      // Fallback hardcoded
      contas = [
        { codigo: '1.01.001', nome: 'Receita de Serviços', tipo: 'RECEITA' },
        { codigo: '1.01.002', nome: 'Receita de Produtos', tipo: 'RECEITA' },
        { codigo: '2.01.001', nome: 'Fornecedores', tipo: 'DESPESA' },
        { codigo: '2.01.002', nome: 'Salários', tipo: 'DESPESA' },
        { codigo: '2.01.003', nome: 'Impostos', tipo: 'DESPESA' },
        { codigo: '2.01.004', nome: 'Aluguel', tipo: 'DESPESA' },
        { codigo: '2.01.005', nome: 'Utilities', tipo: 'DESPESA' },
      ];
    }

    return {
      filiais: filiais.filter((f: any) => f[0]).map((f: any) => ({ codigo: String(f[0]), nome: String(f[1]) })),
      canais: canais.filter((c: any) => c[0]).map((c: any) => ({ codigo: String(c[0]), nome: String(c[1]) })),
      contas: contas,
      centrosCusto: centrosCusto,
    };
  } catch (error: any) {
    Logger.log(`Erro ao carregar dados de referência: ${error.message}`);
    // Retornar dados vazios em caso de erro
    return {
      filiais: [],
      canais: [],
      contas: [],
      centrosCusto: [],
    };
  }
}

// ============================================================================
// CRUD CONFIGURAÇÕES
// ============================================================================

// Centros de Custo
export function salvarCentroCusto(centroCusto: { codigo: string; nome: string }, editIndex: number): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_REF_CCUSTO);

    // Criar aba se não existir
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_REF_CCUSTO);
      sheet.getRange('A1:B1').setValues([['Código', 'Nome']]);
      sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#00a8e8').setFontColor('#ffffff');
    }

    // Verificar código duplicado
    if (editIndex < 0) {
      const data = sheet.getDataRange().getValues().slice(1);
      const codigoExiste = data.some((row: any) => String(row[0]).toUpperCase() === centroCusto.codigo.toUpperCase());
      if (codigoExiste) {
        return { success: false, message: 'Código já existe' };
      }
    }

    if (editIndex >= 0) {
      // Editar (linha = editIndex + 2)
      sheet.getRange(editIndex + 2, 1, 1, 2).setValues([[centroCusto.codigo, centroCusto.nome]]);
    } else {
      // Novo
      sheet.appendRow([centroCusto.codigo, centroCusto.nome]);
    }

    return { success: true, message: 'Centro de custo salvo com sucesso' };
  } catch (error: any) {
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function excluirCentroCusto(index: number): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_CCUSTO);

    if (!sheet) {
      throw new Error('Aba de centros de custo não encontrada');
    }

    // Deletar linha (index + 2)
    sheet.deleteRow(index + 2);

    return { success: true, message: 'Centro de custo excluído' };
  } catch (error: any) {
    return { success: false, message: `Erro: ${error.message}` };
  }
}

// Plano de Contas
export function salvarContaContabil(conta: { codigo: string; nome: string; tipo: string }, editIndex: number): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_REF_PLANO_CONTAS);

    // Criar aba se não existir
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_REF_PLANO_CONTAS);
      sheet.getRange('A1:C1').setValues([['Código', 'Nome', 'Tipo']]);
      sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#00a8e8').setFontColor('#ffffff');
    }

    // Verificar código duplicado
    if (editIndex < 0) {
      const data = sheet.getDataRange().getValues().slice(1);
      const codigoExiste = data.some((row: any) => String(row[0]) === conta.codigo);
      if (codigoExiste) {
        return { success: false, message: 'Código já existe' };
      }
    }

    if (editIndex >= 0) {
      // Editar (linha = editIndex + 2)
      sheet.getRange(editIndex + 2, 1, 1, 3).setValues([[conta.codigo, conta.nome, conta.tipo]]);
    } else {
      // Novo
      sheet.appendRow([conta.codigo, conta.nome, conta.tipo]);
    }

    return { success: true, message: 'Conta contábil salva com sucesso' };
  } catch (error: any) {
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function excluirConta(index: number): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_PLANO_CONTAS);

    if (!sheet) {
      throw new Error('Aba de plano de contas não encontrada');
    }

    // Deletar linha (index + 2)
    sheet.deleteRow(index + 2);

    return { success: true, message: 'Conta excluída' };
  } catch (error: any) {
    return { success: false, message: `Erro: ${error.message}` };
  }
}

// Canais
export function salvarCanal(canal: { codigo: string; nome: string }, editIndex: number): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_CANAIS);

    if (!sheet) {
      throw new Error('Aba de canais não encontrada');
    }

    // Verificar código duplicado
    if (editIndex < 0) {
      const data = sheet.getDataRange().getValues().slice(1);
      const codigoExiste = data.some((row: any) => String(row[0]).toUpperCase() === canal.codigo.toUpperCase());
      if (codigoExiste) {
        return { success: false, message: 'Código já existe' };
      }
    }

    if (editIndex >= 0) {
      // Editar (linha = editIndex + 2)
      sheet.getRange(editIndex + 2, 1, 1, 2).setValues([[canal.codigo, canal.nome]]);
    } else {
      // Novo
      sheet.appendRow([canal.codigo, canal.nome]);
    }

    return { success: true, message: 'Canal salvo com sucesso' };
  } catch (error: any) {
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function excluirCanal(index: number): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_CANAIS);

    if (!sheet) {
      throw new Error('Aba de canais não encontrada');
    }

    // Deletar linha (index + 2 pois +1 header + 1 base-0)
    sheet.deleteRow(index + 2);

    return { success: true, message: 'Canal excluído' };
  } catch (error: any) {
    return { success: false, message: `Erro: ${error.message}` };
  }
}

// Filiais
export function salvarFilial(filial: { codigo: string; nome: string }, editIndex: number): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_FILIAIS);

    if (!sheet) {
      throw new Error('Aba de filiais não encontrada');
    }

    // Verificar código duplicado
    if (editIndex < 0) {
      const data = sheet.getDataRange().getValues().slice(1);
      const codigoExiste = data.some((row: any) => String(row[0]).toUpperCase() === filial.codigo.toUpperCase());
      if (codigoExiste) {
        return { success: false, message: 'Código já existe' };
      }
    }

    if (editIndex >= 0) {
      sheet.getRange(editIndex + 2, 1, 1, 2).setValues([[filial.codigo, filial.nome]]);
    } else {
      sheet.appendRow([filial.codigo, filial.nome]);
    }

    return { success: true, message: 'Filial salva com sucesso' };
  } catch (error: any) {
    return { success: false, message: `Erro: ${error.message}` };
  }
}

export function excluirFilial(index: number): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REF_FILIAIS);

    if (!sheet) {
      throw new Error('Aba de filiais não encontrada');
    }

    sheet.deleteRow(index + 2);

    return { success: true, message: 'Filial excluída' };
  } catch (error: any) {
    return { success: false, message: `Erro: ${error.message}` };
  }
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
// SALVAR LANÇAMENTO
// ============================================================================

export function salvarLancamento(lancamento: any): { success: boolean; message: string; id?: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TB_LANCAMENTOS);
    if (!sheet) throw new Error('Aba de lançamentos não encontrada');

    // Converter objeto lancamento para array de valores (seguindo a ordem das colunas)
    const row = [
      lancamento.id,                    // ID
      lancamento.dataCompetencia,       // Data Competência
      lancamento.dataVencimento,        // Data Vencimento
      lancamento.dataPagamento || '',   // Data Pagamento
      lancamento.tipo,                  // Tipo (RECEITA/DESPESA)
      lancamento.filial,                // Filial
      lancamento.centroCusto,           // Centro de Custo
      lancamento.contaGerencial,        // Conta Gerencial
      lancamento.contaContabil,         // Conta Contábil
      lancamento.grupoReceita || '',    // Grupo Receita
      lancamento.canal || '',           // Canal
      lancamento.descricao,             // Descrição
      lancamento.valorBruto,            // Valor Bruto
      lancamento.desconto || 0,         // Desconto
      lancamento.juros || 0,            // Juros
      lancamento.multa || 0,            // Multa
      lancamento.valorLiquido,          // Valor Líquido
      lancamento.status,                // Status
      lancamento.idExtratoBanco || '',  // ID Extrato Banco
      lancamento.origem || 'MANUAL',    // Origem
      lancamento.observacoes || '',     // Observações
    ];

    // Adicionar linha à planilha
    sheet.appendRow(row);

    return {
      success: true,
      message: lancamento.tipo === 'RECEITA' ? 'Conta a receber salva com sucesso' : 'Conta a pagar salva com sucesso',
      id: lancamento.id,
    };
  } catch (error: any) {
    return {
      success: false,
      message: `Erro ao salvar lançamento: ${error.message}`,
    };
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

// ============================================================================
// DRE (Demonstração do Resultado do Exercício)
// ============================================================================

export function getDREMensal(mes: number, ano: number, filial?: string): any {
  try {
    const lancamentos = getLancamentosFromSheet();

    // Filtrar por período e filial
    const lancamentosMes = lancamentos.filter(l => {
      const data = new Date(l.dataCompetencia);
      const mesLanc = data.getMonth() + 1; // JavaScript months are 0-indexed
      const anoLanc = data.getFullYear();

      const matchPeriodo = mesLanc === mes && anoLanc === ano;
      const matchFilial = !filial || l.filial === filial;

      return matchPeriodo && matchFilial;
    });

    // Separar receitas e despesas
    const receitas = lancamentosMes.filter(l => l.tipo === 'RECEITA');
    const despesas = lancamentosMes.filter(l => l.tipo === 'DESPESA');

    // Calcular valores
    const receitaBruta = sumValues(receitas.map(r => ({ valorLiquido: r.valorBruto })));
    const deducoes = sumValues(receitas.map(r => ({ valorLiquido: r.desconto })));
    const receitaLiquida = receitaBruta - deducoes;

    // Separar custos e despesas operacionais (baseado na conta contábil)
    const custos = despesas.filter(d =>
      d.contaContabil && (d.contaContabil.startsWith('2.01') || d.contaContabil.startsWith('2.02'))
    );
    const despesasOp = despesas.filter(d =>
      d.contaContabil && d.contaContabil.startsWith('3.')
    );
    const despesasFinanceiras = despesas.filter(d =>
      d.contaContabil && d.contaContabil.startsWith('4.')
    );

    const totalCustos = sumValues(custos);
    const margemBruta = receitaLiquida - totalCustos;
    const percMargemBruta = receitaLiquida > 0 ? (margemBruta / receitaLiquida) * 100 : 0;

    // Despesas operacionais por categoria (baseado em centro de custo)
    const despPessoal = despesasOp.filter(d => d.centroCusto === 'ADM' || d.centroCusto === 'OPS');
    const despMarketing = despesasOp.filter(d => d.centroCusto === 'MKT' || d.centroCusto === 'COM');
    const despAdministrativas = despesasOp.filter(d => d.centroCusto === 'FIN' || d.centroCusto === 'TI');

    const totalDespOp = sumValues(despesasOp);
    const ebitda = margemBruta - totalDespOp;
    const percEbitda = receitaLiquida > 0 ? (ebitda / receitaLiquida) * 100 : 0;

    // Resultado Financeiro
    const receitasFinanceiras = receitas.filter(r => r.contaContabil && r.contaContabil.startsWith('4.02'));
    const totalResultadoFinanceiro = sumValues(receitasFinanceiras) - sumValues(despesasFinanceiras);

    const lucroLiquido = ebitda + totalResultadoFinanceiro;
    const percLucroLiquido = receitaLiquida > 0 ? (lucroLiquido / receitaLiquida) * 100 : 0;

    return {
      periodo: {
        mes,
        ano,
        mesNome: getMesNome(mes),
        filial: filial || 'Consolidado'
      },
      valores: {
        receitaBruta,
        deducoes,
        receitaLiquida,
        custos: totalCustos,
        margemBruta,
        despesasOperacionais: {
          pessoal: sumValues(despPessoal),
          marketing: sumValues(despMarketing),
          administrativas: sumValues(despAdministrativas),
          total: totalDespOp
        },
        ebitda,
        resultadoFinanceiro: totalResultadoFinanceiro,
        lucroLiquido
      },
      percentuais: {
        margemBruta: percMargemBruta,
        ebitda: percEbitda,
        lucroLiquido: percLucroLiquido
      },
      classificacao: {
        margemBruta: classificarIndicador(percMargemBruta, 'margem_bruta'),
        ebitda: classificarIndicador(percEbitda, 'ebitda'),
        lucroLiquido: classificarIndicador(percLucroLiquido, 'lucro_liquido')
      },
      transacoes: {
        totalReceitas: receitas.length,
        totalDespesas: despesas.length
      }
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular DRE: ${error.message}`);
    throw new Error(`Erro ao calcular DRE: ${error.message}`);
  }
}

export function getDREComparativo(meses: Array<{ mes: number; ano: number }>, filial?: string): any {
  try {
    const dres = meses.map(periodo => getDREMensal(periodo.mes, periodo.ano, filial));

    return {
      periodos: dres.map(d => d.periodo),
      comparativo: dres,
      evolucao: {
        receitaLiquida: calcularEvolucao(dres.map(d => d.valores.receitaLiquida)),
        margemBruta: calcularEvolucao(dres.map(d => d.valores.margemBruta)),
        ebitda: calcularEvolucao(dres.map(d => d.valores.ebitda)),
        lucroLiquido: calcularEvolucao(dres.map(d => d.valores.lucroLiquido))
      }
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular DRE comparativo: ${error.message}`);
    throw new Error(`Erro ao calcular DRE comparativo: ${error.message}`);
  }
}

export function getDREPorFilial(mes: number, ano: number): any {
  try {
    const lancamentos = getLancamentosFromSheet();

    // Obter lista única de filiais
    const filiais = [...new Set(lancamentos.map(l => l.filial))].filter(f => f);

    // Calcular DRE para cada filial
    const dresPorFilial = filiais.map(filial => ({
      filial,
      dre: getDREMensal(mes, ano, filial)
    }));

    // DRE consolidado
    const dreConsolidado = getDREMensal(mes, ano);

    return {
      consolidado: dreConsolidado,
      porFilial: dresPorFilial,
      filiais
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular DRE por filial: ${error.message}`);
    throw new Error(`Erro ao calcular DRE por filial: ${error.message}`);
  }
}

// Helper functions
function getMesNome(mes: number): string {
  const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
  return meses[mes - 1] || '';
}

function classificarIndicador(percentual: number, tipo: string): string {
  // Benchmarks baseados no PDF do comitê
  const benchmarks: any = {
    margem_bruta: { sensacional: 60, excelente: 50, bom: 40, ruim: 30 },
    ebitda: { sensacional: 25, excelente: 20, bom: 15, ruim: 10 },
    lucro_liquido: { sensacional: 20, excelente: 15, bom: 10, ruim: 5 }
  };

  const bench = benchmarks[tipo] || benchmarks.ebitda;

  if (percentual >= bench.sensacional) return 'Sensacional';
  if (percentual >= bench.excelente) return 'Excelente';
  if (percentual >= bench.bom) return 'Bom';
  if (percentual >= bench.ruim) return 'Ruim';
  return 'Péssimo';
}

function calcularEvolucao(valores: number[]): any {
  if (valores.length < 2) return { percentual: 0, tendencia: 'estavel' };

  const primeiro = valores[0];
  const ultimo = valores[valores.length - 1];

  if (primeiro === 0) return { percentual: 0, tendencia: 'estavel' };

  const percentual = ((ultimo - primeiro) / Math.abs(primeiro)) * 100;
  const tendencia = percentual > 5 ? 'crescimento' : percentual < -5 ? 'queda' : 'estavel';

  return { percentual, tendencia };
}

// ============================================================================
// FLUXO DE CAIXA (DFC)
// ============================================================================

export function getFluxoCaixaMensal(mes: number, ano: number, filial?: string): any {
  try {
    const lancamentos = getLancamentosFromSheet();

    // Filtrar por período e filial
    const lancamentosMes = lancamentos.filter(l => {
      const data = new Date(l.dataCompetencia);
      const mesLanc = data.getMonth() + 1;
      const anoLanc = data.getFullYear();

      const matchPeriodo = mesLanc === mes && anoLanc === ano;
      const matchFilial = !filial || l.filial === filial;

      return matchPeriodo && matchFilial;
    });

    // Separar por tipo e status
    const entradas = lancamentosMes.filter(l =>
      l.tipo === 'RECEITA' && (l.status === 'PAGO' || l.status === 'RECEBIDO')
    );
    const saidas = lancamentosMes.filter(l =>
      l.tipo === 'DESPESA' && l.status === 'PAGO'
    );

    // Calcular totais
    const totalEntradas = sumValues(entradas);
    const totalSaidas = sumValues(saidas);

    // Saldo inicial (simplificado - pode ser melhorado para buscar do mês anterior)
    const saldoInicial = 0; // TODO: Buscar saldo final do mês anterior
    const saldoFinal = saldoInicial + totalEntradas - totalSaidas;
    const variacao = saldoInicial !== 0 ? ((saldoFinal - saldoInicial) / Math.abs(saldoInicial)) * 100 : 0;

    // Agrupar entradas por categoria (conta contábil)
    const entradasPorCategoria = agruparPorCategoria(entradas);
    const saidasPorCategoria = agruparPorCategoria(saidas);

    return {
      periodo: {
        mes,
        ano,
        mesNome: getMesNome(mes),
        filial: filial || 'Consolidado'
      },
      valores: {
        saldoInicial,
        totalEntradas,
        totalSaidas,
        saldoFinal,
        variacao
      },
      transacoes: {
        qtdEntradas: entradas.length,
        qtdSaidas: saidas.length
      },
      entradas: entradasPorCategoria,
      saidas: saidasPorCategoria
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular Fluxo de Caixa: ${error.message}`);
    throw new Error(`Erro ao calcular Fluxo de Caixa: ${error.message}`);
  }
}

export function getFluxoCaixaProjecao(meses: number, filial?: string): any {
  try {
    const hoje = new Date();
    const periodos = [];

    for (let i = 0; i < meses; i++) {
      const data = new Date(hoje.getFullYear(), hoje.getMonth() + i, 1);
      periodos.push({
        mes: data.getMonth() + 1,
        ano: data.getFullYear()
      });
    }

    const fluxos = periodos.map(p => getFluxoCaixaMensal(p.mes, p.ano, filial));

    return {
      periodos: fluxos.map(f => f.periodo),
      fluxos,
      evolucao: {
        entradas: calcularEvolucao(fluxos.map(f => f.valores.totalEntradas)),
        saidas: calcularEvolucao(fluxos.map(f => f.valores.totalSaidas)),
        saldo: calcularEvolucao(fluxos.map(f => f.valores.saldoFinal))
      }
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular projeção de Fluxo de Caixa: ${error.message}`);
    throw new Error(`Erro ao calcular projeção de Fluxo de Caixa: ${error.message}`);
  }
}

// Helper function para agrupar por categoria
function agruparPorCategoria(lancamentos: any[]): any[] {
  const categorias: any = {};

  lancamentos.forEach(l => {
    const categoria = l.categoria || 'Outros';
    if (!categorias[categoria]) {
      categorias[categoria] = {
        categoria,
        valor: 0,
        quantidade: 0
      };
    }
    categorias[categoria].valor += parseFloat(String(l.valorLiquido || l.valor || 0));
    categorias[categoria].quantidade++;
  });

  return Object.values(categorias).sort((a: any, b: any) => b.valor - a.valor);
}

// ============================================================================
// KPIs FINANCEIROS
// ============================================================================

export function getKPIsMensal(mes: number, ano: number, filial?: string): any {
  try {
    const lancamentos = getLancamentosFromSheet();

    // Filtrar lançamentos do mês atual
    const lancamentosMes = lancamentos.filter(l => {
      const data = new Date(l.dataCompetencia);
      const mesLanc = data.getMonth() + 1;
      const anoLanc = data.getFullYear();
      const matchPeriodo = mesLanc === mes && anoLanc === ano;
      const matchFilial = !filial || l.filial === filial;
      return matchPeriodo && matchFilial;
    });

    // Filtrar lançamentos do mês anterior
    const dataAnterior = new Date(ano, mes - 2, 1); // mes - 2 porque JavaScript months são 0-indexed
    const mesAnterior = dataAnterior.getMonth() + 1;
    const anoAnterior = dataAnterior.getFullYear();

    const lancamentosMesAnterior = lancamentos.filter(l => {
      const data = new Date(l.dataCompetencia);
      const mesLanc = data.getMonth() + 1;
      const anoLanc = data.getFullYear();
      const matchPeriodo = mesLanc === mesAnterior && anoLanc === anoAnterior;
      const matchFilial = !filial || l.filial === filial;
      return matchPeriodo && matchFilial;
    });

    // Calcular DRE do mês atual e anterior
    const dreAtual = getDREMensal(mes, ano, filial);
    const dreAnterior = getDREMensal(mesAnterior, anoAnterior, filial);
    const fcAtual = getFluxoCaixaMensal(mes, ano, filial);

    // Separar receitas e despesas
    const receitas = lancamentosMes.filter(l => l.tipo === 'RECEITA');
    const despesas = lancamentosMes.filter(l => l.tipo === 'DESPESA');
    const receitasAnterior = lancamentosMesAnterior.filter(l => l.tipo === 'RECEITA');

    // KPIs de Rentabilidade
    const margemBruta = dreAtual.percentuais.margemBruta;
    const margemEbitda = dreAtual.percentuais.ebitda;
    const margemLiquida = dreAtual.percentuais.lucroLiquido;
    const roi = dreAtual.valores.receitaLiquida > 0
      ? (dreAtual.valores.lucroLiquido / dreAtual.valores.receitaLiquida) * 100
      : 0;

    // KPIs de Liquidez
    const contasReceber = lancamentosMes.filter(l => l.tipo === 'RECEITA' && l.status === 'PENDENTE');
    const contasPagar = lancamentosMes.filter(l => l.tipo === 'DESPESA' && l.status === 'PENDENTE');
    const ativoCirculante = sumValues(contasReceber) + fcAtual.valores.saldoFinal;
    const passivoCirculante = sumValues(contasPagar);
    const liquidezCorrente = passivoCirculante > 0 ? ativoCirculante / passivoCirculante : 0;

    const saldoCaixa = fcAtual.valores.saldoFinal;
    const burnRate = Math.abs(dreAtual.valores.lucroLiquido < 0 ? dreAtual.valores.lucroLiquido : 0);
    const runway = burnRate > 0 ? saldoCaixa / burnRate : 999;

    // KPIs de Crescimento
    const receitaAtual = dreAtual.valores.receitaLiquida;
    const receitaAnteriorVal = dreAnterior.valores.receitaLiquida;
    const crescimentoReceita = receitaAnteriorVal > 0
      ? ((receitaAtual - receitaAnteriorVal) / receitaAnteriorVal) * 100
      : 0;

    const ticketMedio = receitas.length > 0 ? receitaAtual / receitas.length : 0;

    const receitasVencidas = lancamentosMes.filter(l => {
      if (l.tipo !== 'RECEITA' || l.status !== 'PENDENTE') return false;
      const vencimento = new Date(l.dataVencimento);
      return vencimento < new Date();
    });
    const taxaInadimplencia = receitas.length > 0
      ? (receitasVencidas.length / receitas.length) * 100
      : 0;

    // Prazo médio de recebimento
    const receitasRecebidas = lancamentosMes.filter(l =>
      l.tipo === 'RECEITA' && (l.status === 'PAGO' || l.status === 'RECEBIDO')
    );
    let prazoMedioRecebimento = 0;
    if (receitasRecebidas.length > 0) {
      const prazos = receitasRecebidas.map(r => {
        const venc = new Date(r.dataVencimento);
        const pag = new Date(r.dataPagamento || r.dataCompetencia);
        return Math.floor((pag.getTime() - venc.getTime()) / (1000 * 60 * 60 * 24));
      });
      prazoMedioRecebimento = prazos.reduce((a, b) => a + b, 0) / prazos.length;
    }

    // KPIs Operacionais
    const despesasMarketing = despesas.filter(d => d.centroCusto === 'MKT' || d.centroCusto === 'COM');
    const cac = receitas.length > 0 ? sumValues(despesasMarketing) / receitas.length : 0;

    const despOperacionaisPerc = dreAtual.valores.receitaLiquida > 0
      ? (dreAtual.valores.despesasOperacionais.total / dreAtual.valores.receitaLiquida) * 100
      : 0;

    const breakEven = dreAtual.valores.margemBruta > 0
      ? dreAtual.valores.despesasOperacionais.total / (dreAtual.valores.margemBruta / dreAtual.valores.receitaLiquida)
      : 0;

    // Prazo médio de pagamento
    const despesasPagas = lancamentosMes.filter(l =>
      l.tipo === 'DESPESA' && l.status === 'PAGO'
    );
    let prazoMedioPagamento = 0;
    if (despesasPagas.length > 0) {
      const prazos = despesasPagas.map(d => {
        const venc = new Date(d.dataVencimento);
        const pag = new Date(d.dataPagamento || d.dataCompetencia);
        return Math.floor((pag.getTime() - venc.getTime()) / (1000 * 60 * 60 * 24));
      });
      prazoMedioPagamento = prazos.reduce((a, b) => a + b, 0) / prazos.length;
    }

    return {
      periodo: {
        mes,
        ano,
        mesNome: getMesNome(mes),
        filial: filial || 'Consolidado'
      },
      rentabilidade: {
        margemBruta: {
          valor: margemBruta,
          classificacao: classificarIndicador(margemBruta, 'margem_bruta')
        },
        margemEbitda: {
          valor: margemEbitda,
          classificacao: classificarIndicador(margemEbitda, 'ebitda')
        },
        margemLiquida: {
          valor: margemLiquida,
          classificacao: classificarIndicador(margemLiquida, 'lucro_liquido')
        },
        roi: {
          valor: roi,
          descricao: roi > 0 ? 'Positivo' : 'Negativo'
        }
      },
      liquidez: {
        liquidezCorrente: {
          valor: liquidezCorrente,
          classificacao: classificarLiquidez(liquidezCorrente)
        },
        saldoCaixa: {
          valor: saldoCaixa,
          descricao: saldoCaixa > 0 ? 'Saudável' : 'Atenção'
        },
        burnRate: {
          valor: burnRate,
          descricao: `${burnRate > 0 ? 'Queimando' : 'Gerando'} caixa`
        },
        runway: {
          valor: runway,
          descricao: runway < 6 ? 'Crítico' : runway < 12 ? 'Atenção' : 'Saudável'
        }
      },
      crescimento: {
        crescimentoReceita: {
          valor: crescimentoReceita,
          mesAnterior: receitaAnteriorVal
        },
        ticketMedio: {
          valor: ticketMedio,
          qtdTransacoes: receitas.length
        },
        inadimplencia: {
          valor: taxaInadimplencia,
          classificacao: taxaInadimplencia < 5 ? 'Excelente' : taxaInadimplencia < 10 ? 'Bom' : 'Ruim'
        },
        prazoMedioRecebimento: {
          valor: prazoMedioRecebimento,
          descricao: `${prazoMedioRecebimento.toFixed(0)} dias`
        }
      },
      operacional: {
        cac: {
          valor: cac,
          descricao: `Custo por cliente`
        },
        despOperacionaisReceita: {
          valor: despOperacionaisPerc,
          classificacao: despOperacionaisPerc < 30 ? 'Excelente' : despOperacionaisPerc < 50 ? 'Bom' : 'Ruim'
        },
        breakEven: {
          valor: breakEven,
          descricao: `Ponto de equilíbrio`
        },
        prazoMedioPagamento: {
          valor: prazoMedioPagamento,
          descricao: `${prazoMedioPagamento.toFixed(0)} dias`
        }
      },
      comparativo: {
        mesAnterior: {
          receita: receitaAnteriorVal,
          margemBruta: dreAnterior.percentuais.margemBruta,
          margemLiquida: dreAnterior.percentuais.lucroLiquido
        }
      }
    };
  } catch (error: any) {
    Logger.log(`Erro ao calcular KPIs: ${error.message}`);
    throw new Error(`Erro ao calcular KPIs: ${error.message}`);
  }
}

// Helper function para classificar liquidez
function classificarLiquidez(valor: number): string {
  if (valor >= 2) return 'Excelente';
  if (valor >= 1.5) return 'Bom';
  if (valor >= 1) return 'Aceitável';
  return 'Ruim';
}

// ============================================================================
// USUÁRIOS E PERMISSÕES
// ============================================================================

const SHEET_USUARIOS = 'TB_Usuarios';

interface Usuario {
  id?: string;
  email: string;
  nome: string;
  perfil: 'ADMIN' | 'GESTOR' | 'OPERACIONAL' | 'VISUALIZADOR';
  status: 'ATIVO' | 'INATIVO';
  ultimoAcesso?: string;
  permissoes?: {
    criarLancamentos: boolean;
    editarLancamentos: boolean;
    excluirLancamentos: boolean;
    aprovarPagamentos: boolean;
    visualizarRelatorios: boolean;
    gerenciarConfig: boolean;
  };
}

export function getUsuarios(): Usuario[] {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_USUARIOS);

    // Se a planilha não existe, criar
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_USUARIOS);
      sheet.getRange('A1:H1').setValues([[
        'ID', 'Email', 'Nome', 'Perfil', 'Status', 'Último Acesso', 'Permissões', 'Data Criação'
      ]]);
      sheet.getRange('A1:H1').setFontWeight('bold');
      sheet.setFrozenRows(1);

      // Adicionar usuário administrador padrão
      const userEmail = Session.getActiveUser().getEmail();
      sheet.appendRow([
        Utilities.getUuid(),
        userEmail,
        'Administrador',
        'ADMIN',
        'ATIVO',
        new Date().toISOString(),
        JSON.stringify({
          criarLancamentos: true,
          editarLancamentos: true,
          excluirLancamentos: true,
          aprovarPagamentos: true,
          visualizarRelatorios: true,
          gerenciarConfig: true
        }),
        new Date().toISOString()
      ]);

      return [{
        id: Utilities.getUuid(),
        email: userEmail,
        nome: 'Administrador',
        perfil: 'ADMIN',
        status: 'ATIVO',
        ultimoAcesso: new Date().toISOString(),
        permissoes: {
          criarLancamentos: true,
          editarLancamentos: true,
          excluirLancamentos: true,
          aprovarPagamentos: true,
          visualizarRelatorios: true,
          gerenciarConfig: true
        }
      }];
    }

    const data = sheet.getDataRange().getValues();
    const usuarios: Usuario[] = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // Skip empty rows

      usuarios.push({
        id: String(row[0]),
        email: String(row[1]),
        nome: String(row[2]),
        perfil: String(row[3]) as any,
        status: String(row[4]) as any,
        ultimoAcesso: row[5] ? String(row[5]) : undefined,
        permissoes: row[6] ? JSON.parse(String(row[6])) : getPermissoesPadrao(String(row[3]))
      });
    }

    return usuarios;
  } catch (error: any) {
    Logger.log(`Erro ao buscar usuários: ${error.message}`);
    return [];
  }
}

export function salvarUsuario(usuario: Usuario): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_USUARIOS);

    if (!sheet) {
      // Criar planilha se não existe
      getUsuarios();
      sheet = ss.getSheetByName(SHEET_USUARIOS);
    }

    const data = sheet!.getDataRange().getValues();
    let rowIndex = -1;

    // Procurar usuário existente
    if (usuario.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === usuario.id) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    const permissoes = usuario.permissoes || getPermissoesPadrao(usuario.perfil);
    const rowData = [
      usuario.id || Utilities.getUuid(),
      usuario.email,
      usuario.nome,
      usuario.perfil,
      usuario.status,
      usuario.ultimoAcesso || '',
      JSON.stringify(permissoes),
      rowIndex === -1 ? new Date().toISOString() : data[rowIndex - 1][7]
    ];

    if (rowIndex > 0) {
      // Atualizar existente
      sheet!.getRange(rowIndex, 1, 1, 8).setValues([rowData]);
      return { success: true, message: 'Usuário atualizado com sucesso!' };
    } else {
      // Criar novo
      sheet!.appendRow(rowData);
      return { success: true, message: 'Usuário criado com sucesso!' };
    }
  } catch (error: any) {
    Logger.log(`Erro ao salvar usuário: ${error.message}`);
    return { success: false, message: `Erro ao salvar usuário: ${error.message}` };
  }
}

export function excluirUsuario(id: string): { success: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USUARIOS);

    if (!sheet) {
      return { success: false, message: 'Planilha de usuários não encontrada' };
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id || data[i][1] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Usuário excluído com sucesso!' };
      }
    }

    return { success: false, message: 'Usuário não encontrado' };
  } catch (error: any) {
    Logger.log(`Erro ao excluir usuário: ${error.message}`);
    return { success: false, message: `Erro ao excluir usuário: ${error.message}` };
  }
}

function getPermissoesPadrao(perfil: string): any {
  const permissoes: any = {
    'ADMIN': {
      criarLancamentos: true,
      editarLancamentos: true,
      excluirLancamentos: true,
      aprovarPagamentos: true,
      visualizarRelatorios: true,
      gerenciarConfig: true
    },
    'GESTOR': {
      criarLancamentos: true,
      editarLancamentos: true,
      excluirLancamentos: false,
      aprovarPagamentos: true,
      visualizarRelatorios: true,
      gerenciarConfig: false
    },
    'OPERACIONAL': {
      criarLancamentos: true,
      editarLancamentos: true,
      excluirLancamentos: false,
      aprovarPagamentos: false,
      visualizarRelatorios: true,
      gerenciarConfig: false
    },
    'VISUALIZADOR': {
      criarLancamentos: false,
      editarLancamentos: false,
      excluirLancamentos: false,
      aprovarPagamentos: false,
      visualizarRelatorios: true,
      gerenciarConfig: false
    }
  };

  return permissoes[perfil] || permissoes['VISUALIZADOR'];
}
