/**
 * sheets-client.ts
 *
 * Módulo centralizado de acesso ao Google Sheets.
 * TODO acesso aos serviços deve passar por este módulo.
 *
 * IMPORTANTE:
 * - Usar operações em lote (batch) sempre que possível
 * - Evitar chamadas célula por célula
 * - Considerar limites de quota do Apps Script
 */

/**
 * Interface para operações batch
 */
export interface BatchOperation {
  sheetName: string;
  range: string;
  values: any[][];
}

/**
 * Opções para getSheetValues
 */
export interface GetSheetOptions {
  range?: string; // ex: 'A1:Z1000' ou null para toda a aba
  skipHeader?: boolean; // se true, pula primeira linha
  limit?: number; // limite de linhas a retornar
}

/**
 * Retorna a planilha ativa ou a planilha por ID
 * TODO: Implementar busca por ID configurável via config-service
 */
function getSpreadsheet(spreadsheetId?: string): GoogleAppsScript.Spreadsheet.Spreadsheet {
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Obtém valores de uma aba do Google Sheets
 *
 * @param sheetName - Nome da aba
 * @param options - Opções de leitura
 * @returns Array 2D com os valores
 *
 * TODO: Implementar cache em memória para dados de referência
 * TODO: Implementar paginação para abas muito grandes
 */
export function getSheetValues(
  sheetName: string,
  options: GetSheetOptions = {}
): any[][] {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Aba "${sheetName}" não encontrada`);
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) {
      return [];
    }

    // Define range
    let range: GoogleAppsScript.Spreadsheet.Range;
    if (options.range) {
      range = sheet.getRange(options.range);
    } else {
      const startRow = options.skipHeader ? 2 : 1;
      const numRows = options.limit
        ? Math.min(options.limit, lastRow - startRow + 1)
        : lastRow - startRow + 1;
      range = sheet.getRange(startRow, 1, numRows, lastCol);
    }

    const values = range.getValues();
    return values;
  } catch (error) {
    console.error(`Erro ao ler aba ${sheetName}:`, error);
    throw error;
  }
}

/**
 * Define valores em uma aba do Google Sheets
 *
 * @param sheetName - Nome da aba
 * @param range - Range no formato A1 (ex: 'A2:D10')
 * @param values - Array 2D com os valores
 *
 * TODO: Implementar validação de tamanho (quota limits)
 * TODO: Implementar retry em caso de falha
 */
export function setSheetValues(
  sheetName: string,
  range: string,
  values: any[][]
): void {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Aba "${sheetName}" não encontrada`);
    }

    const targetRange = sheet.getRange(range);
    targetRange.setValues(values);
  } catch (error) {
    console.error(`Erro ao escrever na aba ${sheetName}:`, error);
    throw error;
  }
}

/**
 * Adiciona linhas ao final de uma aba
 *
 * @param sheetName - Nome da aba
 * @param rows - Array 2D com as linhas a adicionar
 *
 * TODO: Implementar validação de schema antes de inserir
 */
export function appendRows(sheetName: string, rows: any[][]): void {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Aba "${sheetName}" não encontrada`);
    }

    if (rows.length === 0) {
      return;
    }

    const lastRow = sheet.getLastRow();
    const numCols = rows[0].length;
    const startRow = lastRow + 1;

    sheet.getRange(startRow, 1, rows.length, numCols).setValues(rows);
  } catch (error) {
    console.error(`Erro ao adicionar linhas na aba ${sheetName}:`, error);
    throw error;
  }
}

/**
 * Atualiza uma linha específica
 *
 * @param sheetName - Nome da aba
 * @param rowIndex - Índice da linha (1-based)
 * @param rowData - Array com os valores da linha
 */
export function updateRow(
  sheetName: string,
  rowIndex: number,
  rowData: any[]
): void {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Aba "${sheetName}" não encontrada`);
    }

    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } catch (error) {
    console.error(`Erro ao atualizar linha ${rowIndex} na aba ${sheetName}:`, error);
    throw error;
  }
}

/**
 * Deleta linhas de uma aba
 *
 * @param sheetName - Nome da aba
 * @param startRow - Linha inicial (1-based)
 * @param numRows - Quantidade de linhas a deletar
 */
export function deleteRows(
  sheetName: string,
  startRow: number,
  numRows: number
): void {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Aba "${sheetName}" não encontrada`);
    }

    sheet.deleteRows(startRow, numRows);
  } catch (error) {
    console.error(`Erro ao deletar linhas na aba ${sheetName}:`, error);
    throw error;
  }
}

/**
 * Limpa o conteúdo de um range (mantém formatação)
 *
 * @param sheetName - Nome da aba
 * @param range - Range no formato A1
 */
export function clearRange(sheetName: string, range: string): void {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Aba "${sheetName}" não encontrada`);
    }

    sheet.getRange(range).clearContent();
  } catch (error) {
    console.error(`Erro ao limpar range ${range} na aba ${sheetName}:`, error);
    throw error;
  }
}

/**
 * Procura uma linha onde uma coluna contém um valor específico
 *
 * @param sheetName - Nome da aba
 * @param columnIndex - Índice da coluna (0-based)
 * @param value - Valor a procurar
 * @returns Índice da linha (1-based) ou null se não encontrado
 *
 * TODO: Otimizar para grandes volumes (usar TextFinder API)
 */
export function findRowByColumnValue(
  sheetName: string,
  columnIndex: number,
  value: any
): number | null {
  try {
    const values = getSheetValues(sheetName);

    for (let i = 0; i < values.length; i++) {
      if (values[i][columnIndex] === value) {
        return i + 1; // +1 porque getSheetValues pode ter skipHeader
      }
    }

    return null;
  } catch (error) {
    console.error(`Erro ao procurar valor na aba ${sheetName}:`, error);
    throw error;
  }
}

/**
 * Executa operações em lote (batch)
 *
 * @param operations - Array de operações batch
 *
 * TODO: Implementar lógica de batch real usando batchUpdate API
 */
export function executeBatch(operations: BatchOperation[]): void {
  try {
    for (const op of operations) {
      setSheetValues(op.sheetName, op.range, op.values);
    }
  } catch (error) {
    console.error('Erro ao executar operações batch:', error);
    throw error;
  }
}

/**
 * Verifica se uma aba existe
 */
export function sheetExists(sheetName: string): boolean {
  try {
    const ss = getSpreadsheet();
    return ss.getSheetByName(sheetName) !== null;
  } catch (error) {
    return false;
  }
}

/**
 * Cria uma nova aba se não existir
 *
 * @param sheetName - Nome da aba
 * @param headers - Cabeçalhos (primeira linha)
 */
export function createSheetIfNotExists(
  sheetName: string,
  headers?: string[]
): void {
  try {
    if (sheetExists(sheetName)) {
      return;
    }

    const ss = getSpreadsheet();
    const sheet = ss.insertSheet(sheetName);

    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  } catch (error) {
    console.error(`Erro ao criar aba ${sheetName}:`, error);
    throw error;
  }
}
