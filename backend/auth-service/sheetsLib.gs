var sheetCache = (function () {
  var cache = {};
  return {
    getTable: function (spreadsheetId, sheetName) {
      var key = spreadsheetId + '::' + sheetName;
      if (cache[key]) return cache[key];
      var sheet = getSheetByName(spreadsheetId, sheetName);
      var values = sheet.getDataRange().getValues();
      cache[key] = values;
      return values;
    },
    clearCache: function () {
      cache = {};
    }
  };
})();

/**
 * Returns a Sheet by name with error handling.
 * @param {string} spreadsheetId
 * @param {string} sheetName
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheetByName(spreadsheetId, sheetName) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" not found in spreadsheet ' + spreadsheetId);
  }
  return sheet;
}

/**
 * Reads an entire sheet as an array of objects keyed by header row.
 * @param {string} spreadsheetId
 * @param {string} sheetName
 * @return {Object[]}
 */
function readTable(spreadsheetId, sheetName) {
  var values = sheetCache.getTable(spreadsheetId, sheetName);
  if (!values.length) {
    return [];
  }

  var header = values[0];
  if (!header.length || header.every(function (h) { return h === '' || h === null; })) {
    throw new Error('Header row is empty for sheet "' + sheetName + '"');
  }

  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    // Skip entirely empty rows.
    var allEmpty = row.every(function (cell) { return cell === '' || cell === null; });
    if (allEmpty) {
      continue;
    }
    var obj = {};
    for (var c = 0; c < header.length; c++) {
      var key = String(header[c]).trim();
      if (key) {
        obj[key] = row[c];
      }
    }
    rows.push(obj);
  }
  return rows;
}

/**
 * Appends a row object to the end of the sheet respecting header order.
 * Unknown keys are ignored; missing keys become empty cells.
 * @param {string} spreadsheetId
 * @param {string} sheetName
 * @param {Object} rowObject
 */
function writeRow(spreadsheetId, sheetName, rowObject) {
  var sheet = getSheetByName(spreadsheetId, sheetName);
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!header.length || header.every(function (h) { return h === '' || h === null; })) {
    throw new Error('Header row is empty for sheet "' + sheetName + '"');
  }

  var row = [];
  for (var i = 0; i < header.length; i++) {
    var key = String(header[i]).trim();
    row.push(key ? rowObject[key] : '');
  }

  sheet.appendRow(row);
}
