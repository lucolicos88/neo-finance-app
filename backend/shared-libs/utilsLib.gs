/**
 * Parses a date from "dd/MM/yyyy" string (Brazilian format).
 * Returns a Date or throws on invalid input.
 * @param {string} str
 * @return {Date}
 */
function parseDateFromString(str) {
  if (!str || typeof str !== 'string') {
    throw new Error('Invalid date string');
  }
  var parts = str.split('/');
  if (parts.length !== 3) {
    throw new Error('Date must be in dd/MM/yyyy format');
  }
  var day = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1; // zero-based
  var year = parseInt(parts[2], 10);
  var date = new Date(year, month, day);
  if (isNaN(date.getTime()) || date.getDate() !== day || date.getMonth() !== month || date.getFullYear() !== year) {
    throw new Error('Invalid date value');
  }
  return date;
}

/**
 * Formats a Date as "dd/MM/yyyy".
 * @param {Date} date
 * @return {string}
 */
function formatDateToString(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    throw new Error('Invalid Date object');
  }
  var day = ('0' + date.getDate()).slice(-2);
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var year = date.getFullYear();
  return day + '/' + month + '/' + year;
}

/**
 * Parses Brazilian money string like "1.234,56" into number 1234.56.
 * @param {string} str
 * @return {number}
 */
function parseMoney(str) {
  if (str === null || str === undefined) {
    throw new Error('Invalid money string');
  }
  var normalized = String(str).trim().replace(/\./g, '').replace(',', '.');
  var value = parseFloat(normalized);
  if (isNaN(value)) {
    throw new Error('Invalid money value');
  }
  return value;
}

/**
 * Formats a number into Brazilian money string "1.234,56".
 * @param {number} value
 * @return {string}
 */
function formatMoney(value) {
  if (typeof value !== 'number' || isNaN(value)) {
    throw new Error('Invalid money number');
  }
  var fixed = value.toFixed(2); // "1234.56"
  var parts = fixed.split('.');
  var integerPart = parts[0];
  var decimalPart = parts[1];
  var withSeparators = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return withSeparators + ',' + decimalPart;
}
