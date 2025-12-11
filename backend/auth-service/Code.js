/**
 * Roteia requisições para WebApp (HTML) ou APIs (path).
 */
function doGet(e) {
  var path = e && e.parameter && e.parameter.path;
  if (!path) {
    // Serve HTML do próprio Apps Script para evitar CORS.
    return HtmlService.createHtmlOutputFromFile('WebApp')
      .setTitle('Neo Finance App');
  }
  return routeRequest(e, 'GET');
}

function doPost(e) {
  return routeRequest(e, 'POST');
}

function routeRequest(e, method) {
  var path = (e && e.parameter && e.parameter.path) ? e.parameter.path : '';
  if (path.indexOf('/ap') === 0 && typeof handleRequest === 'function') {
    return handleRequest(e);
  }
  if (path.indexOf('/ar') === 0 && typeof handleArRequest === 'function') {
    return handleArRequest(e);
  }
  if (path.indexOf('/cashflow') === 0 && typeof handleCashflowRequest === 'function') {
    return handleCashflowRequest(e);
  }
  if (path.indexOf('/dre') === 0 && typeof handleDreRequest === 'function') {
    return handleDreRequest(e);
  }
  if (path.indexOf('/reports') === 0 && typeof handleReportsRequest === 'function') {
    return handleReportsRequest(e);
  }
  var payload = { status: 'ok', message: 'Hello World from Apps Script' };
  var out = ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
  if (out.setHeader) {
    out.setHeader('Access-Control-Allow-Origin', '*');
    out.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    out.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
  return out;
}

function parseBody(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (err) { return {}; }
  }
  return {};
}

function jsonResponseSafe(obj) {
  var out = ContentService.createTextOutput(JSON.stringify(obj || {}))
    .setMimeType(ContentService.MimeType.JSON);
  if (out.setHeader) {
    out.setHeader('Access-Control-Allow-Origin', '*');
    out.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    out.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
  return out;
}
