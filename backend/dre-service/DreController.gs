/**
 * Controlador de DRE.
 * Rota: /dre/mensal
 */
function handleDreRequest(e) {
  var path = (e && e.parameter && e.parameter.path) ? e.parameter.path : '/dre/mensal';
  var payload = parseBody(e);
  var router = {
    '/dre/mensal': function (ctx) {
      var p = ctx.payload || {};
      var dre = calcularDreMensal(p.dadosReceitas || [], p.dadosDespesas || [], p.parametros || {});
      var cls = classificarIndicadores(dre, p.benchmarks || {});
      return { dre: dre, classificacao: cls };
    }
  };
  if (router[path]) {
    return withAuth(function (ctx) { return router[path](ctx); }, e);
  }
  return jsonResponse({ error: 'Rota não encontrada' }, 404);
}

function parseBody(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (err) { return {}; }
  }
  return {};
}

function withAuth(handler, e) {
  try {
    var token = null;
    if (e && e.parameter && e.parameter['X-Auth-Token']) token = e.parameter['X-Auth-Token'];
    else if (e && e.postData && e.postData.contents) {
      try { var b = JSON.parse(e.postData.contents); token = b.token || b.authToken || null; } catch (err2) {}
    }
    var user = authLib.getUserByToken(token);
    if (!user) return jsonResponse({ error: 'Não autorizado' }, 401);
    var ctx = { user: user, payload: parseBody(e) };
    var result = handler(ctx) || {};
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message || 'Erro interno' }, 500);
  }
}

var authLib = {
  getUserByToken: function (token) {
    if (typeof getUserByToken === 'function') return getUserByToken(token);
    return null;
  }
};

function jsonResponse(obj, code) {
  var output = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  if (code && output.setResponseCode) return output.setResponseCode(code);
  return output;
}
