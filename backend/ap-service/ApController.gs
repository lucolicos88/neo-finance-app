/**
 * Controlador de Contas a Pagar.
 * Roteia chamadas em handleRequest(e) usando path:
 * - /ap/list
 * - /ap/create
 * - /ap/{id}/pagar
 */
function handleRequest(e) {
  var path = (e && e.parameter && e.parameter.path) ? e.parameter.path : '/ap/list';
  var body = {};
  if (e && e.postData && e.postData.contents) {
    try {
      body = JSON.parse(e.postData.contents);
    } catch (err) {
      return jsonResponse({ error: 'Body inválido' }, 400);
    }
  }

  var router = {
    '/ap/list': function (ctx) { return listAp(ctx.payload || {}); },
    '/ap/create': function (ctx) { return createAp(ctx.payload || {}); }
  };

  if (path && path.indexOf('/ap/') === 0 && path.indexOf('/pagar') > -1) {
    var parts = path.split('/');
    var id = parts[2];
    return withAuth(function (ctx) {
      var p = ctx.payload || {};
      var ok = pagarAp(id, p.data_pagamento, p.valor_pago);
      return { success: ok };
    }, e);
  }

  if (router[path]) {
    return withAuth(function (ctx) {
      return router[path](ctx);
    }, e);
  }

  return jsonResponse({ error: 'Rota não encontrada' }, 404);
}

/**
 * Middleware simples reutilizando authLib.getUserByToken (token via header ou body).
 */
function withAuth(handler, e) {
  try {
    var token = null;
    if (e && e.parameter && e.parameter['X-Auth-Token']) {
      token = e.parameter['X-Auth-Token'];
    } else if (e && e.postData && e.postData.contents) {
      try {
        var b = JSON.parse(e.postData.contents);
        token = b.token || b.authToken || null;
      } catch (errParse) {
        // ignore
      }
    }
    var user = authLib.getUserByToken(token);
    if (!user) {
      return jsonResponse({ error: 'Não autorizado' }, 401);
    }
    var ctx = { user: user, payload: (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : null };
    var result = handler(ctx) || {};
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message || 'Erro interno' }, 500);
  }
}

var authLib = {
  getUserByToken: function (token) {
    if (typeof getUserByToken === 'function') {
      return getUserByToken(token);
    }
    return null;
  }
};

function jsonResponse(obj, code) {
  var output = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  if (code) {
    return output.setResponseCode ? output.setResponseCode(code) : output;
  }
  return output;
}
