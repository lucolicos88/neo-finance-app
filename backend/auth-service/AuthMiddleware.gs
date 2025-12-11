/**
 * Middleware simples para endpoints autenticados.
 * Exemplo de uso:
 *   return withAuth(function(ctx) {
 *     return { message: 'ok', user: ctx.user };
 *   }, e);
 */
function withAuth(handler, e) {
  try {
    var token = null;
    if (e && e.parameter && e.parameter['X-Auth-Token']) {
      token = e.parameter['X-Auth-Token'];
    } else if (e && e.postData && e.postData.contents) {
      // Caso o token venha no body JSON
      try {
        var body = JSON.parse(e.postData.contents);
        token = body.token || body.authToken || null;
      } catch (parseErr) {
        // ignore
      }
    }

    var user = authLib.getUserByToken(token);
    if (!user) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: 'Não autorizado' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var context = {
      user: user,
      payload: (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : null
    };
    var result = handler(context) || {};
    return ContentService.createTextOutput(
      JSON.stringify(result)
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    var message = (err && err.message) ? err.message : 'Erro interno';
    return ContentService.createTextOutput(
      JSON.stringify({ error: message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Stub de authLib.getUserByToken para integração com CacheService.
 * Aqui delegamos para o repositório existente.
 */
var authLib = {
  getUserByToken: function (token) {
    return getUserByToken(token);
  }
};
