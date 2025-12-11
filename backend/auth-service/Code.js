/**
 * GET handler: returns a simple JSON hello.
 */
function doGet(e) {
  try {
    var payload = {
      status: 'ok',
      message: 'Hello World from Apps Script'
    };
    return ContentService.createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return createErrorResponse(err);
  }
}

/**
 * POST handler: echoes JSON body with a receivedAt timestamp.
 */
function doPost(e) {
  try {
    var rawBody = e && e.postData && e.postData.contents;
    var data = rawBody ? JSON.parse(rawBody) : {};
    data.receivedAt = new Date().toISOString();

    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return createErrorResponse(err);
  }
}

/**
 * Builds an error JSON response.
 * @param {Error} err
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function createErrorResponse(err) {
  var message = (err && err.message) ? err.message : 'Unexpected error';
  var payload = { error: message };
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
