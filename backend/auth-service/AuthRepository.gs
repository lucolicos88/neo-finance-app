/**
 * Repositório de usuários baseado em planilha tb_usuarios.
 * Colunas esperadas: id, nome, email, senha_hash, papel, filial_padrao, ativo
 */
var USERS_SPREADSHEET_ID = '1e-u2qTehu-iT4P68wP8nOQgcWlKU32VxAHsDzqP7Vnc';
var USERS_SHEET_NAME = 'tb_usuarios';

/**
 * Retorna usuário pelo email (case-insensitive) ou null.
 * @param {string} email
 * @return {Object|null}
 */
function findUserByEmail(email) {
  if (!email) return null;
  var rows = readTable(USERS_SPREADSHEET_ID, USERS_SHEET_NAME);
  var lower = String(email).trim().toLowerCase();
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row.email && String(row.email).toLowerCase() === lower) {
      return {
        id: row.id,
        nome: row.nome,
        email: row.email,
        senhaHash: row.senha_hash,
        papel: row.papel,
        filialPadrao: row.filial_padrao,
        ativo: row.ativo
      };
    }
  }
  return null;
}

/**
 * Hash simples de senha usando SHA-256 + base64 (ambiente interno).
 * @param {string} plain
 * @return {string}
 */
function hashPassword(plain) {
  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    plain,
    Utilities.Charset.UTF_8
  );
  var bytes = digest.map(function (b) { return (b + 256) % 256; });
  return Utilities.base64Encode(bytes);
}

/**
 * Verifica senha comparando hash.
 * @param {string} plainPassword
 * @param {string} storedHash
 * @return {boolean}
 */
function verifyPassword(plainPassword, storedHash) {
  if (!plainPassword || !storedHash) return false;
  var computed = hashPassword(plainPassword);
  return computed === storedHash;
}

/**
 * Cria uma sessão (token) com expiração de 8 horas no CacheService.
 * @param {string} userId
 * @return {string} token
 */
function createSession(userId) {
  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  var payload = JSON.stringify({ userId: userId });
  cache.put(token, payload, 8 * 60 * 60); // 8 horas em segundos
  return token;
}

/**
 * Recupera usuário pelo token salvo no cache. Retorna objeto ou null.
 * @param {string} token
 * @return {Object|null}
 */
function getUserByToken(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var cached = cache.get(token);
  if (!cached) return null;
  var parsed = {};
  try {
    parsed = JSON.parse(cached);
  } catch (err) {
    return null;
  }
  if (!parsed.userId) return null;

  var rows = readTable(USERS_SPREADSHEET_ID, USERS_SHEET_NAME);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (String(row.id) === String(parsed.userId)) {
      return {
        id: row.id,
        nome: row.nome,
        email: row.email,
        papel: row.papel,
        filialPadrao: row.filial_padrao,
        ativo: row.ativo
      };
    }
  }
  return null;
}
