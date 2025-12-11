/**
 * Controlador de Auth para Web App.
 */

/**
 * Processa login com payload { email, senha }.
 * Retorna { token, usuario: { id, nome, papel, filialPadrao } }
 */
function handleLogin(payload) {
  if (!payload || !payload.email || !payload.senha) {
    return { error: 'Credenciais inválidas' };
  }

  var user = findUserByEmail(payload.email);
  if (!user || String(user.ativo).toLowerCase() === 'false') {
    return { error: 'Usuário não encontrado ou inativo' };
  }

  var ok = verifyPassword(payload.senha, user.senhaHash);
  if (!ok) {
    return { error: 'Credenciais inválidas' };
  }

  var token = createSession(user.id);
  return {
    token: token,
    usuario: {
      id: user.id,
      nome: user.nome,
      papel: user.papel,
      filialPadrao: user.filialPadrao
    }
  };
}

/**
 * Retorna usuário logado a partir do token.
 * @param {string} token
 * @return {Object}
 */
function handleMe(token) {
  if (!token) {
    return { error: 'Token ausente' };
  }
  var user = getUserByToken(token);
  if (!user) {
    return { error: 'Token inválido' };
  }
  return { usuario: user };
}
