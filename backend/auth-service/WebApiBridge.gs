/**
 * Funções expostas para uso via google.script.run (elimina CORS).
 */
function webListAp(filtros) {
  return listAp(filtros || {});
}

function webListAr(filtros) {
  return listAr(filtros || {});
}
