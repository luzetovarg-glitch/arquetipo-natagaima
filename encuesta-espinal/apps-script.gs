/**
 * Encuesta Mesas JAC El Espinal → Hoja de Google
 * Pegue TODO este código en: Hoja "Votos Encuesta Mesas JAC El Espinal" > Extensiones > Apps Script
 * Luego: Implementar > Nueva implementación > Aplicación web
 *        Ejecutar como: Yo   ·   Quién tiene acceso: Cualquier usuario
 */
const HOJA = 'Votos';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA);
    const id = limpiar(d.id);
    // No guardar dos veces el mismo voto
    if (id && sh.getLastRow() > 1) {
      const repetido = sh.getRange(2, 13, sh.getLastRow() - 1, 1)
        .createTextFinder(id).matchEntireCell(true).findNext();
      if (repetido) return salida({ ok: true, repetido: true });
    }
    sh.appendRow([
      new Date(),
      limpiar(d.barrio), limpiar(d.nombre), limpiar(d.rol),
      limpiar(d.tema1), limpiar(d.sub1),
      limpiar(d.tema2), limpiar(d.sub2),
      limpiar(d.tema3), limpiar(d.sub3),
      limpiar(d.comentario), limpiar(d.codigo), id
    ]);
    return salida({ ok: true });
  } catch (err) {
    return salida({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Para probar: al abrir la dirección /exec en el navegador debe decir "conectada"
function doGet() {
  return salida({ ok: true, mensaje: 'La encuesta está conectada a la hoja.' });
}

function limpiar(v) {
  v = String(v == null ? '' : v).slice(0, 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // evita fórmulas escritas por terceros
}

function salida(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
