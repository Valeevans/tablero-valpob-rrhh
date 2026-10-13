/* ════════════════════════════════════════════════════════════════════
   MÓDULO STORE — Tablero RRHH Grupo VALPOB
   Guarda perfiles, catálogos (competencias y requisitos), nómina de
   empleados y asesores externos en una hoja "Store" del mismo Sheets.

   CÓMO INSTALARLO (5 minutos):
   1) En el Google Sheets del tablero: Extensiones → Apps Script.
   2) Pegá TODO este archivo al final del código que ya existe (no borres nada).
   3) En la función doGet(e), agregá como PRIMERA línea:
          if (e && e.parameter && e.parameter.store) return getStore_();
   4) En la función doPost(e), agregá como PRIMERA línea:
          var _st = storeDesdePost_(e); if (_st) return _st;
   5) Guardar → Implementar → Gestionar implementaciones → ✏️ (lápiz)
      → Versión: "Nueva versión" → Implementar.
      (Así la URL del tablero no cambia.)
   La hoja "Store" se crea sola la primera vez.
   ════════════════════════════════════════════════════════════════════ */

var STORE_SHEET = 'Store';
var STORE_CHUNK = 45000; // una celda admite 50.000 caracteres: se parte en varias columnas

function storeSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(STORE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STORE_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['clave', 'actualizado (ms)', 'valor JSON (continúa en columnas siguientes)']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function storeJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getStore_() {
  var sh = storeSheet_();
  var data = sh.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    var partes = [];
    for (var j = 2; j < r.length; j++) if (r[j] !== '' && r[j] !== null) partes.push(String(r[j]));
    out[String(r[0])] = { ts: Number(r[1]) || 0, value: partes.join('') };
  }
  return storeJson_({ store: out });
}

function storeDesdePost_(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body && body.action === 'saveStore') return saveStore_(body);
  } catch (err) {}
  return null;
}

function saveStore_(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = storeSheet_();
    var key = String(body.key || '');
    if (!key) return storeJson_({ ok: false, error: 'Falta la clave' });
    var ts = Number(body.ts) || Date.now();
    var val = String(body.value || '');

    var data = sh.getDataRange().getValues();
    var fila = -1;
    for (var i = 1; i < data.length; i++) if (String(data[i][0]) === key) { fila = i + 1; break; }
    if (fila > 0 && (Number(data[fila - 1][1]) || 0) > ts) {
      return storeJson_({ ok: true, skipped: 'ya hay una versión más nueva' });
    }
    if (fila < 0) fila = sh.getLastRow() + 1;

    // partir el JSON en trozos; el apóstrofo fuerza texto (evita que Sheets lo convierta en número o fórmula)
    var partes = [];
    for (var p = 0; p < val.length; p += STORE_CHUNK) partes.push("'" + val.substring(p, p + STORE_CHUNK));
    if (!partes.length) partes.push("'");
    var filaVals = [key, ts].concat(partes);

    if (sh.getMaxColumns() < filaVals.length) sh.insertColumnsAfter(sh.getMaxColumns(), filaVals.length - sh.getMaxColumns());
    var ultCol = Math.max(sh.getLastColumn(), filaVals.length);
    sh.getRange(fila, 1, 1, ultCol).clearContent();
    sh.getRange(fila, 1, 1, filaVals.length).setValues([filaVals]);
    return storeJson_({ ok: true, ts: ts });
  } finally {
    lock.releaseLock();
  }
}
