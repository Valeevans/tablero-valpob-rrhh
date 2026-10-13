function doGet(e) {
  // ── Módulo Store (perfiles, catálogos, nómina) ──
  if (e && e.parameter && e.parameter.store) return getStore_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ['Formacion','Perfiles','Desempenio','KPIs','Alertas'];
  const out = {};
  sheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    const headers = data[0];
    const rows = data.slice(1);
    if (name === 'Alertas') {
      // Alertas: una fila por alerta
      out['Alertas'] = rows
        .filter(r => r[0])
        .map(r => ({ titulo: r[0], detalle: r[1], tipo: r[2] }));
    } else if (name === 'Perfiles') {
      out['Perfiles'] = rows
        .filter(r => r[0])
        .map(r => ({
          puesto: r[0], area: r[1], tipo: r[2],
          formacion: r[3], exp: r[4],
          comps: r[5] ? JSON.parse(r[5]) : {},
          responsabilidades: r[6] || '',
          tareas: r[7] || ''
        }));
    } else {
      out[name] = rows.map(r => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = r[i]);
        return obj;
      });
    }
  });
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // ── Módulo Store (perfiles, catálogos, nómina) ──
  var _st = storeDesdePost_(e); if (_st) return _st;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  if (action === 'updateFormacion') {
    const sh = ss.getSheetByName('Formacion');
    const vals = sh.getDataRange().getValues();
    const headers = vals[0];
    const mesCol = headers.indexOf(getMesName(data.mes));
    for (let i = 1; i < vals.length; i++) {
      if (parseInt(vals[i][0]) === parseInt(data.id)) {
        if (mesCol > -1) sh.getRange(i+1, mesCol+1).setValue(data.estado);
        sh.getRange(i+1, headers.indexOf('Obs')+1).setValue(
          `Fecha: ${data.fecha} | Asistentes: ${data.asistentes} (${data.cantidad}) | ${data.obs}`
        );
        break;
      }
    }
  }

  if (action === 'editFormacion') {
    const sh = ss.getSheetByName('Formacion');
    const vals = sh.getDataRange().getValues();
    const headers = vals[0];
    const meses = JSON.parse(data.meses || '{}');
    for (let i = 1; i < vals.length; i++) {
      if (parseInt(vals[i][0]) === parseInt(data.id)) {
        if (headers.indexOf('Tema') > -1) sh.getRange(i+1, headers.indexOf('Tema')+1).setValue(data.tema);
        if (headers.indexOf('Categoria') > -1) sh.getRange(i+1, headers.indexOf('Categoria')+1).setValue(data.bloque);
        if (headers.indexOf('Responsable') > -1) sh.getRange(i+1, headers.indexOf('Responsable')+1).setValue(data.resp);
        Object.entries(meses).forEach(([mes, val]) => {
          const col = headers.indexOf(getMesName(parseInt(mes)));
          if (col > -1) sh.getRange(i+1, col+1).setValue(val);
        });
        break;
      }
    }
  }

  if (action === 'updateDesempenio') {
    const sh = ss.getSheetByName('Desempenio');
    const vals = sh.getDataRange().getValues();
    const headers = vals[0];
    let found = false;
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][0] === data.puesto) {
        sh.getRange(i+1, headers.indexOf('Puntaje')+1).setValue(data.puntaje);
        sh.getRange(i+1, headers.indexOf('Estado')+1).setValue(data.estado);
        sh.getRange(i+1, headers.indexOf('Observaciones')+1).setValue(data.obs);
        sh.getRange(i+1, headers.indexOf('Competencias')+1).setValue(data.competencias);
        if (data.fecha && headers.indexOf('Fecha') > -1) sh.getRange(i+1, headers.indexOf('Fecha')+1).setValue(data.fecha);
        found = true; break;
      }
    }
    if (!found) {
      sh.appendRow([data.puesto, data.area || '', data.puntaje, data.competencias, data.estado, data.obs, data.fecha || '']);
    }
  }

  if (action === 'addEficacia') {
    const sh = ss.getSheetByName('KPIs') || ss.insertSheet('KPIs');
    sh.appendRow([data.id, data.tema, data.fecha, data.cantidad, data.resultado, data.obs, new Date()]);
  }

  if (action === 'editEficacia') {
    const sh = ss.getSheetByName('KPIs');
    if (sh) {
      const vals = sh.getDataRange().getValues();
      const idx = parseInt(data.idx) + 2; // +1 header +1 base-1
      if (idx < vals.length + 1) {
        sh.getRange(idx, 1, 1, 6).setValues([[
          vals[idx-1][0], data.tema, data.fecha, data.cantidad, data.resultado, data.obs
        ]]);
      }
    }
  }

  // ── ALERTAS ──
  if (action === 'updateAlertas') {
    let sh = ss.getSheetByName('Alertas');
    if (!sh) {
      sh = ss.insertSheet('Alertas');
      sh.appendRow(['Titulo', 'Detalle', 'Tipo', 'FechaActualizacion']);
    }
    const alertas = JSON.parse(data.alertas || '[]');
    // Limpiar y reescribir
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 4).clearContent();
    alertas.forEach(a => {
      sh.appendRow([a.titulo, a.detalle, a.tipo, new Date()]);
    });
  }

  // ── PERFILES (versión anterior; el tablero nuevo guarda los perfiles en "Store") ──
  if (action === 'updatePerfil') {
    let sh = ss.getSheetByName('Perfiles');
    if (!sh) {
      sh = ss.insertSheet('Perfiles');
      sh.appendRow(['Puesto','Area','Tipo','Formacion','Experiencia','Competencias','Responsabilidades','Tareas','FechaActualizacion']);
    }
    const perfil = JSON.parse(data.perfil || '{}');
    const vals = sh.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][0] === perfil.puesto) {
        sh.getRange(i+1, 1, 1, 9).setValues([[
          perfil.puesto, perfil.area, perfil.tipo, perfil.formacion, perfil.exp,
          JSON.stringify(perfil.comps || {}),
          perfil.responsabilidades || '',
          perfil.tareas || '',
          new Date()
        ]]);
        found = true; break;
      }
    }
    if (!found) {
      sh.appendRow([
        perfil.puesto, perfil.area, perfil.tipo, perfil.formacion, perfil.exp,
        JSON.stringify(perfil.comps || {}),
        perfil.responsabilidades || '',
        perfil.tareas || '',
        new Date()
      ]);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getMesName(num) {
  return ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'][parseInt(num)-1];
}

/* ===== MÓDULO STORE — Tablero RRHH VALPOB ===== */
var STORE_SHEET = 'Store';
var STORE_CHUNK = 45000;

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
    if (fila > 0 && (Number(data[fila - 1][1]) || 0) > ts) return storeJson_({ ok: true, skipped: true });
    if (fila < 0) fila = sh.getLastRow() + 1;
    var partes = [];
    for (var p = 0; p < val.length; p += STORE_CHUNK) partes.push("'" + val.substring(p, p + STORE_CHUNK));
    if (!partes.length) partes.push("'");
    var filaVals = [key, ts].concat(partes);
    if (sh.getMaxColumns() < filaVals.length) sh.insertColumnsAfter(sh.getMaxColumns(), filaVals.length - sh.getMaxColumns());
    sh.getRange(fila, 1, 1, Math.max(sh.getLastColumn(), filaVals.length)).clearContent();
    sh.getRange(fila, 1, 1, filaVals.length).setValues([filaVals]);
    return storeJson_({ ok: true, ts: ts });
  } finally {
    lock.releaseLock();
  }
}
