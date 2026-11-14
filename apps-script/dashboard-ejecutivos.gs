/**
 * Dashboard Comercial Open — copias por ejecutivo (Google Apps Script)
 * ---------------------------------------------------------------------------------------
 * Qué hace: para cada ejecutivo de la hoja «Accesos» crea (o actualiza) un archivo de Google
 * Sheets llamado «Dashboard Ejecutivo - <Nombre>» con SOLO su información, y lo comparte en
 * modo lectura con los correos autorizados. El dashboard, cuando alguien entra sin acceso a
 * los Sheets principales, lee únicamente los archivos que le compartieron. La seguridad la da
 * Google (permisos del archivo): nadie puede abrir lo de otro ejecutivo.
 *
 * Qué lleva cada archivo:
 *   · Pestañas «2025», «2026»… : las líneas de venta del ejecutivo (todas las columnas) y, de
 *     sus CLIENTES ACTUALES, el historial vendido por otros ejecutivos pero SOLO con venta y
 *     fechas (sin costos, utilidad, ROI, comisiones ni proveedor; el ejecutivo aparece como
 *     «Otro ejecutivo»).
 *   · «Concentrado 2024»: lo mismo para 2024 (solo líneas con Año 2024).
 *   · «Forecast 2026», «Forecast 2027»…: solo las filas de ese ejecutivo.
 *   · «Info»: fecha de la última actualización.
 *
 * Cómo instalarlo (una vez): ver apps-script/LEEME.md.
 * ---------------------------------------------------------------------------------------
 */

var CFG = {
  DATA_ID: '18cmJNgQn-mgJSiN7p154bLmc3zKJ0va0p3cNBJAVt0Y',       // Data Comercial Final (pestañas por año)
  FORECAST_ID: '1GNm0czUzY-WF5S8BtV-jqvxjNZ6KNbpZnI1O-nlvsb0',   // Forecast por vendedor (pestañas por año)
  CONC24_ID: '1j1LQeU6NDxLn4y3My5vahTFUxVUtjAqazV6i6R2Fs1Y',     // Concentrado de ventas 2024
  CONC24_TAB: 'Data',
  ACCESOS_TAB: 'Accesos',
  FOLDER_NAME: 'Dashboard Ejecutivos',
  FILE_PREFIX: 'Dashboard Ejecutivo - ',                          // el dashboard busca archivos con este prefijo
  OTRO: 'Otro ejecutivo',
  HORA_DIARIA: 5,                                                 // actualización automática diaria (hora local)
};

// =======================================================================================
// LÓGICA PURA (sin servicios de Google: se prueba fuera de Apps Script)
// =======================================================================================
function normTxt(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}
// Mismas equivalencias de nombre que usa el dashboard (normEjecutivo)
var EJE_ALIAS = { 'viani': 'viani paz' };
function normEje(s) { var n = normTxt(s); return EJE_ALIAS[n] || n; }

var COLS_EJE = ['nombre ejecutivo comercial', 'ejecutivo comercial', 'ejecutivo'];
var COLS_CLIENTE = ['cliente', 'cliente marca', 'cliente comercial', 'cliente nombre comercial'];
var COLS_ANIO = ['ano'];
var COLS_FC_EJE = ['nuevo ejecutivo', 'ejecutivo', 'vendedor'];
// Columnas que se conservan en las líneas que vendió OTRO ejecutivo (todo lo demás queda en blanco)
var PERMITIDAS_AJENAS = {};
['cm', 'empresa', 'cliente', 'cliente marca', 'cliente comercial', 'cliente nombre comercial', 'holding', 'agencia',
 'fecha inicio', 'fecha fin', 'mes', 'ano', 'estado', 'localidad', 'campana', 'tipo de compra', 'canales', 'medio',
 'soporte', 'centro comercial o aeropuerto', 'tarifa total cliente', 'venta sin roi', 'status open', 'tipo',
 'de medios', 'total m2', 'id de sitio', 'suma forecast', 'intercambio', 'intercompania'
].forEach(function (k) { PERMITIDAS_AJENAS[k] = true; });

function indiceEncabezados(header) {
  var idx = {};
  header.forEach(function (h, j) { var n = normTxt(h); if (idx[n] == null) idx[n] = j; });
  return idx;
}
function primerIndice(idx, nombres) {
  for (var i = 0; i < nombres.length; i++) if (idx[nombres[i]] != null) return idx[nombres[i]];
  return -1;
}
// Busca (en las primeras filas) la fila de encabezados que trae alguna de las columnas dadas
function filaEncabezado(values, alguna) {
  for (var i = 0; i < Math.min(10, values.length); i++) {
    var idx = indiceEncabezados(values[i] || []);
    if (primerIndice(idx, alguna) >= 0) return i;
  }
  return -1;
}
function anioDe(v, porDefecto) {
  var n = parseInt(String(v == null ? '' : v).replace(/[^0-9]/g, ''), 10);
  return n > 0 ? n : (porDefecto || 0);
}

/** Convierte los valores de una pestaña en { header, rows, iEje, iCli, iAnio }. */
function prepararTab(values, anioPestana) {
  var h = filaEncabezado(values, COLS_EJE);
  if (h < 0) return null;
  var header = values[h];
  var idx = indiceEncabezados(header);
  return {
    header: header,
    rows: values.slice(h + 1),
    iEje: primerIndice(idx, COLS_EJE),
    iCli: primerIndice(idx, COLS_CLIENTE),
    iAnio: primerIndice(idx, COLS_ANIO),
    anioPestana: anioPestana || 0,
  };
}

/** Clientes "actuales" del ejecutivo: los que tienen línea suya del año `desde` en adelante. */
function clientesActuales(tabs, ejeNorm, desde) {
  var set = {};
  tabs.forEach(function (t) {
    if (!t || t.iEje < 0 || t.iCli < 0) return;
    t.rows.forEach(function (r) {
      if (normEje(r[t.iEje]) !== ejeNorm) return;
      var a = t.iAnio >= 0 ? anioDe(r[t.iAnio], t.anioPestana) : t.anioPestana;
      if (a >= desde) { var c = normTxt(r[t.iCli]); if (c) set[c] = true; }
    });
  });
  return set;
}

/**
 * Filas que lleva el archivo de un ejecutivo para una pestaña de ventas:
 *  · líneas suyas → completas;
 *  · líneas de otros ejecutivos a sus clientes actuales → solo columnas permitidas y «Otro ejecutivo».
 * soloAnio (opcional): para el concentrado 2024, solo líneas de ese año.
 */
function filasEjecutivo(tab, ejeNorm, clientes, soloAnio) {
  var out = [], propias = 0, ajenas = 0;
  // Una columna se conserva solo si su encabezado normalizado está en la lista de permitidas. Ojo: un
  // encabezado como «%» se normaliza a '' y NO debe pasar (son porcentajes de ROI / comisión); solo se
  // conserva la columna cuyo encabezado está realmente vacío (en algunas pestañas es la del CM).
  var permitidas = tab.header.map(function (h) {
    var crudo = String(h == null ? '' : h).trim(), n = normTxt(h);
    return crudo === '' ? true : (n !== '' && !!PERMITIDAS_AJENAS[n]);
  });
  tab.rows.forEach(function (r) {
    if (soloAnio) {
      var a = tab.iAnio >= 0 ? anioDe(r[tab.iAnio], soloAnio) : soloAnio;
      if (a !== soloAnio) return;
    }
    var esMia = normEje(r[tab.iEje]) === ejeNorm;
    if (esMia) { out.push(r); propias++; return; }
    var c = tab.iCli >= 0 ? normTxt(r[tab.iCli]) : '';
    if (!c || !clientes[c]) return;
    var fila = tab.header.map(function (_, j) { return permitidas[j] && r[j] !== undefined ? r[j] : ''; });
    fila[tab.iEje] = CFG.OTRO;
    out.push(fila); ajenas++;
  });
  return { rows: out, propias: propias, ajenas: ajenas };
}

/** Forecast: solo las filas de ese ejecutivo. */
function filasForecast(values, ejeNorm) {
  var h = filaEncabezado(values, COLS_FC_EJE);
  if (h < 0) return null;
  var header = values[h], idx = indiceEncabezados(header), iE = primerIndice(idx, COLS_FC_EJE);
  var rows = values.slice(h + 1).filter(function (r) { return normEje(r[iE]) === ejeNorm; });
  return { header: header, rows: rows };
}

/** Lee la hoja Accesos: [{correo, rol, ejecutivos:[...]}] (ignora filas incompletas). */
function parsearAccesos(values) {
  var out = [];
  values.slice(1).forEach(function (r, i) {
    var correo = String(r[0] || '').trim().toLowerCase();
    if (!correo || correo.indexOf('@') < 0) return;
    var ejes = String(r[2] || '').split(/[;,\n]/).map(function (s) { return s.trim(); }).filter(Boolean);
    out.push({ fila: i + 2, correo: correo, rol: String(r[1] || '').trim(), ejecutivos: ejes });
  });
  return out;
}
/** Agrupa por ejecutivo: { ejeNorm: { nombre, correos:[...] } } */
function agruparPorEjecutivo(accesos) {
  var mapa = {};
  accesos.forEach(function (a) {
    a.ejecutivos.forEach(function (nombre) {
      var k = normEje(nombre);
      if (!mapa[k]) mapa[k] = { nombre: nombre, correos: [] };
      if (mapa[k].correos.indexOf(a.correo) < 0) mapa[k].correos.push(a.correo);
    });
  });
  return mapa;
}

if (typeof module !== 'undefined') {
  module.exports = { CFG: CFG, normTxt: normTxt, normEje: normEje, prepararTab: prepararTab, clientesActuales: clientesActuales,
    filasEjecutivo: filasEjecutivo, filasForecast: filasForecast, parsearAccesos: parsearAccesos, agruparPorEjecutivo: agruparPorEjecutivo, anioDe: anioDe };
}

// =======================================================================================
// SERVICIOS DE GOOGLE (solo corre dentro de Apps Script)
// =======================================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dashboard')
    .addItem('Actualizar copias ahora', 'actualizarTodo')
    .addItem('Programar actualización diaria', 'programarDiaria')
    .addItem('Revisar accesos (sin actualizar)', 'revisarAccesos')
    .addToUi();
}

function programarDiaria() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'actualizarTodo') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('actualizarTodo').timeBased().everyDays(1).atHour(CFG.HORA_DIARIA).create();
  SpreadsheetApp.getUi().alert('Listo: se actualizará todos los días alrededor de las ' + CFG.HORA_DIARIA + ':00.');
}

function leerAccesos_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(CFG.ACCESOS_TAB);
  if (!hoja) throw new Error('Falta la pestaña «' + CFG.ACCESOS_TAB + '» (Correo | Rol | Ejecutivo(s) | Estado).');
  return { hoja: hoja, accesos: parsearAccesos(hoja.getDataRange().getValues()) };
}

function revisarAccesos() {
  var a = leerAccesos_(), mapa = agruparPorEjecutivo(a.accesos);
  var n = Object.keys(mapa).length;
  SpreadsheetApp.getUi().alert('Accesos leídos: ' + a.accesos.length + ' correos, ' + n + ' ejecutivos con archivo.');
}

/** Lee todo lo necesario de las fuentes UNA vez (las fuentes no se modifican). */
function leerFuentes_() {
  var ventas = [], fcs = [];
  var data = SpreadsheetApp.openById(CFG.DATA_ID);
  data.getSheets().forEach(function (s) {
    var m = s.getName().match(/^(\d{4})$/);
    if (!m) return;
    var t = prepararTab(s.getDataRange().getValues(), parseInt(m[1], 10));
    if (t) { t.nombre = s.getName(); ventas.push(t); }
  });
  try {
    var c = SpreadsheetApp.openById(CFG.CONC24_ID).getSheetByName(CFG.CONC24_TAB);
    if (c) { var t24 = prepararTab(c.getDataRange().getValues(), 2024); if (t24) { t24.nombre = 'Concentrado 2024'; t24.soloAnio = 2024; ventas.push(t24); } }
  } catch (e) { Logger.log('Sin Concentrado 2024: ' + e); }
  var fc = SpreadsheetApp.openById(CFG.FORECAST_ID);
  fc.getSheets().forEach(function (s) {
    var m = s.getName().match(/(20\d{2})/);
    if (m) fcs.push({ nombre: 'Forecast ' + m[1], values: s.getDataRange().getValues() });
  });
  return { ventas: ventas, forecasts: fcs };
}

function carpeta_() {
  var it = DriveApp.getFoldersByName(CFG.FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CFG.FOLDER_NAME);
}
function archivoDe_(nombre, carpeta) {
  var it = carpeta.getFilesByName(CFG.FILE_PREFIX + nombre);
  if (it.hasNext()) return SpreadsheetApp.open(it.next());
  var ss = SpreadsheetApp.create(CFG.FILE_PREFIX + nombre);
  var f = DriveApp.getFileById(ss.getId());
  carpeta.addFile(f); DriveApp.getRootFolder().removeFile(f);
  return ss;
}
function escribirTab_(ss, nombre, header, rows) {
  var hoja = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  hoja.clearContents();
  var matriz = [header].concat(rows);
  if (matriz.length && header.length) hoja.getRange(1, 1, matriz.length, header.length).setValues(matriz);
}
function sincronizarLectores_(ss, correos) {
  var f = DriveApp.getFileById(ss.getId());
  var deseados = {}; correos.forEach(function (c) { deseados[c.toLowerCase()] = true; });
  var actuales = {};
  f.getViewers().forEach(function (u) { actuales[u.getEmail().toLowerCase()] = u; });
  Object.keys(deseados).forEach(function (c) { if (!actuales[c]) f.addViewer(c); });
  Object.keys(actuales).forEach(function (c) { if (!deseados[c]) f.removeViewer(c); });
  // Editores que no sean el dueño: no debería haber
  f.getEditors().forEach(function (u) { if (u.getEmail() !== Session.getEffectiveUser().getEmail()) f.removeEditor(u); });
  f.setShareableByEditors(false);
}

/** Si un ejecutivo ya no está en Accesos, su archivo se queda pero se le quitan todos los lectores. */
function revocarHuerfanos_(carp, mapa) {
  var it = carp.getFiles(), revocados = 0;
  while (it.hasNext()) {
    var f = it.next(), n = f.getName();
    if (n.indexOf(CFG.FILE_PREFIX) !== 0) continue;
    if (mapa[normEje(n.substring(CFG.FILE_PREFIX.length))]) continue;
    f.getViewers().forEach(function (u) { f.removeViewer(u); revocados++; });
  }
  return revocados;
}

function actualizarTodo() {
  var a = leerAccesos_(), mapa = agruparPorEjecutivo(a.accesos);
  var fuentes = leerFuentes_();
  var desde = new Date().getFullYear() - 1;   // «clientes actuales»: venta del ejecutivo este año o el anterior
  var carp = carpeta_();
  var estado = {};
  Object.keys(mapa).forEach(function (k) {
    var e = mapa[k];
    try {
      var clientes = clientesActuales(fuentes.ventas, k, desde);
      var ss = archivoDe_(e.nombre, carp);
      var propias = 0, ajenas = 0, usadas = {};
      fuentes.ventas.forEach(function (t) {
        var r = filasEjecutivo(t, k, clientes, t.soloAnio);
        propias += r.propias; ajenas += r.ajenas;
        escribirTab_(ss, t.nombre, t.header, r.rows); usadas[t.nombre] = true;
      });
      var nFc = 0;
      fuentes.forecasts.forEach(function (f) {
        var r = filasForecast(f.values, k);
        if (!r) return;
        escribirTab_(ss, f.nombre, r.header, r.rows); usadas[f.nombre] = true; nFc += r.rows.length;
      });
      escribirTab_(ss, 'Info', ['Actualizado', 'Ejecutivo'], [[new Date().toISOString(), e.nombre]]); usadas['Info'] = true;
      ss.getSheets().forEach(function (s) { if (!usadas[s.getName()] && ss.getSheets().length > 1) ss.deleteSheet(s); });
      sincronizarLectores_(ss, e.correos);
      estado[k] = propias === 0 && nFc === 0 ? 'Sin líneas para «' + e.nombre + '»: revisa que el nombre sea igual al de Data Comercial'
        : 'OK · ' + propias + ' líneas propias, ' + ajenas + ' de historial, ' + nFc + ' de forecast · ' + new Date().toLocaleString();
    } catch (err) { estado[k] = 'ERROR: ' + err.message; }
  });
  revocarHuerfanos_(carp, mapa);
  // Estado por correo en la columna D de Accesos
  a.accesos.forEach(function (acc) {
    var msgs = acc.ejecutivos.map(function (n) { return estado[normEje(n)] || ''; }).filter(Boolean);
    a.hoja.getRange(acc.fila, 4).setValue(msgs.join(' | ') || 'Sin ejecutivo asignado');
  });
}
