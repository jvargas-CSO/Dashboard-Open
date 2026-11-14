// =========================================================================
// ejecutivo.js · Vista por ejecutivo (acceso restringido)
//
// Quien entra SIN acceso a los Sheets principales entra en "modo ejecutivo": el dashboard solo
// lee los archivos «Dashboard Ejecutivo - <Nombre>» que le compartieron (los genera el script
// de Apps Script de la carpeta apps-script/). La seguridad la dan los permisos de Google sobre
// esos archivos; aquí solo se descubren y se muestran. Nada de correos ni de listas de acceso
// vive en este código (el repositorio es público).
//
//  · Se descubren con la API de Drive: archivos compartidos conmigo cuyo nombre empieza con
//    «Dashboard Ejecutivo - ».
//  · "Mi venta" = solo líneas suyas. "Historial de mis clientes" = además lo que otros
//    ejecutivos vendieron a sus clientes actuales (el script ya lo trae sin costos, utilidad,
//    ROI ni comisiones, y con el nombre «Otro ejecutivo»).
//  · Una ADV o un mentor reciben varios archivos: selector de a quién ver o "Todos mis
//    ejecutivos" (en ese caso solo líneas propias, para no duplicar historial).
// =========================================================================
const EXEC_FILE_PREFIX = 'Dashboard Ejecutivo - ';
const EXEC_OTRO = 'Otro Ejecutivo';          // normEjecutivo('Otro ejecutivo')
let execMode = false;
let execChecked = false;
let execFiles = [];                           // [{ id, name, ejecutivo, data: { records, fc, warnings, info } }]
let execSel = 'ALL';                          // id de archivo, o 'ALL' = todos mis ejecutivos
let execView = 'mia';                         // 'mia' | 'historial'
let execWired = false;

// ¿Esta cuenta puede abrir el Sheet principal? Si no (403/404) → modo ejecutivo.
async function execDetectar() {
  if (execChecked) return execMode;
  try {
    await driveApiRequest(`https://www.googleapis.com/drive/v3/files/${DRIVE_SHEETS.dataComercial}?fields=id&supportsAllDrives=true`);
    execMode = false;
  } catch (e) {
    if (e && e.code === 'AUTH_REQUIRED') throw e;
    if (e && (e.code === 'NO_ACCESS' || /HTTP 404/.test(String(e.message)))) execMode = true;
    else throw e;
  }
  execChecked = true;
  document.body.classList.toggle('exec-mode', execMode);
  if (!execMode) document.getElementById('execBar').style.display = 'none';
  return execMode;
}

// Libro de un archivo de ejecutivo -> registros, forecast y fecha de actualización.
function execProcesar(wb) {
  const records = processDataWorkbook(wb);
  const s24 = wb.Sheets['Concentrado 2024'];
  if (s24) {
    const vals = XLSX.utils.sheet_to_json(s24, { header: 1, defval: null, raw: true });
    processConcentrado2024Values(vals).forEach(r => records.push(r));
  }
  const fcNames = wb.SheetNames.filter(n => /forecast/i.test(n));
  const fc = fcNames.length ? processForecastWorkbook({ SheetNames: fcNames, Sheets: wb.Sheets }, 2026) : { records: [], warnings: [] };
  let info = '';
  if (wb.Sheets['Info']) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets['Info'], { header: 1, defval: null, raw: true });
    info = (rows[1] && rows[1][0]) || '';
  }
  return { records, fc: fc.records, warnings: fc.warnings, info };
}

async function loadExecModeAndBoot({ silent = false } = {}) {
  if (execPreview) { await execVerComo(execPreview, { silent: true }); return; }   // dirección viendo como otra persona
  const loading = document.getElementById('loading');
  if (!silent) setDzStatus('Buscando tus archivos de seguimiento…');
  else { loading.classList.remove('hide'); document.getElementById('loadingText').textContent = 'Actualizando tu información…'; }
  try {
    const q = encodeURIComponent("sharedWithMe = true and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false and name contains 'Dashboard'");
    const res = await driveApiRequest(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=${encodeURIComponent('files(id,name,modifiedTime)')}&pageSize=100&supportsAllDrives=true&includeItemsFromAllDrives=true`);
    const list = ((await res.json()).files || []).filter(f => String(f.name).startsWith(EXEC_FILE_PREFIX));
    if (!list.length) {
      loading.classList.add('hide');
      showSignInScreen('Esta cuenta de Google no tiene acceso al dashboard. Pide que te den de alta o inicia sesión con otra cuenta.');
      return;
    }
    if (!silent) setDzStatus(`Cargando tu información (${list.length} archivo${list.length === 1 ? '' : 's'})…`);
    const loaded = await Promise.all(list.map(async f => {
      const wb = await fetchWorkbookFromDrive(f.id);
      return { id: f.id, name: f.name, ejecutivo: f.name.slice(EXEC_FILE_PREFIX.length).trim(), data: execProcesar(wb) };
    }));
    loaded.sort((a, b) => a.ejecutivo.localeCompare(b.ejecutivo, 'es'));
    execFiles = loaded;
    if (execSel !== 'ALL' && !execFiles.some(f => f.id === execSel)) execSel = 'ALL';
    if (execFiles.length === 1) execSel = execFiles[0].id;
    execAplicar({ boot: !silent });
  } catch (err) {
    console.error(err);
    loading.classList.add('hide');
    if (err && err.code === 'AUTH_REQUIRED') showSignInScreen('Tu sesión de Google expiró. Inicia sesión de nuevo.');
    else setDzStatus('No se pudo cargar tu información. ' + (err && err.message ? err.message : ''), { error: true });
  }
}

// Arma lo que se ve según el selector (a quién) y la vista (mi venta / historial).
function execAplicar({ boot = false } = {}) {
  const sel = execSel === 'ALL' ? execFiles : execFiles.filter(f => f.id === execSel);
  const soloMias = execSel === 'ALL' || execView === 'mia';
  const records = [], fc = []; let warnings = [];
  sel.forEach(f => {
    f.data.records.forEach(r => { if (soloMias && r.eje === EXEC_OTRO) return; records.push(r); });
    f.data.fc.forEach(x => fc.push(x));
    warnings = warnings.concat(f.data.warnings || []);
  });
  if (execPreview) { Engine.operaciones = null; Engine.opsMatchByCM = {}; Engine.patrimonial = null; Engine.contactos = []; }
  Engine.load(records, { name: 'Mi información' });
  dataLoaded = true;
  Engine.loadForecast(fc, { name: 'Mi forecast' });
  forecastLoaded = fc.length > 0;
  forecastWarnings = warnings;
  contactosLoaded = false; operacionesLoaded = false; patrimonialLoaded = false; concentrado2024Loaded = false;
  filters.eje = '';
  if (filters.anio && !Engine.yearsAvailable.includes(+filters.anio)) filters.anio = '';
  if (boot) bootDashboard();
  else { populateFilters(); updateHeaderInfo(); render(); }
  execRenderBar();
  document.getElementById('loading').classList.add('hide');
}

function execRenderBar() {
  document.getElementById('execBar').style.display = 'flex';
  const todos = execSel === 'ALL' && execFiles.length > 1;
  const actual = execFiles.find(f => f.id === execSel) || execFiles[0];
  document.getElementById('execWho').textContent = todos ? 'Todos mis ejecutivos' : (actual ? actual.ejecutivo : '');
  const wrap = document.getElementById('execPickWrap'), pick = document.getElementById('execPick');
  if (execFiles.length > 1) {
    wrap.style.display = 'flex';
    pick.innerHTML = '<option value="ALL">Todos mis ejecutivos</option>' + execFiles.map(f => `<option value="${escapeHtml(f.id)}">${escapeHtml(f.ejecutivo)}</option>`).join('');
    pick.value = execSel;
  } else wrap.style.display = 'none';
  document.querySelectorAll('#execView button').forEach(b => {
    b.classList.toggle('on', b.dataset.v === (execSel === 'ALL' ? 'mia' : execView));
    b.classList.toggle('dis', execSel === 'ALL' && b.dataset.v === 'historial');
  });
  const fechas = execFiles.map(f => f.data.info).filter(Boolean).map(s => new Date(s)).filter(d => !isNaN(d));
  const info = document.getElementById('execInfo');
  info.textContent = fechas.length ? 'Datos al ' + new Date(Math.min(...fechas)).toLocaleString('es-MX', { dateStyle: 'medium', timeStyle: 'short' }) : '';
}

function execWire() {
  if (execWired) return;
  execWired = true;
  document.getElementById('execPick').addEventListener('change', e => { execSel = e.target.value; execAplicar(); });
  document.getElementById('execView').addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b || b.classList.contains('dis')) return;
    execView = b.dataset.v; execAplicar();
  });
}
execWire();


// =========================================================================
// "VISTA DE" — solo dirección: ver exactamente lo que ve una persona.
// Carga el archivo REAL «Dashboard Ejecutivo - …» que generó el script (la dirección es la dueña
// de esos archivos, así que puede abrirlos) y arma la misma vista que tendría esa persona. Sirve
// para probar sin necesitar su cuenta. No prueba los permisos de Google (eso lo hace Google).
// =========================================================================
const ACCESOS_SHEET_ID = '1yoRIFp8wpP9XwoG8-ouqY-dItCHF3h6OOGVdNX61LsY';   // hoja "accesos dashboard" (protegida por permisos de Google)
let execPreview = null;        // { nombre, puesto, ejecutivos:[] } mientras se ve como otra persona
let execPersonas = [];

const execKey = n => { const k = normTexto(n); return k === 'viani' ? 'viani paz' : k; };
const execPlaceholder = x => ['pendiente', 'por definir', 'tbd', 'na', 'n a', 'ninguno', 'no aplica'].includes(normTexto(x)) || /^[-–—\s]*$/.test(String(x));

// Misma regla que parsearAccesos() del script de Apps Script (sin correos: aquí solo se necesitan nombres)
function execParsearPersonas(values) {
  let h = -1, idx = {};
  for (let i = 0; i < Math.min(10, values.length); i++) {
    const m = {}; (values[i] || []).forEach((c, j) => { const n = normTexto(c); if (n && m[n] == null) m[n] = j; });
    if (m['correo'] != null) { h = i; idx = m; break; }
  }
  if (h < 0) return [];
  const col = names => { for (const n of names) if (idx[n] != null) return idx[n]; return -1; };
  const iNom = col(['nombre']), iPue = col(['puesto', 'rol']), iCor = col(['correo']), iAcc = col(['acceso']), iVe = col(['ve a', 've tambien a', 'ejecutivo s', 'ejecutivos', 'ejecutivo']);
  const out = [];
  values.slice(h + 1).forEach(r => {
    if (!String(r[iCor] || '').includes('@')) return;
    const nombre = iNom >= 0 ? String(r[iNom] || '').trim() : '', puesto = iPue >= 0 ? String(r[iPue] || '').trim() : '';
    const total = normTexto(iAcc >= 0 ? r[iAcc] : '') === 'total';
    const esAdv = /\badv\b/i.test(puesto);
    const extras = iVe >= 0 ? String(r[iVe] || '').split(/[;,\n]/).map(x => x.trim()).filter(x => x && !execPlaceholder(x)) : [];
    const ejecutivos = [];
    if (!total) {
      if (!esAdv && nombre) ejecutivos.push(nombre);
      extras.forEach(e => { if (!ejecutivos.some(x => execKey(x) === execKey(e))) ejecutivos.push(e); });
    }
    if (nombre) out.push({ nombre, puesto, total, ejecutivos });
  });
  return out;
}

async function fetchAccesosValues() {
  const meta = await (await driveApiRequest(`https://sheets.googleapis.com/v4/spreadsheets/${ACCESOS_SHEET_ID}?fields=sheets.properties.title`)).json();
  const tabs = (meta.sheets || []).map(s => s.properties.title);
  const tab = tabs.find(t => normTexto(t) === 'accesos') || tabs[0];
  const res = await driveApiRequest(`https://sheets.googleapis.com/v4/spreadsheets/${ACCESOS_SHEET_ID}/values/${encodeURIComponent("'" + tab + "'")}?valueRenderOption=UNFORMATTED_VALUE`);
  return (await res.json()).values || [];
}

// Se llama al terminar la carga de dirección. Si la cuenta no puede leer la hoja de accesos, el menú no aparece.
async function execInitPreview() {
  if (execMode || execPreview) return;
  try {
    execPersonas = execParsearPersonas(await fetchAccesosValues()).filter(p => !p.total && p.ejecutivos.length);
  } catch (e) { execPersonas = []; }
  const bar = document.getElementById('previewBar');
  if (!execPersonas.length) { bar.style.display = 'none'; return; }
  const pick = document.getElementById('previewPick');
  pick.innerHTML = '<option value="">Mi vista (completa)</option>' + execPersonas.map((p, i) => `<option value="${i}">${escapeHtml(p.nombre)} · ${escapeHtml(p.puesto)}</option>`).join('');
  pick.value = '';
  bar.style.display = 'flex';
  document.getElementById('previewInfo').textContent = 'Para probar lo que ve cada persona (con los archivos reales que genera el script).';
  if (!pick.dataset.wired) {
    pick.dataset.wired = '1';
    pick.addEventListener('change', () => { const v = pick.value; if (v === '') execVolverAMiVista(); else execVerComo(execPersonas[+v]); });
    document.getElementById('previewBack').addEventListener('click', () => execVolverAMiVista());
  }
}

async function execVerComo(persona, { silent = false } = {}) {
  const loading = document.getElementById('loading');
  loading.classList.remove('hide'); document.getElementById('loadingText').textContent = `Cargando la vista de ${persona.nombre}…`;
  try {
    const q = encodeURIComponent("'me' in owners and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false and name contains 'Dashboard'");
    const res = await driveApiRequest(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=${encodeURIComponent('files(id,name)')}&pageSize=200&supportsAllDrives=true&includeItemsFromAllDrives=true`);
    const archivos = ((await res.json()).files || []).filter(f => String(f.name).startsWith(EXEC_FILE_PREFIX));
    const porClave = new Map(archivos.map(f => [execKey(f.name.slice(EXEC_FILE_PREFIX.length)), f]));
    const hallados = [], faltan = [];
    persona.ejecutivos.forEach(n => { const f = porClave.get(execKey(n)); if (f) hallados.push({ f, n }); else faltan.push(n); });
    if (!hallados.length) {
      loading.classList.add('hide');
      alert(`Aún no existen los archivos de ${persona.ejecutivos.join(', ')}.\n\nEjecuta «Dashboard → Actualizar copias ahora» en la hoja de accesos y vuelve a intentar.`);
      document.getElementById('previewPick').value = execPreview ? String(execPersonas.indexOf(execPreview)) : '';
      return;
    }
    const cargados = await Promise.all(hallados.map(async ({ f, n }) => ({ id: f.id, name: f.name, ejecutivo: n, data: execProcesar(await fetchWorkbookFromDrive(f.id)) })));
    cargados.sort((a, b) => a.ejecutivo.localeCompare(b.ejecutivo, 'es'));
    execFiles = cargados;
    execPreview = persona;
    execMode = true; execChecked = true;
    document.body.classList.add('exec-mode');
    execSel = execFiles.length === 1 ? execFiles[0].id : 'ALL';
    execView = 'mia';
    const entrando = !silent;
    execAplicar();
    if (entrando) document.querySelector('.nav-item[data-tab="resumen"]').click();
    const info = document.getElementById('previewInfo');
    info.textContent = `Estás viendo lo que ve ${persona.nombre} (${persona.puesto}).` + (faltan.length ? ` Sin archivo aún: ${faltan.join(', ')}.` : '');
    document.getElementById('previewBack').style.display = '';
    document.getElementById('previewPick').value = String(execPersonas.indexOf(persona));
  } catch (err) {
    console.error(err);
    loading.classList.add('hide');
    alert('No se pudo cargar esa vista: ' + (err && err.message ? err.message : err));
  }
}

async function execVolverAMiVista() {
  execPreview = null; execMode = false; execChecked = true; execFiles = [];
  document.body.classList.remove('exec-mode');
  document.getElementById('execBar').style.display = 'none';
  document.getElementById('previewBack').style.display = 'none';
  document.getElementById('previewPick').value = '';
  document.getElementById('previewInfo').textContent = 'Para probar lo que ve cada persona (con los archivos reales que genera el script).';
  // Proyección: recuperar los supuestos de dirección (en la vista previa se usaron los de ejecutivo)
  if (typeof pxLoad === 'function') { PX = pxLoad(); pxExecInit = false; pxPanelSig = ''; }
  await loadFromDriveAndBoot({ silent: true });
}
