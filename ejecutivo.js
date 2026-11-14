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
