// =========================================================================
// app.js · Dashboard Comercial Open · v3
// =========================================================================
const MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const MESES_FULL = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const fmtMoney = v => '$' + Math.round(v||0).toLocaleString('es-MX');
const fmtMoneyShort = v => {
  const n = v||0; const a = Math.abs(n);
  if (a >= 1e9) return '$'+(n/1e9).toFixed(1)+'B';
  if (a >= 1e6) return '$'+(n/1e6).toFixed(1)+'M';
  if (a >= 1e3) return '$'+(n/1e3).toFixed(1)+'K';
  return '$'+Math.round(n||0);
};
const fmtPct = v => (v||0).toFixed(1)+'%';
const fmtNum = v => Math.round(v||0).toLocaleString('es-MX');
const PALETTE = ['#d9662c','#2563eb','#1f9d6e','#b45309','#db2777','#7c3aed','#dc4747','#0891b2','#64748b','#c2410c','#16a34a','#ca8a04','#9333ea','#0284c7'];

// =========================================================================
// GOOGLE DRIVE — fuente de datos en vivo (Sheets privados, requieren login de Google)
// =========================================================================
const DRIVE_SHEETS = {
  dataComercial: '18cmJNgQn-mgJSiN7p154bLmc3zKJ0va0p3cNBJAVt0Y',
  forecast: '1GNm0czUzY-WF5S8BtV-jqvxjNZ6KNbpZnI1O-nlvsb0',
};
// Client ID de OAuth (Google Cloud Console > APIs & Services > Credentials > OAuth client ID).
const GOOGLE_CLIENT_ID = '929626244128-ft82mdkvt7rftvg9ajg5kqiocams1a72.apps.googleusercontent.com';
const GOOGLE_SCOPE = 'https://www.googleapis.com/auth/drive.readonly';

let googleAccessToken = null;
let tokenClient = null;

function initGoogleAuth() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: GOOGLE_SCOPE,
    callback: () => {}, // se sobreescribe en cada solicitud individual
  });
}

// Pide un access token a Google. silent=true reintenta sin mostrar UI (para refrescar un token vencido).
function requestGoogleToken({ silent = false } = {}) {
  return new Promise((resolve, reject) => {
    tokenClient.callback = (resp) => {
      if (resp.error) { reject(resp); return; }
      googleAccessToken = resp.access_token;
      resolve(resp.access_token);
    };
    tokenClient.requestAccessToken({ prompt: silent ? '' : 'consent' });
  });
}

async function fetchWorkbookFromDrive(sheetId, { retried = false } = {}) {
  const mime = encodeURIComponent('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  const res = await fetch(`https://www.googleapis.com/drive/v3/files/${sheetId}/export?mimeType=${mime}`, {
    headers: { Authorization: `Bearer ${googleAccessToken}` },
  });
  if (res.status === 401 && !retried) {
    try { await requestGoogleToken({ silent: true }); }
    catch (_) { const e = new Error('Tu sesión de Google expiró.'); e.code = 'AUTH_REQUIRED'; throw e; }
    return fetchWorkbookFromDrive(sheetId, { retried: true });
  }
  if (res.status === 401) { const e = new Error('Tu sesión de Google expiró.'); e.code = 'AUTH_REQUIRED'; throw e; }
  if (res.status === 403) { const e = new Error('Tu cuenta de Google no tiene acceso a este Sheet. Pide que te compartan el archivo.'); e.code = 'NO_ACCESS'; throw e; }
  if (!res.ok) throw new Error('HTTP ' + res.status);
  const buf = await res.arrayBuffer();
  return XLSX.read(buf, { type: 'array', cellDates: true });
}

const filters = { anio:'', mes:'', emp:'', eje:'', est:'', loc:'', cat:'', hol:'', cli:'', tp:'', cc:'', statusOpen: ['Suma Forecast'] };
let charts = {};
let detSortKey = null, detSortDir = -1;
let selectedVendedor = '';
let acVendedor = '';
let dataLoaded = false;
let forecastLoaded = false;

// =========================================================================
// INIT
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
  Chart.defaults.color = '#6b665a';
  Chart.defaults.borderColor = 'rgba(36,33,28,0.06)';
  Chart.defaults.font.family = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
  setupTopbarActions();
  initGoogleAuth();
  document.getElementById('btnGoogleSignIn').addEventListener('click', () => signInAndBoot());
  showSignInScreen();
});

function setupTopbarActions() {
  document.getElementById('btnRefreshDrive').addEventListener('click', () => loadFromDriveAndBoot({ silent: dataLoaded }));
  document.getElementById('btnRetryDrive').addEventListener('click', () => loadFromDriveAndBoot());
  document.getElementById('fileInputFC').addEventListener('change', e => { if (e.target.files[0]) loadForecastFile(e.target.files[0]); });
}

function showSignInScreen(text = 'Inicia sesión con tu cuenta de Google para cargar los datos.') {
  document.getElementById('dropzone').classList.remove('hide');
  document.getElementById('app').style.display = 'none';
  document.getElementById('dzSpinner').style.display = 'none';
  document.getElementById('dzStatusText').textContent = text;
  document.getElementById('dzErrorActions').style.display = 'none';
  document.getElementById('dzSignInActions').style.display = 'flex';
}

function signInAndBoot() {
  document.getElementById('dzSignInActions').style.display = 'none';
  setDzStatus('Conectando con Google…');
  requestGoogleToken({ silent: false })
    .then(() => loadFromDriveAndBoot())
    .catch(err => {
      console.error(err);
      showSignInScreen('No se pudo iniciar sesión con Google. Intenta de nuevo.');
    });
}

function setDzStatus(text, { error = false } = {}) {
  document.getElementById('dropzone').classList.remove('hide');
  document.getElementById('app').style.display = 'none';
  document.getElementById('dzSpinner').style.display = error ? 'none' : 'block';
  document.getElementById('dzStatusText').textContent = text;
  document.getElementById('dzErrorActions').style.display = error ? 'flex' : 'none';
  document.getElementById('dzSignInActions').style.display = 'none';
}

// Carga Data Comercial + Forecast desde Google Drive.
// silent=true: refresco en caliente (dashboard ya visible) sin pasar por la pantalla de estado.
async function loadFromDriveAndBoot({ silent = false } = {}) {
  if (!silent) {
    setDzStatus('Conectando con Data Comercial (Google Drive)…');
  } else {
    document.getElementById('loading').classList.remove('hide');
    document.getElementById('loadingText').textContent = 'Actualizando desde Drive...';
  }
  try {
    const wbData = await fetchWorkbookFromDrive(DRIVE_SHEETS.dataComercial);
    const records = processDataWorkbook(wbData);
    if (!records.length) throw new Error('El Sheet de Data Comercial no tiene hojas con nombre de año (2025, 2026...) o está vacío.');
    Engine.load(records, { name: 'Data Comercial (Drive)' });
    dataLoaded = true;

    if (!silent) setDzStatus('Cargando Forecast (Google Drive)…');
    try {
      const wbFc = await fetchWorkbookFromDrive(DRIVE_SHEETS.forecast);
      const fcRecords = processForecastWorkbook(wbFc, 2026);
      Engine.loadForecast(fcRecords, { name: 'Forecast (Drive)' });
      forecastLoaded = true;
    } catch (fcErr) {
      console.error('No se pudo cargar el Forecast desde Drive:', fcErr);
      forecastLoaded = false;
    }

    if (silent) {
      populateFilters();
      updateHeaderInfo();
      render();
      document.getElementById('loading').classList.add('hide');
    } else {
      bootDashboard();
    }
  } catch (err) {
    console.error(err);
    if (silent) document.getElementById('loading').classList.add('hide');
    if (err.code === 'AUTH_REQUIRED') {
      showSignInScreen('Tu sesión de Google expiró. Inicia sesión de nuevo.');
    } else {
      setDzStatus('No se pudo cargar Data Comercial desde Drive. ' + err.message, { error: true });
    }
  }
}

function updateHeaderInfo() {
  document.getElementById('dataInfo').textContent = `${Engine.records.length.toLocaleString('es-MX')} registros · ${Engine.yearsAvailable.join(', ')}`;
  const fc = document.getElementById('fcInfo');
  fc.style.display = 'inline-flex';
  if (forecastLoaded) {
    fc.classList.add('green');
    fc.textContent = `Forecast: ${Engine.forecast.length} entradas`;
  } else {
    fc.classList.remove('green');
    fc.textContent = 'Forecast no disponible — sube uno manual';
  }
}

function loadForecastFile(file) {
  const loading = document.getElementById('loading');
  loading.classList.remove('hide');
  document.getElementById('loadingText').textContent = `Procesando ${file.name}...`;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type:'array', cellDates:true });
      const records = processForecastWorkbook(wb, 2026);
      Engine.loadForecast(records, { name: file.name });
      forecastLoaded = true;
      populateFilters();
      updateHeaderInfo();
      render();
      loading.classList.add('hide');
    } catch (err) { console.error(err); alert('Error: '+err.message); loading.classList.add('hide'); }
  };
  reader.readAsArrayBuffer(file);
}

function bootDashboard() {
  document.getElementById('dropzone').classList.add('hide');
  document.getElementById('app').style.display = 'flex';
  // Default activo: Resumen
  const first = document.querySelector('.nav-item[data-tab="resumen"]');
  if (first) {
    document.querySelectorAll('.nav-item').forEach(x => x.classList.remove('active'));
    first.classList.add('active');
    const crumb = document.getElementById('crumbActive');
    if (crumb) crumb.textContent = first.querySelector('.nav-item-text')?.textContent || 'Resumen';
  }
  populateFilters();
  attachListeners();
  updateHeaderInfo();
  // Año más reciente como default
  if (Engine.yearsAvailable.length > 0) {
    filters.anio = String(Math.max(...Engine.yearsAvailable));
    document.getElementById('f-anio').value = filters.anio;
  }
  // Vendedor inicial
  const ejes = Engine.groupBy(Engine.facturable(Engine.records), 'eje').sort((a,b)=>b.vb-a.vb);
  if (ejes.length > 0) {
    selectedVendedor = ejes[0].key;
    document.getElementById('vendedorSel').value = selectedVendedor;
  }
  render();
}

function populateFilters() {
  const uniq = (key) => [...new Set(Engine.records.map(r=>r[key]).filter(v=>v && v!=='Sin definir' && v!=='Otros'))].sort((a,b)=>String(a).localeCompare(String(b),'es'));
  const fillSel = (id, opts, includeAll=true, placeholder='Todos') => {
    const el = document.getElementById(id);
    el.innerHTML = includeAll ? `<option value="">${placeholder}</option>` : '';
    opts.forEach(o => {
      const op = document.createElement('option');
      op.value = o.v != null ? o.v : o; op.textContent = o.t || o;
      el.appendChild(op);
    });
  };
  fillSel('f-anio', Engine.yearsAvailable.map(a => ({v:a, t:String(a)})));
  fillSel('f-mes', MESES.map((m,i)=>({v:i+1, t:MESES_FULL[i]})));
  fillSel('f-emp', uniq('emp').map(v=>({v,t:v})));
  // Ejecutivos: solo los activos en 2026 (con forecast). Si no hay forecast cargado, todos.
  const ejesActivos = forecastLoaded ? Engine.vendedoresActivos2026() : uniq('eje');
  fillSel('f-eje', ejesActivos.map(v=>({v,t:v})));
  fillSel('f-est', uniq('est').map(v=>({v,t:v})));
  fillSel('f-loc', uniq('loc').map(v=>({v,t:v})));
  // Categorías ordenadas según CATEGORIAS_MIX
  const cats = [...new Set(Engine.records.map(r=>r.cat).filter(v=>v && v!=='Otros'))];
  const catsOrdered = CATEGORIAS_MIX.filter(c => cats.includes(c)).concat(cats.filter(c => !CATEGORIAS_MIX.includes(c)));
  fillSel('f-cat', catsOrdered.map(v=>({v,t:v})));
  fillSel('f-hol', uniq('hol').map(v=>({v,t:v})));
  fillSel('f-cli', uniq('cli').map(v=>({v,t:v})));
  fillSel('f-tp', uniq('tp').map(v=>({v,t:v})));

  // Multi-select Status OPEN
  setupStatusMultiSelect();

  const ejesAll = Engine.groupBy(Engine.facturable(Engine.records), 'eje').sort((a,b)=>b.vb-a.vb);
  // Filtrar a activos 2026 si hay forecast
  const ejesShow = forecastLoaded ? ejesAll.filter(e => Engine.isActive2026(e.key)) : ejesAll;
  const sel = document.getElementById('vendedorSel');
  sel.innerHTML = '';
  ejesShow.forEach(e => {
    const op = document.createElement('option');
    op.value = e.key; op.textContent = `${e.key} · ${fmtMoneyShort(e.vb)}`;
    sel.appendChild(op);
  });

  const acSel = document.getElementById('acVendedorSel');
  acSel.innerHTML = '<option value="">Todos</option>';
  ejesShow.forEach(e => {
    const op = document.createElement('option');
    op.value = e.key; op.textContent = e.key;
    acSel.appendChild(op);
  });
}

function attachListeners() {
  Object.keys(filters).forEach(k => {
    if (k === 'statusOpen') return; // multi-select se maneja aparte
    const el = document.getElementById('f-'+k);
    if (el) el.addEventListener('change', e => { filters[k] = e.target.value; render(); });
  });
  document.querySelectorAll('.nav-item').forEach(t => {
    t.addEventListener('click', () => {
      document.querySelectorAll('.nav-item').forEach(x=>x.classList.remove('active'));
      document.querySelectorAll('.section').forEach(x=>x.classList.remove('active'));
      t.classList.add('active');
      document.getElementById('sec-'+t.dataset.tab).classList.add('active');
      // Update breadcrumb
      const label = t.querySelector('.nav-item-text')?.textContent || '';
      const crumb = document.getElementById('crumbActive');
      if (crumb) crumb.textContent = label;
      // Close mobile sidebar
      document.getElementById('sidebar')?.classList.remove('mobile-open');
      render();
    });
  });
  // Sidebar toggle (collapse)
  const sbToggle = document.getElementById('sbToggle');
  if (sbToggle) sbToggle.addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('collapsed');
    // Resize charts after collapse animation
    setTimeout(() => {
      Object.values(charts).forEach(c => { try { c.resize(); } catch(e){} });
    }, 220);
  });
  // Mobile burger
  const burger = document.getElementById('mobileBurger');
  if (burger) burger.addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('mobile-open');
  });
  document.querySelectorAll('#tblDetalle th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const k = th.dataset.sort;
      if (detSortKey === k) detSortDir = -detSortDir;
      else { detSortKey = k; detSortDir = -1; }
      renderDetalle();
    });
  });
  document.getElementById('vendedorSel').addEventListener('change', e => {
    selectedVendedor = e.target.value;
    if (document.querySelector('.nav-item.active').dataset.tab === 'vendedor') renderVendedor();
  });
  document.getElementById('acVendedorSel').addEventListener('change', e => {
    acVendedor = e.target.value;
    if (document.querySelector('.nav-item.active').dataset.tab === 'alcanceCliente') renderAlcanceCliente();
  });
}
function resetFilters() {
  Object.keys(filters).forEach(k => {
    if (k === 'statusOpen') return; // se resetea aparte
    filters[k] = '';
  });
  Object.keys(filters).forEach(k => {
    if (k === 'statusOpen') return;
    const el = document.getElementById('f-'+k); if (el) el.value = '';
  });
  // Reset Status OPEN al default: solo "Suma Forecast"
  filters.statusOpen = ['Suma Forecast'];
  updateStatusMultiSelectUI();
  render();
}

// =========================================================================
// MULTI-SELECT Status OPEN
// =========================================================================
function setupStatusMultiSelect() {
  const panel = document.getElementById('msStatusPanel');
  const trigger = document.getElementById('msStatusTrigger');
  if (!panel || !trigger) return;

  // Conteo por status
  const counts = {};
  Engine.records.forEach(r => { counts[r.st] = (counts[r.st] || 0) + 1; });

  // Construir opciones
  const opts = Engine.statusOpenValues || [];
  panel.innerHTML = opts.map(s => `
    <label class="ms-option">
      <input type="checkbox" value="${s}" ${filters.statusOpen.includes(s) ? 'checked' : ''}>
      <span class="ms-label">${s}</span>
      <span class="ms-count-val">${(counts[s] || 0).toLocaleString('es-MX')}</span>
    </label>
  `).join('') + `
    <div class="ms-actions">
      <button type="button" id="msStatusAll">Todos</button>
      <button type="button" id="msStatusForecast">Solo Suma Forecast</button>
      <button type="button" id="msStatusClose">Cerrar</button>
    </div>
  `;

  // Toggle panel
  trigger.onclick = (e) => {
    e.stopPropagation();
    panel.classList.toggle('open');
  };
  // Cerrar al hacer click fuera
  document.addEventListener('click', (e) => {
    if (!document.getElementById('msStatus').contains(e.target)) {
      panel.classList.remove('open');
    }
  });

  // Listener checkboxes
  panel.querySelectorAll('input[type=checkbox]').forEach(cb => {
    cb.addEventListener('change', () => {
      const checked = [...panel.querySelectorAll('input[type=checkbox]:checked')].map(c => c.value);
      filters.statusOpen = checked;
      updateStatusMultiSelectUI();
      render();
    });
  });

  // Acciones
  document.getElementById('msStatusAll').onclick = () => {
    filters.statusOpen = [...opts];
    panel.querySelectorAll('input[type=checkbox]').forEach(c => c.checked = true);
    updateStatusMultiSelectUI();
    render();
  };
  document.getElementById('msStatusForecast').onclick = () => {
    filters.statusOpen = ['Suma Forecast'];
    panel.querySelectorAll('input[type=checkbox]').forEach(c => c.checked = (c.value === 'Suma Forecast'));
    updateStatusMultiSelectUI();
    render();
  };
  document.getElementById('msStatusClose').onclick = () => {
    panel.classList.remove('open');
  };

  updateStatusMultiSelectUI();
}

function updateStatusMultiSelectUI() {
  const trigger = document.getElementById('msStatusTrigger');
  const panel = document.getElementById('msStatusPanel');
  if (!trigger || !panel) return;
  const opts = Engine.statusOpenValues || [];
  const sel = filters.statusOpen || [];
  // Sync checkboxes
  panel.querySelectorAll('input[type=checkbox]').forEach(cb => {
    cb.checked = sel.includes(cb.value);
  });
  // Texto del trigger
  if (sel.length === 0) {
    trigger.textContent = '(ninguno)';
    trigger.classList.remove('has-selection');
  } else if (sel.length === opts.length) {
    trigger.textContent = 'Todos';
    trigger.classList.remove('has-selection');
  } else if (sel.length === 1) {
    trigger.textContent = sel[0];
    trigger.classList.add('has-selection');
  } else {
    trigger.innerHTML = `${sel.length} seleccionados <span class="ms-count">${sel.length}</span>`;
    trigger.classList.add('has-selection');
  }
}

// =========================================================================
// RENDER
// =========================================================================
function render() {
  const tab = document.querySelector('.nav-item.active').dataset.tab;
  switch(tab) {
    case 'resumen': renderResumen(); break;
    case 'empresa': renderEmpresa(); break;
    case 'vendedor': renderVendedor(); break;
    case 'alcanceCliente': renderAlcanceCliente(); break;
    case 'comparativo': renderComparativo(); break;
    case 'estrategia': renderEstrategia(); break;
    case 'proveedor': renderProveedor(); break;
    case 'medios': renderMedios(); break;
    case 'clientes': renderClientes(); break;
    case 'holdings': renderHoldings(); break;
    case 'geografia': renderGeografia(); break;
    case 'estacionalidad': renderEstacionalidad(); break;
    case 'rentabilidad': renderRentabilidad(); break;
    case 'detalle': renderDetalle(); break;
  }
  // Post-render: inyectar botones de export PNG y wire search inputs
  requestAnimationFrame(() => {
    addExportButtons();
    setupTableSearch();
  });
}

// =========================================================================
// HELPERS GENERALES
// =========================================================================
const MARGEN_OBJETIVO = 0.32;  // 32% sobre venta neta forecast

// Devuelve grupo del año anterior para comparativo
function groupByPrev(key, valKey='vb') {
  const yActual = parseInt(filters.anio) || 2026;
  const yPrev = yActual - 1;
  const factPrev = Engine.facturable(Engine.applyFilters({...filters, anio: yPrev}, {ignoreYear:false}));
  const grpPrev = Engine.groupBy(factPrev, key);
  const map = {};
  grpPrev.forEach(g => { map[g.key] = g; });
  return map;
}
function renderYoYCell(actual, previous, valKey='vb') {
  const a = actual?.[valKey] || 0;
  const p = previous?.[valKey] || 0;
  const yoy = p ? ((a-p)/p*100) : null;
  const cls = yoy === null ? '' : yoy>=0 ? 'pos' : 'neg';
  const txt = yoy === null ? (a>0?'🆕':'—') : (yoy>=0?'+':'')+fmtPct(yoy);
  const tooltip = `${valKey==='vb'?'V. Bruta':valKey==='vn'?'V. Neta':valKey==='ut'?'Utilidad':'Costo'}\n2026: ${fmtMoney(a)}\n2025: ${fmtMoney(p)}\nDiff: ${fmtMoney(a-p)}`;
  return `<td class="num" style="color:var(--text-muted)" title="${tooltip}">${fmtMoney(p)}</td><td class="num ${cls}" title="${tooltip}">${txt}</td>`;
}
function alcanceColor(p) {
  if (p === null || isNaN(p)) return '';
  if (p >= 100) return 'alc-good';
  if (p >= 80) return 'alc-warn';
  return 'alc-bad';
}
function fmtTooltipDiff(real, forecast, ventaPrev) {
  const diff = real - forecast;
  const sign = diff >= 0 ? '+' : '';
  let txt = `Real: ${fmtMoney(real)}\nForecast: ${fmtMoney(forecast)}\nDiff: ${sign}${fmtMoney(diff)}`;
  if (ventaPrev !== null && ventaPrev !== undefined) txt += `\nAño anterior: ${fmtMoney(ventaPrev)}`;
  return txt;
}

// =========================================================================
// RESUMEN — V7
// 3 KPI grandes (V.Bruta, V.Neta, Utilidad) con Real / vs Forecast / vs Año Pasado
// 3 gráficas lineales mensuales + Directo vs Agencia (barras Venta + Utilidad)
// =========================================================================
function renderResumen() {
  const data = Engine.applyFilters(filters);
  const fact = Engine.facturable(data);
  const t = Engine.totalize(fact);

  const yActual = parseInt(filters.anio) || (Engine.yearsAvailable.length ? Math.max(...Engine.yearsAvailable) : 2026);
  const yPrev = yActual ? yActual - 1 : null;
  const factPrev = yPrev ? Engine.facturable(Engine.applyFilters({...filters, anio:yPrev}, {ignoreYear:false})) : [];
  const tP = Engine.totalize(factPrev);

  // Forecast: solo si está cargado y es 2026
  const showFC = forecastLoaded && yActual === 2026;
  const fcNetoMonthly = showFC ? Engine.forecastMonthlyTotal(yActual) : Array(12).fill(0);
  const fcNetoTotal = fcNetoMonthly.reduce((a,b)=>a+b,0);
  const roiPct = Engine.roiPctParaForecastBruto(yActual);
  const fcBrutoMonthly = fcNetoMonthly.map(v => roiPct < 1 ? v / (1 - roiPct) : v);
  const fcBrutoTotal = fcBrutoMonthly.reduce((a,b)=>a+b,0);
  const fcUtilMonthly = fcNetoMonthly.map(v => v * MARGEN_OBJETIVO);
  const fcUtilTotal = fcUtilMonthly.reduce((a,b)=>a+b,0);

  // Mensuales del año actual y anterior
  const realVB = Engine.monthly(fact, 'vb');
  const realVN = Engine.monthly(fact, 'vn');
  const realUT = Engine.monthly(fact, 'ut');
  const prevVB = Engine.monthly(factPrev, 'vb');
  const prevVN = Engine.monthly(factPrev, 'vn');
  const prevUT = Engine.monthly(factPrev, 'ut');

  // === Helper para formatear comparativa con clase de color y signo ===
  function buildCmp(real, comp, label) {
    if (comp === null || comp === undefined || comp === 0) {
      return `
        <div class="kpi-cmp">
          <div class="kpi-cmp-label">${label}</div>
          <div class="kpi-cmp-pct muted">—</div>
          <div class="kpi-cmp-diff">Sin dato</div>
        </div>`;
    }
    const diff = real - comp;
    const pct = (diff / comp * 100);
    const cls = pct >= 0 ? 'pos' : 'neg';
    const sign = pct >= 0 ? '+' : '';
    const signD = diff >= 0 ? '+' : '';
    return `
      <div class="kpi-cmp">
        <div class="kpi-cmp-label">${label}</div>
        <div class="kpi-cmp-pct ${cls}">${sign}${fmtPct(pct)}</div>
        <div class="kpi-cmp-diff ${cls}">${signD}${fmtMoneyShort(diff)}</div>
      </div>`;
  }

  // === Helper para construir un KPI big completo ===
  function buildKPIBig(label, real, fc, prev, color, spark, sub) {
    return `
      <div class="kpi-big" style="--accent:${color}">
        <div class="kpi-big-head">
          <div class="kpi-big-label">${label}</div>
          ${sub ? `<div class="kpi-big-sub">${sub}</div>` : ''}
        </div>
        <div class="kpi-big-value">${fmtMoneyShort(real)}</div>
        <svg class="kpi-big-spark" viewBox="0 0 200 36" preserveAspectRatio="none">
          ${sparklinePath(spark, color, 200, 36)}
        </svg>
        <div class="kpi-big-compare">
          ${buildCmp(real, showFC ? fc : null, `vs Forecast ${yActual}`)}
          ${buildCmp(real, prev, `vs ${yPrev}`)}
        </div>
      </div>`;
  }

  const kpisHTML = [
    buildKPIBig('Venta Bruta', t.vb, fcBrutoTotal, tP.vb, '#d9662c', realVB, `${fmtNum(t.n)} líneas`),
    buildKPIBig('Venta Neta (sin ROI)', t.vn, fcNetoTotal, tP.vn, '#2563eb', realVN, `ROI ${fmtPct(t.vb?t.roi/t.vb*100:0)}`),
    buildKPIBig('Utilidad por Línea', t.ut, fcUtilTotal, tP.ut, '#1f9d6e', realUT, `Margen ${fmtPct(t.margen)}`),
  ].join('');
  document.getElementById('kpis').innerHTML = kpisHTML;

  // === 3 Gráficas Lineales: Real vs Forecast vs Año Pasado ===
  const buildLineDatasets = (real, fc, prev) => {
    const ds = [
      { label: 'Real ' + yActual, data: real, borderColor: '#d9662c', backgroundColor: 'rgba(217,102,44,0.12)', tension: 0.35, borderWidth: 3, pointRadius: 4, pointBackgroundColor:'#d9662c', fill: true },
    ];
    if (showFC) ds.push({
      label: 'Forecast ' + yActual,
      data: fc,
      borderColor: '#7c3aed',
      backgroundColor: 'transparent',
      tension: 0.35,
      borderWidth: 2.5,
      borderDash: [6, 4],
      pointRadius: 3,
      pointBackgroundColor: '#7c3aed',
      fill: false
    });
    if (yPrev && prev.some(v => v > 0)) ds.push({
      label: String(yPrev),
      data: prev,
      borderColor: '#6b7280',
      backgroundColor: 'transparent',
      tension: 0.35,
      borderWidth: 2,
      pointRadius: 3,
      pointBackgroundColor: '#6b7280',
      fill: false
    });
    return ds;
  };

  drawChart('ch-vbMes', {
    type:'line',
    data:{ labels: MESES, datasets: buildLineDatasets(realVB, fcBrutoMonthly, prevVB) },
    options: lineOpts({money:true})
  });
  drawChart('ch-vnMes', {
    type:'line',
    data:{ labels: MESES, datasets: buildLineDatasets(realVN, fcNetoMonthly, prevVN) },
    options: lineOpts({money:true})
  });
  drawChart('ch-utMes', {
    type:'line',
    data:{ labels: MESES, datasets: buildLineDatasets(realUT, fcUtilMonthly, prevUT) },
    options: lineOpts({money:true})
  });

  // === Directo vs Agencia: barras agrupadas con Venta Bruta + Utilidad ===
  const dirRows = fact.filter(r => r.cc === 'Directo');
  const agRows = fact.filter(r => r.cc === 'Agencia');
  const dirVB = dirRows.reduce((a,r)=>a+r.vb,0);
  const agVB = agRows.reduce((a,r)=>a+r.vb,0);
  const dirUT = dirRows.reduce((a,r)=>a+r.ut,0);
  const agUT = agRows.reduce((a,r)=>a+r.ut,0);
  const dirVN = dirRows.reduce((a,r)=>a+r.vn,0);
  const agVN = agRows.reduce((a,r)=>a+r.vn,0);
  const dirMg = dirVN ? (dirUT/dirVN*100) : 0;
  const agMg = agVN ? (agUT/agVN*100) : 0;

  drawChart('ch-canal', {
    type:'bar',
    data:{
      labels: [`Directo · Mgn ${fmtPct(dirMg)}`, `Agencia · Mgn ${fmtPct(agMg)}`],
      datasets:[
        {label:'Venta Bruta', data:[dirVB, agVB], backgroundColor:'rgba(217,102,44,0.85)', borderRadius:6, borderSkipped: false},
        {label:'Utilidad por línea', data:[dirUT, agUT], backgroundColor:'rgba(31,157,110,0.85)', borderRadius:6, borderSkipped: false},
      ]
    },
    options: barOpts({money:true, onClickFilter: (label) => {
      // El label viene con "Directo · Mgn X%" o "Agencia · Mgn X%" — extraer
      const canal = label.startsWith('Directo') ? 'Directo' : 'Agencia';
      filters.cc = filters.cc === canal ? '' : canal;
      const sel = document.getElementById('f-cc'); if (sel) sel.value = filters.cc;
      render();
    }})
  });
}

// === Sparkline como path SVG (sin viewBox interno, para que sea responsive) ===
function sparklinePath(data, color='#d9662c', w=200, h=36) {
  if (!data || data.length === 0) return '';
  const max = Math.max(...data, 1);
  const min = Math.min(...data, 0);
  const range = max - min || 1;
  const step = w / (data.length - 1);
  const points = data.map((v, i) => `${(i*step).toFixed(1)},${(h - ((v-min)/range)*h).toFixed(1)}`);
  const polyline = points.join(' ');
  const areaPath = `M0,${h} L${points.join(' L')} L${w},${h} Z`;
  const id = 'g' + Math.random().toString(36).slice(2,8);
  return `
    <defs><linearGradient id="${id}" x1="0" x2="0" y1="0" y2="1">
      <stop offset="0%" stop-color="${color}" stop-opacity="0.35"/>
      <stop offset="100%" stop-color="${color}" stop-opacity="0"/>
    </linearGradient></defs>
    <path d="${areaPath}" fill="url(#${id})"/>
    <polyline points="${polyline}" fill="none" stroke="${color}" stroke-width="2"/>`;
}

// =========================================================================
// POR EMPRESA · 3 tablas + Grupo Open
// =========================================================================
function renderEmpresa() {
  const data = Engine.applyFilters(filters);
  const fact = Engine.facturable(data);
  const empresas = ['Entretenimiento','Servicios'].filter(e => fact.some(r=>r.emp===e));
  const yActual = parseInt(filters.anio) || (Engine.yearsAvailable.length ? Math.max(...Engine.yearsAvailable) : 2026);
  const showFC = forecastLoaded && yActual === 2026;
  const roiPct = Engine.roiPctParaForecastBruto(yActual);

  // Helper: forecast NETO mensual del Grupo Open o Empresa específica
  // Como el forecast es por vendedor sin distinguir empresa, asumimos:
  //  - Forecast Grupo Open = total del forecast del año
  //  - Para split Entretenimiento/Servicios usamos proporción de ventas reales del año
  function forecastNetoMensual(emp) {
    if (!showFC) return null;
    const totFC = Engine.forecastMonthlyTotal(yActual);
    if (emp === 'Grupo Open') return totFC;
    // Calcular proporción de empresa basado en ventas reales del año
    const totVN = fact.filter(r=>r.anio===yActual).reduce((a,r)=>a+r.vn,0);
    const empVN = fact.filter(r=>r.anio===yActual && r.emp===emp).reduce((a,r)=>a+r.vn,0);
    const ratio = totVN ? empVN/totVN : 0;
    return totFC.map(v => v * ratio);
  }

  function buildTable(title, valKey, fcType) {
    let html = `<div class="card-header"><div class="card-title">${title}</div><div class="card-meta">${filters.anio || 'Todos los años'}${showFC ? ` · ROI ${yActual-1}: ${(roiPct*100).toFixed(2)}%` : ''}</div></div>`;
    html += '<div class="table-wrap"><table>';
    html += '<thead class="top"><tr><th>Empresa</th>';
    MESES_FULL.forEach(m=>html+=`<th class="num">${m}</th>`);
    html += '<th class="num">Total</th></tr></thead><tbody>';

    const empresasMostrar = [...empresas, 'Grupo Open'];
    empresasMostrar.forEach((emp) => {
      // Filas: Forecast (si aplica), Venta, Alcance %
      const sub = emp === 'Grupo Open' ? fact : fact.filter(r=>r.emp===emp);
      const valsReal = Engine.monthly(sub, valKey);
      const totReal = valsReal.reduce((a,b)=>a+b,0);

      // Sección de empresa con sub-filas
      let fcRow = null;
      let alcRow = null;
      if (showFC && fcType !== null) {
        const fcNeto = forecastNetoMensual(emp) || Array(12).fill(0);
        const fcVals = fcType === 'bruto' ? fcNeto.map(v=>Engine.forecastBrutoFromNeto(v, yActual))
                       : fcType === 'neto' ? fcNeto
                       : fcType === 'rent' ? null : null;
        if (fcVals) {
          fcRow = fcVals;
          alcRow = fcVals.map((f,i) => f ? valsReal[i]/f*100 : null);
        }
      }

      // Forecast (si aplica)
      if (fcRow) {
        html += `<tr><td class="row-label">Forecast ${title.includes('Brutas')?'Bruto':title.includes('Netas')?'Neto':''} ${emp==='Grupo Open'?'<b>(Grupo Open)</b>':emp}</td>`;
        fcRow.forEach(v => html += `<td class="num" style="color:var(--text-muted)">${fmtMoney(v)}</td>`);
        html += `<td class="num" style="color:var(--text-muted)"><b>${fmtMoney(fcRow.reduce((a,b)=>a+b,0))}</b></td></tr>`;
      }
      // Venta
      const labelVenta = title.includes('Brutas') ? 'Venta Bruta' : title.includes('Netas') ? 'Venta Neta' : 'Rentabilidad $';
      html += `<tr ${emp==='Grupo Open'?'class="subtotal-row"':''}><td class="row-label">${labelVenta} ${emp==='Grupo Open'?'<b>(Grupo Open)</b>':emp}</td>`;
      valsReal.forEach(v => html += `<td class="num">${fmtMoney(v)}</td>`);
      html += `<td class="num"><b>${fmtMoney(totReal)}</b></td></tr>`;

      // Alcance
      if (alcRow) {
        html += `<tr><td class="row-label">Alcance Forecast ${emp==='Grupo Open'?'<b>(Grupo Open)</b>':emp}</td>`;
        const totFC = fcRow.reduce((a,b)=>a+b,0);
        alcRow.forEach(p => {
          const cls = p === null ? '' : p>=100?'alc-good':p>=80?'alc-warn':'alc-bad';
          html += `<td class="num ${cls}">${p===null?'—':fmtPct(p)}</td>`;
        });
        const totAlc = totFC ? totReal/totFC*100 : null;
        const cls = totAlc===null?'' : totAlc>=100?'alc-good':totAlc>=80?'alc-warn':'alc-bad';
        html += `<td class="num ${cls}"><b>${totAlc===null?'—':fmtPct(totAlc)}</b></td></tr>`;
      }

      // Para Rentabilidad %, agregar fila adicional
      if (title.includes('Rentabilidad')) {
        const valsBruta = Engine.monthly(sub, 'vb');
        const valsNeta = Engine.monthly(sub, 'vn');
        const margenes = valsNeta.map((vn,i) => vn ? valsReal[i]/vn*100 : 0);
        html += `<tr><td class="row-label">Rentabilidad %  ${emp==='Grupo Open'?'<b>(Grupo Open)</b>':emp}</td>`;
        margenes.forEach(p => {
          const cls = p>=30?'alc-good':p>=15?'alc-warn':p<0?'alc-bad':'';
          html += `<td class="num ${cls}">${fmtPct(p)}</td>`;
        });
        const totVN = valsNeta.reduce((a,b)=>a+b,0);
        const totMg = totVN ? totReal/totVN*100 : 0;
        const clsTot = totMg>=30?'alc-good':totMg>=15?'alc-warn':totMg<0?'alc-bad':'';
        html += `<td class="num ${clsTot}"><b>${fmtPct(totMg)}</b></td></tr>`;
      }
    });

    html += '</tbody></table></div>';
    return html;
  }

  document.getElementById('cardBrutaEmp').innerHTML = buildTable('✅ Ventas Brutas (con ROI)', 'vb', 'bruto');
  document.getElementById('cardNetaEmp').innerHTML = buildTable('💵 Ventas Netas (sin ROI)', 'vn', 'neto');
  document.getElementById('cardRentEmp').innerHTML = buildTable('💰 Rentabilidad (Utilidad por línea)', 'ut', null);
}

// =========================================================================
// POR VENDEDOR
// =========================================================================
function renderVendedor() {
  if (!selectedVendedor) {
    const ejes = Engine.groupBy(Engine.facturable(Engine.records), 'eje').sort((a,b)=>b.vb-a.vb);
    if (ejes.length > 0) selectedVendedor = ejes[0].key;
  }
  const data = Engine.applyFilters(filters);
  const fact = Engine.facturable(data);
  const vData = fact.filter(r => r.eje === selectedVendedor);

  const yActual = parseInt(filters.anio) || (Engine.yearsAvailable.length ? Math.max(...Engine.yearsAvailable) : 2026);
  const showFC = forecastLoaded && yActual === 2026;

  const vb = Engine.monthly(vData, 'vb'), vn = Engine.monthly(vData, 'vn'), ut = Engine.monthly(vData, 'ut'), com = Engine.monthly(vData, 'com');
  const totVB = vb.reduce((a,b)=>a+b,0), totVN = vn.reduce((a,b)=>a+b,0), totUT = ut.reduce((a,b)=>a+b,0), totCOM = com.reduce((a,b)=>a+b,0);

  // Forecast del vendedor
  const fcNeto = showFC ? Engine.forecastMonthlyByEje(selectedVendedor, yActual) : Array(12).fill(0);
  const fcBruto = fcNeto.map(v => Engine.forecastBrutoFromNeto(v, yActual));
  const totFCNeto = fcNeto.reduce((a,b)=>a+b,0);
  const totFCBruto = fcBruto.reduce((a,b)=>a+b,0);

  let html = `<div class="table-wrap"><table>`;
  html += `<thead class="top"><tr><th>Vendedor</th><th class="center">${selectedVendedor}</th>`;
  MESES_FULL.forEach(m=>html+=`<th class="num">${m}</th>`);
  html += `<th class="num">Total</th></tr></thead><tbody>`;

  // Forecast Bruto
  html += `<tr><td class="row-label" colspan="2">Forecast Bruto</td>`;
  if (showFC) {
    fcBruto.forEach(v => html += `<td class="num">${fmtMoney(v)}</td>`);
    html += `<td class="num"><b>${fmtMoney(totFCBruto)}</b></td></tr>`;
  } else { for (let i=0;i<13;i++) html += `<td class="num" style="color:var(--text-muted)">—</td>`; html += `</tr>`; }

  // Venta Bruta
  html += `<tr><td class="row-label" colspan="2">Venta Bruta</td>`;
  vb.forEach(v => html += `<td class="num"><b>${fmtMoney(v)}</b></td>`);
  html += `<td class="num"><b>${fmtMoney(totVB)}</b></td></tr>`;

  // Alcance Forecast (Bruto)
  html += `<tr><td class="row-label" colspan="2">Alcance Forecast</td>`;
  if (showFC) {
    fcBruto.forEach((f,i) => {
      const p = f ? vb[i]/f*100 : null;
      const cls = p===null?'':p>=100?'alc-good':p>=80?'alc-warn':'alc-bad';
      html += `<td class="num ${cls}">${p===null?'—':fmtPct(p)}</td>`;
    });
    const tA = totFCBruto?totVB/totFCBruto*100:null;
    const clsT = tA===null?'':tA>=100?'alc-good':tA>=80?'alc-warn':'alc-bad';
    html += `<td class="num ${clsT}"><b>${tA===null?'—':fmtPct(tA)}</b></td></tr>`;
  } else { for (let i=0;i<13;i++) html += `<td class="num" style="color:var(--text-muted)">—</td>`; html += `</tr>`; }

  // Forecast Neto
  html += `<tr><td class="row-label" colspan="2">Forecast Neto</td>`;
  if (showFC) {
    fcNeto.forEach(v => html += `<td class="num">${fmtMoney(v)}</td>`);
    html += `<td class="num"><b>${fmtMoney(totFCNeto)}</b></td></tr>`;
  } else { for (let i=0;i<13;i++) html += `<td class="num" style="color:var(--text-muted)">—</td>`; html += `</tr>`; }

  // Venta Neta
  html += `<tr><td class="row-label" colspan="2">Venta Neta</td>`;
  vn.forEach(v => html += `<td class="num"><b>${fmtMoney(v)}</b></td>`);
  html += `<td class="num"><b>${fmtMoney(totVN)}</b></td></tr>`;

  // Alcance Neto
  html += `<tr><td class="row-label" colspan="2">Alcance Forecast</td>`;
  if (showFC) {
    fcNeto.forEach((f,i) => {
      const p = f ? vn[i]/f*100 : null;
      const cls = p===null?'':p>=100?'alc-good':p>=80?'alc-warn':'alc-bad';
      html += `<td class="num ${cls}">${p===null?'—':fmtPct(p)}</td>`;
    });
    const tA = totFCNeto?totVN/totFCNeto*100:null;
    const clsT = tA===null?'':tA>=100?'alc-good':tA>=80?'alc-warn':'alc-bad';
    html += `<td class="num ${clsT}"><b>${tA===null?'—':fmtPct(tA)}</b></td></tr>`;
  } else { for (let i=0;i<13;i++) html += `<td class="num" style="color:var(--text-muted)">—</td>`; html += `</tr>`; }

  // Rentabilidad $
  html += `<tr><td class="row-label" colspan="2">Rentabilidad $</td>`;
  ut.forEach(v => html += `<td class="num pos">${fmtMoney(v)}</td>`);
  html += `<td class="num pos"><b>${fmtMoney(totUT)}</b></td></tr>`;

  // Rentabilidad %
  html += `<tr><td class="row-label" colspan="2">Rentabilidad %</td>`;
  vn.forEach((vnVal, i) => {
    const p = vnVal ? ut[i]/vnVal*100 : 0;
    const cls = p>=30?'alc-good':p>=15?'alc-warn':p<0?'alc-bad':'';
    html += `<td class="num ${vnVal>0?cls:''}" style="${vnVal>0?'':'color:var(--text-muted)'}">${vnVal>0?fmtPct(p):'—'}</td>`;
  });
  const tp = totVN ? totUT/totVN*100 : 0;
  html += `<td class="num ${tp>=30?'alc-good':tp>=15?'alc-warn':'alc-bad'}"><b>${fmtPct(tp)}</b></td></tr>`;

  // Comisión
  html += `<tr><td class="row-label" colspan="2">Comisión</td>`;
  com.forEach(v => html += `<td class="num">${fmtMoney(v)}</td>`);
  html += `<td class="num"><b>${fmtMoney(totCOM)}</b></td></tr>`;

  // AAA y Gastos en 0
  html += `<tr><td class="row-label" colspan="2">AAA</td>` + Array(13).fill('<td class="num" style="color:var(--text-muted)">$0</td>').join('') + `</tr>`;
  html += `<tr><td class="row-label" colspan="2">Gastos</td>` + Array(13).fill('<td class="num" style="color:var(--text-muted)">$0</td>').join('') + `</tr>`;

  html += '</tbody></table></div>';
  document.getElementById('vendedorTable').innerHTML = html;

  // Ranking general con alcance global
  const ejes = Engine.groupBy(fact, 'eje').sort((a,b)=>b.vn-a.vn);
  let r = `<div class="table-wrap"><table class="table-default"><thead class="top"><tr>
    <th>Vendedor</th><th class="num">V. Bruta</th><th class="num">V. Neta</th><th class="num">Utilidad</th><th class="num">Margen %</th><th class="num">Comisión</th>`;
  if (showFC) r += `<th class="num">Forecast Neto</th><th class="num">Alcance %</th>`;
  r += `<th class="num"># Líneas</th><th class="num"># Clientes</th></tr></thead><tbody>`;
  ejes.forEach(e => {
    const cls = e.margen>=30?'alc-good':e.margen>=15?'alc-warn':'alc-bad';
    let row = `<tr><td><b>${e.key}</b></td>
      <td class="num">${fmtMoney(e.vb)}</td>
      <td class="num">${fmtMoney(e.vn)}</td>
      <td class="num pos">${fmtMoney(e.ut)}</td>
      <td class="num ${cls}">${fmtPct(e.margen)}</td>
      <td class="num">${fmtMoney(e.com)}</td>`;
    if (showFC) {
      const fcN = Engine.forecast.filter(f=>f.anio===yActual && f.eje===e.key).reduce((a,f)=>a+f.fcNeto, 0);
      const alc = fcN ? e.vn/fcN*100 : null;
      const acl = alc===null?'':alc>=100?'alc-good':alc>=80?'alc-warn':'alc-bad';
      row += `<td class="num">${fmtMoney(fcN)}</td><td class="num ${acl}">${alc===null?'—':fmtPct(alc)}</td>`;
    }
    row += `<td class="num">${fmtNum(e.n)}</td><td class="num">${fmtNum(e.nClientes)}</td></tr>`;
    r += row;
  });
  r += '</tbody></table></div>';
  document.getElementById('vendedorRanking').innerHTML = r;
}

// =========================================================================
// ALCANCE POR CLIENTE · v4 con Nuevo Negocio + mensual + trimestral
// =========================================================================
function renderAlcanceCliente() {
  if (!forecastLoaded) {
    const msg = '<div class="alert warning">⚠️ No hay archivo de Forecast cargado. Esta vista requiere el Excel "Forecast por Vendedor.xlsx". Recarga el dashboard y sube ambos archivos.</div>';
    document.getElementById('alcanceClienteMensual').innerHTML = msg;
    document.getElementById('alcanceClienteTrimestral').innerHTML = '';
    return;
  }
  const yActual = parseInt(filters.anio) || 2026;
  const fact = Engine.facturable(Engine.applyFilters({...filters, anio:yActual}, {ignoreYear:false}));

  // 1. Recolectar clientes con forecast (por vendedor o todos)
  let forecastEntries = Engine.forecast.filter(f => f.anio === yActual);
  if (acVendedor) forecastEntries = forecastEntries.filter(f => f.eje === acVendedor);

  // 2. Construir mapa: vendedor → set(clientes con forecast)
  const fcByEje = {};   // ejecutivo → { cli → [12 forecast neto] }
  forecastEntries.forEach(f => {
    if (!fcByEje[f.eje]) fcByEje[f.eje] = {};
    if (!fcByEje[f.eje][f.cli]) fcByEje[f.eje][f.cli] = Array(12).fill(0);
    if (f.mes>=1 && f.mes<=12) fcByEje[f.eje][f.cli][f.mes-1] += f.fcNeto;
  });

  // 3. Recolectar TODAS las ventas (incluyendo clientes que NO están en forecast)
  // Filtrar fact por vendedor si aplica
  const factVendedor = acVendedor ? fact.filter(r => r.eje === acVendedor) : fact;
  const ventasByEjeCli = {};   // ejecutivo → { cli → [12 venta neta], [12 utilidad] }
  factVendedor.forEach(r => {
    if (!ventasByEjeCli[r.eje]) ventasByEjeCli[r.eje] = {};
    if (!ventasByEjeCli[r.eje][r.cli]) ventasByEjeCli[r.eje][r.cli] = { vn: Array(12).fill(0), ut: Array(12).fill(0) };
    if (r.mes>=1 && r.mes<=12) {
      ventasByEjeCli[r.eje][r.cli].vn[r.mes-1] += r.vn || 0;
      ventasByEjeCli[r.eje][r.cli].ut[r.mes-1] += r.ut || 0;
    }
  });

  // 4. Lista de vendedores a mostrar (los que tengan forecast O ventas en filtro)
  const ejesShow = acVendedor ? [acVendedor] : [...new Set([...Object.keys(fcByEje), ...Object.keys(ventasByEjeCli)])].sort();

  // Generar tablas mensual y trimestral
  document.getElementById('alcanceClienteMensual').innerHTML = buildAlcanceTable(ejesShow, fcByEje, ventasByEjeCli, yActual, 'mensual');
  document.getElementById('alcanceClienteTrimestral').innerHTML = buildAlcanceTable(ejesShow, fcByEje, ventasByEjeCli, yActual, 'trimestral');
}

function buildAlcanceTable(ejes, fcByEje, ventasByEjeCli, yActual, modo) {
  const periodos = modo === 'trimestral' ? ['Q1','Q2','Q3','Q4'] : MESES_FULL;

  let html = '<div class="table-wrap"><table><thead class="top"><tr><th>Vendedor</th><th>Cliente</th>';
  periodos.forEach(p => html += `<th class="num">${p}</th>`);
  html += '<th class="num">Total</th></tr></thead><tbody>';

  ejes.forEach((eje, idx) => {
    const fcCliMap = fcByEje[eje] || {};
    const ventasCliMap = ventasByEjeCli[eje] || {};

    // Clientes con forecast
    const clientesConFC = Object.keys(fcCliMap).sort();
    // Clientes con venta pero SIN forecast (estos arman "Nuevo Negocio")
    const clientesSoloVenta = Object.keys(ventasCliMap).filter(c => !(c in fcCliMap)).sort();

    // Header del vendedor
    html += `<tr class="subtotal-row"><td colspan="${periodos.length+2}">👤 ${eje}</td></tr>`;

    // Acumuladores para totales del vendedor
    const totFC = Array(12).fill(0);
    const totVN = Array(12).fill(0);
    const totUT = Array(12).fill(0);
    const totNN_VN = Array(12).fill(0);  // ventas de nuevo negocio
    const totNN_UT = Array(12).fill(0);  // utilidad de nuevo negocio

    // 1. Filas por cliente con forecast
    clientesConFC.forEach(cli => {
      const fcMonths = fcCliMap[cli];
      const vn = ventasCliMap[cli]?.vn || Array(12).fill(0);
      const ut = ventasCliMap[cli]?.ut || Array(12).fill(0);
      const ventaPrev = Engine.ventaMensualAnioAnterior(cli, yActual, 'vn', filters.statusOpen);

      // Acumular para totales (solo clientes con forecast cuentan al forecast total)
      fcMonths.forEach((v,i) => totFC[i] += v);
      vn.forEach((v,i) => totVN[i] += v);
      ut.forEach((v,i) => totUT[i] += v);

      html += renderClienteFila(cli, fcMonths, vn, ut, ventaPrev, modo, false);
    });

    // 2. Filas por cliente SIN forecast (Nuevo Negocio individual)
    clientesSoloVenta.forEach(cli => {
      const fcMonths = Array(12).fill(0);
      const vn = ventasCliMap[cli].vn;
      const ut = ventasCliMap[cli].ut;
      const ventaPrev = Engine.ventaMensualAnioAnterior(cli, yActual, 'vn', filters.statusOpen);

      vn.forEach((v,i) => totNN_VN[i] += v);
      ut.forEach((v,i) => totNN_UT[i] += v);

      html += renderClienteFila(cli, fcMonths, vn, ut, ventaPrev, modo, true);
    });

    // 3. Fila resumen "Nuevo Negocio" (suma de marcas sin forecast)
    if (clientesSoloVenta.length > 0) {
      html += renderClienteFila('🆕 TOTAL Nuevo Negocio', Array(12).fill(0), totNN_VN, totNN_UT, Array(12).fill(0), modo, true, true);
    }

    // 4. TOTAL del vendedor: Forecast + Nuevo Negocio
    // Total venta = totVN (clientes con forecast) + totNN_VN (nuevo negocio)
    const totalVN = totVN.map((v,i) => v + totNN_VN[i]);
    const totalUT = totUT.map((v,i) => v + totNN_UT[i]);
    html += renderClienteFila(`📊 TOTAL ${eje}`, totFC, totalVN, totalUT, Array(12).fill(0), modo, false, true);
  });

  html += '</tbody></table></div>';
  return html;
}

function renderClienteFila(cli, fcMonths, vnMonths, utMonths, prevMonths, modo, esNuevoNegocio, esTotal=false) {
  let valuesFC, valuesVN, valuesUT, valuesPrev;
  if (modo === 'trimestral') {
    valuesFC = Engine.monthlyToQuarterly(fcMonths);
    valuesVN = Engine.monthlyToQuarterly(vnMonths);
    valuesUT = Engine.monthlyToQuarterly(utMonths);
    valuesPrev = Engine.monthlyToQuarterly(prevMonths);
  } else {
    valuesFC = fcMonths; valuesVN = vnMonths; valuesUT = utMonths; valuesPrev = prevMonths;
  }
  const totFC = valuesFC.reduce((a,b)=>a+b,0);
  const totVN = valuesVN.reduce((a,b)=>a+b,0);
  const totUT = valuesUT.reduce((a,b)=>a+b,0);
  const totPrev = valuesPrev.reduce((a,b)=>a+b,0);

  const labelClass = esTotal ? 'total-row' : '';
  let html = '';

  // Fila 1: Venta Neta vs Forecast
  html += `<tr class="${labelClass}"><td class="row-label" rowspan="2">${cli}</td><td class="row-label">Venta Neta</td>`;
  valuesVN.forEach((v, i) => {
    const fc = valuesFC[i];
    const alc = fc ? v/fc*100 : (v > 0 ? 999 : null);
    const cls = esNuevoNegocio && !esTotal ? 'alc-good' : alcanceColor(alc);
    const tooltip = fmtTooltipDiff(v, fc, valuesPrev[i]);
    html += `<td class="num ${cls}" title="${tooltip}">${fmtMoney(v)}</td>`;
  });
  // Total
  const totAlc = totFC ? totVN/totFC*100 : (totVN > 0 ? 999 : null);
  const totCls = esNuevoNegocio && !esTotal ? 'alc-good' : alcanceColor(totAlc);
  const totTooltip = fmtTooltipDiff(totVN, totFC, totPrev);
  html += `<td class="num ${totCls}" title="${totTooltip}"><b>${fmtMoney(totVN)}</b></td></tr>`;

  // Fila 2: Margen $ vs Margen objetivo (32% del Forecast Neto)
  html += `<tr class="${labelClass}"><td class="row-label">Margen</td>`;
  valuesUT.forEach((u, i) => {
    const margenObj = valuesFC[i] * MARGEN_OBJETIVO;
    const alc = margenObj ? u/margenObj*100 : (u > 0 ? 999 : null);
    const cls = esNuevoNegocio && !esTotal ? 'alc-good' : alcanceColor(alc);
    const tooltip = `Margen real: ${fmtMoney(u)}\nMargen objetivo (32%): ${fmtMoney(margenObj)}\nDiff: ${u-margenObj>=0?'+':''}${fmtMoney(u-margenObj)}`;
    html += `<td class="num ${cls}" title="${tooltip}">${fmtMoney(u)}</td>`;
  });
  const margenObjTot = totFC * MARGEN_OBJETIVO;
  const totMgAlc = margenObjTot ? totUT/margenObjTot*100 : (totUT > 0 ? 999 : null);
  const totMgCls = esNuevoNegocio && !esTotal ? 'alc-good' : alcanceColor(totMgAlc);
  const totMgTooltip = `Margen real: ${fmtMoney(totUT)}\nObjetivo (32%): ${fmtMoney(margenObjTot)}\nDiff: ${totUT-margenObjTot>=0?'+':''}${fmtMoney(totUT-margenObjTot)}`;
  html += `<td class="num ${totMgCls}" title="${totMgTooltip}"><b>${fmtMoney(totUT)}</b></td></tr>`;

  return html;
}

// =========================================================================
// COMPARATIVO ANUAL
// =========================================================================
function renderComparativo() {
  const yrs = Engine.yearsAvailable;
  const datasetsBy = (valKey) => yrs.map((y, i) => {
    const f = Engine.facturable(Engine.applyFilters({...filters, anio:y}, {ignoreYear:false}));
    return { label: String(y), data: Engine.monthly(f, valKey), backgroundColor: PALETTE[i], borderColor: PALETTE[i], borderRadius:4 };
  });
  drawChart('ch-cmpBruta', { type:'bar', data:{ labels: MESES, datasets: datasetsBy('vb')}, options: barOpts({money:true}) });
  drawChart('ch-cmpNeta', { type:'bar', data:{ labels: MESES, datasets: datasetsBy('vn')}, options: barOpts({money:true}) });
  drawChart('ch-cmpUtil', { type:'bar', data:{ labels: MESES, datasets: datasetsBy('ut')}, options: barOpts({money:true}) });

  const dsM = yrs.map((y, i) => {
    const f = Engine.facturable(Engine.applyFilters({...filters, anio:y}, {ignoreYear:false}));
    const vn = Engine.monthly(f, 'vn'); const ut = Engine.monthly(f, 'ut');
    const m = vn.map((v, idx) => v ? ut[idx]/v*100 : 0);
    return { label: String(y), data: m, borderColor: PALETTE[i], backgroundColor: PALETTE[i]+'33', tension:0.3, fill:false };
  });
  drawChart('ch-cmpMargen', { type:'line', data:{ labels: MESES, datasets: dsM}, options: lineOpts({pct:true}) });

  const meses = Array.from({length:12},(_,i)=>i+1);
  let html = `<div class="table-wrap"><table><thead class="top"><tr><th>Mes</th>`;
  yrs.forEach(y => html += `<th class="num">V. Bruta ${y}</th><th class="num">V. Neta ${y}</th><th class="num">Util ${y}</th><th class="num">Margen ${y}</th>`);
  if (yrs.length >= 2) html += `<th class="num">Δ Bruta vs anterior</th>`;
  html += `</tr></thead><tbody>`;

  const yearTotals = yrs.map(y => Engine.totalize(Engine.facturable(Engine.applyFilters({...filters, anio:y}, {ignoreYear:false}))));
  meses.forEach(m => {
    html += `<tr><td><b>${MESES_FULL[m-1]}</b></td>`;
    let prev = null;
    yrs.forEach((y, i) => {
      const f = Engine.facturable(Engine.applyFilters({...filters, anio:y, mes:m}, {ignoreYear:false})).filter(r=>r.mes===m);
      const vb = f.reduce((a,r)=>a+r.vb,0); const vn = f.reduce((a,r)=>a+r.vn,0); const ut = f.reduce((a,r)=>a+r.ut,0);
      const mg = vn ? ut/vn*100 : 0;
      html += `<td class="num">${fmtMoney(vb)}</td><td class="num">${fmtMoney(vn)}</td><td class="num pos">${fmtMoney(ut)}</td><td class="num">${fmtPct(mg)}</td>`;
      if (i === yrs.length - 1 && yrs.length >= 2 && prev != null) {
        const d = vb - prev; const p = prev ? d/prev*100 : 0;
        html += `<td class="num ${d>=0?'pos':'neg'}">${d>=0?'+':''}${fmtPct(p)}</td>`;
      }
      prev = vb;
    });
    html += `</tr>`;
  });

  html += `<tr class="total-row"><td><b>TOTAL</b></td>`;
  let prevTot = null;
  yrs.forEach((y, i) => {
    const t = yearTotals[i];
    html += `<td class="num">${fmtMoney(t.vb)}</td><td class="num">${fmtMoney(t.vn)}</td><td class="num pos">${fmtMoney(t.ut)}</td><td class="num">${fmtPct(t.margen)}</td>`;
    if (i === yrs.length - 1 && yrs.length >= 2 && prevTot != null) {
      const d = t.vb - prevTot; const p = prevTot ? d/prevTot*100 : 0;
      html += `<td class="num ${d>=0?'pos':'neg'}">${d>=0?'+':''}${fmtPct(p)}</td>`;
    }
    prevTot = t.vb;
  });
  html += `</tr></tbody></table></div>`;
  document.getElementById('tblComparativo').innerHTML = html;
}

// =========================================================================
// ESTRATEGIA
// =========================================================================
function renderEstrategia() {
  const data = Engine.applyFilters(filters);
  const fact = Engine.facturable(data);
  const t = Engine.totalize(fact);
  const ejes = Engine.groupBy(fact, 'eje').sort((a,b)=>b.vb-a.vb);
  const clis = Engine.groupBy(fact, 'cli').sort((a,b)=>b.vb-a.vb);
  const top1 = clis[0] ? clis[0].vb / t.vb * 100 : 0;
  const top3 = clis.slice(0,3).reduce((a,c)=>a+c.vb,0) / (t.vb||1) * 100;
  const top10 = clis.slice(0,10).reduce((a,c)=>a+c.vb,0) / (t.vb||1) * 100;
  const monoCat = clis.filter(c => {
    const cats = new Set(fact.filter(r=>r.cli===c.key).map(r=>r.cat));
    return cats.size === 1 && c.vb > 50000;
  }).length;

  document.getElementById('estKpis').innerHTML = [
    {label:'Cliente #1 concentra', val: fmtPct(top1), sub: clis[0]?.key || '—', cls: top1>=20?'danger':top1>=10?'warning':'success'},
    {label:'Top 3 clientes', val: fmtPct(top3), sub:'del total', cls: top3>=50?'danger':top3>=35?'warning':'success'},
    {label:'Top 10 clientes', val: fmtPct(top10), sub:'del total', cls: top10>=80?'warning':'info'},
    {label:'Clientes con 1 sola categoría', val: fmtNum(monoCat), sub:'oportunidades de upsell', cls:'warning'},
  ].map(k=>`<div class="kpi ${k.cls}"><div class="kpi-label">${k.label}</div><div class="kpi-value">${k.val}</div><div class="kpi-sub">${k.sub}</div></div>`).join('');

  // Mix Vendedor × Categoría
  const cats = [...new Set(fact.map(r=>r.cat))].filter(c=>c && c!=='Otros');
  const catRank = cats.map(c => ({c, v: fact.filter(r=>r.cat===c).reduce((a,r)=>a+r.vb,0)})).sort((a,b)=>b.v-a.v);
  const topCats = catRank.slice(0, 8).map(o=>o.c);

  let html = '<div class="table-wrap"><table><thead class="top"><tr><th>Vendedor</th><th class="num">Total</th>';
  topCats.forEach(c => html += `<th class="num">${c}</th>`);
  html += '<th class="num">Otros</th><th class="num"># Cats</th></tr></thead><tbody>';
  ejes.forEach(e => {
    const vData = fact.filter(r => r.eje === e.key);
    const total = vData.reduce((a,r)=>a+r.vb,0);
    const cs = new Set(vData.map(r=>r.cat));
    html += `<tr><td><b>${e.key}</b></td><td class="num">${fmtMoneyShort(total)}</td>`;
    let topSum = 0;
    topCats.forEach(c => {
      const v = vData.filter(r=>r.cat===c).reduce((a,r)=>a+r.vb,0);
      topSum += v;
      const p = total ? v/total*100 : 0;
      const cls = p > 50 ? 'alc-bad' : p > 30 ? 'alc-warn' : '';
      html += `<td class="num ${cls}" style="${p===0?'color:var(--text-dim)':''}">${p > 0 ? fmtPct(p) : '—'}</td>`;
    });
    const otros = total - topSum;
    html += `<td class="num">${otros>0 ? fmtPct(otros/total*100) : '—'}</td>`;
    html += `<td class="num">${cs.size}</td></tr>`;
  });
  html += '</tbody></table></div>';
  document.getElementById('mixVendedorMedio').innerHTML = html;

  // Concentración
  let html2 = '<div class="table-wrap"><table><thead class="top"><tr><th>Vendedor</th><th class="num">V. Bruta</th><th>Top 1 cliente</th><th class="num">% T1</th><th class="num">% T3</th><th class="num">% T5</th><th class="num"># Clientes</th><th>Riesgo</th></tr></thead><tbody>';
  ejes.forEach(e => {
    const vData = fact.filter(r => r.eje === e.key);
    const cli = Engine.groupBy(vData, 'cli').sort((a,b)=>b.vb-a.vb);
    const total = e.vb;
    const t1 = cli[0]?.vb || 0; const t3 = cli.slice(0,3).reduce((a,c)=>a+c.vb,0); const t5 = cli.slice(0,5).reduce((a,c)=>a+c.vb,0);
    const p1 = total?t1/total*100:0, p3 = total?t3/total*100:0, p5 = total?t5/total*100:0;
    const riesgo = p1>=50?'🔴 Alto':p1>=30?'🟡 Medio':'🟢 Bajo';
    const cls = p1>=50?'alc-bad':p1>=30?'alc-warn':'alc-good';
    html2 += `<tr><td><b>${e.key}</b></td><td class="num">${fmtMoney(total)}</td><td>${cli[0]?.key || '—'}</td><td class="num ${cls}">${fmtPct(p1)}</td><td class="num">${fmtPct(p3)}</td><td class="num">${fmtPct(p5)}</td><td class="num">${e.nClientes}</td><td>${riesgo}</td></tr>`;
  });
  html2 += '</tbody></table></div>';
  document.getElementById('concentracionVendedor').innerHTML = html2;

  // === Desempeño por Cuenta (sustituye Upsell) ===
  const yActual = parseInt(filters.anio) || 2026;
  const yPrev = yActual - 1;
  // Para cada cliente: ventas YoY, # categorías, margen, # ODVs, alcance forecast, status
  const cliRows = clis.slice(0, 50).map(c => {
    const recsActual = fact.filter(r=>r.cli===c.key);
    const sOpen = filters.statusOpen && filters.statusOpen.length ? filters.statusOpen : null;
    const recsPrev = Engine.records.filter(r => r.anio===yPrev && r.cli===c.key && r.tc==='Facturable' && (!sOpen || sOpen.includes(r.st)));
    const eje = [...new Set(recsActual.map(r=>r.eje))].join(', ') || '—';
    const hol = recsActual[0]?.hol || '—';
    const vnPrev = recsPrev.reduce((a,r)=>a+r.vn,0);
    const utPrev = recsPrev.reduce((a,r)=>a+r.ut,0);
    const cats = new Set(recsActual.map(r=>r.cat));
    const odvs = new Set(recsActual.map(r=>r.odv).filter(Boolean));
    const ultimaCampania = [...new Set(recsActual.map(r=>r.cmp).filter(Boolean))].slice(-1)[0] || '—';
    // Alcance forecast (solo aplica a 2026)
    let alc = null, fcNeto = 0;
    if (yActual === 2026 && forecastLoaded) {
      fcNeto = Engine.forecast.filter(f=>f.anio===2026 && f.cli===c.key).reduce((a,f)=>a+f.fcNeto, 0);
      alc = fcNeto ? c.vn/fcNeto*100 : null;
    }
    // YoY
    const yoy = vnPrev ? ((c.vn - vnPrev)/vnPrev*100) : null;
    // Status
    let status;
    if (yPrev && vnPrev > 0 && c.vn === 0) status = '<span class="pill danger">🔻 Perdida</span>';
    else if (vnPrev === 0 && c.vn > 0) status = '<span class="pill success">🆕 Nueva</span>';
    else if (yoy !== null && yoy >= 20) status = '<span class="pill success">📈 Crecimiento</span>';
    else if (yoy !== null && yoy <= -20) status = '<span class="pill warning">📉 Caída</span>';
    else if (alc !== null && alc < 80) status = '<span class="pill warning">🟡 Rezagado</span>';
    else if (alc !== null && alc >= 100) status = '<span class="pill success">✅ Cumple</span>';
    else status = '<span class="pill info">↔️ Recurrente</span>';
    return { ...c, eje, hol, vnPrev, utPrev, yoy, cats, odvs, ultimaCampania, alc, fcNeto, status };
  });

  let html3 = '<div class="table-wrap"><table><thead class="top"><tr>';
  html3 += '<th>Cliente</th><th>Vendedor</th><th>Holding</th><th class="num">V. Neta '+yActual+'</th><th class="num">V. Neta '+yPrev+'</th><th class="num">YoY %</th><th class="num">Margen %</th><th class="num">Mgn vs Vend.</th><th class="num"># Cats</th><th class="num"># ODVs</th>';
  if (yActual === 2026 && forecastLoaded) html3 += '<th class="num">Alcance FC</th>';
  html3 += '<th>Status</th></tr></thead><tbody>';
  cliRows.forEach(c => {
    const yoyCls = c.yoy === null ? '' : c.yoy>=0 ? 'pos' : 'neg';
    const yoyTxt = c.yoy === null ? '—' : (c.yoy>=0?'+':'')+fmtPct(c.yoy);
    const mgVendor = ejes.find(e => e.key === c.eje.split(',')[0])?.margen || 0;
    const mgDelta = c.margen - mgVendor;
    const mgCls = c.margen>=30?'alc-good':c.margen>=15?'alc-warn':'alc-bad';
    const alcCls = c.alc === null ? '' : alcanceColor(c.alc);
    html3 += `<tr>
      <td><b>${c.key}</b></td>
      <td>${c.eje}</td>
      <td>${c.hol}</td>
      <td class="num">${fmtMoney(c.vn)}</td>
      <td class="num" style="color:var(--text-muted)">${fmtMoney(c.vnPrev)}</td>
      <td class="num ${yoyCls}">${yoyTxt}</td>
      <td class="num ${mgCls}">${fmtPct(c.margen)}</td>
      <td class="num ${mgDelta>=0?'pos':'neg'}">${mgDelta>=0?'+':''}${fmtPct(mgDelta)}</td>
      <td class="num">${c.cats.size}</td>
      <td class="num">${c.odvs.size}</td>`;
    if (yActual === 2026 && forecastLoaded) html3 += `<td class="num ${alcCls}">${c.alc===null?'—':fmtPct(c.alc)}</td>`;
    html3 += `<td>${c.status}</td></tr>`;
  });
  html3 += '</tbody></table></div>';
  document.getElementById('desempenoCuenta').innerHTML = html3;

  // === Estacionalidad comparativa 2025 vs 2026 ===
  document.getElementById('estacionalidadVendedor').innerHTML = buildHeatmapComparativo('eje', 'vb', 8);
  document.getElementById('estacionalidadMedio').innerHTML = buildHeatmapComparativo('cat', 'vb', 10);

  // Holdings que erosionan margen
  const grpH = Engine.groupBy(fact.filter(r=>r.hol && r.hol !== 'Sin definir'), 'hol').filter(g=>g.vb>100000).sort((a,b)=>a.margen-b.margen);
  const promMargen = t.margen;
  let html4 = `<div class="alert info">Margen promedio: ${fmtPct(promMargen)}.</div>`;
  html4 += '<div class="table-wrap"><table><thead class="top"><tr><th>Holding</th><th class="num">V. Bruta</th><th class="num">V. Neta</th><th class="num">Utilidad</th><th class="num">Margen %</th><th class="num">Δ vs prom.</th><th class="num"># Clientes</th></tr></thead><tbody>';
  grpH.forEach(h => {
    const delta = h.margen - promMargen;
    const cls = delta < -10 ? 'alc-bad' : delta < 0 ? 'alc-warn' : 'alc-good';
    html4 += `<tr><td><b>${h.key}</b></td><td class="num">${fmtMoney(h.vb)}</td><td class="num">${fmtMoney(h.vn)}</td><td class="num pos">${fmtMoney(h.ut)}</td><td class="num ${cls}">${fmtPct(h.margen)}</td><td class="num ${delta<0?'neg':'pos'}">${delta>=0?'+':''}${fmtPct(delta)}</td><td class="num">${h.nClientes}</td></tr>`;
  });
  html4 += '</tbody></table></div>';
  document.getElementById('holdingsRiesgo').innerHTML = html4;

  // Combinaciones top
  const combos = {};
  fact.forEach(r => {
    const k = `${r.cli}||${r.cat}||${r.loc}`;
    if (!combos[k]) combos[k] = { cli:r.cli, cat:r.cat, loc:r.loc, eje:new Set(), vb:0, vn:0, ut:0, n:0 };
    combos[k].vb += r.vb; combos[k].vn += r.vn; combos[k].ut += r.ut; combos[k].n++;
    combos[k].eje.add(r.eje);
  });
  const top = Object.values(combos).map(c => ({...c, margen: c.vn?c.ut/c.vn*100:0})).sort((a,b)=>b.ut-a.ut).slice(0,15);
  let html5 = '<div class="table-wrap"><table><thead class="top"><tr><th>Cliente</th><th>Categoría</th><th>Localidad</th><th>Vendedor</th><th class="num">V. Bruta</th><th class="num">Utilidad</th><th class="num">Margen %</th><th class="num"># Líneas</th></tr></thead><tbody>';
  top.forEach(c => {
    const cls = c.margen>=30?'alc-good':c.margen>=15?'alc-warn':'alc-bad';
    html5 += `<tr><td><b>${c.cli}</b></td><td>${c.cat}</td><td>${c.loc}</td><td>${[...c.eje].join(', ')}</td><td class="num">${fmtMoney(c.vb)}</td><td class="num pos">${fmtMoney(c.ut)}</td><td class="num ${cls}">${fmtPct(c.margen)}</td><td class="num">${c.n}</td></tr>`;
  });
  html5 += '</tbody></table></div>';
  document.getElementById('combinacionesTop').innerHTML = html5;
}

function buildHeatmap(data, key, valKey, opts={}) {
  const grp = Engine.groupBy(data, key).sort((a,b)=>b[valKey]-a[valKey]);
  const top = grp.slice(0, opts.topN || 10);
  const matrix = top.map(g => Engine.monthly(data.filter(r=>r[key]===g.key), valKey));
  const max = Math.max(...matrix.flat(), 1);
  let html = `<div class="heatmap-row header" style="grid-template-columns:160px repeat(12, 1fr) 100px"><div></div>`;
  MESES.forEach(m => html += `<div>${m}</div>`);
  html += `<div>Total</div></div>`;
  top.forEach((g, i) => {
    const total = matrix[i].reduce((a,b)=>a+b,0);
    html += `<div class="heatmap-row" style="grid-template-columns:160px repeat(12, 1fr) 100px">`;
    html += `<div class="heatmap-label">${g.key}</div>`;
    matrix[i].forEach(v => {
      const intensity = v/max;
      const bg = `rgba(217,102,44,${(0.05+intensity*0.85).toFixed(2)})`;
      html += `<div class="heatmap-cell" style="background:${bg};color:${intensity>0.5?'#fff':'#24211c'}" title="${fmtMoney(v)}">${fmtMoneyShort(v)}</div>`;
    });
    html += `<div class="heatmap-cell" style="background:rgba(36,33,28,0.06);font-weight:700">${fmtMoneyShort(total)}</div></div>`;
  });
  return html;
}

// Heatmap COMPARATIVO 2026 vs 2025: dos filas por item, separadas visualmente con borde
function buildHeatmapComparativo(key, valKey, topN=10) {
  const data = Engine.applyFilters(filters, {ignoreYear:true});
  const fact = Engine.facturable(data);
  const fact2026 = fact.filter(r => r.anio === 2026);
  const fact2025 = fact.filter(r => r.anio === 2025);

  // Top items basado en 2026
  const grp26 = Engine.groupBy(fact2026, key).sort((a,b)=>b[valKey]-a[valKey]);
  const top = grp26.slice(0, topN);
  if (top.length === 0) return '<div style="padding:20px;color:var(--text-muted);text-align:center">Sin datos</div>';

  const m26 = top.map(g => Engine.monthly(fact2026.filter(r=>r[key]===g.key), valKey));
  const m25 = top.map(g => Engine.monthly(fact2025.filter(r=>r[key]===g.key), valKey));
  const max = Math.max(...m26.flat(), ...m25.flat(), 1);

  let html = '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:11px">';
  html += '<thead><tr style="background:var(--bg-card)"><th style="padding:6px;text-align:left;width:160px">'+(key==='eje'?'Vendedor':'Categoría')+'</th><th style="padding:6px;text-align:center">Año</th>';
  MESES.forEach(m => html += `<th style="padding:6px;text-align:center;font-weight:600">${m}</th>`);
  html += '<th style="padding:6px;text-align:center">Total</th><th style="padding:6px;text-align:center">YoY %</th></tr></thead><tbody>';

  top.forEach((g, i) => {
    const t26 = m26[i].reduce((a,b)=>a+b,0);
    const t25 = m25[i].reduce((a,b)=>a+b,0);
    const yoy = t25 ? (t26-t25)/t25*100 : null;
    const yoyCls = yoy === null ? '' : yoy>=0 ? 'pos' : 'neg';

    // Fila 2026
    html += `<tr style="border-top:2px solid var(--border)"><td rowspan="2" style="padding:8px;font-weight:700;border-right:1px solid var(--border);vertical-align:middle">${g.key}</td><td style="padding:6px;text-align:center;color:var(--primary);font-weight:700">2026</td>`;
    m26[i].forEach(v => {
      const intensity = v/max;
      const bg = `rgba(217,102,44,${(0.05+intensity*0.85).toFixed(2)})`;
      html += `<td style="padding:6px;text-align:center;background:${bg};color:${intensity>0.5?'#fff':'#24211c'}" title="${fmtMoney(v)}">${fmtMoneyShort(v)}</td>`;
    });
    html += `<td style="padding:6px;text-align:center;background:rgba(36,33,28,0.06);font-weight:700">${fmtMoneyShort(t26)}</td>`;
    html += `<td rowspan="2" style="padding:6px;text-align:center;vertical-align:middle" class="${yoyCls}"><b>${yoy===null?'—':(yoy>=0?'+':'')+fmtPct(yoy)}</b></td></tr>`;

    // Fila 2025
    html += `<tr><td style="padding:6px;text-align:center;color:var(--info);font-weight:700">2025</td>`;
    m25[i].forEach(v => {
      const intensity = v/max;
      const bg = `rgba(37,99,235,${(0.05+intensity*0.85).toFixed(2)})`;
      html += `<td style="padding:6px;text-align:center;background:${bg};color:${intensity>0.5?'#fff':'#24211c'}" title="${fmtMoney(v)}">${fmtMoneyShort(v)}</td>`;
    });
    html += `<td style="padding:6px;text-align:center;background:rgba(36,33,28,0.06);font-weight:700">${fmtMoneyShort(t25)}</td></tr>`;
  });
  html += '</tbody></table></div>';
  return html;
}

// =========================================================================
// POR PROVEEDOR
// =========================================================================
function renderProveedor() {
  const data = Engine.applyFilters(filters);
  const fact = Engine.facturable(data);
  const yActual = parseInt(filters.anio) || 2026;
  const yPrev = yActual - 1;

  // Datos del año anterior (mismo Status OPEN filter)
  const factPrev = Engine.facturable(Engine.applyFilters({...filters, anio: yPrev}, {ignoreYear:false}));

  // Helper: construir stats enriquecidas por proveedor para un set de records
  function statsByProv(records) {
    const map = {};
    records.filter(r => r.prov && r.prov !== 'Sin definir').forEach(r => {
      if (!map[r.prov]) map[r.prov] = {
        key: r.prov,
        cst: 0, vb: 0, vn: 0, ut: 0,
        n: 0,
        cats: new Set(), ejes: new Set(), locs: new Set(), clientes: new Set(), campanas: new Set(),
        meses: new Set(),
        firstMes: 13, lastMes: 0,
      };
      const o = map[r.prov];
      o.cst += r.cst; o.vb += r.vb; o.vn += r.vn; o.ut += r.ut;
      o.n += 1;
      if (r.cat) o.cats.add(r.cat);
      if (r.eje) o.ejes.add(r.eje);
      if (r.loc) o.locs.add(r.loc);
      if (r.cli) o.clientes.add(r.cli);
      if (r.cmp) o.campanas.add(r.cmp);
      if (r.mes >= 1 && r.mes <= 12) {
        o.meses.add(r.mes);
        if (r.mes < o.firstMes) o.firstMes = r.mes;
        if (r.mes > o.lastMes) o.lastMes = r.mes;
      }
    });
    return Object.values(map).map(o => ({
      ...o,
      margen: o.vn ? (o.ut/o.vn*100) : 0,
      costoPromCompra: o.n ? (o.cst / o.n) : 0,
      nCats: o.cats.size, nEjes: o.ejes.size, nLocs: o.locs.size, nClientes: o.clientes.size, nMeses: o.meses.size,
      nCampanas: o.campanas.size,
    }));
  }

  const provs = statsByProv(fact);
  const provsPrev = statsByProv(factPrev);
  const provPrevMap = Object.fromEntries(provsPrev.map(p => [p.key, p]));
  const provActualKeys = new Set(provs.map(p => p.key));

  // === SCORE 0-100 (Recurrencia 40% + Margen 30% + Monto 30%) ===
  // Normalizo por percentile rank para mitigar outliers
  function percentileRank(arr, getVal) {
    const sorted = arr.map(getVal).slice().sort((a,b) => a-b);
    const N = sorted.length;
    return v => {
      // # de valores <= v (excluyendo v actual)
      let count = 0;
      for (const x of sorted) if (x < v) count++;
      return N > 1 ? count / (N - 1) * 100 : 50;
    };
  }
  if (provs.length > 0) {
    const rankN = percentileRank(provs, p => p.n);
    const rankMg = percentileRank(provs, p => p.margen);
    const rankCst = percentileRank(provs, p => p.cst);
    provs.forEach(p => {
      const rN = rankN(p.n);
      const rM = rankMg(p.margen);
      const rC = rankCst(p.cst);
      p.score = (rN * 0.40 + rM * 0.30 + rC * 0.30);
      p.rankN = rN; p.rankMg = rM; p.rankCst = rC;
    });
  }

  // === Status: Nuevo / Existente / Perdido ===
  provs.forEach(p => {
    p.status = provPrevMap[p.key] ? 'Existente' : 'Nuevo';
  });
  const perdidos = provsPrev.filter(p => !provActualKeys.has(p.key)).map(p => ({...p, status: 'Perdido'}));

  // === KPIs principales ===
  const totCosto = provs.reduce((a,p)=>a+p.cst,0);
  const totCostoPrev = provsPrev.reduce((a,p)=>a+p.cst,0);
  const yoyCostoTotal = totCostoPrev ? ((totCosto-totCostoPrev)/totCostoPrev*100) : null;
  const top3Cost = [...provs].sort((a,b)=>b.cst-a.cst).slice(0,3).reduce((a,p)=>a+p.cst,0);
  const concentracion = totCosto ? top3Cost/totCosto*100 : 0;
  const nNuevos = provs.filter(p => p.status === 'Nuevo').length;
  const topScore = [...provs].sort((a,b)=>b.score-a.score)[0];
  const utilidadTotal = provs.reduce((a,p)=>a+p.ut,0);

  document.getElementById('provKpis').innerHTML = [
    {label:'Total Proveedores', val: fmtNum(provs.length), sub:`${yPrev}: ${provsPrev.length} · Existentes ${provs.length-nNuevos}/${provs.length}`, cls:'info'},
    {label:'Total Costo', val: fmtMoneyShort(totCosto), sub: yoyCostoTotal === null ? `${yPrev}: —` : `vs ${yPrev}: ${yoyCostoTotal>=0?'+':''}${fmtPct(yoyCostoTotal)} (${fmtMoneyShort(totCosto-totCostoPrev)})`, cls: yoyCostoTotal===null?'':yoyCostoTotal<0?'success':'warning'},
    {label:'Concentración Top 3', val: fmtPct(concentracion), sub:`Top 3 prov = ${fmtPct(concentracion)} del costo`, cls: concentracion>=60?'warning':'success'},
    {label:'Utilidad aportada', val: fmtMoneyShort(utilidadTotal), sub:`${nNuevos} nuevos · ${perdidos.length} perdidos`, cls:'success'},
    {label:'Top Score', val: topScore ? `${topScore.score.toFixed(0)}/100` : '—', sub: topScore?.key || '—', cls:'success'},
    {label:'Cobertura promedio', val: provs.length ? (provs.reduce((a,p)=>a+p.nMeses,0)/provs.length).toFixed(1) + ' meses' : '—', sub:'Meses con uso × proveedor', cls:'info'},
  ].map(k=>`<div class="kpi ${k.cls}"><div class="kpi-label">${k.label}</div><div class="kpi-value">${k.val}</div><div class="kpi-sub">${k.sub}</div></div>`).join('');

  // === INSIGHTS automáticos ===
  const insights = [];
  // 1. Proveedores que subieron mucho costo con menos uso
  provs.forEach(p => {
    const prev = provPrevMap[p.key];
    if (!prev) return;
    const yoyCsto = prev.cst ? (p.cst-prev.cst)/prev.cst*100 : 0;
    const yoyN = prev.n ? (p.n-prev.n)/prev.n*100 : 0;
    if (yoyCsto > 30 && yoyN < 0 && p.cst > 200000) {
      insights.push({cls:'warning', icon:'⚠️', txt:`<b>${p.key}</b> aumentó ${yoyCsto>=0?'+':''}${fmtPct(yoyCsto)} en costo (${fmtMoneyShort(p.cst-prev.cst)}) con <b>${yoyN.toFixed(0)}% menos compras</b>. Posible incremento de tarifa.`});
    }
    if (yoyCsto < -30 && p.cst > 200000) {
      insights.push({cls:'info', icon:'⬇️', txt:`<b>${p.key}</b> bajó ${fmtPct(yoyCsto)} en costo (${fmtMoneyShort(p.cst-prev.cst)}). ${yoyN.toFixed(0)}% menos compras. ¿Reducción de uso o cancelación?`});
    }
    // Cambio de margen
    if (Math.abs(p.margen - prev.margen) > 8 && p.cst > 300000) {
      const dir = p.margen > prev.margen ? '⬆️' : '⬇️';
      const cls = p.margen > prev.margen ? 'success' : 'warning';
      insights.push({cls, icon:dir, txt:`<b>${p.key}</b> cambió margen: ${fmtPct(prev.margen)} → ${fmtPct(p.margen)} (${(p.margen-prev.margen>=0?'+':'')}${fmtPct(p.margen-prev.margen)} pts)`});
    }
  });
  // 2. Concentración por proveedor único
  provs.forEach(p => {
    if (p.cst > 500000 && p.nEjes === 1) {
      const eje = [...p.ejes][0];
      insights.push({cls:'info', icon:'👤', txt:`<b>${p.key}</b> (${fmtMoneyShort(p.cst)}) lo usa SOLO <b>${eje}</b>. Riesgo si rota.`});
    }
  });
  // 3. Nuevos con monto relevante
  const nuevosTop = provs.filter(p=>p.status==='Nuevo' && p.cst > 500000).sort((a,b)=>b.cst-a.cst).slice(0,3);
  nuevosTop.forEach(p => {
    insights.push({cls:'success', icon:'🆕', txt:`Nuevo proveedor relevante: <b>${p.key}</b> · ${fmtMoneyShort(p.cst)} en ${p.n} compras · ${fmtPct(p.margen)} margen`});
  });
  // 4. Perdidos importantes
  perdidos.filter(p => p.cst > 500000).slice(0,3).forEach(p => {
    insights.push({cls:'danger', icon:'🔻', txt:`Proveedor perdido: <b>${p.key}</b> aportó ${fmtMoneyShort(p.cst)} en ${yPrev} con ${fmtPct(p.margen)} margen. ¿Sustituido?`});
  });
  // Pintar insights
  const insightsHTML = insights.length === 0
    ? '<div style="padding:14px;color:var(--text-muted);text-align:center">Sin movimientos relevantes detectados</div>'
    : insights.slice(0, 12).map(i => `<div class="alert ${i.cls}" style="margin-bottom:6px">${i.icon} ${i.txt}</div>`).join('');
  document.getElementById('provInsights').innerHTML = insightsHTML;

  // === TOP 15 por SCORE — tabla ===
  const top15Score = [...provs].sort((a,b)=>b.score-a.score).slice(0,15);
  let htmlTop = '<div class="table-wrap"><table class="table-default"><thead class="top"><tr>';
  htmlTop += '<th>#</th><th>Proveedor</th><th class="num">Score</th><th class="num">Recurrencia</th><th class="num">Margen</th><th class="num">Monto</th><th class="num"># Compras</th><th class="num">Costo</th><th class="num">Margen %</th><th class="num">Utilidad</th><th>Status</th>';
  htmlTop += '</tr></thead><tbody>';
  top15Score.forEach((p, i) => {
    const scoreCls = p.score >= 70 ? 'alc-good' : p.score >= 50 ? 'alc-warn' : 'alc-bad';
    const mgCls = p.margen >= 30 ? 'alc-good' : p.margen >= 15 ? 'alc-warn' : 'alc-bad';
    const statusPill = p.status === 'Nuevo' ? '<span class="pill success">🆕 Nuevo</span>' : '<span class="pill info">Existente</span>';
    htmlTop += `<tr>
      <td><b>${i+1}</b></td>
      <td><b>${p.key}</b></td>
      <td class="num ${scoreCls}"><b>${p.score.toFixed(0)}</b></td>
      <td class="num" style="color:var(--text-muted)" title="Percentile rank en # compras">${p.rankN.toFixed(0)}</td>
      <td class="num" style="color:var(--text-muted)" title="Percentile rank en margen %">${p.rankMg.toFixed(0)}</td>
      <td class="num" style="color:var(--text-muted)" title="Percentile rank en monto $">${p.rankCst.toFixed(0)}</td>
      <td class="num">${p.n}</td>
      <td class="num">${fmtMoney(p.cst)}</td>
      <td class="num ${mgCls}">${fmtPct(p.margen)}</td>
      <td class="num pos">${fmtMoney(p.ut)}</td>
      <td>${statusPill}</td>
    </tr>`;
  });
  htmlTop += '</tbody></table></div>';
  document.getElementById('tblProvTop15').innerHTML = htmlTop;

  // === Gráficas existentes ===
  const top15Cost = [...provs].sort((a,b)=>b.cst-a.cst).slice(0,15);
  drawChart('ch-provCosto', {
    type:'bar',
    data:{ labels: top15Cost.map(g=>g.key), datasets:[{label:'Costo', data:top15Cost.map(g=>g.cst), backgroundColor:'#b45309', borderRadius:4}]},
    options: barOpts({money:true, horizontal:true})
  });
  // Top 15 por margen con +3 compras
  const top15Margen = provs.filter(g=>g.n >= 3).sort((a,b)=>b.margen-a.margen).slice(0,15);
  drawChart('ch-provMargen', {
    type:'bar',
    data:{ labels: top15Margen.map(g=>g.key), datasets:[{label:'Margen %', data:top15Margen.map(g=>g.margen), backgroundColor: top15Margen.map(g=>g.margen>=30?'#1f9d6e':g.margen>=15?'#b45309':'#dc4747'), borderRadius:4}]},
    options: barOpts({pct:true, horizontal:true})
  });

  // === PARETO ===
  const provsSorted = [...provs].sort((a,b)=>b.cst-a.cst);
  const topPareto = provsSorted.slice(0, 30);
  const total = provsSorted.reduce((a,p)=>a+p.cst,0);
  let acum = 0;
  const acumPct = topPareto.map(p => { acum += p.cst; return total ? (acum/total*100) : 0; });
  const indivPct = topPareto.map(p => total ? (p.cst/total*100) : 0);
  drawChart('ch-provPareto', {
    type:'bar',
    data:{
      labels: topPareto.map(p=>p.key),
      datasets:[
        {type:'bar', label:'% del total', data: indivPct, backgroundColor:'#d9662c', borderRadius:4, yAxisID:'y'},
        {type:'line', label:'Acumulado %', data: acumPct, borderColor:'#1f9d6e', backgroundColor:'rgba(31,157,110,0.1)', tension:0.25, pointRadius:3, yAxisID:'y1', fill:false},
      ]
    },
    options: {
      maintainAspectRatio:false, responsive:true,
      interaction:{mode:'index', intersect:false},
      plugins:{ legend:{position:'top'}, tooltip:{ ...tooltipBase, callbacks:{ label:c => ' ' + c.dataset.label + ': ' + fmtPct(c.parsed.y) }}},
      scales:{
        y:{ position:'left', ticks:{callback:v=>v+'%', color:'#6b665a'}, grid:{color:'rgba(36,33,28,0.05)'}, border:{display:false}},
        y1:{ position:'right', min:0, max:100, ticks:{callback:v=>v+'%', color:'#1f9d6e'}, grid:{display:false}, border:{display:false}},
        x:{ ticks:{font:{size:9}, autoSkip:false, color:'#6b665a', maxRotation:60, minRotation:60}, grid:{display:false}, border:{display:false}}
      }
    }
  });

  // === HEATMAP mensual del top 12 proveedores ===
  const top12Heat = [...provs].sort((a,b)=>b.cst-a.cst).slice(0, 12);
  // Calculate mensual cost per proveedor
  const top12HeatData = top12Heat.map(p => {
    const monthly = Array(12).fill(0);
    fact.filter(r => r.prov === p.key).forEach(r => {
      if (r.mes >=1 && r.mes <= 12) monthly[r.mes-1] += r.cst;
    });
    const maxM = Math.max(...monthly, 1);
    return { p, monthly, maxM };
  });
  let htmlHeat = '<div class="table-wrap"><table style="font-size:11px"><thead class="top"><tr><th>Proveedor</th>';
  MESES.forEach(m => htmlHeat += `<th class="num">${m}</th>`);
  htmlHeat += '<th class="num">Total</th></tr></thead><tbody>';
  // Para colorear, usar max global entre todos
  const globalMax = Math.max(...top12HeatData.flatMap(d => d.monthly), 1);
  top12HeatData.forEach(({p, monthly}) => {
    const totRow = monthly.reduce((a,b)=>a+b,0);
    htmlHeat += `<tr><td><b>${p.key}</b></td>`;
    monthly.forEach(v => {
      const intensity = v / globalMax;
      const bg = v === 0 ? 'transparent' : `rgba(217,102,44,${0.08 + intensity * 0.72})`;
      const color = intensity > 0.5 ? '#fff' : 'var(--text-muted)';
      htmlHeat += `<td class="num" style="background:${bg};color:${color}" title="${p.key} · ${MESES_FULL[monthly.indexOf(v)]}: ${fmtMoney(v)}">${v === 0 ? '—' : fmtMoneyShort(v)}</td>`;
    });
    htmlHeat += `<td class="num"><b>${fmtMoneyShort(totRow)}</b></td></tr>`;
  });
  htmlHeat += '</tbody></table></div>';
  document.getElementById('provHeatmap').innerHTML = htmlHeat;

  // === DETALLE enriquecido ===
  let html = '<div class="table-wrap"><table class="table-default"><thead class="top"><tr>';
  html += '<th>Proveedor</th><th>Status</th><th class="num">Score</th><th class="num"># Compras</th><th class="num">Costo '+yActual+'</th><th class="num">Costo '+yPrev+'</th><th class="num">Δ%</th><th class="num">Venta</th><th class="num">Utilidad</th><th class="num">Margen %</th><th class="num">$ Prom/compra</th><th class="num"># Categorías</th><th class="num"># Vendedores</th><th class="num"># Clientes</th><th class="num"># Localidades</th><th class="num">Meses uso</th><th class="num">% del total</th>';
  html += '</tr></thead><tbody>';
  const provsSortedAll = [...provs].sort((a,b)=>b.cst-a.cst);
  provsSortedAll.forEach(p => {
    const prev = provPrevMap[p.key];
    const costPrev = prev?.cst || 0;
    const yoy = costPrev ? (p.cst-costPrev)/costPrev*100 : null;
    const yoyCls = yoy === null ? '' : yoy>=0 ? 'pos' : 'neg';
    const yoyTxt = yoy === null ? '🆕' : (yoy>=0?'+':'')+fmtPct(yoy);
    const cls = p.margen>=30?'alc-good':p.margen>=15?'alc-warn':'alc-bad';
    const scoreCls = p.score >= 70 ? 'alc-good' : p.score >= 50 ? 'alc-warn' : 'alc-bad';
    const status = p.status === 'Nuevo' ? '<span class="pill success">🆕 Nuevo</span>' : '<span class="pill info">Existente</span>';
    const pctTotal = totCosto ? p.cst/totCosto*100 : 0;
    const tooltip = `${p.key}\n${yActual}: ${fmtMoney(p.cst)} (${p.n} compras)\n${yPrev}: ${fmtMoney(costPrev)} (${prev?.n||0} compras)\nMargen ${yActual}: ${fmtPct(p.margen)} · Margen ${yPrev}: ${prev?fmtPct(prev.margen):'—'}`;
    html += `<tr title="${tooltip}">
      <td><b>${p.key}</b></td>
      <td>${status}</td>
      <td class="num ${scoreCls}"><b>${p.score.toFixed(0)}</b></td>
      <td class="num">${p.n}</td>
      <td class="num"><b>${fmtMoney(p.cst)}</b></td>
      <td class="num" style="color:var(--text-muted)">${fmtMoney(costPrev)}</td>
      <td class="num ${yoyCls}">${yoyTxt}</td>
      <td class="num">${fmtMoney(p.vb)}</td>
      <td class="num pos">${fmtMoney(p.ut)}</td>
      <td class="num ${cls}">${fmtPct(p.margen)}</td>
      <td class="num">${fmtMoney(p.costoPromCompra)}</td>
      <td class="num">${p.nCats}</td>
      <td class="num">${p.nEjes}</td>
      <td class="num">${p.nClientes}</td>
      <td class="num">${p.nLocs}</td>
      <td class="num">${p.nMeses}/12</td>
      <td class="num">${fmtPct(pctTotal)}</td>
    </tr>`;
  });
  html += '</tbody></table></div>';
  document.getElementById('tblProveedores').innerHTML = html;

  // === Nuevos ===
  const nuevos = provs.filter(p => p.status === 'Nuevo').sort((a,b)=>b.cst-a.cst);
  let htmlNuevos = '<div class="table-wrap"><table><thead class="top"><tr><th>Proveedor</th><th class="num">Score</th><th class="num">Costo</th><th class="num">Venta</th><th class="num">Margen %</th><th class="num"># Compras</th><th class="num"># Vendedores</th></tr></thead><tbody>';
  if (nuevos.length === 0) htmlNuevos += '<tr><td colspan="7" style="text-align:center;color:var(--text-muted);padding:12px">Sin proveedores nuevos</td></tr>';
  else nuevos.forEach(p => {
    const cls = p.margen>=30?'alc-good':p.margen>=15?'alc-warn':'alc-bad';
    const scoreCls = p.score >= 70 ? 'alc-good' : p.score >= 50 ? 'alc-warn' : 'alc-bad';
    htmlNuevos += `<tr><td><b>${p.key}</b></td><td class="num ${scoreCls}">${p.score.toFixed(0)}</td><td class="num">${fmtMoney(p.cst)}</td><td class="num">${fmtMoney(p.vb)}</td><td class="num ${cls}">${fmtPct(p.margen)}</td><td class="num">${p.n}</td><td class="num">${p.nEjes}</td></tr>`;
  });
  htmlNuevos += '</tbody></table></div>';
  document.getElementById('tblProvNuevos').innerHTML = htmlNuevos;

  // === Perdidos ===
  const perdidosSorted = perdidos.sort((a,b)=>b.cst-a.cst);
  let htmlPerdidos = '<div class="table-wrap"><table><thead class="top"><tr><th>Proveedor</th><th class="num">Costo '+yPrev+'</th><th class="num">Venta '+yPrev+'</th><th class="num">Margen %</th><th class="num"># Compras</th><th class="num"># Vendedores</th></tr></thead><tbody>';
  if (perdidosSorted.length === 0) htmlPerdidos += '<tr><td colspan="6" style="text-align:center;color:var(--text-muted);padding:12px">Sin proveedores perdidos</td></tr>';
  else perdidosSorted.forEach(p => {
    const cls = p.margen>=30?'alc-good':p.margen>=15?'alc-warn':'alc-bad';
    htmlPerdidos += `<tr><td><b>${p.key}</b></td><td class="num">${fmtMoney(p.cst)}</td><td class="num">${fmtMoney(p.vb)}</td><td class="num ${cls}">${fmtPct(p.margen)}</td><td class="num">${p.n}</td><td class="num">${p.nEjes}</td></tr>`;
  });
  htmlPerdidos += '</tbody></table></div>';
  document.getElementById('tblProvPerdidos').innerHTML = htmlPerdidos;

  // === Cruce Proveedor × Categoría con hover 2025 ===
  const top15Provs = [...provs].sort((a,b)=>b.cst-a.cst).slice(0,15);
  const allCats = [...new Set(fact.map(r=>r.cat))].filter(c=>c && c!=='Otros');
  const catRank = allCats.map(c=>({c, v:fact.filter(r=>r.cat===c).reduce((a,r)=>a+r.cst,0)})).sort((a,b)=>b.v-a.v);
  const topCats = catRank.slice(0,6).map(o=>o.c);

  let htmlPC = '<div class="table-wrap"><table><thead class="top"><tr><th>Proveedor</th><th class="num">Total Costo</th>';
  topCats.forEach(c => htmlPC += `<th class="num">${c}</th>`);
  htmlPC += '<th class="num">Otras</th></tr></thead><tbody>';
  top15Provs.forEach(p => {
    const recs = fact.filter(r=>r.prov===p.key);
    const recsP = factPrev.filter(r=>r.prov===p.key);
    htmlPC += `<tr><td><b>${p.key}</b></td><td class="num">${fmtMoney(p.cst)}</td>`;
    let inTop = 0;
    topCats.forEach(c => {
      const v = recs.filter(r=>r.cat===c).reduce((a,r)=>a+r.cst,0);
      const vP = recsP.filter(r=>r.cat===c).reduce((a,r)=>a+r.cst,0);
      inTop += v;
      const diff = v - vP;
      const tooltip = `${p.key} × ${c}\n${yActual}: ${fmtMoney(v)}\n${yPrev}: ${fmtMoney(vP)}\nDiff: ${diff>=0?'+':''}${fmtMoney(diff)}`;
      const cls = v === 0 ? 'style="color:var(--text-dim)"' : '';
      htmlPC += `<td class="num" ${cls} title="${tooltip}">${v>0?fmtMoney(v):'—'}</td>`;
    });
    const otras = p.cst - inTop;
    const otrasP = recsP.filter(r=>!topCats.includes(r.cat)).reduce((a,r)=>a+r.cst,0);
    const otrasTip = `Otras categorías\n${yActual}: ${fmtMoney(otras)}\n${yPrev}: ${fmtMoney(otrasP)}\nDiff: ${(otras-otrasP)>=0?'+':''}${fmtMoney(otras-otrasP)}`;
    htmlPC += `<td class="num" title="${otrasTip}">${otras>0?fmtMoney(otras):'—'}</td></tr>`;
  });
  htmlPC += '</tbody></table></div>';
  document.getElementById('tblProvCategoria').innerHTML = htmlPC;

  // === Cruce Proveedor × Vendedor con hover 2025 ===
  const top12Provs = [...provs].sort((a,b)=>b.cst-a.cst).slice(0,12);
  const ejesActual = [...new Set(fact.map(r=>r.eje))].sort();
  const ejesByVenta = ejesActual.map(e=>({e, v:fact.filter(r=>r.eje===e).reduce((a,r)=>a+r.vb,0)})).sort((a,b)=>b.v-a.v);
  const topEjes = ejesByVenta.slice(0,8).map(o=>o.e);

  let htmlPV = '<div class="table-wrap"><table><thead class="top"><tr><th>Proveedor</th><th class="num">Total Costo</th>';
  topEjes.forEach(e => htmlPV += `<th class="num">${e}</th>`);
  htmlPV += '</tr></thead><tbody>';
  top12Provs.forEach(p => {
    const recs = fact.filter(r=>r.prov===p.key);
    const recsP = factPrev.filter(r=>r.prov===p.key);
    htmlPV += `<tr><td><b>${p.key}</b></td><td class="num">${fmtMoney(p.cst)}</td>`;
    topEjes.forEach(e => {
      const v = recs.filter(r=>r.eje===e).reduce((a,r)=>a+r.cst,0);
      const vP = recsP.filter(r=>r.eje===e).reduce((a,r)=>a+r.cst,0);
      const diff = v - vP;
      const tooltip = `${p.key} × ${e}\n${yActual}: ${fmtMoney(v)}\n${yPrev}: ${fmtMoney(vP)}\nDiff: ${diff>=0?'+':''}${fmtMoney(diff)}`;
      const cls = v === 0 ? 'style="color:var(--text-dim)"' : '';
      htmlPV += `<td class="num" ${cls} title="${tooltip}">${v>0?fmtMoney(v):'—'}</td>`;
    });
    htmlPV += `</tr>`;
  });
  htmlPV += '</tbody></table></div>';
  document.getElementById('tblProvVendedor').innerHTML = htmlPV;
}

// =========================================================================
// MEDIOS / CATEGORÍAS
// =========================================================================
function renderMedios() {
  const data = Engine.applyFilters(filters);
  const fact = Engine.facturable(data);
  const grpCat = Engine.groupBy(fact, 'cat').sort((a,b)=>b.vb-a.vb);

  drawChart('ch-medBruta', {
    type:'bar',
    data:{ labels: grpCat.map(g=>g.key), datasets:[{label:'V. Bruta', data:grpCat.map(g=>g.vb), backgroundColor:'#d9662c', borderRadius:4}]},
    options: barOpts({money:true, horizontal:true})
  });
  const margenes = grpCat.filter(g=>g.vn>0).sort((a,b)=>b.margen-a.margen);
  drawChart('ch-medMargen', {
    type:'bar',
    data:{ labels: margenes.map(g=>g.key), datasets:[{label:'Margen %', data:margenes.map(g=>g.margen), backgroundColor:margenes.map(g=>g.margen>=30?'#1f9d6e':g.margen>=15?'#b45309':'#dc4747'), borderRadius:4}]},
    options: barOpts({pct:true, horizontal:true})
  });

  const tP = Engine.totalize(fact.filter(r=>r.tp==='Propio'));
  const tA = Engine.totalize(fact.filter(r=>r.tp==='Arrendado'));
  drawChart('ch-propio', { type:'doughnut', data:{ labels:['Propio','Arrendado'], datasets:[{data:[tP.vb, tA.vb], backgroundColor:['#1f9d6e','#b45309'], borderWidth:0}]}, options: pieOpts() });
  drawChart('ch-propioMargen', { type:'bar', data:{ labels:['Propio','Arrendado'], datasets:[{label:'Margen %', data:[tP.margen, tA.margen], backgroundColor:['#1f9d6e','#b45309'], borderRadius:4}]}, options: barOpts({pct:true}) });

  const totalV = grpCat.reduce((a,g)=>a+g.vb,0);
  const prevMapMed = groupByPrev('cat');
  let html = '<div class="table-wrap"><table class="table-default"><thead class="top"><tr><th>Categoría</th><th class="num">V. Bruta</th><th class="num">V. Neta</th><th class="num">Utilidad</th><th class="num">Margen %</th><th class="num"># Líneas</th><th class="num">% Total</th><th class="num">V. Bruta Año Ant.</th><th class="num">YoY %</th></tr></thead><tbody>';
  grpCat.forEach(g => {
    const p = totalV?g.vb/totalV*100:0;
    const cls = g.margen>=30?'alc-good':g.margen>=15?'alc-warn':'alc-bad';
    html += `<tr><td><b>${g.key}</b></td><td class="num">${fmtMoney(g.vb)}</td><td class="num">${fmtMoney(g.vn)}</td><td class="num pos">${fmtMoney(g.ut)}</td><td class="num ${cls}">${fmtPct(g.margen)}</td><td class="num">${g.n}</td><td class="num">${fmtPct(p)}</td>${renderYoYCell(g, prevMapMed[g.key], 'vb')}</tr>`;
  });
  html += '</tbody></table></div>';
  document.getElementById('tblMedios').innerHTML = html;

  // Subcategorías: por Soporte (2026) o Medio (2025)
  const subKey = r => `${r.cat} · ${r.anio===2026 ? (r.sop || r.med || '—') : (r.med || '—')}`;
  const subMap = {};
  fact.forEach(r => {
    const k = subKey(r);
    if (!subMap[k]) subMap[k] = { key:k, cat:r.cat, sub:r.anio===2026?(r.sop||r.med||'—'):(r.med||'—'), vb:0, vn:0, ut:0, n:0 };
    subMap[k].vb += r.vb; subMap[k].vn += r.vn; subMap[k].ut += r.ut; subMap[k].n++;
  });
  const sub = Object.values(subMap).map(s=>({...s, margen:s.vn?s.ut/s.vn*100:0})).sort((a,b)=>b.vb-a.vb);
  let html2 = '<div class="table-wrap"><table class="table-default"><thead class="top"><tr><th>Categoría</th><th>Subcategoría</th><th class="num">V. Bruta</th><th class="num">V. Neta</th><th class="num">Utilidad</th><th class="num">Margen %</th><th class="num"># Líneas</th></tr></thead><tbody>';
  sub.forEach(s => {
    const cls = s.margen>=30?'alc-good':s.margen>=15?'alc-warn':'alc-bad';
    html2 += `<tr><td>${s.cat}</td><td><b>${s.sub}</b></td><td class="num">${fmtMoney(s.vb)}</td><td class="num">${fmtMoney(s.vn)}</td><td class="num pos">${fmtMoney(s.ut)}</td><td class="num ${cls}">${fmtPct(s.margen)}</td><td class="num">${s.n}</td></tr>`;
  });
  html2 += '</tbody></table></div>';
  document.getElementById('tblSubMedios').innerHTML = html2;
}

// =========================================================================
// CLIENTES
// =========================================================================
function renderClientes() {
  const data = Engine.applyFilters(filters);
  const fact = Engine.facturable(data);
  const grp = Engine.groupBy(fact, 'cli').sort((a,b)=>b.vb-a.vb);
  const top20 = grp.slice(0,20);
  drawChart('ch-cliBruta', { type:'bar', data:{ labels: top20.map(g=>g.key), datasets:[{label:'V. Bruta', data:top20.map(g=>g.vb), backgroundColor:'#d9662c', borderRadius:4}]}, options: barOpts({money:true, horizontal:true}) });

  const total = grp.reduce((a,g)=>a+g.vb,0);
  let acum = 0;
  const par = grp.slice(0,30).map(g => { acum += g.vb; return { key:g.key, p: total?g.vb/total*100:0, a: total?acum/total*100:0 }; });
  drawChart('ch-pareto', {
    type:'bar',
    data:{ labels: par.map(p=>p.key),
      datasets:[
        {label:'% del total', data:par.map(p=>p.p), backgroundColor:'#d9662c', borderRadius:4, yAxisID:'y'},
        {label:'Acumulado %', data:par.map(p=>p.a), type:'line', borderColor:'#1f9d6e', backgroundColor:'rgba(31,157,110,0.1)', tension:0.2, yAxisID:'y1', fill:false}
      ]},
    options:{ maintainAspectRatio:false, responsive:true, scales:{ y:{ position:'left', ticks:{callback:v=>v+'%'}, grid:{color:'rgba(36,33,28,0.06)'}}, y1:{ position:'right', max:100, ticks:{callback:v=>v+'%'}, grid:{display:false}}, x:{ ticks:{maxRotation:45, minRotation:45, autoSkip:false, font:{size:9}}, grid:{display:false}}}, plugins:{ legend:{position:'top'}}}
  });

  acum = 0;
  const prevMapCli = groupByPrev('cli');
  let html = '<div class="table-wrap"><table class="table-default"><thead class="top"><tr><th>Cliente</th><th>Holding</th><th class="num">V. Bruta</th><th class="num">V. Neta</th><th class="num">Utilidad</th><th class="num">Margen %</th><th class="num"># Líneas</th><th class="num">% Total</th><th class="num">Acum %</th><th class="num">V. Bruta Año Ant.</th><th class="num">YoY %</th></tr></thead><tbody>';
  grp.forEach(g => {
    acum += g.vb;
    const p = total?g.vb/total*100:0; const ac = total?acum/total*100:0;
    const cls = g.margen>=30?'alc-good':g.margen>=15?'alc-warn':'alc-bad';
    const hol = mainHolding(fact, g.key);
    html += `<tr><td><b>${g.key}</b></td><td>${hol}</td><td class="num">${fmtMoney(g.vb)}</td><td class="num">${fmtMoney(g.vn)}</td><td class="num pos">${fmtMoney(g.ut)}</td><td class="num ${cls}">${fmtPct(g.margen)}</td><td class="num">${g.n}</td><td class="num">${fmtPct(p)}</td><td class="num">${fmtPct(ac)}</td>${renderYoYCell(g, prevMapCli[g.key], 'vb')}</tr>`;
  });
  html += '</tbody></table></div>';
  document.getElementById('tblClientes').innerHTML = html;
}
function mainHolding(data, cli) {
  const map = {};
  data.filter(r=>r.cli===cli).forEach(r=>{ map[r.hol] = (map[r.hol]||0) + r.vb; });
  let best=''; let bv=0;
  for (const k in map) if (map[k]>bv) { bv=map[k]; best=k; }
  return best;
}

// =========================================================================
// HOLDINGS
// =========================================================================
function renderHoldings() {
  const data = Engine.applyFilters(filters);
  const fact = Engine.facturable(data);
  const grp = Engine.groupBy(fact, 'hol').sort((a,b)=>b.vb-a.vb);
  const top = grp.slice(0,15);
  drawChart('ch-holNeta', { type:'bar', data:{ labels: top.map(g=>g.key), datasets:[{label:'V. Neta', data:top.map(g=>g.vn), backgroundColor:'#2563eb', borderRadius:4}]}, options: barOpts({money:true, horizontal:true}) });
  drawChart('ch-holMargen', { type:'bar', data:{ labels: top.map(g=>g.key), datasets:[{label:'Margen %', data:top.map(g=>g.margen), backgroundColor: top.map(g=>g.margen>=30?'#1f9d6e':g.margen>=15?'#b45309':'#dc4747'), borderRadius:4}]}, options: barOpts({pct:true, horizontal:true}) });

  const prevMapHol = groupByPrev('hol');
  let html = '<div class="table-wrap"><table class="table-default"><thead class="top"><tr><th>Holding</th><th class="num">V. Bruta</th><th class="num">V. Neta</th><th class="num">Utilidad</th><th class="num">Margen %</th><th class="num"># Clientes</th><th class="num"># Líneas</th><th class="num">V. Bruta Año Ant.</th><th class="num">YoY %</th></tr></thead><tbody>';
  grp.forEach(g => {
    const cls = g.margen>=30?'alc-good':g.margen>=15?'alc-warn':'alc-bad';
    html += `<tr><td><b>${g.key}</b></td><td class="num">${fmtMoney(g.vb)}</td><td class="num">${fmtMoney(g.vn)}</td><td class="num pos">${fmtMoney(g.ut)}</td><td class="num ${cls}">${fmtPct(g.margen)}</td><td class="num">${g.nClientes}</td><td class="num">${g.n}</td>${renderYoYCell(g, prevMapHol[g.key], 'vb')}</tr>`;
  });
  html += '</tbody></table></div>';
  document.getElementById('tblHoldings').innerHTML = html;

  const ag = Engine.groupBy(fact, 'age').sort((a,b)=>b.vb-a.vb).slice(0,30);
  const prevMapAg = groupByPrev('age');
  let html2 = '<div class="table-wrap"><table class="table-default"><thead class="top"><tr><th>Agencia</th><th class="num">V. Bruta</th><th class="num">V. Neta</th><th class="num">Utilidad</th><th class="num">Margen %</th><th class="num"># Clientes</th><th class="num">V. Bruta Año Ant.</th><th class="num">YoY %</th></tr></thead><tbody>';
  ag.forEach(g => {
    const cls = g.margen>=30?'alc-good':g.margen>=15?'alc-warn':'alc-bad';
    html2 += `<tr><td><b>${g.key}</b></td><td class="num">${fmtMoney(g.vb)}</td><td class="num">${fmtMoney(g.vn)}</td><td class="num pos">${fmtMoney(g.ut)}</td><td class="num ${cls}">${fmtPct(g.margen)}</td><td class="num">${g.nClientes}</td>${renderYoYCell(g, prevMapAg[g.key], 'vb')}</tr>`;
  });
  html2 += '</tbody></table></div>';
  document.getElementById('tblAgencias').innerHTML = html2;
}

// =========================================================================
// GEOGRAFIA
// =========================================================================
function renderGeografia() {
  const data = Engine.applyFilters(filters);
  const fact = Engine.facturable(data);
  const grpE = Engine.groupBy(fact, 'est').sort((a,b)=>b.vb-a.vb).slice(0,15);
  drawChart('ch-estado', { type:'bar', data:{ labels: grpE.map(g=>g.key), datasets:[{label:'V. Bruta', data:grpE.map(g=>g.vb), backgroundColor:'#d9662c', borderRadius:4}]}, options: barOpts({money:true, horizontal:true}) });
  const grpL = Engine.groupBy(fact, 'loc').sort((a,b)=>b.vb-a.vb).slice(0,15);
  drawChart('ch-localidad', { type:'bar', data:{ labels: grpL.map(g=>g.key), datasets:[{label:'V. Bruta', data:grpL.map(g=>g.vb), backgroundColor:'#2563eb', borderRadius:4}]}, options: barOpts({money:true, horizontal:true}) });

  const totV = fact.reduce((a,r)=>a+r.vb,0);
  const prevMapEst = groupByPrev('est');
  let html = '<div class="table-wrap"><table class="table-default"><thead class="top"><tr><th>Estado</th><th class="num">V. Bruta</th><th class="num">V. Neta</th><th class="num">Utilidad</th><th class="num">Margen %</th><th class="num"># Localidades</th><th class="num"># Líneas</th><th class="num">% Total</th><th class="num">V. Bruta Año Ant.</th><th class="num">YoY %</th></tr></thead><tbody>';
  Engine.groupBy(fact,'est').sort((a,b)=>b.vb-a.vb).forEach(g => {
    const cls = g.margen>=30?'alc-good':g.margen>=15?'alc-warn':'alc-bad';
    html += `<tr><td><b>${g.key}</b></td><td class="num">${fmtMoney(g.vb)}</td><td class="num">${fmtMoney(g.vn)}</td><td class="num pos">${fmtMoney(g.ut)}</td><td class="num ${cls}">${fmtPct(g.margen)}</td><td class="num">${g.nLoc}</td><td class="num">${g.n}</td><td class="num">${fmtPct(totV?g.vb/totV*100:0)}</td>${renderYoYCell(g, prevMapEst[g.key], 'vb')}</tr>`;
  });
  html += '</tbody></table></div>';
  document.getElementById('tblEstados').innerHTML = html;

  const prevMapLoc = groupByPrev('loc');
  let html2 = '<div class="table-wrap"><table class="table-default"><thead class="top"><tr><th>Localidad</th><th class="num">V. Bruta</th><th class="num">V. Neta</th><th class="num">Utilidad</th><th class="num">Margen %</th><th class="num"># Líneas</th><th class="num">V. Bruta Año Ant.</th><th class="num">YoY %</th></tr></thead><tbody>';
  Engine.groupBy(fact,'loc').sort((a,b)=>b.vb-a.vb).slice(0,30).forEach(g => {
    const cls = g.margen>=30?'alc-good':g.margen>=15?'alc-warn':'alc-bad';
    html2 += `<tr><td><b>${g.key}</b></td><td class="num">${fmtMoney(g.vb)}</td><td class="num">${fmtMoney(g.vn)}</td><td class="num pos">${fmtMoney(g.ut)}</td><td class="num ${cls}">${fmtPct(g.margen)}</td><td class="num">${g.n}</td>${renderYoYCell(g, prevMapLoc[g.key], 'vb')}</tr>`;
  });
  html2 += '</tbody></table></div>';
  document.getElementById('tblLocalidades').innerHTML = html2;
}

// =========================================================================
// ESTACIONALIDAD
// =========================================================================
function renderEstacionalidad() {
  const data = Engine.applyFilters(filters, {ignoreYear:true});
  const fact = Engine.facturable(data);
  const yrs = Engine.yearsAvailable;
  const matrix = yrs.map(y => Engine.monthly(fact.filter(r=>r.anio===y), 'vn'));
  const max = Math.max(...matrix.flat(), 1);
  let html = `<div class="heatmap-row header" style="grid-template-columns:120px repeat(12, 1fr) 110px"><div></div>`;
  MESES.forEach(m=>html+=`<div>${m}</div>`);
  html += `<div>Total</div></div>`;
  yrs.forEach((y, i) => {
    const total = matrix[i].reduce((a,b)=>a+b,0);
    html += `<div class="heatmap-row" style="grid-template-columns:120px repeat(12, 1fr) 110px">`;
    html += `<div class="heatmap-label">${y}</div>`;
    matrix[i].forEach(v => {
      const intensity = v/max;
      const bg = `rgba(217,102,44,${(0.05+intensity*0.85).toFixed(2)})`;
      html += `<div class="heatmap-cell" style="background:${bg};color:${intensity>0.5?'#fff':'#24211c'}" title="${fmtMoney(v)}">${fmtMoneyShort(v)}</div>`;
    });
    html += `<div class="heatmap-cell" style="background:rgba(36,33,28,0.06);font-weight:700">${fmtMoneyShort(total)}</div></div>`;
  });
  document.getElementById('heatmapAnioMes').innerHTML = html;

  const datasets = yrs.map((y, i) => ({ label: String(y), data: matrix[i], borderColor: PALETTE[i], backgroundColor: PALETTE[i]+'33', tension:0.3, fill: false }));
  drawChart('ch-stTend', { type:'line', data:{ labels: MESES, datasets }, options: lineOpts({money:true}) });

  const latest = matrix[matrix.length-1];
  const variacion = latest.map((v,i,arr)=> i===0||arr[i-1]===0?0:((v-arr[i-1])/arr[i-1]*100));
  drawChart('ch-stVar', { type:'bar', data:{ labels: MESES, datasets:[{label:`Δ% mes anterior (${yrs[yrs.length-1]})`, data:variacion, backgroundColor: variacion.map(v=>v>=0?'#1f9d6e':'#dc4747'), borderRadius:4}]}, options: barOpts({pct:true}) });
}

// =========================================================================
// RENTABILIDAD
// =========================================================================
function renderRentabilidad() {
  const data = Engine.applyFilters(filters);
  const fact = Engine.facturable(data);
  const t = Engine.totalize(fact);
  const margenes = fact.filter(r=>r.vn>0).map(r=>r.ut/r.vn*100);
  const negs = fact.filter(r=>r.vn>0 && r.ut/r.vn*100 < 0);
  const bajos = fact.filter(r=>r.vn>0 && r.ut/r.vn*100 >= 0 && r.ut/r.vn*100 < 15);
  const buenos = fact.filter(r=>r.vn>0 && r.ut/r.vn*100 >= 30);

  document.getElementById('rentKpis').innerHTML = [
    {label:'Margen Promedio', val: fmtPct(t.margen), sub:`Sobre venta neta`, cls:'success'},
    {label:'Líneas con margen ≥ 30%', val: fmtNum(buenos.length), sub: fmtMoneyShort(buenos.reduce((a,r)=>a+r.vb,0)), cls:'success'},
    {label:'Líneas con margen 0-15%', val: fmtNum(bajos.length), sub: fmtMoneyShort(bajos.reduce((a,r)=>a+r.vb,0)), cls:'warning'},
    {label:'Líneas con margen negativo', val: fmtNum(negs.length), sub: fmtMoneyShort(negs.reduce((a,r)=>a+r.vb,0)), cls:'danger'},
  ].map(k=>`<div class="kpi ${k.cls}"><div class="kpi-label">${k.label}</div><div class="kpi-value">${k.val}</div><div class="kpi-sub">${k.sub}</div></div>`).join('');

  const buckets = ['<0%','0-10%','10-20%','20-30%','30-40%','40-50%','>50%'];
  const counts = [0,0,0,0,0,0,0];
  margenes.forEach(m => {
    if (m<0) counts[0]++;
    else if (m<10) counts[1]++;
    else if (m<20) counts[2]++;
    else if (m<30) counts[3]++;
    else if (m<40) counts[4]++;
    else if (m<50) counts[5]++;
    else counts[6]++;
  });
  drawChart('ch-rentDist', { type:'bar', data:{ labels: buckets, datasets:[{label:'# líneas', data:counts, backgroundColor:['#dc4747','#c2410c','#b45309','#ca8a04','#65a30d','#16a34a','#1f9d6e'], borderRadius:4}]}, options: barOpts({}) });

  // Consolidar líneas duplicadas: agrupar por cliente+vendedor+categoría+localidad
  const conRiesgoRaw = [...negs, ...bajos];
  const consolidated = {};
  conRiesgoRaw.forEach(r => {
    const k = `${r.cli}||${r.eje}||${r.cat}||${r.loc}`;
    if (!consolidated[k]) consolidated[k] = { cli: r.cli, eje: r.eje, cat: r.cat, loc: r.loc, vb:0, vn:0, ut:0, n:0 };
    consolidated[k].vb += r.vb;
    consolidated[k].vn += r.vn;
    consolidated[k].ut += r.ut;
    consolidated[k].n++;
  });
  const conRiesgo = Object.values(consolidated)
    .map(o => ({...o, mg: o.vn ? o.ut/o.vn*100 : 0}))
    .sort((a,b)=>b.vb-a.vb)
    .slice(0, 80);

  // Actualizar header para incluir # líneas
  const tbl = document.querySelector('#sec-rentabilidad #tbl-bajos').closest('table');
  if (tbl) tbl.querySelector('thead tr').innerHTML = '<th>Cliente</th><th>Vendedor</th><th>Categoría</th><th>Localidad</th><th class="num">Venta</th><th class="num">Margen %</th><th class="num"># Líneas</th>';

  document.getElementById('tbl-bajos').innerHTML = conRiesgo.length ? conRiesgo.map(r => {
    return `<tr><td>${r.cli}</td><td>${r.eje}</td><td>${r.cat}</td><td>${r.loc}</td><td class="num">${fmtMoney(r.vb)}</td><td class="num ${r.mg<0?'neg':'alc-warn'}">${fmtPct(r.mg)}</td><td class="num">${r.n}</td></tr>`;
  }).join('') : '<tr><td colspan="7" style="text-align:center;color:var(--text-muted)">Sin líneas con margen bajo</td></tr>';
}

// =========================================================================
// DETALLE
// =========================================================================
function renderDetalle() {
  let data = Engine.applyFilters(filters);
  if (detSortKey) {
    data = [...data].sort((a,b) => {
      let av = a[detSortKey], bv = b[detSortKey];
      if (detSortKey === 'mg') { av = a.vn?a.ut/a.vn*100:0; bv = b.vn?b.ut/b.vn*100:0; }
      if (typeof av === 'number' || typeof bv === 'number') return ((av||0) - (bv||0)) * detSortDir;
      return String(av).localeCompare(String(bv), 'es') * detSortDir;
    });
  }
  document.getElementById('detCount').textContent = data.length.toLocaleString('es-MX') + ' líneas';
  const limit = data.slice(0, 1500);
  let html = limit.map(r => {
    const mg = r.vn ? r.ut/r.vn*100 : 0;
    const cls = mg>=30?'alc-good':mg>=15?'alc-warn':mg<0?'alc-bad':'';
    return `<tr><td>${r.anio}</td><td>${MESES[r.mes-1]||''}</td><td>${r.emp}</td><td>${r.eje}</td><td>${r.cli}</td><td>${r.hol}</td><td>${r.age}</td><td>${r.est}</td><td>${r.loc}</td><td><span class="pill info">${r.cat}</span></td><td>${r.med}</td><td>${r.sop}</td><td><span class="pill ${r.tp==='Propio'?'success':'info'}">${r.tp}</span></td><td class="num">${fmtMoney(r.vb)}</td><td class="num">${fmtMoney(r.vn)}</td><td class="num pos">${fmtMoney(r.ut)}</td><td class="num ${cls}">${fmtPct(mg)}</td></tr>`;
  }).join('');
  if (data.length > 1500) html += `<tr><td colspan="17" style="text-align:center;color:var(--text-muted);padding:16px">Mostrando 1,500 de ${data.length.toLocaleString('es-MX')}.</td></tr>`;
  document.getElementById('tbl-det').innerHTML = html;
}

// =========================================================================
// CHART HELPERS
// =========================================================================
function drawChart(id, cfg) {
  if (charts[id]) charts[id].destroy();
  const ctx = document.getElementById(id);
  if (!ctx) return;
  // Inyectar animación por defecto si no se pasó
  cfg.options = cfg.options || {};
  if (cfg.options.animation === undefined) {
    cfg.options.animation = { duration: 400, easing: 'easeOutCubic' };
  }
  charts[id] = new Chart(ctx, cfg);
}

// Tooltip enriquecido base (común a todos los charts)
const tooltipBase = {
  backgroundColor: 'rgba(13, 17, 23, 0.96)',
  titleColor: '#fff',
  titleFont: { weight: '700', size: 12 },
  bodyColor: '#e8eaed',
  bodyFont: { size: 12 },
  borderColor: 'rgba(217,102,44,0.4)',
  borderWidth: 1,
  padding: 10,
  cornerRadius: 6,
  displayColors: true,
  boxPadding: 4,
};

function barOpts({money=false, pct=false, horizontal=false, onClickFilter=null}={}) {
  return {
    indexAxis: horizontal?'y':'x',
    maintainAspectRatio:false, responsive:true,
    interaction: { mode: 'index', intersect: false },
    onClick: onClickFilter ? (evt, els) => {
      if (els.length === 0) return;
      const label = evt.chart.data.labels[els[0].index];
      onClickFilter(label);
    } : undefined,
    plugins:{
      legend:{display:true, position:'top', labels:{ usePointStyle:true, padding:14, font:{size:11}}},
      tooltip:{ ...tooltipBase, callbacks:{ label:(c)=>{ const v = c.parsed[horizontal?'x':'y']; return ' ' + c.dataset.label+': '+(money?fmtMoney(v):pct?fmtPct(v):fmtNum(v)); }}}
    },
    scales:{
      [horizontal?'x':'y']:{ beginAtZero:true, ticks:{ callback: v=> money?fmtMoneyShort(v):pct?v+'%':fmtNum(v), color:'#6b665a', font:{size:10}}, grid:{color:'rgba(36,33,28,0.05)'}, border:{display:false}},
      [horizontal?'y':'x']:{ ticks:{ font:{size:11}, autoSkip:false, color:'#6b665a' }, grid:{display:false}, border:{display:false}}
    }
  };
}
function lineOpts({money=false, pct=false}={}) {
  return {
    maintainAspectRatio:false, responsive:true,
    interaction: { mode: 'index', intersect: false },
    plugins:{
      legend:{position:'top', labels:{ usePointStyle:true, padding:14, font:{size:11}}},
      tooltip:{ ...tooltipBase, callbacks:{ label:c=>' ' + c.dataset.label+': '+(money?fmtMoney(c.parsed.y):pct?fmtPct(c.parsed.y):fmtNum(c.parsed.y))}}
    },
    scales:{
      y:{ ticks:{callback:v=>money?fmtMoneyShort(v):pct?v+'%':fmtNum(v), color:'#6b665a', font:{size:10}}, grid:{color:'rgba(36,33,28,0.05)'}, border:{display:false}},
      x:{ grid:{display:false}, ticks:{color:'#6b665a', font:{size:10}}, border:{display:false}}
    }
  };
}
function pieOpts({onClickFilter=null}={}) {
  return {
    maintainAspectRatio:false, responsive:true,
    onClick: onClickFilter ? (evt, els) => {
      if (els.length === 0) return;
      const label = evt.chart.data.labels[els[0].index];
      onClickFilter(label);
    } : undefined,
    onHover: (e, els) => { e.native.target.style.cursor = els[0] && onClickFilter ? 'pointer' : 'default'; },
    plugins:{
      legend:{position:'right', labels:{ usePointStyle:true, padding:10, font:{size:11}, color:'#24211c'}},
      tooltip:{ ...tooltipBase, callbacks:{ label:c=>{ const total = c.dataset.data.reduce((a,b)=>a+b,0); const pct = total?(c.parsed/total*100).toFixed(1):0; return ' '+c.label+': '+fmtMoney(c.parsed)+' ('+pct+'%)'; }}}
    }
  };
}

// =========================================================================
// SPARKLINES en KPI cards
// Dibuja un SVG mini chart (12 puntos) dentro del KPI
// =========================================================================
function sparklineSVG(data, color='#d9662c', w=64, h=24) {
  if (!data || data.length === 0) return '';
  const max = Math.max(...data, 1);
  const min = Math.min(...data, 0);
  const range = max - min || 1;
  const step = w / (data.length - 1);
  const points = data.map((v, i) => `${(i*step).toFixed(1)},${(h - ((v-min)/range)*h).toFixed(1)}`).join(' ');
  // Area path
  const areaPath = `M0,${h} L${points.split(' ').join(' L')} L${w},${h} Z`;
  return `<svg class="kpi-sparkline" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="none">
    <defs><linearGradient id="g${Math.random().toString(36).slice(2,8)}" x1="0" x2="0" y1="0" y2="1">
      <stop offset="0%" stop-color="${color}" stop-opacity="0.4"/>
      <stop offset="100%" stop-color="${color}" stop-opacity="0"/>
    </linearGradient></defs>
    <path d="${areaPath}" fill="${color}" fill-opacity="0.15"/>
    <polyline points="${points}" fill="none" stroke="${color}" stroke-width="1.5"/>
  </svg>`;
}

// =========================================================================
// EXPORT PNG de cualquier chart
// =========================================================================
function exportChartPNG(chartId, filename) {
  const c = charts[chartId];
  if (!c) return;
  const url = c.toBase64Image('image/png', 1);
  const a = document.createElement('a');
  a.href = url;
  a.download = (filename || chartId) + '.png';
  a.click();
}

// =========================================================================
// BÚSQUEDA en tablas
// El input con class="tbl-search" data-target="ID" filtra rows del tbody
// =========================================================================
function setupTableSearch() {
  document.querySelectorAll('.tbl-search').forEach(inp => {
    if (inp.dataset.wired) return;
    inp.dataset.wired = '1';
    inp.addEventListener('input', e => {
      const target = document.querySelector(inp.dataset.target);
      if (!target) return;
      const q = e.target.value.toLowerCase().trim();
      target.querySelectorAll('tbody tr').forEach(row => {
        const txt = row.textContent.toLowerCase();
        row.style.display = (!q || txt.includes(q)) ? '' : 'none';
      });
    });
  });
}

// =========================================================================
// Inyectar botones de Exportar PNG en cada card que contiene canvas
// =========================================================================
function addExportButtons() {
  document.querySelectorAll('.section.active .card').forEach(card => {
    const canvas = card.querySelector('canvas');
    if (!canvas) return;
    const cardHeader = card.querySelector('.card-header');
    if (!cardHeader) return;
    if (cardHeader.querySelector('.card-export-btn')) return; // ya existe
    const btn = document.createElement('button');
    btn.className = 'card-icon-btn card-export-btn';
    btn.title = 'Exportar gráfica como PNG';
    btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>`;
    btn.onclick = (e) => {
      e.stopPropagation();
      const title = cardHeader.querySelector('.card-title')?.textContent || 'chart';
      exportChartPNG(canvas.id, title.replace(/\s+/g, '_').toLowerCase());
    };
    // Insertar dentro de un wrapper de acciones (al final del header)
    let actions = cardHeader.querySelector('.card-actions');
    if (!actions) {
      actions = document.createElement('div');
      actions.className = 'card-actions';
      cardHeader.appendChild(actions);
    }
    actions.appendChild(btn);
  });
}

// =========================================================================
// EXPORTAR CSV
// =========================================================================
function exportCSV() {
  const data = Engine.applyFilters(filters);
  const headers = ['Año','Mes','Empresa','Ejecutivo','Cliente','Holding','Agencia','Estado','Localidad','Categoría','Medio','Soporte','Tipo','TipoCompra','Status','Canal','Campaña','VentaBruta','VentaNeta','TotalROIs','Costos','Utilidad','ComisionVendedor','ComisionADV'];
  const rows = data.map(r => [r.anio, r.mes, r.emp, r.eje, r.cli, r.hol, r.age, r.est, r.loc, r.cat, r.med, r.sop, r.tp, r.tc, r.st, r.cc, r.cmp, r.vb, r.vn, r.roi, r.cst, r.ut, r.com, r.comAdv]);
  let csv = '﻿' + headers.join(',') + '\n';
  rows.forEach(row => {
    csv += row.map(v => {
      if (typeof v === 'string' && (v.includes(',')||v.includes('"')||v.includes('\n'))) return '"'+v.replace(/"/g,'""')+'"';
      return v;
    }).join(',') + '\n';
  });
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `dashboard_open_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
}
