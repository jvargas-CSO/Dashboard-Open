// =========================================================================
// proyeccion.js · Pestaña Proyección
//
//  1) Cierre del año en curso: real a la fecha + lo ya comprometido + venta nueva esperada,
//     contra el Forecast.
//  2) Proyección del año siguiente comparada con el Forecast del equipo, con supuestos y
//     eventos que se activan/desactivan (elecciones, Mundial, salidas de ejecutivos, holdings,
//     cuentas clave, macro, ejecutivos nuevos).
//
// Método (sin cajas negras; todo se ve en el "puente"):
//   Cierre = max( A comprometido , promedio( B real+forecast del resto , C avance estacional ) )
//   Proyección = Base del cierre − no recurrente − salidas
//                × (retención y crecimiento de cuentas actuales + clientes nuevos, moderado por
//                   un % de persistencia del crecimiento histórico)
//                con piso en lo ya vendido para ese año, + eventos.
//
// Los supuestos se guardan en localStorage (solo en el navegador de cada persona).
// =========================================================================
const PX_STORE = 'proyeccion_supuestos_v1';
const PX_ELEC_W = [0.08, 0.12, 0.15, 0.27, 0.28, 0.10, 0, 0, 0, 0, 0, 0]; // dónde cae la venta electoral (ene-jun)
const PX_HOLD_HINTS = {
  ipg: 'Integrada con OMG (Omnicom); más revisiones de cuentas',
  omg: 'Integrándose con IPG; más revisiones de cuentas',
  omnicom: 'Integrándose con IPG; más revisiones de cuentas',
  wpp: 'WPP Media México reestructura; busca crecer a doble dígito en 2027 con pocas cuentas',
  havas: 'Señal débil de Latam en 2026 (dato no oficial)',
  dentsu: 'Reestructura global; OXXO pasó a iProspect',
  publicis: 'Líder en nuevo negocio global; PepsiCo global pasó de Omnicom a Publicis',
  starcom: 'Publicis Media; sin plan específico para México',
};

function pxDefaults() {
  return {
    metrica: 'vn',
    metaCierre: '', metaProx: '',
    persistencia: 50,                // % del crecimiento histórico que se asume que se repite
    calibrar: true,                  // calibrar el forecast con el cumplimiento del año en curso
    elec: { on: true, partidos: 2.0, desplaza: -1.5, prima: 0.5 },
    mundial: { on: true, override: '' },
    salidas: { on: true, retencion: 70, noSalen: [] },
    nuevosEj: { on: false, monto: '' },
    macro: { on: false, pct: 0 },
    holdings: { on: false, pct: {} },
    cuentas: { on: false, perdidas: [] },
  };
}
function pxLoad() {
  const d = pxDefaults();
  try {
    const s = JSON.parse(localStorage.getItem(PX_STORE) || 'null');
    if (s && typeof s === 'object') {
      Object.keys(d).forEach(k => {
        if (s[k] === undefined) return;
        d[k] = (d[k] && typeof d[k] === 'object' && !Array.isArray(d[k])) ? { ...d[k], ...s[k] } : s[k];
      });
    }
  } catch (e) { /* sin almacenamiento: se usan los valores por defecto */ }
  return d;
}
const pxIsExec = () => typeof execMode !== 'undefined' && execMode;
let pxExecInit = false;
function pxSave() { if (pxIsExec()) return; try { localStorage.setItem(PX_STORE, JSON.stringify(PX)); } catch (e) { /* ignorar */ } }
let PX = pxLoad();
let pxCache = {};
let pxPanelSig = '';
let pxWired = false;
let pxLast = null;

// ---------- utilidades ----------
const pxNum = v => { const n = parseFloat(String(v == null ? '' : v).replace(/[,$\s]/g, '')); return isFinite(n) ? n : NaN; };
const pxSum = (a, from = 0, to = 12) => a.slice(from, to).reduce((x, y) => x + y, 0);
const pxAvg = a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0;
const pxSigned = v => (v < 0 ? '−' : '+') + fmtMoneyShort(Math.abs(v));
const pxCls = v => v > 0 ? 'pos' : v < 0 ? 'neg' : '';
const pxPctTxt = v => v == null || !isFinite(v) ? '—' : (v >= 0 ? '+' : '−') + Math.abs(v).toFixed(1) + '%';
const pxEsc = s => (typeof escapeHtml === 'function' ? escapeHtml(String(s)) : String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])));
const pxKey = s => String(s).replace(/[^a-zA-Z0-9]+/g, '_');

function pxRecs(y) {
  const key = 'r' + y;
  if (!pxCache[key]) pxCache[key] = Engine.facturable(Engine.applyFilters({ ...filters, anio: y, mes: '' }, { ignoreYear: false }));
  return pxCache[key];
}
function pxMonthly(y, mk) {
  const a = Array(12).fill(0);
  pxRecs(y).forEach(r => { if (r.mes >= 1 && r.mes <= 12) a[r.mes - 1] += r[mk] || 0; });
  return a;
}
function pxFcNeto(y) {
  return Engine.hasForecast(y) ? Engine.forecastMonthlyFilteredArr({ ...filters, mes: '' }, y) : null;
}
// Venta neta -> la métrica elegida (misma convención que el resto del dashboard)
function pxConv(neto, y, mk) {
  if (mk === 'vn') return neto;
  if (mk === 'ut') return neto * MARGEN_OBJETIVO;
  const roi = Engine.roiPctParaForecastBruto(y);
  return roi < 1 ? neto / (1 - roi) : neto;
}
// Forma estacional: participación promedio de cada mes en los últimos años cerrados (datos reales;
// si un año no está cargado se usa la tabla histórica manual, solo para la forma).
function pxShares(yC, mk) {
  const arrs = [];
  for (let y = yC - 1; y >= yC - 4; y--) {
    let a = Engine.yearsAvailable.includes(y) ? pxMonthly(y, mk) : null;
    if (!pxIsExec() && (!a || pxSum(a) <= 0) && typeof HISTORICAL_SALES_BRUTA !== 'undefined' && HISTORICAL_SALES_BRUTA[y]) a = HISTORICAL_SALES_BRUTA[y];
    if (a) { const t = pxSum(a); if (t > 0) arrs.push(a.map(v => v / t)); }
  }
  if (!arrs.length) return Array(12).fill(1 / 12);
  return Array.from({ length: 12 }, (_, i) => pxAvg(arrs.map(a => a[i])));
}

// =========================================================================
// MODELO
// =========================================================================
function pxModel() {
  const mk = PX.metrica;
  const yC = projAnioCierre(), yP = yC - 1, yN = yC + 1;
  const now = new Date();
  const k = now.getFullYear() === yC ? now.getMonth() : 12;      // meses completos transcurridos
  const M = { mk, yC, yP, yN, k };

  // ---------- 1) CIERRE ----------
  const real = pxMonthly(yC, mk), prev = pxMonthly(yP, mk), s = pxShares(yC, mk);
  const YTD = pxSum(real, 0, k), A = pxSum(real), commRem = pxSum(real, k, 12);
  const fcNeto = pxFcNeto(yC);
  const fcArr = fcNeto ? fcNeto.map(v => pxConv(v, yC, mk)) : null;
  const cumK = pxSum(s, 0, k);
  const C = k >= 12 ? YTD : (k > 0 && cumK > 0.05 ? YTD / cumK : null);
  const B = fcArr ? YTD + real.slice(k).reduce((acc, v, i) => acc + Math.max(v, fcArr[k + i]), 0) : null;
  const cands = [B, C].filter(v => v != null && isFinite(v));
  let est = cands.length ? pxAvg(cands) : null;
  if (est == null) { est = pxIsExec() ? A : projBuildForecast(mk).totalCierre; M.cierreFallback = true; }
  const central = Math.max(A, est);
  const low = cands.length ? Math.max(A, Math.min(...cands)) : central;
  const high = cands.length ? Math.max(central, Math.max(...cands)) : central;
  const newExpTot = Math.max(0, central - A);
  const wOpen = pxSum(s, k, 12) || 1;
  const newExp = s.map((v, i) => i >= k ? newExpTot * v / wOpen : 0);
  Object.assign(M, { real, prev, s, YTD, A, commRem, fcArr, fcTot: fcArr ? pxSum(fcArr) : null, B, C, central, low, high, newExpTot, newExp,
    prevTot: pxSum(prev), prevYTD: pxSum(prev, 0, k), cumK });

  // Cumplimiento del forecast del año en curso (siempre en venta neta: la razón es la misma)
  const realVN = pxMonthly(yC, 'vn');
  let att = null;
  if (fcNeto && k > 0) {
    const rv = pxSum(realVN, 0, k), fv = pxSum(fcNeto, 0, k);
    if (fv > 0) { M.attRaw = rv / fv; att = Math.min(1.4, Math.max(0.7, rv / fv)); }
  }
  M.att = att;

  // ---------- 2) PROYECCIÓN DEL AÑO SIGUIENTE ----------
  // Retención / crecimiento de cuentas actuales y clientes nuevos, mismo periodo (meses cerrados)
  const kk = (k >= 1 && k <= 11) ? k : 12;
  const cmap = y => {
    const m = new Map();
    pxRecs(y).forEach(r => { if (r.mes >= 1 && r.mes <= kk) { const v = r[mk] || 0; if (v > 0) m.set(r.cli, (m.get(r.cli) || 0) + v); } });
    return m;
  };
  const trans = [];
  [[yC - 2, yP], [yP, yC]].forEach(([a, b]) => {
    if (!Engine.yearsAvailable.includes(a) || !Engine.yearsAvailable.includes(b)) return;
    const ma = cmap(a), mb = cmap(b);
    let sumA = 0, kept = 0, ret = 0, nw = 0;
    ma.forEach((v, c) => { sumA += v; const vb = mb.get(c) || 0; kept += vb; if (vb > 0) ret += v; });
    mb.forEach((v, c) => { if (!ma.has(c)) nw += v; });
    if (sumA > 0) trans.push({ a, b, nrr: kept / sumA, nw: nw / sumA, logo: ret / sumA });
  });
  const nrr = trans.length ? pxAvg(trans.map(t => t.nrr)) : 1;
  const nw = trans.length ? pxAvg(trans.map(t => t.nw)) : 0;
  const Gs = trans.map(t => t.nrr + t.nw);
  const G = nrr + nw;
  let gLow = Gs.length ? Math.min(...Gs) : G, gHigh = Gs.length ? Math.max(...Gs) : G;
  if (Gs.length < 2) { gLow = G - 0.05; gHigh = G + 0.05; }
  const lam = Math.max(0, Math.min(1, (PX.persistencia || 0) / 100));
  const damp = g => 1 + lam * (g - 1);
  const nrrL = 1 + lam * (nrr - 1), nwL = lam * nw, GL = nrrL + nwL;
  Object.assign(M, { trans, nrr, nw, G, gLow, gHigh, lam, nrrL, nwL, GL });

  // No recurrente (Mundial 2026): solo aplica si el año en curso es 2026 y ya pasaron jun-jul
  const mundialAplica = yC === 2026 && k >= 7;
  let Wauto = 0;
  if (mundialAplica) {
    const otros = real.map((v, i) => ({ v, i })).filter(o => o.i < k && o.i !== 5 && o.i !== 6);
    const sumO = otros.reduce((a, o) => a + o.v, 0), shO = otros.reduce((a, o) => a + s[o.i], 0);
    const esperado = shO > 0 ? sumO * (s[5] + s[6]) / shO : 0;
    Wauto = Math.max(0, real[5] + real[6] - esperado);
  }
  const wOv = pxNum(PX.mundial.override);
  const Wbase = isFinite(wOv) ? pxFromNeto(wOv * 1e6, yN, mk) : Wauto;
  const W = (PX.mundial.on && mundialAplica) ? Math.min(Wbase, central) : 0;
  const Wfrac = central > 0 ? W / central : 0;
  Object.assign(M, { mundialAplica, Wauto, W });

  // Salidas de ejecutivos: quienes vendieron en el año en curso pero no tienen forecast del año siguiente
  const hasFcN = Engine.hasForecast(yN);
  const activeN = new Set(Engine.forecast.filter(f => f.anio === yN).map(f => f.eje));
  const detDep = e => hasFcN && e && e !== 'Sin asignar' && !activeN.has(e);
  const excl = PX.salidas.noSalen || [];
  const isDep = e => detDep(e) && !excl.includes(e);
  const retS = Math.max(0, Math.min(1, (PX.salidas.retencion || 0) / 100));

  // Base del cierre repartida por cliente / vendedor / holding (con no recurrente y salidas)
  const scale = A > 0 ? central / A : 1;
  const byCli = new Map(), byEje = new Map(), byHol = new Map(), depAll = new Map();
  const add = (map, key, base, adj) => { const o = map.get(key) || { base: 0, adj: 0 }; o.base += base; o.adj += adj; map.set(key, o); };
  let B0 = 0, Wsum = 0, Lsum = 0;
  pxRecs(yC).forEach(r => {
    const v = (r[mk] || 0) * scale;
    if (!v) return;
    const afterW = v * (1 - Wfrac);
    const dep = PX.salidas.on && isDep(r.eje);
    const adj = dep ? afterW * retS : afterW;
    B0 += v; Wsum += v * Wfrac; Lsum += afterW - adj;
    add(byCli, r.cli, v, adj); add(byEje, r.eje, v, adj); add(byHol, r.hol || 'Sin definir', v, adj);
    if (detDep(r.eje)) depAll.set(r.eje, (depAll.get(r.eje) || 0) + v);
  });
  const B1 = B0 - Wsum - Lsum;
  Object.assign(M, { B0, B1, Lsum, byCli, byEje, byHol, depAll, hasFcN });

  const S0 = B1 * GL;                                         // proyección estadística sin piso ni eventos
  const Y27m = pxMonthly(yN, mk), Y27 = pxSum(Y27m);           // ya vendido para el año siguiente
  const S0adj = Math.max(S0, Y27);
  M.S0 = S0; M.S0adj = S0adj; M.Y27 = Y27; M.Y27m = Y27m;

  // Eventos (fracciones sobre la proyección o montos)
  const ev = { elec: { partidos: 0, desplaza: 0, prima: 0 }, macro: 0, holds: [], cuentas: [], nuevos: 0 };
  if (PX.elec.on) {
    ev.elec.partidos = S0adj * (PX.elec.partidos || 0) / 100;
    ev.elec.desplaza = S0adj * (PX.elec.desplaza || 0) / 100;
    ev.elec.prima = S0adj * (PX.elec.prima || 0) / 100;
  }
  if (PX.macro.on) ev.macro = S0adj * (PX.macro.pct || 0) / 100;
  if (PX.holdings.on) {
    Object.keys(PX.holdings.pct || {}).forEach(h => {
      const pct = PX.holdings.pct[h] || 0, o = byHol.get(h);
      if (!pct || !o || B1 <= 0) return;
      ev.holds.push({ h, pct, delta: S0adj * (o.adj / B1) * pct / 100 });
    });
  }
  if (PX.cuentas.on) {
    (PX.cuentas.perdidas || []).forEach(c => {
      const o = byCli.get(c);
      if (o) ev.cuentas.push({ c, delta: -o.adj * nrrL });
    });
  }
  if (PX.nuevosEj.on) {
    const m = pxNum(PX.nuevosEj.monto);
    if (isFinite(m)) ev.nuevos = pxFromNeto(m * 1e6, yN, mk);
  }
  const elecTot = ev.elec.partidos + ev.elec.desplaza + ev.elec.prima;
  const holdTot = ev.holds.reduce((a, x) => a + x.delta, 0);
  const ctaTot = ev.cuentas.reduce((a, x) => a + x.delta, 0);
  const evTot = elecTot + ev.macro + holdTot + ctaTot + ev.nuevos;
  const P = S0adj + evTot;
  Object.assign(M, { ev, elecTot, holdTot, ctaTot, evTot, P });

  // Rango: con el crecimiento histórico más bajo / más alto observado (los eventos escalan igual)
  const fracEv = S0adj > 0 ? (elecTot + ev.macro + holdTot) / S0adj : 0;
  const lostB1 = (PX.cuentas.on ? (PX.cuentas.perdidas || []) : []).reduce((a, c) => a + (byCli.get(c)?.adj || 0), 0);
  const pAt = g => { const s0 = Math.max(B1 * damp(g), Y27); return s0 * (1 + fracEv) - lostB1 * (1 + lam * ((g - nw) - 1)) + ev.nuevos; };
  M.pLow = Math.min(pAt(gLow), P); M.pHigh = Math.max(pAt(gHigh), P);

  // Forecast del equipo para el año siguiente y su versión calibrada
  const fcN2 = pxFcNeto(yN);
  M.fcArrN = fcN2 ? fcN2.map(v => pxConv(v, yN, mk)) : null;
  M.F27 = M.fcArrN ? pxSum(M.fcArrN) : null;
  M.Fcal = (M.F27 != null && PX.calibrar && att != null) ? M.F27 * att : null;

  // Metas capturadas (venta neta, millones)
  const mC = pxNum(PX.metaCierre), mP = pxNum(PX.metaProx);
  M.metaC = isFinite(mC) && mC > 0 ? pxFromNeto(mC * 1e6, yC, mk) : null;
  M.metaP = isFinite(mP) && mP > 0 ? pxFromNeto(mP * 1e6, yN, mk) : null;

  // ---------- Distribución mensual de la proyección ----------
  const sN = s;                                               // forma estacional
  let s0m = sN.map(w => S0adj * w);
  // Un mes nunca baja de lo ya vendido: se fija y el resto se reparte en los demás meses
  const fixed = s0m.map((v, i) => Y27m[i] > v);
  const fixedSum = Y27m.reduce((a, v, i) => a + (fixed[i] ? v : 0), 0);
  const freeW = sN.reduce((a, w, i) => a + (fixed[i] ? 0 : w), 0);
  if (fixed.some(Boolean) && freeW > 0) s0m = sN.map((w, i) => fixed[i] ? Y27m[i] : Math.max(0, S0adj - fixedSum) * w / freeW);
  const evShape = elecTot * 1;
  const restEv = ev.macro + holdTot + ctaTot + ev.nuevos;
  M.s0m = s0m;
  M.Pm = s0m.map((v, i) => v + evShape * PX_ELEC_W[i] + restEv * sN[i]);
  return M;
}

// Venta neta (pesos) -> métrica elegida, para montos capturados a mano
function pxFromNeto(x, y, mk) { return pxConv(x, y, mk); }

// =========================================================================
// RENDER
// =========================================================================
function kpiBig(label, sub, value, accent, cmps) {
  return `<div class="kpi-big" style="--accent:${accent}"><div class="kpi-big-head"><div class="kpi-big-label">${label}</div><div class="kpi-big-sub">${sub || ''}</div></div>
    <div class="kpi-big-value">${value}</div>
    <div class="kpi-big-compare">${cmps.map(c => `<div class="kpi-cmp"><div class="kpi-cmp-label">${c[0]}</div><div class="kpi-cmp-pct ${c[2] || ''}">${c[1]}</div></div>`).join('')}</div></div>`;
}
const pxTd = (v, cls = '') => `<td class="num ${cls}" title="${fmtMoney(v)}">${fmtMoneyShort(v)}</td>`;
const pxTdS = v => `<td class="num ${pxCls(v)}" title="${fmtMoney(v)}">${v === 0 ? '—' : pxSigned(v)}</td>`;
const MLAB = { vn: 'venta neta', vb: 'venta bruta', ut: 'utilidad' };

function renderProyeccion() {
  renderProyeccionLegacy();              // etiquetas de año + histórico
  if (pxIsExec() && !pxExecInit) {
    // Vista por ejecutivo: todos los eventos apagados, sin metas de empresa
    PX = pxDefaults();
    ['elec', 'mundial', 'salidas', 'nuevosEj', 'macro', 'holdings', 'cuentas'].forEach(k => { PX[k].on = false; });
    pxExecInit = true;
  }
  if (!pxWired) pxWire();
  pxUpdate();
}

function pxUpdate() {
  pxCache = {};
  let M;
  try { M = pxModel(); } catch (e) {
    console.error('Proyección:', e);
    document.getElementById('pxKpisCierre').innerHTML = `<div class="alert info" style="grid-column:1/-1">No se pudo calcular la proyección: ${pxEsc(e.message)}</div>`;
    return;
  }
  pxLast = M;
  pxSyncToolbar(M);
  pxBuildPanel(M);
  pxRenderCierre(M);
  pxRenderProx(M);
}

function pxSyncToolbar(M) {
  document.querySelectorAll('#pxMetrica button').forEach(b => b.classList.toggle('on', b.dataset.m === PX.metrica));
  const mc = document.getElementById('pxMetaC'), mp = document.getElementById('pxMetaP');
  if (mc && document.activeElement !== mc) mc.value = PX.metaCierre;
  if (mp && document.activeElement !== mp) mp.value = PX.metaProx;
  const f = [];
  if (filters.eje) f.push('vendedor: ' + filters.eje);
  if (filters.cli) f.push('cliente: ' + filters.cli);
  if (filters.hol) f.push('holding: ' + filters.hol);
  if (filters.emp) f.push('empresa: ' + filters.emp);
  const st = (filters.statusOpen || []).join(', ') || 'todos los status';
  document.getElementById('pxFiltrosNota').textContent = `Respeta los filtros del panel superior (excepto Año y Mes) · Status: ${st}` + (f.length ? ' · ' + f.join(' · ') : '') + ` · Cifras en ${MLAB[M.mk]}`;
}

// ---------- 1) Cierre ----------
function pxRenderCierre(M) {
  const { yC, yP, k } = M;
  const mesTxt = k > 0 ? (k === 1 ? 'enero' : `enero–${MESES_FULL[k - 1].toLowerCase()}`) : '—';
  document.getElementById('pxCierreNota').textContent = k >= 12
    ? `Año completo: ${yC} ya cerró.`
    : `Meses cerrados: ${mesTxt}. Del mes en curso (${MESES_FULL[Math.min(k, 11)].toLowerCase()}) en adelante cuenta lo ya comprometido más la venta nueva esperada.`;
  const yoy = M.prevTot > 0 ? (M.central / M.prevTot - 1) * 100 : null;
  const yoyYTD = M.prevYTD > 0 ? (M.YTD / M.prevYTD - 1) * 100 : null;
  const vsFc = M.fcTot ? (M.central / M.fcTot - 1) * 100 : null;
  const vsMeta = M.metaC ? (M.central / M.metaC - 1) * 100 : null;
  const cards = [];
  cards.push(kpiBig(`Cierre proyectado ${yC}`, MLAB[M.mk], fmtMoneyShort(M.central), '#d9662c', [
    ['Rango', `${fmtMoneyShort(M.low)} – ${fmtMoneyShort(M.high)}`],
    [`vs ${yP}`, pxPctTxt(yoy), yoy == null ? 'muted' : yoy >= 0 ? 'pos' : 'neg'],
  ]));
  cards.push(kpiBig('Real a la fecha', mesTxt, fmtMoneyShort(M.YTD), '#1f9d6e', [
    [`vs mismo periodo ${yP}`, pxPctTxt(yoyYTD), yoyYTD == null ? 'muted' : yoyYTD >= 0 ? 'pos' : 'neg'],
    ['Avance del cierre', M.central > 0 ? fmtPct(M.YTD / M.central * 100) : '—'],
  ]));
  cards.push(kpiBig('Resto del año', 'meses abiertos', fmtMoneyShort(M.commRem + M.newExpTot), '#2563eb', [
    ['Ya comprometido', fmtMoneyShort(M.commRem)],
    ['Venta nueva esperada', fmtMoneyShort(M.newExpTot)],
  ]));
  if (M.fcTot != null) {
    cards.push(kpiBig(`Contra Forecast ${yC}`, M.metaC ? 'y meta' : '', fmtMoneyShort(M.fcTot), '#7c3aed', [
      ['Cierre vs Forecast', pxPctTxt(vsFc), vsFc == null ? 'muted' : vsFc >= 0 ? 'pos' : 'neg'],
      M.metaC ? ['Cierre vs meta', pxPctTxt(vsMeta), vsMeta >= 0 ? 'pos' : 'neg'] : ['Cumplimiento a la fecha', M.attRaw != null ? fmtPct(M.attRaw * 100) : '—'],
    ]));
  } else {
    cards.push(kpiBig(`Forecast ${yC}`, '', 'No cargado', '#7c3aed', [['Cierre vs meta', M.metaC ? pxPctTxt(vsMeta) : '—', M.metaC ? (vsMeta >= 0 ? 'pos' : 'neg') : 'muted'], ['', '—', 'muted']]));
  }
  document.getElementById('pxKpisCierre').innerHTML = cards.join('');

  document.getElementById('pxCierreMeta').textContent = `Líneas = Forecast ${yC} y real ${yP}`;
  const closed = M.real.map((v, i) => i < k ? v : 0);
  const comm = M.real.map((v, i) => i >= k ? v : 0);
  const baseOpts = lineOpts({ money: true });
  baseOpts.scales.x.stacked = true; baseOpts.scales.y.stacked = true;
  drawChart('ch-pxCierre', {
    type: 'bar',
    data: { labels: MESES, datasets: [
      { label: 'Real (meses cerrados)', data: closed, backgroundColor: '#1f9d6e', stack: 's', borderRadius: 3, order: 3 },
      { label: 'Ya comprometido (meses abiertos)', data: comm, backgroundColor: '#2563eb', stack: 's', borderRadius: 3, order: 3 },
      { label: 'Venta nueva esperada', data: M.newExp, backgroundColor: 'rgba(217,102,44,0.55)', stack: 's', borderRadius: 3, order: 3 },
      ...(M.fcArr ? [{ type: 'line', label: `Forecast ${yC}`, data: M.fcArr, borderColor: '#7c3aed', borderDash: [6, 4], backgroundColor: 'transparent', pointRadius: 3, borderWidth: 2, tension: 0.3, order: 1 }] : []),
      { type: 'line', label: `${yP} real`, data: M.prev, borderColor: '#9a948a', backgroundColor: 'transparent', pointRadius: 2, borderWidth: 2, tension: 0.3, order: 2 },
    ] },
    options: baseOpts,
  });

  const fila = (id, nombre, val, txt) => `<tr><td><b>${id}</b> · ${nombre}</td>${pxTd(val)}<td style="color:var(--text-muted)">${txt}</td></tr>`;
  let html = '<div class="table-wrap"><table><thead class="top"><tr><th>Método</th><th class="num">Cierre</th><th>Cómo se calcula</th></tr></thead><tbody>';
  html += fila('A', 'Comprometido', M.A, `Todo lo que ya está cargado en Data Comercial para ${yC} (piso: no puede cerrar por debajo).`);
  html += M.B != null
    ? fila('B', 'Real + Forecast del resto', M.B, `Lo real de los meses cerrados + en cada mes abierto el mayor entre lo comprometido y el Forecast.`)
    : `<tr><td><b>B</b> · Real + Forecast del resto</td><td class="num">—</td><td style="color:var(--text-muted)">No hay Forecast cargado para ${yC}.</td></tr>`;
  html += M.C != null
    ? fila('C', 'Avance estacional', M.C, `Real de los meses cerrados ÷ el % del año que esos meses suelen representar (${fmtPct(M.cumK * 100)}, promedio de los últimos años cerrados).`)
    : `<tr><td><b>C</b> · Avance estacional</td><td class="num">—</td><td style="color:var(--text-muted)">Aún no hay meses cerrados en ${yC}.</td></tr>`;
  html += `<tr class="px-puente tot"><td><b>Cierre proyectado</b></td>${pxTd(M.central)}<td style="color:var(--text-muted)">${M.cierreFallback ? 'Sin B ni C: se usa la proyección por tendencia histórica.' : 'Promedio de B y C, nunca menor que A.'} Rango: ${fmtMoneyShort(M.low)} – ${fmtMoneyShort(M.high)}.</td></tr>`;
  html += '</tbody></table></div>';
  document.getElementById('pxCierreMetodos').innerHTML = html;
}

// ---------- 2) Proyección año siguiente ----------
function pxRenderProx(M) {
  const { yC, yN } = M;
  document.getElementById('pxProxNota').textContent = M.hasFcN
    ? `Base: cierre ${yC} proyectado, ajustado por retención, clientes nuevos y los eventos que tengas activos.`
    : `Aún no hay Forecast ${yN} cargado: se muestra solo la proyección estadística.`;
  const gapF = M.F27 != null ? M.P - M.F27 : null;
  const gapPct = M.F27 ? gapF / M.F27 * 100 : null;
  const growth = M.central > 0 ? (M.P / M.central - 1) * 100 : null;
  const cov = M.F27 ? M.Y27 / M.F27 * 100 : null;
  const vsMeta = M.metaP ? (M.P / M.metaP - 1) * 100 : null;
  const cards = [];
  cards.push(kpiBig(`Proyección ${yN}`, 'con tus supuestos', fmtMoneyShort(M.P), '#d9662c', [
    ['Rango', `${fmtMoneyShort(M.pLow)} – ${fmtMoneyShort(M.pHigh)}`],
    [`vs cierre ${yC}`, pxPctTxt(growth), growth == null ? 'muted' : growth >= 0 ? 'pos' : 'neg'],
  ]));
  if (M.F27 != null) {
    cards.push(kpiBig(`Forecast del equipo ${yN}`, M.Fcal != null ? 'y calibrado' : '', fmtMoneyShort(M.F27), '#7c3aed', [
      ['Proyección − Forecast', `${pxSigned(gapF)} (${pxPctTxt(gapPct)})`, gapF >= 0 ? 'pos' : 'neg'],
      M.Fcal != null ? [`Forecast calibrado (×${M.att.toFixed(2)})`, fmtMoneyShort(M.Fcal)] : ['Calibrado', 'apagado', 'muted'],
    ]));
  } else {
    cards.push(kpiBig(`Forecast del equipo ${yN}`, '', 'No cargado', '#7c3aed', [['Proyección − Forecast', '—', 'muted'], ['Calibrado', '—', 'muted']]));
  }
  cards.push(kpiBig(`Ya vendido para ${yN}`, 'piso seguro', fmtMoneyShort(M.Y27), '#2563eb', [
    ['Cobertura del Forecast', cov == null ? '—' : fmtPct(cov)],
    ['Cobertura de la proyección', M.P > 0 ? fmtPct(M.Y27 / M.P * 100) : '—'],
  ]));
  if (!pxIsExec()) cards.push(kpiBig('Efecto de eventos', 'sobre la proyección', pxSigned(M.evTot), M.evTot >= 0 ? '#1f9d6e' : '#dc4747', [
    ['Elecciones', pxSigned(M.elecTot), pxCls(M.elecTot)],
    ['Otros eventos', pxSigned(M.evTot - M.elecTot), pxCls(M.evTot - M.elecTot)],
  ]));
  if (M.metaP) cards.push(kpiBig(`Meta ${yN}`, 'venta neta capturada', fmtMoneyShort(M.metaP), '#0891b2', [
    ['Proyección vs meta', pxPctTxt(vsMeta), vsMeta >= 0 ? 'pos' : 'neg'],
    ['Forecast vs meta', M.F27 != null ? pxPctTxt((M.F27 / M.metaP - 1) * 100) : '—', M.F27 != null ? (M.F27 >= M.metaP ? 'pos' : 'neg') : 'muted'],
  ]));
  document.getElementById('pxKpisProx').innerHTML = cards.join('');

  // gráfica mensual
  document.getElementById('pxProxMeta').textContent = `Barras = ya vendido ${yN} · líneas = proyecciones y forecast`;
  drawChart('ch-pxProx', {
    type: 'bar',
    data: { labels: MESES, datasets: [
      { label: `Ya vendido ${yN}`, data: M.Y27m, backgroundColor: 'rgba(37,99,235,0.55)', borderRadius: 3, order: 4 },
      { type: 'line', label: `Proyección ${yN} (con eventos)`, data: M.Pm, borderColor: '#d9662c', backgroundColor: 'rgba(217,102,44,0.10)', fill: true, borderWidth: 3, pointRadius: 4, tension: 0.3, order: 1 },
      { type: 'line', label: 'Proyección sin eventos', data: M.s0m, borderColor: '#d9662c', borderDash: [5, 4], backgroundColor: 'transparent', borderWidth: 2, pointRadius: 0, tension: 0.3, order: 2 },
      ...(M.fcArrN ? [{ type: 'line', label: `Forecast ${yN}`, data: M.fcArrN, borderColor: '#7c3aed', backgroundColor: 'transparent', borderWidth: 2, pointRadius: 3, tension: 0.3, order: 2 }] : []),
      ...(M.fcArrN && M.Fcal != null ? [{ type: 'line', label: 'Forecast calibrado', data: M.fcArrN.map(v => v * M.att), borderColor: '#7c3aed', borderDash: [2, 3], backgroundColor: 'transparent', borderWidth: 2, pointRadius: 0, tension: 0.3, order: 2 }] : []),
      { type: 'line', label: `Cierre ${yC}`, data: M.real.map((v, i) => v + M.newExp[i]), borderColor: '#9a948a', backgroundColor: 'transparent', borderWidth: 2, pointRadius: 2, tension: 0.3, order: 3 },
    ] },
    options: lineOpts({ money: true }),
  });

  pxRenderPuente(M);
  pxRenderTablas(M);
  pxRenderMetodos(M);
  pxUpdateImpacts(M);
}

function pxRenderPuente(M) {
  const { yC, yN } = M;
  const row = (txt, v, cls = '', note = '', sub = false) => `<tr class="${cls}"><td class="${sub ? 'sub' : ''}">${txt}</td>${cls === 'mid' || cls === 'tot' ? pxTd(v) : pxTdS(v)}<td style="color:var(--text-muted)">${note}</td></tr>`;
  const t = M.trans.map(x => `${x.a}→${x.b}: retención ${fmtPct(x.nrr * 100)}, nuevos +${fmtPct(x.nw * 100)}`).join(' · ') || 'sin historia suficiente de clientes';
  let h = '<div class="table-wrap px-puente"><table><thead class="top"><tr><th>Concepto</th><th class="num">Monto</th><th>Nota</th></tr></thead><tbody>';
  h += row(`Cierre ${yC} proyectado`, M.B0, 'mid', 'Punto de partida (ver sección 1)');
  if (!pxIsExec()) h += row('− No recurrente (Mundial)', -M.W, '', PX.mundial.on ? (M.mundialAplica ? (isFinite(pxNum(PX.mundial.override)) ? 'Monto capturado a mano' : `Exceso de jun–jul sobre la estacionalidad normal: ${fmtMoneyShort(M.Wauto)} (estimado)`) : 'No aplica en este año') : 'Apagado');
  if (!pxIsExec()) h += row('− Salidas de ejecutivos', -M.Lsum, '', PX.salidas.on ? (M.hasFcN ? `Cartera de quienes no tienen forecast ${yN}; se asume que el equipo retiene ${PX.salidas.retencion}%` : `Requiere Forecast ${yN}`) : 'Apagado');
  h += row('Base repetible', M.B1, 'mid', 'Lo que se espera que repita');
  h += row('± Retención y crecimiento de cuentas actuales', M.B1 * (M.nrrL - 1), '', `Histórico: ${fmtPct(M.nrr * 100)} de retención neta; se asume que persiste ${PX.persistencia}% de ese crecimiento`);
  h += row('+ Clientes nuevos', M.B1 * M.nwL, '', `Histórico: clientes nuevos aportan ${fmtPct(M.nw * 100)} de la base; persiste ${PX.persistencia}%`);
  if (M.S0adj > M.S0 + 1) h += row('+ Piso por ya vendido', M.S0adj - M.S0, '', `Lo ya vendido para ${yN} (${fmtMoneyShort(M.Y27)}) supera la proyección estadística`);
  h += row('Proyección estadística (sin eventos)', M.S0adj, 'mid', `Crecimiento histórico observado: ${t}`);
  if (!pxIsExec()) {
  h += row('🗳️ Elecciones', M.elecTot, '', PX.elec.on ? 'Partidos + desplazamiento de marcas + prima de tarifa' : 'Apagado');
  if (PX.elec.on) {
    h += row('Venta a partidos y gobiernos', M.ev.elec.partidos, '', '', true);
    h += row('Marcas desplazadas / que evitan compartir', M.ev.elec.desplaza, '', '', true);
    h += row('Prima de tarifa por demanda alta', M.ev.elec.prima, '', '', true);
  }
  h += row('📉 Macro (PIB / T-MEC)', M.ev.macro, '', PX.macro.on ? `${PX.macro.pct}% sobre la proyección` : 'Apagado');
  h += row('🏢 Holdings', M.holdTot, '', PX.holdings.on ? (M.ev.holds.length ? M.ev.holds.map(x => `${pxEsc(x.h)} ${x.pct > 0 ? '+' : ''}${x.pct}%`).join(' · ') : 'Sin ajustes capturados') : 'Apagado');
  h += row('🎯 Cuentas clave', M.ctaTot, '', PX.cuentas.on ? (M.ev.cuentas.length ? 'Se simula perder: ' + M.ev.cuentas.map(x => pxEsc(x.c)).join(', ') : 'Ninguna seleccionada') : 'Apagado');
  h += row('🆕 Ejecutivos nuevos', M.ev.nuevos, '', PX.nuevosEj.on ? 'Venta capturada a mano' : 'Apagado');
  }
  h += row(`Proyección ${yN}`, M.P, 'tot', M.F27 != null ? `Forecast del equipo: ${fmtMoneyShort(M.F27)} · diferencia ${pxSigned(M.P - M.F27)}` : '');
  h += '</tbody></table></div>';
  document.getElementById('pxPuente').innerHTML = h;
}

function pxRenderTablas(M) {
  const { yN } = M, mk = M.mk;
  // Forecast del año siguiente agrupado (respeta filtros de vendedor/cliente/holding)
  const fcEntries = Engine.hasForecast(yN) ? Engine.applyForecastFilters({ ...filters, anio: yN, mes: '' }) : [];
  const isNuevo = c => /nuevo negocio/i.test(String(c));
  const fEje = new Map(), fCli = new Map(); let fNuevo = 0;
  fcEntries.forEach(f => {
    const v = pxConv(f.fcNeto, yN, mk);
    if (isNuevo(f.cli)) { fNuevo += v; return; }
    fEje.set(f.eje, (fEje.get(f.eje) || 0) + v);
    fCli.set(f.cli, (fCli.get(f.cli) || 0) + v);
  });
  const yEje = new Map(), yCli = new Map();
  pxRecs(yN).forEach(r => { const v = r[mk] || 0; yEje.set(r.eje, (yEje.get(r.eje) || 0) + v); yCli.set(r.cli, (yCli.get(r.cli) || 0) + v); });

  const mkRows = (keys, byMap, fMap, yMap, extra) => keys.map(key => {
    const o = byMap.get(key) || { base: 0, adj: 0 };
    const proj = o.adj * M.nrrL, F = fMap.get(key) || 0, Y = yMap.get(key) || 0;
    return { key, base: o.base, proj, F, Y, gap: proj - F, extra: extra ? extra(key) : '' };
  });
  const head = lbl => `<div class="table-wrap"><table><thead class="top"><tr><th>${lbl}</th><th class="num">Base ${M.yC}</th><th class="num">Proyección ${yN}</th><th class="num">Forecast ${yN}</th><th class="num">Ya vendido ${yN}</th><th class="num">Proy. − Forecast</th><th class="num">%</th></tr></thead><tbody>`;
  const line = r => `<tr><td><b>${pxEsc(r.key)}</b>${r.extra}</td>${pxTd(r.base)}${pxTd(r.proj)}${pxTd(r.F)}${pxTd(r.Y)}${pxTdS(r.gap)}<td class="num ${pxCls(r.gap)}">${r.F > 0 ? pxPctTxt(r.gap / r.F * 100) : '—'}</td></tr>`;
  const tot = (lbl, r, cls = '') => `<tr class="px-puente ${cls}"><td><b>${lbl}</b></td>${pxTd(r.base)}${pxTd(r.proj)}${pxTd(r.F)}${pxTd(r.Y)}${pxTdS(r.gap)}<td class="num ${pxCls(r.gap)}">${r.F > 0 ? pxPctTxt(r.gap / r.F * 100) : '—'}</td></tr>`;

  const newProj = M.B1 * M.nwL;
  const newRow = { key: 'Nuevo negocio (clientes nuevos)', base: 0, proj: newProj, F: fNuevo, Y: 0, gap: newProj - fNuevo, extra: '<span class="px-tag">no asignado</span>' };

  // ---- por vendedor ----
  const ejeKeys = [...new Set([...M.byEje.keys(), ...fEje.keys(), ...yEje.keys()])];
  let ejeRows = mkRows(ejeKeys, M.byEje, fEje, yEje, k => (M.depAll.has(k) && !(PX.salidas.noSalen || []).includes(k) && PX.salidas.on) ? '<span class="px-tag">sin forecast</span>' : '')
    .filter(r => r.base || r.proj || r.F || r.Y).sort((a, b) => Math.max(b.proj, b.F) - Math.max(a.proj, a.F));
  let h = head('Vendedor');
  ejeRows.forEach(r => { h += line(r); });
  h += line(newRow);
  const sum = (rows, f) => rows.reduce((a, r) => a + r[f], 0);
  const allE = [...ejeRows, newRow];
  const tE = { base: sum(allE, 'base'), proj: sum(allE, 'proj'), F: sum(allE, 'F'), Y: sum(allE, 'Y') }; tE.gap = tE.proj - tE.F;
  h += tot('Total estadístico', tE, 'mid');
  const pisoEv = (M.S0adj - tE.proj) + M.evTot;
  h += `<tr><td>Piso por ya vendido + eventos activos</td><td></td>${pxTdS(pisoEv)}<td></td><td></td><td></td><td></td></tr>`;
  h += tot('Proyección con eventos', { base: tE.base, proj: M.P, F: tE.F, Y: tE.Y, gap: M.P - tE.F }, 'tot');
  h += '</tbody></table></div>';
  document.getElementById('pxTblEje').innerHTML = h;

  // ---- por cliente (top 15) ----
  const cliKeys = [...new Set([...M.byCli.keys(), ...fCli.keys(), ...yCli.keys()])];
  let cliRows = mkRows(cliKeys, M.byCli, fCli, yCli).filter(r => r.base || r.proj || r.F || r.Y).sort((a, b) => Math.max(b.proj, b.F) - Math.max(a.proj, a.F));
  const top = cliRows.slice(0, 15), rest = cliRows.slice(15);
  h = head('Cliente');
  top.forEach(r => { h += line(r); });
  if (rest.length) {
    const rr = { key: `Resto (${rest.length} clientes)`, base: sum(rest, 'base'), proj: sum(rest, 'proj'), F: sum(rest, 'F'), Y: sum(rest, 'Y'), extra: '' }; rr.gap = rr.proj - rr.F;
    h += line(rr);
  }
  h += line(newRow);
  h += tot('Total estadístico', tE, 'mid');
  h += '</tbody></table></div>';
  document.getElementById('pxTblCli').innerHTML = h;
}

function pxRenderMetodos(M) {
  const { yN, mk } = M;
  const old = pxIsExec() ? { totalProjBase: null, linReg: null } : projBuildForecast(mk);
  const row = (n, v, note) => `<tr><td>${n}</td>${v == null ? '<td class="num">—</td>' : pxTd(v)}<td class="num ${v != null && M.P ? pxCls(v - M.P) : ''}">${v != null && M.P ? pxPctTxt((v / M.P - 1) * 100) : '—'}</td><td style="color:var(--text-muted)">${note}</td></tr>`;
  let h = `<div class="table-wrap"><table><thead class="top"><tr><th>Método</th><th class="num">${yN}</th><th class="num">vs mi proyección</th><th>Qué es</th></tr></thead><tbody>`;
  if (pxIsExec()) h += row('<b>Proyección</b>', M.P, 'Base repetible × retención/crecimiento moderado + clientes nuevos');
  else {
    h += row('<b>Proyección con eventos</b>', M.P, 'Base repetible × retención/crecimiento moderado + clientes nuevos + eventos activos');
    h += row('Proyección sin eventos', M.S0adj, 'Lo mismo, con todos los eventos apagados');
  }
  h += row('Forecast del equipo', M.F27, 'Lo que capturaron vendedores y dirección');
  h += row('Forecast calibrado', M.Fcal, M.att != null ? `Forecast × cumplimiento de ${M.yC} a la fecha (${fmtPct(M.attRaw * 100)}, acotado a 70%–140%)` : 'Necesita Forecast del año en curso');
  if (!pxIsExec()) {
    h += row('Estacional + tendencia (histórico 10+ años)', old.totalProjBase, 'Crecimiento promedio año a año del histórico completo, repartido por estacionalidad (método anterior)');
    h += row('Regresión lineal', old.linReg, 'Tendencia lineal simple sobre el histórico completo (contraste)');
  }
  h += '</tbody></table></div>';
  document.getElementById('pxMetodos27').innerHTML = h;
}

// =========================================================================
// PANEL DE SUPUESTOS Y EVENTOS
// =========================================================================
function pxBuildPanel(M) {
  if (pxIsExec()) return;
  const holds = [...M.byHol.entries()].sort((a, b) => b[1].adj - a[1].adj).slice(0, 8).map(e => e[0]);
  const clis = [...M.byCli.entries()].sort((a, b) => b[1].base - a[1].base).slice(0, 6).map(e => e[0]);
  const deps = [...M.depAll.entries()].sort((a, b) => b[1] - a[1]).map(e => e[0]);
  const sig = JSON.stringify([M.yC, M.mk, holds, clis, deps, M.mundialAplica, M.hasFcN]);
  if (sig === pxPanelSig) return;
  pxPanelSig = sig;
  const yN = M.yN;
  const sw = (path, checked) => `<label class="px-sw"><input type="checkbox" data-k="${path}" ${checked ? 'checked' : ''}><span></span></label>`;
  const num = (path, val, step = '0.1', cls = 'sm') => `<input type="number" class="px-in ${cls}" step="${step}" data-k="${path}" value="${val === '' || val == null ? '' : val}">`;
  const ev = (id, on, title, body, note) => `<div class="px-ev ${on ? '' : 'off'}" data-ev="${id}"><div class="px-ev-h">${sw(id + '.on', on)}<div class="px-ev-t">${title}</div><div class="px-ev-imp" id="pxImp-${id}"></div></div><div class="px-ev-b">${body}</div><div class="px-ev-note">${note}</div></div>`;
  const blocks = [];

  blocks.push(`<div class="px-ev" data-ev="base"><div class="px-ev-h"><div class="px-ev-t">📊 Base estadística</div><div class="px-ev-imp" id="pxImp-base"></div></div><div class="px-ev-b">
      <div class="px-row"><span>Persistencia del crecimiento histórico</span><input type="range" min="0" max="100" step="5" data-k="persistencia" value="${PX.persistencia}"><span class="px-val" id="pxVal-persistencia">${PX.persistencia}%</span></div>
      <label class="px-chk"><input type="checkbox" data-k="calibrar" ${PX.calibrar ? 'checked' : ''}> Calibrar el Forecast con el cumplimiento de ${M.yC}</label></div>
      <div class="px-ev-note">El crecimiento observado (retención de cuentas + clientes nuevos) rara vez se repite entero: 100% asume que se repite igual, 0% asume que ${yN} queda en la base repetible. El forecast calibrado lo multiplica por lo que el equipo ha cumplido a la fecha.</div></div>`);

  blocks.push(ev('elec', PX.elec.on, '🗳️ Elecciones intermedias (6 jun 2027)',
    `<div class="px-row"><span>Venta a partidos y gobiernos (% de la proyección)</span>${num('elec.partidos', PX.elec.partidos)}</div>
     <div class="px-row"><span>Marcas desplazadas o que evitan compartir (%)</span>${num('elec.desplaza', PX.elec.desplaza)}</div>
     <div class="px-row"><span>Prima de tarifa por demanda alta (%)</span>${num('elec.prima', PX.elec.prima)}</div>`,
    'Precampañas 4 ene–12 feb, campañas 4 abr–2 jun, veda 3–5 jun. 17 gubernaturas, diputados federales, congresos locales y municipios. El efecto cae sobre ene–jun (más fuerte abr–may). Si hay inventario libre los partidos suman; si está lleno desplazan marcas. Valores iniciales = supuestos a calibrar (en 2024 las ventas claramente políticas fueron ~2.6% de la bruta).'));

  blocks.push(ev('mundial', PX.mundial.on, '⚽ Mundial 2026 (no recurrente)',
    M.mundialAplica
      ? `<div class="px-row"><span>Estimado automático (exceso jun–jul)</span><b>${fmtMoneyShort(M.Wauto)}</b></div>
         <div class="px-row"><span>Corregir a mano (venta neta, $M)</span>${num('mundial.override', PX.mundial.override, '0.1', '')}</div>`
      : `<div class="px-row"><span>Solo aplica cuando el año en curso es 2026 y ya pasó julio.</span></div>`,
    'Quita de la base repetible lo que el torneo infló en junio–julio de 2026. Se estima comparando esos meses contra la estacionalidad de los demás; déjalo vacío para usar el estimado.'));

  blocks.push(ev('salidas', PX.salidas.on, '🚪 Salidas de ejecutivos',
    M.hasFcN
      ? `<div class="px-row"><span>Cartera que retiene el equipo</span><input type="range" min="0" max="100" step="5" data-k="salidas.retencion" value="${PX.salidas.retencion}"><span class="px-val" id="pxVal-retencion">${PX.salidas.retencion}%</span></div>
         ${deps.length ? deps.map(e => `<label class="px-chk"><input type="checkbox" data-sale="${pxEsc(e)}" ${(PX.salidas.noSalen || []).includes(e) ? '' : 'checked'}> ${pxEsc(e)} <span class="px-tag">${fmtMoneyShort(M.depAll.get(e))}</span></label>`).join('') : '<div class="px-row"><span>Todos los vendedores con venta tienen forecast.</span></div>'}`
      : `<div class="px-row"><span>Necesita el Forecast ${yN} cargado.</span></div>`,
    `Vendedores con venta en ${M.yC} que no aparecen en el Forecast ${yN}. Desmarca a quien sí siga (por ejemplo, si cambia de rol).`));

  blocks.push(ev('nuevosEj', PX.nuevosEj.on, '🆕 Ejecutivos nuevos',
    `<div class="px-row"><span>Venta esperada de ejecutivos nuevos (venta neta, $M)</span>${num('nuevosEj.monto', PX.nuevosEj.monto, '0.1', '')}</div>`,
    'Suma a la proyección pero no está en el forecast (colchón). Captura la venta neta anual que esperas de ellos.'));

  blocks.push(ev('macro', PX.macro.on, '📉 Macro: PIB / T-MEC',
    `<div class="px-row"><span>Efecto sobre la proyección</span><input type="range" min="-10" max="10" step="0.5" data-k="macro.pct" value="${PX.macro.pct}"><span class="px-val" id="pxVal-macro">${PX.macro.pct}%</span></div>`,
    'Crecimiento esperado de 2027 ~1.7–2.2% (Banco Mundial, Banxico, FMI); riesgo principal: revisión del T-MEC. Úsalo como escenario: negativo si ves recorte de inversión, positivo si se resuelve bien.'));

  blocks.push(ev('holdings', PX.holdings.on, '🏢 Holdings y agencias',
    holds.map(h => {
      const hint = Object.keys(PX_HOLD_HINTS).find(k => String(h).toLowerCase().includes(k));
      const val = (PX.holdings.pct || {})[h] || 0;
      return `<div class="px-row"><span title="${hint ? pxEsc(PX_HOLD_HINTS[hint]) : ''}">${pxEsc(h)}${hint ? '<span class="px-tag">ⓘ</span>' : ''}</span><input type="range" min="-15" max="15" step="1" data-hold="${pxEsc(h)}" value="${val}"><span class="px-val" id="pxVal-h-${pxKey(h)}">${val}%</span></div>`;
    }).join('') || '<div class="px-row"><span>Sin datos de holding.</span></div>',
    'Ajuste (±%) sobre lo que cada holding aporta a la proyección. Pasa el mouse por ⓘ para ver qué se sabe de cada grupo (Omnicom+IPG integrados, WPP reestructurado, Dentsu en reestructura, etc.).'));

  blocks.push(ev('cuentas', PX.cuentas.on, '🎯 Cuentas clave: ¿y si se pierde una?',
    clis.map(c => `<label class="px-chk"><input type="checkbox" data-cta="${pxEsc(c)}" ${(PX.cuentas.perdidas || []).includes(c) ? 'checked' : ''}> ${pxEsc(c)} <span class="px-tag">${fmtMoneyShort(M.byCli.get(c).base)}</span></label>`).join(''),
    'Las 6 cuentas con más venta en el año en curso. Marca las que quieras simular como perdidas: se resta su proyección.'));

  document.getElementById('pxPanel').innerHTML = `<div class="px-events">${blocks.join('')}</div>
    <div class="px-ev-note" style="margin-top:10px">Eventos recurrentes (Buen Fin, GP de México de F1 el 29–31 oct 2027, Navidad, Semana Santa que cae en marzo) ya están dentro de la estacionalidad histórica: no se suman aparte.</div>`;
}

function pxUpdateImpacts(M) {
  const set = (id, txt, v) => { const el = document.getElementById('pxImp-' + id); if (el) { el.textContent = txt; el.className = 'px-ev-imp ' + (v == null ? '' : pxCls(v)); } };
  set('base', `${fmtMoneyShort(M.S0adj)}`, null);
  set('elec', PX.elec.on ? pxSigned(M.elecTot) : '', M.elecTot);
  set('mundial', PX.mundial.on && M.mundialAplica ? pxSigned(-M.W) : '', -M.W);
  set('salidas', PX.salidas.on && M.hasFcN ? pxSigned(-M.Lsum) : '', -M.Lsum);
  set('nuevosEj', PX.nuevosEj.on ? pxSigned(M.ev.nuevos) : '', M.ev.nuevos);
  set('macro', PX.macro.on ? pxSigned(M.ev.macro) : '', M.ev.macro);
  set('holdings', PX.holdings.on ? pxSigned(M.holdTot) : '', M.holdTot);
  set('cuentas', PX.cuentas.on ? pxSigned(M.ctaTot) : '', M.ctaTot);
}

// =========================================================================
// EVENTOS DE LA INTERFAZ
// =========================================================================
function pxSetPath(path, val) {
  const parts = path.split('.'); let o = PX;
  for (let i = 0; i < parts.length - 1; i++) { if (o[parts[i]] == null || typeof o[parts[i]] !== 'object') o[parts[i]] = {}; o = o[parts[i]]; }
  o[parts[parts.length - 1]] = val;
}
function pxWire() {
  pxWired = true;
  document.getElementById('pxMetrica').addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    PX.metrica = b.dataset.m; pxSave(); pxPanelSig = ''; pxUpdate();
  });
  const meta = (id, key) => document.getElementById(id).addEventListener('input', e => { PX[key] = e.target.value; pxSave(); pxUpdate(); });
  meta('pxMetaC', 'metaCierre'); meta('pxMetaP', 'metaProx');
  document.getElementById('pxReset').addEventListener('click', () => {
    if (!confirm('¿Restablecer todos los supuestos y eventos a sus valores iniciales?')) return;
    const m = PX.metrica; PX = pxDefaults(); PX.metrica = m; pxSave(); pxPanelSig = ''; pxUpdate();
  });
  const onChange = e => {
    const t = e.target; if (!(t instanceof HTMLInputElement)) return;
    if (t.dataset.hold != null && t.dataset.hold !== '') {
      const v = parseFloat(t.value) || 0;
      PX.holdings.pct = { ...(PX.holdings.pct || {}), [t.dataset.hold]: v };
      const lab = document.getElementById('pxVal-h-' + pxKey(t.dataset.hold)); if (lab) lab.textContent = v + '%';
    } else if (t.dataset.sale != null && t.dataset.sale !== '') {
      const set = new Set(PX.salidas.noSalen || []);
      if (t.checked) set.delete(t.dataset.sale); else set.add(t.dataset.sale);
      PX.salidas.noSalen = [...set];
    } else if (t.dataset.cta != null && t.dataset.cta !== '') {
      const set = new Set(PX.cuentas.perdidas || []);
      if (t.checked) set.add(t.dataset.cta); else set.delete(t.dataset.cta);
      PX.cuentas.perdidas = [...set];
    } else if (t.dataset.k) {
      const k = t.dataset.k;
      const val = t.type === 'checkbox' ? t.checked : (t.type === 'number' || t.type === 'range') ? (t.value === '' ? '' : parseFloat(t.value)) : t.value;
      pxSetPath(k, val);
      if (k === 'persistencia') document.getElementById('pxVal-persistencia').textContent = val + '%';
      if (k === 'salidas.retencion') document.getElementById('pxVal-retencion').textContent = val + '%';
      if (k === 'macro.pct') document.getElementById('pxVal-macro').textContent = val + '%';
      if (t.type === 'checkbox' && k.endsWith('.on')) { const box = t.closest('.px-ev'); if (box) box.classList.toggle('off', !t.checked); }
    } else return;
    pxSave(); pxUpdate();
  };
  const panel = document.getElementById('pxPanel');
  panel.addEventListener('input', e => { if (e.target.type !== 'checkbox') onChange(e); });
  panel.addEventListener('change', e => { if (e.target.type === 'checkbox') onChange(e); });
}
