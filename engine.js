// =========================================================================
// engine.js · Motor de datos del Dashboard Comercial Open · v3
// Procesa Data Comercial + Forecast por Vendedor. Detecta hojas por año.
// =========================================================================

// =========================================================================
// Mapeo de columnas del archivo de Datos
// =========================================================================
const COL_MAP = {
  ejecutivo: ['Ejecutivo Comercial', 'Nombre ejecutivo Comercial', 'Nombre Ejecutivo Comercial'],
  empresa: ['Empresa'],
  cliente: ['Cliente', 'Cliente/Marca', 'Cliente / Marca', 'Cliente Comercial'],
  holding: ['Holding'],
  agencia: ['Agencia'],
  proveedor: ['Proveedor Comercial'],
  proveedorRSF: ['Proveedor RSF', 'Proveedor Razón Social'],
  estado: ['Estado'],
  localidad: ['Localidad'],
  campana: ['Campaña', 'Campana'],
  tipoCompra: ['Tipo de Compra'],
  medio: ['Medio'],
  soporte: ['Soporte'],
  centroOAero: ['Centro Comercial o Aeropuerto'],
  mes: ['Mes'],
  anio: ['Año', 'Anio', 'Ano'],
  fechaInicio: ['Fecha Inicio'],
  fechaFin: ['Fecha Fin'],
  ventaBruta: ['Tarifa Total cliente', 'Tarifa Total Cliente'],
  ventaNeta: ['VENTA SIN ROI', 'Venta Sin ROI'],
  totalCostos: ['Total Costos'],
  totalROIs: ['Total ROIs'],
  utilidadLinea: ['Utilidad por línea \n$', 'Utilidad por línea\n$', 'Utilidad por linea $', 'Utilidad por línea $'],
  utilidadOpen: ['Utilidad final\nOPEN', 'Utilidad final OPEN'],
  comisionVendedor: ['Comision final\nVendedor', 'Comision final Vendedor', 'Comisión final Vendedor'],
  comisionADV: ['Comisión ADV', 'Comision ADV'],
  status: ['Status OPEN', 'Status'],
  tipo: ['Tipo'],
  m2: ['Total m2', 'Total M2'],
  odv: ['ODV'],
  odc: ['ODC'],
  // --- Ampliación: columnas adicionales del Excel para análisis más profundo ---
  cm: ['CM'],
  adv: ['ADV'],
  financiamiento: ['Financiamiento'],
  costoUnitProveedor: ['Costo unitario proveedor'],
  totalTarifaCostos: ['Total tarifa de costos'],
  incrementoVentaUnit: ['Incremento Venta Unitario $'],
  pctROIAgencia: ['% \nROI Agencia', '% \n ROI Agencia'],
  montoROIAgencia: ['Monto \nROI Agencia', 'Monto \n ROI Agencia'],
  pctROIPersonal: ['%\nROI Personal', '%\n ROI Personal'],
  montoROIPersonal: ['Monto\nROI Personal', 'Monto\n ROI Personal'],
  pctComisionVendedor: ['%\nComision\nVendedor', '%\n Comision\n Vendedor'],
  gastos: ['Gastos'],
  utilidadTotalProyecto: ['Utilidad Total Proyecto'],
  pctComisionPagar: ['% de Comisión a Pagar'],
};

// =========================================================================
// Mapeo Soporte (2026) → Categoría del Mix
// Cuando Soporte está vacío o ambiguo, fallback a Medio
// =========================================================================
const SOPORTE_TO_CATEGORIA = {
  // Espectaculares
  'unipolar': 'Espectaculares',
  'azotea': 'Espectaculares',
  'bajopuente': 'Espectaculares',
  'estructura de piso': 'Espectaculares',
  'totem': 'Espectaculares',
  'tótem': 'Espectaculares',
  'panoramico': 'Espectaculares',
  'panóramico': 'Espectaculares',
  'panorámico': 'Espectaculares',
  'espectacular': 'Espectaculares',

  // Pantalla
  'espectacular - pantalla': 'Pantalla',
  'totem - espectacular': 'Pantalla',
  'tótem - espectacular': 'Pantalla',
  'pantalla digital': 'Pantalla',  // Default si Medio no especifica otro
  'videowall': 'Pantalla',
  'video wall': 'Pantalla',

  // Take Over
  'take over - muro': 'Take Over',
  'take over - pantalla': 'Take Over',
  'take over - valla': 'Take Over',
  'take over - espectacular': 'Take Over',
  'take over - mixto': 'Take Over',

  // Murales
  'muro': 'Murales',
  'mural': 'Murales',
  'vinil': 'Murales',
  'barda': 'Murales',
  'barda pintada': 'Murales',
  'mapping': 'Murales',
  'pintado': 'Murales',
  'espectacular - muro': 'Murales',

  // Vallas
  'fija': 'Vallas',
  'movil': 'Vallas',
  'móvil': 'Vallas',
  'megavalla': 'Vallas',
  'fija - digital': 'Vallas',
  'movil-digital': 'Vallas',
  'móvil-digital': 'Vallas',
  'valla digital': 'Vallas',
  'valla': 'Vallas',
  'valla móvil': 'Vallas',
  'valla movil': 'Vallas',

  // Puentes
  'puente': 'Puentes',
  'puentes': 'Puentes',

  // Transporte Público
  'integral': 'Transporte Público',
  'semintegral': 'Transporte Público',
  'semi integral': 'Transporte Público',
  'medio medallon': 'Transporte Público',
  'medio medallón': 'Transporte Público',
  'medellón': 'Transporte Público',
  'medellon': 'Transporte Público',
  'madallon': 'Transporte Público',
  'medallón': 'Transporte Público',
  'camion': 'Transporte Público',
  'camión': 'Transporte Público',
  'camion urbano': 'Transporte Público',

  // Aeropuertos (cuando Medio lo confirma)
  'mupie': 'Aeropuertos',
  'cartelera': 'Aeropuertos',
  'mega pantalla': 'Aeropuertos',

  // Centros Comerciales
  'activacion': 'Centros Comerciales',
  'activación': 'Centros Comerciales',
  'carteleras': 'Centros Comerciales',
  'pendones': 'Centros Comerciales',
  'escaleras': 'Centros Comerciales',
  'elevadores': 'Centros Comerciales',
  'estacionamiento': 'Centros Comerciales',

  // Producción
  'impresion': 'Producción',
  'impresión': 'Producción',
  'impresos': 'Producción',
  'instalacion': 'Producción',
  'instalación': 'Producción',
  'instalaciones': 'Producción',
  'retiros': 'Producción',
  'punto de banda': 'Producción',
  'toques especiales': 'Producción',
  'implementacion': 'Producción',
  'implementación': 'Producción',
  'implementaciones': 'Producción',
  'implementación especial': 'Producción',
  'producción': 'Producción',
  'lona': 'Producción',
  'monitoreo': 'Producción',
  'cuadruple': 'Producción',
  'cuádruple': 'Producción',

  // Mobiliario Urbano
  'para periodicos': 'Mobiliario Urbano',
  'para periódicos': 'Mobiliario Urbano',
  'telefonos': 'Mobiliario Urbano',
  'teléfonos': 'Mobiliario Urbano',
  'pantalla interior': 'Mobiliario Urbano',

  // Paradas de Autobús
  'pister': 'Paradas de Autobús',
  'paradas de autobus': 'Paradas de Autobús',
  'paradas de autobús': 'Paradas de Autobús',

  // Vitrina Móvil
  'vitrinamovil': 'Vitrina Móvil',
  'vitrina movil': 'Vitrina Móvil',
  'vitrinamóvil': 'Vitrina Móvil',
  'vitrina móvil': 'Vitrina Móvil',
};

// Mapeo de Medio (cuando se usa como categoría primaria - 2025 o fallback)
const MEDIO_TO_CATEGORIA = {
  'espectacular': 'Espectaculares',
  'espectaculares': 'Espectaculares',
  'pantalla digital': 'Pantalla',
  'pantalla': 'Pantalla',
  'pantallas': 'Pantalla',
  'valla móvil': 'Vallas',
  'valla movil': 'Vallas',
  'vallas': 'Vallas',
  'valla': 'Vallas',
  'aeropuertos': 'Aeropuertos',
  'aeropuerto': 'Aeropuertos',
  'camion': 'Transporte Público',
  'camión': 'Transporte Público',
  'camion urbano': 'Transporte Público',
  'camión urbano': 'Transporte Público',
  'transporte público': 'Transporte Público',
  'transporte publico': 'Transporte Público',
  'muro': 'Murales',
  'murales': 'Murales',
  'mural': 'Murales',
  'impresión': 'Producción',
  'impresion': 'Producción',
  'instalación': 'Producción',
  'instalacion': 'Producción',
  'producción': 'Producción',
  'produccion': 'Producción',
  'monitoreo': 'Producción',
  'lona': 'Producción',
  'cuadruple': 'Producción',
  'barda pintada': 'Murales',
  'vitrina móvil': 'Vitrina Móvil',
  'vitrina movil': 'Vitrina Móvil',
  'puentes': 'Puentes',
  'puente': 'Puentes',
  'centros comerciales': 'Centros Comerciales',
  'centro comercial': 'Centros Comerciales',
  'mobiliario urbano': 'Mobiliario Urbano',
  'paradas de autobús': 'Paradas de Autobús',
  'paradas de autobus': 'Paradas de Autobús',
};

// Categorías oficiales del mix (orden de visualización)
const CATEGORIAS_MIX = [
  'Espectaculares', 'Pantalla', 'Murales', 'Vallas', 'Take Over',
  'Transporte Público', 'Centros Comerciales', 'Aeropuertos',
  'Mobiliario Urbano', 'Paradas de Autobús', 'Puentes',
  'Vitrina Móvil', 'Producción', 'Otros'
];

function categorizar(medio, soporte) {
  // Prioridad 1: Si Medio especifica un canal de "destino" como Aeropuertos / C. Comerciales /
  // Mobiliario Urbano / Paradas / Vitrina Móvil, usar Medio sin importar Soporte
  // (porque "Pantalla Digital" en un Aeropuerto debe quedar en Aeropuertos).
  const medioLow = medio ? String(medio).toLowerCase().trim() : '';
  const canalesDestino = ['aeropuerto','aeropuertos','centros comerciales','centro comercial','mobiliario urbano','paradas de autobús','paradas de autobus','vitrina móvil','vitrina movil'];
  if (canalesDestino.includes(medioLow)) {
    return MEDIO_TO_CATEGORIA[medioLow] || medio;
  }
  // Prioridad 2: Si Medio = "Vallas", devuelve "Vallas" (porque incluye varios subtipos)
  if (medioLow === 'vallas' || medioLow === 'valla') return 'Vallas';
  // Prioridad 3: si Medio = "Transporte Público" / "Murales" / "Producción" → mantener (subcategorías van a la macro)
  if (['transporte público','transporte publico'].includes(medioLow)) return 'Transporte Público';
  if (['murales','mural'].includes(medioLow)) {
    // Pero si Soporte es Muro (el espectacular tipo muro), va a Murales también
    return 'Murales';
  }
  if (['producción','produccion'].includes(medioLow)) return 'Producción';
  if (['puente','puentes'].includes(medioLow)) return 'Puentes';

  // Prioridad 4: usar Soporte
  const sopLow = soporte ? String(soporte).toLowerCase().trim() : '';
  if (sopLow && SOPORTE_TO_CATEGORIA[sopLow]) return SOPORTE_TO_CATEGORIA[sopLow];

  // Prioridad 5: usar Medio como categoría
  if (medioLow && MEDIO_TO_CATEGORIA[medioLow]) return MEDIO_TO_CATEGORIA[medioLow];

  return 'Otros';
}

// =========================================================================
// Helpers de normalización
// =========================================================================
function stripAccents(s) {
  if (s == null) return '';
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '');
}
function titleCase(s) {
  if (s == null || s === '') return s;
  return String(s).trim().split(/\s+/).map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
}
function normEstado(e) {
  if (e == null || e === '') return 'Sin definir';
  const t = titleCase(e);
  const map = {
    'Cdmx':'CDMX','Ciudad De México':'CDMX','Ciudad De Mexico':'CDMX',
    'Nuevo Léon':'Nuevo León','Edo. De México':'Estado de México','Edo De México':'Estado de México',
    'Estado De México':'Estado de México'
  };
  return map[t] || t;
}
function normLocalidad(l) {
  if (l == null || l === '') return 'Sin definir';
  const t = titleCase(l);
  const map = {'Cdmx':'CDMX','Ciudad De México':'CDMX','Ciudad De Mexico':'CDMX'};
  return map[t] || t;
}
// Normaliza nombre de vendedor para matching (data <-> forecast)
function normEjecutivo(n) {
  if (n == null || n === '') return 'Sin asignar';
  const t = titleCase(n).trim();
  // Mapeos manuales para casos donde data y forecast difieren en acento o nombre
  const map = {
    'Ricardo Ramirez': 'Ricardo Ramírez',
    'Amanda Martinez': 'Amanda Martínez',
    'Hector Muñoz': 'Héctor Muñoz',
    'Efrain Guadalupe': 'Efraín Guadalupe',
    'Fernando Garcia': 'Fernando García',
    'Miguel Garcia': 'Miguel García',
    'Viani Paz': 'Viani Paz',
    'Viani': 'Viani Paz', // En forecast está como "Viani" sin apellido
  };
  return map[t] || t;
}
function num(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  const cleaned = String(v).replace(/[$,\s]/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}
function int(v) { return Math.round(num(v)); }

// =========================================================================
// FUZZY MATCHING de marcas
// fuzzyKey() genera una clave normalizada (lowercase, sin acentos, sin espacios extras)
// para detectar duplicados como "Bacardí" / "Bacardi" / "BACARDI"
// =========================================================================
function fuzzyKey(s) {
  if (s == null || s === '') return '';
  let t = stripAccents(String(s)).toLowerCase().trim();
  // Remover puntuación común
  t = t.replace(/[.,;:'"`´’]/g, '');
  // Colapsar múltiples espacios
  t = t.replace(/\s+/g, ' ');
  return t;
}
// Levenshtein distance para casos avanzados (similitud >= 85%)
function levenshtein(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  const m = a.length, n = b.length;
  let prev = Array(n+1).fill(0).map((_,i)=>i);
  let curr = Array(n+1).fill(0);
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i-1] === b[j-1] ? 0 : 1;
      curr[j] = Math.min(curr[j-1]+1, prev[j]+1, prev[j-1]+cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[n];
}
function similarity(a, b) {
  const maxLen = Math.max(a.length, b.length);
  if (!maxLen) return 1;
  return 1 - levenshtein(a, b) / maxLen;
}
// Construye un mapa fuzzy: nombre_original → clave_canónica
// Si dos nombres tienen similitud >= 0.85, se consolidan al más largo (más completo)
function buildFuzzyMap(names) {
  const unique = [...new Set(names.filter(n => n && n.trim()))];
  // Primero agrupar por fuzzyKey (acentos/mayúsculas/espacios)
  const byKey = {};
  unique.forEach(n => {
    const k = fuzzyKey(n);
    if (!byKey[k]) byKey[k] = [];
    byKey[k].push(n);
  });
  // En cada grupo, escoger el "canónico" (el más común o el más largo)
  const canon = {};
  Object.values(byKey).forEach(group => {
    const best = group.sort((a,b) => b.length - a.length)[0]; // más largo
    group.forEach(g => canon[g] = best);
  });
  return canon;
}

// =========================================================================
// Detección de columnas
// =========================================================================
function buildColumnMapping(headerRow) {
  const mapping = {};
  for (const std in COL_MAP) {
    for (const candidate of COL_MAP[std]) {
      const candNorm = candidate.replace(/\s+/g, ' ').trim().toLowerCase();
      const found = headerRow.find(h => {
        if (h == null) return false;
        return String(h).replace(/\s+/g, ' ').trim().toLowerCase() === candNorm;
      });
      if (found) { mapping[std] = found; break; }
    }
  }
  return mapping;
}

// =========================================================================
// Procesar Excel de Datos (hojas con nombre de año)
// =========================================================================
function processDataWorkbook(workbook) {
  const records = [];
  const yearlySheets = [];
  workbook.SheetNames.forEach(name => {
    const m = name.match(/^(\d{4})$/);
    if (m) yearlySheets.push({ name, year: parseInt(m[1]) });
  });
  yearlySheets.sort((a,b) => a.year - b.year);

  yearlySheets.forEach(({name, year}) => {
    const sheet = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
    if (rows.length === 0) return;
    const headers = Object.keys(rows[0]);
    const mapping = buildColumnMapping(headers);

    rows.forEach(row => {
      const get = (k) => mapping[k] ? row[mapping[k]] : null;
      const vb = num(get('ventaBruta'));
      const status = get('status');
      if (!status || (vb === 0 && num(get('ventaNeta')) === 0)) return;

      const mes = int(get('mes'));
      const anio = year;
      const ejecutivo = normEjecutivo(get('ejecutivo'));
      const cliente = (get('cliente') || 'Sin definir').toString().trim();
      const holding = (get('holding') || 'Sin definir').toString().trim();
      const agencia = (get('agencia') || 'Sin definir').toString().trim();
      const empresa = (get('empresa') || 'Sin definir').toString().trim();
      const medio = get('medio');
      const soporte = get('soporte');

      const proveedor = (get('proveedor') || 'Sin definir').toString().trim();

      const rec = {
        anio, mes,
        emp: empresa,
        eje: ejecutivo,
        cli: titleCase(cliente),
        cliKey: fuzzyKey(cliente),     // clave normalizada para fuzzy matching
        hol: titleCase(holding),
        age: titleCase(agencia),
        prov: titleCase(proveedor),
        est: normEstado(get('estado')),
        loc: normLocalidad(get('localidad')),
        med: medio || 'Sin definir',
        sop: soporte || '',
        cat: categorizar(medio, soporte),
        tp: get('tipo') || 'Sin definir',
        tc: get('tipoCompra') || 'Sin definir',
        st: status,
        cmp: get('campana') || '',
        odv: get('odv') || '',
        cc: (holding.toLowerCase() === 'directo' || agencia.toLowerCase() === 'directo') ? 'Directo' : 'Agencia',
        // Métricas
        vb: vb,
        vn: num(get('ventaNeta')),
        roi: num(get('totalROIs')),
        cst: num(get('totalCostos')),
        ut: num(get('utilidadLinea')),
        utOpen: num(get('utilidadOpen')),
        com: num(get('comisionVendedor')),
        comAdv: num(get('comisionADV')),
        m2: num(get('m2')),
        // --- Ampliación ---
        cm: (get('cm') || '').toString().trim(),
        adv: titleCase(get('adv') || ''),
        provRSF: (get('proveedorRSF') || '').toString().trim(),
        fechaInicio: get('fechaInicio') || null,
        fechaFin: get('fechaFin') || null,
        financiamiento: (get('financiamiento') || '').toString().trim(),
        costoUnitProveedor: num(get('costoUnitProveedor')),
        totalTarifaCostos: num(get('totalTarifaCostos')),
        incrementoVentaUnit: num(get('incrementoVentaUnit')),
        pctROIAgencia: num(get('pctROIAgencia')),
        montoROIAgencia: num(get('montoROIAgencia')),
        pctROIPersonal: num(get('pctROIPersonal')),
        montoROIPersonal: num(get('montoROIPersonal')),
        pctComisionVendedor: num(get('pctComisionVendedor')),
        gastos: num(get('gastos')),
        utilidadTotalProyecto: num(get('utilidadTotalProyecto')),
        pctComisionPagar: num(get('pctComisionPagar')),
        // Fila cruda completa (todas las columnas originales del Excel, incluidas las que no
        // mapeamos a un campo limpio) — permite que el asistente IA consulte literalmente
        // cualquier columna del archivo, no solo las que decidimos nombrar.
        _raw: row,
      };
      records.push(rec);
    });
  });

  return records;
}

// =========================================================================
// Procesar Excel de Forecast (vendedor + cliente + 12 meses)
// =========================================================================
function processForecastWorkbook(workbook, year=2026) {
  const records = [];
  const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  workbook.SheetNames.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
    rows.forEach(r => {
      const eje = normEjecutivo(r['Nuevo ejecutivo'] || r['Ejecutivo'] || r['Vendedor']);
      const cli = titleCase(r['Cliente Comercial'] || r['Cliente'] || '');
      if (!eje || !cli) return;
      meses.forEach((mes, idx) => {
        const v = num(r[mes]);
        if (v > 0) records.push({
          anio: year,
          mes: idx + 1,
          eje, cli,
          fcNeto: v,
        });
      });
    });
  });
  return records;
}

// =========================================================================
// Estado global
// =========================================================================
const Engine = {
  records: [],
  forecast: [],
  yearsAvailable: [],
  metaData: null,
  metaForecast: null,
  // % ROI promedio por año (para calcular forecast Bruto)
  roiPromAnioAnterior: null,

  load(records, fileMeta) {
    this.records = records;
    this.yearsAvailable = [...new Set(records.map(r => r.anio))].sort();
    this.metaData = fileMeta || null;
    this._applyFuzzyMatchingClientes();
    this._computeRoiPromedio();
    this._computeActiveSellers();
    this._computeStatusValues();
  },
  loadForecast(forecast, fileMeta) {
    this.forecast = forecast;
    this.metaForecast = fileMeta || null;
    this._applyFuzzyMatchingForecast();
    this._computeActiveSellers();
  },
  // Aplica fuzzy matching a clientes en records y forecast (canonicaliza nombres similares)
  _applyFuzzyMatchingClientes() {
    // Construir mapa fuzzy con TODOS los nombres (records + forecast)
    const all = this.records.map(r => r.cli).concat(this.forecast.map(f => f.cli));
    this.cliCanonMap = buildFuzzyMap(all);
    this.records.forEach(r => {
      r.cliOrig = r.cli;
      r.cli = this.cliCanonMap[r.cli] || r.cli;
    });
    // También consolidar proveedores con fuzzy match
    this._applyFuzzyMatchingProveedores();
  },
  // Fuzzy match para proveedores (mismo algoritmo que clientes)
  _applyFuzzyMatchingProveedores() {
    const provs = this.records.map(r => r.prov).filter(p => p && p !== 'Sin definir');
    this.provCanonMap = buildFuzzyMap(provs);
    this.records.forEach(r => {
      if (!r.prov || r.prov === 'Sin definir') return;
      r.provOrig = r.prov;
      r.prov = this.provCanonMap[r.prov] || r.prov;
    });
  },
  _applyFuzzyMatchingForecast() {
    // Re-aplicar también si hay records cargados
    const all = this.records.map(r => r.cliOrig || r.cli).concat(this.forecast.map(f => f.cli));
    this.cliCanonMap = buildFuzzyMap(all);
    this.records.forEach(r => {
      r.cli = this.cliCanonMap[r.cliOrig || r.cli] || r.cli;
    });
    this.forecast.forEach(f => {
      f.cliOrig = f.cli;
      f.cli = this.cliCanonMap[f.cli] || f.cli;
    });
  },
  // Vendedores "activos" en 2026 = los que tienen forecast 2026 cargado
  _computeActiveSellers() {
    const sellers = new Set(this.forecast.filter(f => f.anio === 2026).map(f => f.eje));
    this.activeSellers2026 = sellers;
  },
  // Valores únicos de Status OPEN encontrados en los datos
  _computeStatusValues() {
    const s = new Set();
    this.records.forEach(r => { if (r.st) s.add(r.st); });
    // Orden preferido: Suma Forecast primero, luego el resto alfabético
    const arr = [...s];
    const preferred = ['Suma Forecast', 'Intercambio', 'Cancelado', 'Os'];
    const ord = preferred.filter(p => arr.includes(p));
    arr.filter(x => !preferred.includes(x)).sort().forEach(x => ord.push(x));
    this.statusOpenValues = ord;
  },
  isActive2026(eje) {
    return !this.activeSellers2026 || this.activeSellers2026.size === 0
      ? true  // si no hay forecast cargado, todos cuentan
      : this.activeSellers2026.has(eje);
  },
  _computeRoiPromedio() {
    // ROI promedio por año = Total ROIs / Venta Bruta (solo facturable + Suma Forecast)
    // Siempre se calcula sobre Suma Forecast para que el cálculo de Forecast Bruto sea consistente
    this.roiPromPorAnio = {};
    this.yearsAvailable.forEach(y => {
      const fact = this.records.filter(r => r.anio===y && r.tc==='Facturable' && r.st==='Suma Forecast');
      const t = this.totalize(fact);
      this.roiPromPorAnio[y] = t.vb > 0 ? t.roi / t.vb : 0;
    });
  },
  // Para 2026, el ROI% del año anterior es el de 2025
  roiPctParaForecastBruto(anio=2026) {
    const prev = anio - 1;
    return this.roiPromPorAnio?.[prev] ?? 0.165; // fallback 16.5%
  },
  forecastBrutoFromNeto(neto, anio=2026) {
    const roiPct = this.roiPctParaForecastBruto(anio);
    if (roiPct <= 0 || roiPct >= 1) return neto;
    return neto / (1 - roiPct);
  },

  // Filtros
  applyFilters(filters, opts={}) {
    const sOpen = Array.isArray(filters.statusOpen) && filters.statusOpen.length ? filters.statusOpen : null;
    return this.records.filter(r => {
      if (!opts.ignoreYear && filters.anio && r.anio != filters.anio) return false;
      if (filters.mes && r.mes != filters.mes) return false;
      if (filters.emp && r.emp !== filters.emp) return false;
      if (filters.eje && r.eje !== filters.eje) return false;
      if (filters.est && r.est !== filters.est) return false;
      if (filters.loc && r.loc !== filters.loc) return false;
      if (filters.cat && r.cat !== filters.cat) return false;
      if (filters.med && r.med !== filters.med) return false;
      if (filters.hol && r.hol !== filters.hol) return false;
      if (filters.cli && r.cli !== filters.cli) return false;
      if (filters.tp && r.tp !== filters.tp) return false;
      if (filters.cc && r.cc !== filters.cc) return false;
      if (sOpen && !sOpen.includes(r.st)) return false;
      return true;
    });
  },
  // Forecast filtrado igual que datos (vendedor, cliente, mes, año)
  applyForecastFilters(filters, opts={}) {
    return this.forecast.filter(f => {
      if (!opts.ignoreYear && filters.anio && f.anio != filters.anio) return false;
      if (filters.mes && f.mes != filters.mes) return false;
      if (filters.eje && f.eje !== filters.eje) return false;
      if (filters.cli && f.cli !== filters.cli) return false;
      return true;
    });
  },

  // Devuelve los registros "vendibles" según el filtro Status OPEN aplicado.
  // El filtro de Status OPEN ya viene aplicado desde applyFilters() (default ['Suma Forecast']),
  // por lo que aquí solo confirmamos que tienen valores monetarios.
  // Para "Suma Forecast" eso equivale a tc='Facturable' (Bonificable tiene VB=0).
  // Para "Intercambio" incluye tc='Intercambio' o 'Bonificable' que sí tienen monto.
  facturable(data) {
    return data;
  },

  totalize(data) {
    const t = { vb:0, vn:0, roi:0, cst:0, ut:0, com:0, comAdv:0, n:0, cli: new Set(), odv: new Set() };
    data.forEach(r => {
      t.vb += r.vb; t.vn += r.vn; t.roi += r.roi; t.cst += r.cst;
      t.ut += r.ut; t.com += r.com; t.comAdv += r.comAdv; t.n += 1;
      if (r.cli) t.cli.add(r.cli);
      if (r.odv) t.odv.add(r.odv);
    });
    t.margen = t.vn ? (t.ut/t.vn*100) : 0;
    return t;
  },

  // Agrupa por cualquier campo del registro y suma cualquier lista de métricas numéricas.
  // Genérico (a diferencia de groupBy, que solo suma el set fijo de métricas del dashboard) —
  // pensado para el asistente IA, que necesita poder pedir combinaciones dimensión×métrica arbitrarias.
  aggregateBy(data, dimensionKey, metricKeys) {
    const map = {};
    data.forEach(r => {
      const k = r[dimensionKey] || 'Sin definir';
      if (!map[k]) {
        map[k] = { key: k, n: 0 };
        metricKeys.forEach(m => { map[k][m] = 0; });
      }
      map[k].n += 1;
      metricKeys.forEach(m => { map[k][m] += (r[m] || 0); });
    });
    return Object.values(map);
  },

  groupBy(data, key) {
    const map = {};
    data.forEach(r => {
      const k = r[key] || 'Sin definir';
      if (!map[k]) map[k] = { key:k, vb:0, vn:0, roi:0, cst:0, ut:0, com:0, comAdv:0, n:0, cli:new Set(), med:new Set(), loc:new Set(), cat:new Set() };
      const o = map[k];
      o.vb += r.vb; o.vn += r.vn; o.roi += r.roi; o.cst += r.cst;
      o.ut += r.ut; o.com += r.com; o.comAdv += r.comAdv; o.n += 1;
      if (r.cli) o.cli.add(r.cli);
      if (r.med) o.med.add(r.med);
      if (r.loc) o.loc.add(r.loc);
      if (r.cat) o.cat.add(r.cat);
    });
    return Object.values(map).map(o => ({
      ...o,
      nClientes: o.cli.size, nMedios: o.med.size, nLoc: o.loc.size, nCategorias: o.cat.size,
      margen: o.vn ? (o.ut/o.vn*100) : 0,
    }));
  },

  monthly(data, valKey='vb') {
    const arr = Array(12).fill(0);
    data.forEach(r => { if (r.mes >= 1 && r.mes <= 12) arr[r.mes-1] += r[valKey] || 0; });
    return arr;
  },
  forecastMonthlyByEje(eje, anio=2026) {
    const arr = Array(12).fill(0);
    this.forecast.filter(f => f.anio===anio && f.eje===eje).forEach(f => {
      if (f.mes>=1 && f.mes<=12) arr[f.mes-1] += f.fcNeto;
    });
    return arr;
  },
  forecastMonthlyTotal(anio=2026) {
    const arr = Array(12).fill(0);
    this.forecast.filter(f => f.anio===anio).forEach(f => {
      if (f.mes>=1 && f.mes<=12) arr[f.mes-1] += f.fcNeto;
    });
    return arr;
  },
  forecastMonthlyByCli(cli, anio=2026) {
    const arr = Array(12).fill(0);
    this.forecast.filter(f => f.anio===anio && f.cli===cli).forEach(f => {
      if (f.mes>=1 && f.mes<=12) arr[f.mes-1] += f.fcNeto;
    });
    return arr;
  },
  forecastMonthlyByEjeCli(eje, cli, anio=2026) {
    const arr = Array(12).fill(0);
    this.forecast.filter(f => f.anio===anio && f.eje===eje && f.cli===cli).forEach(f => {
      if (f.mes>=1 && f.mes<=12) arr[f.mes-1] += f.fcNeto;
    });
    return arr;
  },
  // Conversión array mensual -> trimestral [Q1, Q2, Q3, Q4]
  monthlyToQuarterly(monthly) {
    return [
      monthly.slice(0,3).reduce((a,b)=>a+b,0),
      monthly.slice(3,6).reduce((a,b)=>a+b,0),
      monthly.slice(6,9).reduce((a,b)=>a+b,0),
      monthly.slice(9,12).reduce((a,b)=>a+b,0),
    ];
  },
  // Devuelve venta del año anterior para una key (cliente, vendedor, etc.)
  // statusFilter es opcional, default ['Suma Forecast'] para consistencia con vista actual
  ventaAnioAnterior(filterFn, anio, valKey='vn', statusFilter=['Suma Forecast']) {
    const yPrev = anio - 1;
    const sOpen = Array.isArray(statusFilter) && statusFilter.length ? statusFilter : null;
    return this.records.filter(r => r.anio === yPrev && r.tc === 'Facturable' && (!sOpen || sOpen.includes(r.st)) && filterFn(r))
      .reduce((a,r) => a + (r[valKey]||0), 0);
  },
  // Mensual venta año anterior para un cliente (todas las dimensiones)
  ventaMensualAnioAnterior(cli, anio=2026, valKey='vn', statusFilter=['Suma Forecast']) {
    const arr = Array(12).fill(0);
    const yPrev = anio - 1;
    const sOpen = Array.isArray(statusFilter) && statusFilter.length ? statusFilter : null;
    this.records.filter(r => r.anio === yPrev && r.cli === cli && r.tc === 'Facturable' && (!sOpen || sOpen.includes(r.st))).forEach(r => {
      if (r.mes>=1 && r.mes<=12) arr[r.mes-1] += r[valKey] || 0;
    });
    return arr;
  },
  // Vendedores activos en 2026 (con forecast)
  vendedoresActivos2026() {
    return this.activeSellers2026 ? [...this.activeSellers2026].sort() : [];
  },
};

window.Engine = Engine;
window.processDataWorkbook = processDataWorkbook;
window.processForecastWorkbook = processForecastWorkbook;
window.CATEGORIAS_MIX = CATEGORIAS_MIX;
