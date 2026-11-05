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
  // --- Controles Maestros: columnas de detalle por sitio/línea ---
  idSitio: ['Id de sitio', 'ID de sitio', 'Id Sitio'],
  domicilio: ['Ubicación y/o Ruta', 'Ubicacion y/o Ruta'],
  nombreRoiPersonal: ['Nombre ROI Personal'],
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
      // El año real de cada renglón viene de su propia columna "Año" — no del nombre de la
      // pestaña. Es común que una pestaña "2026" incluya renglones de campañas multi-año
      // (ej. un contrato firmado en 2026 con servicio programado hasta 2027/2028) — usar el
      // nombre de la pestaña para todos los renglones los cuenta erróneamente como venta del
      // año de la pestaña. Si un renglón no trae su propia columna Año, se usa el nombre de la
      // pestaña como respaldo.
      const rowAnio = int(get('anio'));
      const anio = rowAnio > 0 ? rowAnio : year;
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
        idSitio: (get('idSitio') || '').toString().trim(),
        domicilio: (get('domicilio') || '').toString().trim(),
        nombreRoiPersonal: (get('nombreRoiPersonal') || '').toString().trim(),
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
  const warnings = [];
  const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  workbook.SheetNames.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
    rows.forEach(r => {
      const eje = normEjecutivo(r['Nuevo ejecutivo'] || r['Ejecutivo'] || r['Vendedor']);
      const cli = titleCase(r['Cliente Comercial'] || r['Cliente'] || '');
      if (!eje || !cli) return;
      const hol = titleCase(r['Holding'] || '') || 'Sin definir';
      const age = titleCase(r['Agencia'] || '') || 'Sin definir';
      meses.forEach((mes, idx) => {
        const raw = r[mes];
        const v = num(raw);
        // Si la celda trae algo (no está vacía) pero no es un número real y no truena como 0
        // explícito (ej. un error de fórmula tipo #REF!/#N/A), num() la convierte en 0 en
        // silencio — eso hace que ese mes "desaparezca" del forecast sin ningún aviso. Se
        // detecta aquí y se reporta como warning en vez de dejarlo pasar como si no hubiera
        // forecast ese mes.
        if (v === 0 && raw != null && raw !== '' && typeof raw !== 'number' && !/^\s*0(\.0+)?\s*$/.test(String(raw))) {
          warnings.push(`${eje} · ${cli} · ${mes}: valor no numérico "${raw}"`);
        }
        if (v > 0) records.push({
          anio: year,
          mes: idx + 1,
          eje, cli, hol, age,
          fcNeto: v,
        });
      });
    });
  });
  return { records, warnings };
}

// =========================================================================
// Procesar Excel de Contactos (Relaciones Públicas)
// =========================================================================
function processContactosWorkbook(workbook) {
  const records = [];
  const normHeader = s => String(s || '').trim().toLowerCase();
  // "Sí"/"Si"/"SI"/"X" cuentan como verdadero; "No" (u otro texto) o vacío como falso —
  // a diferencia de "cualquier texto no vacío = verdadero", que marcaría "No" como sí.
  const isAfirmativo = v => {
    const t = String(v || '').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
    return t === 'si' || t === 'x' || t === 'yes' || t === 'true' || t === '1';
  };
  workbook.SheetNames.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    // Se lee como filas crudas (sin asumir encabezados en el renglón 1) porque la hoja
    // puede traer un renglón/banner arriba de los encabezados reales (ej. el nombre de
    // una vista filtrada de Google Sheets) — se busca la fila que de verdad contiene la
    // columna "Contacto" entre las primeras 10, y esa se usa como encabezado real.
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    if (!raw.length) return;
    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(10, raw.length); i++) {
      if ((raw[i] || []).some(cell => normHeader(cell) === 'contacto')) { headerRowIdx = i; break; }
    }
    if (headerRowIdx === -1) return; // esta hoja no tiene la tabla de contactos

    const colIdx = {};
    (raw[headerRowIdx] || []).forEach((h, i) => { colIdx[normHeader(h)] = i; });
    const get = (row, name) => {
      const idx = colIdx[normHeader(name)];
      return idx == null ? null : row[idx];
    };

    for (let i = headerRowIdx + 1; i < raw.length; i++) {
      const row = raw[i];
      if (!row) continue;
      const contacto = titleCase(get(row, 'Contacto') || '');
      if (!contacto) continue; // fila vacía / sin persona
      records.push({
        cliente: titleCase(get(row, 'Cliente') || '') || 'Sin definir',
        aaa: isAfirmativo(get(row, 'AAA')),
        clienteActual: isAfirmativo(get(row, 'Cliente actual')),
        agencia: titleCase(get(row, 'Agencia') || '') || 'Sin definir',
        holding: titleCase(get(row, 'Holding') || '') || 'Sin definir',
        eje: normEjecutivo(get(row, 'Ejecutivo Open') || get(row, 'Ejecutivo') || ''),
        contacto,
        email: String(get(row, 'Email') || '').trim(),
        puesto: String(get(row, 'Puesto de trabajo') || '').trim(),
        perfil: String(get(row, 'Perfil de persona') || '').trim(),
        cumpleanos: get(row, 'Cumpleaños') || null,
      });
    }
  });
  return records;
}

// =========================================================================
// Excelencia Operativa — hoja "Hoja de Control": quién atiende cada campaña
// (Coordinador / Peso / Cliente AAA) y la capacidad por coordinador. Se lee con
// la Sheets API en valores SIN formato: números llegan como número y las fechas
// como serial de Excel, así que no dependen del idioma de la hoja ("4,5", "1/01/2026").
// =========================================================================
// Texto comparable: sin acentos, minúsculas, sin puntuación ("O'REILLY" == "O’reilly").
function normTexto(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}
// Distancia de edición (Levenshtein) con corte rápido; sirve para detectar erratas en nombres.
function distanciaEdicion(a, b) {
  if (a === b) return 0;
  if (Math.abs(a.length - b.length) > 2) return 99;
  let prev = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = cur;
  }
  return prev[b.length];
}
function fechaOperaciones(v) {
  let d = null;
  if (typeof v === 'number') d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
  else if (v instanceof Date) d = v;
  else if (typeof v === 'string') {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(v.trim());
    if (m) d = new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  }
  // Años absurdos (un "14/10/20026" tecleado mal en la hoja) se descartan en vez de
  // romper el cruce o la duración.
  return d && !isNaN(d) && d.getUTCFullYear() >= 2000 && d.getUTCFullYear() <= 2100 ? d : null;
}
function numOperaciones(v) {
  if (typeof v === 'number') return v;
  if (v == null || v === '') return null;
  const n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? null : n;
}
function processOperacionesValues(values) {
  const out = { rows: [], capacidad: [] };
  if (!values || !values.length) return out;
  // La fila de encabezados no está en el renglón 1 (arriba vive la tabla de capacidad):
  // se busca la que trae "Campaña" y "Coordinador".
  let headerIdx = -1;
  for (let i = 0; i < Math.min(30, values.length); i++) {
    const hs = (values[i] || []).map(normTexto);
    if (hs.includes('campana') && hs.includes('coordinador')) { headerIdx = i; break; }
  }
  if (headerIdx === -1) return out;
  const colIdx = {};
  (values[headerIdx] || []).forEach((h, i) => { colIdx[normTexto(h)] = i; });
  const get = (row, name) => { const idx = colIdx[normTexto(name)]; return idx == null ? null : row[idx]; };

  // Tabla de capacidad (arriba del encabezado): primer texto de la fila = coordinador,
  // último número de la fila = capacidad. "Total Sistema" se marca aparte.
  for (let i = 0; i < headerIdx; i++) {
    const row = values[i] || [];
    const nombre = row.find(c => typeof c === 'string' && c.trim() !== '');
    const nums = row.filter(c => typeof c === 'number');
    if (!nombre || !nums.length) continue;
    out.capacidad.push({ nombre: nombre.trim(), capacidad: nums[nums.length - 1], esTotal: normTexto(nombre).startsWith('total') });
  }

  for (let i = headerIdx + 1; i < values.length; i++) {
    const row = values[i];
    if (!row) continue;
    const nombre = String(get(row, 'Campaña') ?? '').trim();
    if (!nombre) continue;
    const aaa = normTexto(get(row, 'Cliente AAA'));
    const status = String(get(row, 'Status') ?? '').trim();
    out.rows.push({
      nombre,
      fechaInicio: fechaOperaciones(get(row, 'Fecha Inicio')),
      fechaFin: fechaOperaciones(get(row, 'Fecha Fin')),
      medios: numOperaciones(get(row, 'Medios')),
      proveedores: numOperaciones(get(row, 'Proveedores')),
      aaa: aaa === 'si' ? 'Si' : aaa === 'no' ? 'No' : '',
      peso: numOperaciones(get(row, 'Peso')),
      status,
      terminada: normTexto(status).includes('terminada'),
      coordinador: String(get(row, 'Coordinador') ?? '').trim(),
      ejecutivo: String(get(row, 'Ejecutivo') ?? '').trim(),
    });
  }
  return out;
}

// =========================================================================
// Estado global
// =========================================================================
const Engine = {
  records: [],
  forecast: [],
  contactos: [],
  operaciones: null,        // { rows, capacidad } — hoja "Hoja de Control" de Excelencia Operativa
  opsMatchByCM: {},         // cm -> { coordinador, peso, aaa, status, nombre }
  opsFilasSinCruce: [],     // filas de Excelencia Operativa que no calzaron con ninguna campaña
  cargaCoordinadores: [],   // [{ nombre, actual, capacidad, esTotal }]
  yearsAvailable: [],
  metaData: null,
  metaForecast: null,
  metaContactos: null,
  metaOperaciones: null,
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
    if (this.operaciones) this._computeOperacionesMatch(); // las campañas cambiaron: recalcular el cruce
  },
  loadOperaciones(data, fileMeta) {
    this.operaciones = data;
    this.metaOperaciones = fileMeta || null;
    this._computeOperacionesMatch();
    this._computeCargaCoordinadores();
  },
  // Cruza cada campaña (CM) de Data Comercial con su fila de Excelencia Operativa. La hoja
  // operativa NO trae CM, solo el nombre, y con datos reales el nombre casi nunca es idéntico
  // ("Gorett" vs "Gorett Noviembre", "Duelo" vs "Duelo - autobuses"), se repite ("ADAS"/"Adas",
  // la misma marca en 2025 y 2026) y una fila de la hoja puede cubrir varios CMs (Jamiroquai =
  // 8 + 1 sitios en dos CMs). Por eso el cruce va del lado de la CAMPAÑA: para cada CM se elige,
  // entre las filas con nombre compatible, la de menor costo. El nombre (igual / contenido / con
  // errata) solo es compuerta; la evidencia son fechas, ejecutivo y # de medios/proveedores.
  // Rangos de fechas que no se traslapan se rechazan (otra corrida de la misma marca). Si dos
  // filas empatan para la misma campaña (roles "Coordinador"/"Administrador") se muestran ambas.
  // Lo que no llega al límite de costo queda sin cruce y se marca para revisión.
  _computeOperacionesMatch() {
    this.opsMatchByCM = {};
    this.opsFilasSinCruce = [];
    const ops = this.operaciones;
    if (!ops) return;
    const DAY = 86400000;
    const LIMITE = 150; // costo máximo para aceptar un cruce
    const INICIO_HOJA = Date.UTC(2026, 0, 1);
    // Fechas fuera de 2020-2100 son typos en la fuente (ej. un inicio "2006-09-18"): se ignoran.
    const fechaOk = v => { const d = v ? new Date(v) : null; return d && !isNaN(d) && d.getUTCFullYear() >= 2020 && d.getUTCFullYear() <= 2100 ? d : null; };

    // Catálogo completo (todos los años, sin filtros del panel): si el cruce dependiera de los
    // filtros activos, cambiar el Año podría reasignar filas a otra campaña.
    const porCM = {};
    this.records.forEach(r => {
      if (!r.cm) return;
      const g = porCM[r.cm] || (porCM[r.cm] = { cm: r.cm, cmp: r.cmp || r.cm, eje: r.eje, fechaInicio: null, fechaFin: null, medios: new Set(), provs: new Set() });
      if (r.idSitio) g.medios.add(r.idSitio);
      if (r.prov) g.provs.add(r.prov);
      const di = fechaOk(r.fechaInicio), df = fechaOk(r.fechaFin);
      if (di && (!g.fechaInicio || di < g.fechaInicio)) g.fechaInicio = di;
      if (df && (!g.fechaFin || df > g.fechaFin)) g.fechaFin = df;
    });
    const camps = Object.values(porCM).map(c => ({ ...c, norm: normTexto(c.cmp), ejeNorm: normTexto(c.eje) }));
    camps.forEach(c => { c.tok = c.norm.split(' ').filter(Boolean); });
    const filas = ops.rows.map(o => {
      const norm = normTexto(o.nombre);
      return { o, norm, tok: norm.split(' ').filter(Boolean), ejeNorm: normTexto(o.ejecutivo) };
    });

    // 0 = nombre igual, 1 = uno contiene al otro ("gorett" ⊂ "gorett noviembre"), 2 = parecido
    // por una errata ("tiguan"/"taigun"), 99 = incompatible.
    const nivelNombre = (f, c) => {
      if (!f.norm || !c.norm) return 99;
      if (f.norm === c.norm) return 0;
      const [corto, largo] = f.tok.length <= c.tok.length ? [f.tok, c.tok] : [c.tok, f.tok];
      if (corto.length && corto.every(t => largo.includes(t))) return 1;
      if (Math.min(f.norm.length, c.norm.length) >= 8 && distanciaEdicion(f.norm, c.norm) <= 2) return 2;
      return 99;
    };
    const pareja = (f, c) => {
      const nivel = nivelNombre(f, c);
      if (nivel === 99) return null;
      const o = f.o, oFin = o.fechaFin, cFin = c.fechaFin;
      // La hoja operativa es de 2026: lo que empezó antes (campañas que vienen del año pasado)
      // se compara desde el 1-ene-2026 en ambos lados, para no penalizar esa diferencia.
      const recorta = d => (d && +d < INICIO_HOJA ? new Date(INICIO_HOJA) : d);
      const oIni = recorta(o.fechaInicio), cIni = recorta(c.fechaInicio);
      const conRangos = oIni && oFin && cIni && cFin;
      // OJO: las fechas se pasan a número (+) antes de sumar; Date + número concatena texto.
      if (conRangos && (+cFin < +oIni - 7 * DAY || +cIni > +oFin + 7 * DAY)) return null; // no se traslapan
      const dIni = oIni && cIni ? Math.abs(oIni - cIni) / DAY : null;
      const dFin = oFin && cFin ? Math.abs(oFin - cFin) / DAY : null;
      let fechas = (dIni ?? 200) + 0.5 * (dFin ?? 200);
      // Fila "agregadora": la campaña de la hoja abarca varios CMs (ej. "Stellantis" todo el año
      // y un CM de marzo a diciembre). El CM queda dentro del rango de la fila.
      const agregadora = conRangos && +cIni >= +oIni - 7 * DAY && +cFin <= +oFin + 7 * DAY && fechas > 25;
      if (agregadora) fechas = 25;
      let costo = nivel * 15 + fechas;
      if (f.ejeNorm !== c.ejeNorm) costo += 120;
      if (agregadora) costo += 10; // los conteos de una fila agregadora suman varios CMs: no se comparan
      else {
        if (o.medios != null && o.medios !== c.medios.size) costo += 60;
        if (o.proveedores != null && o.proveedores !== c.provs.size) costo += 30;
      }
      return { costo, nivel };
    };

    const usadas = new Set(); // filas de la hoja elegidas por al menos una campaña
    camps.forEach(c => {
      const cand = [];
      filas.forEach((f, fi) => {
        const p = pareja(f, c);
        if (p && p.costo <= LIMITE) cand.push({ f, fi, ...p });
      });
      if (!cand.length) return;
      cand.sort((a, b) => a.costo - b.costo);
      const elegidas = cand.filter(x => x.costo - cand[0].costo <= 5);
      elegidas.forEach(x => usadas.add(x.fi));
      const pesos = elegidas.map(x => x.f.o.peso).filter(p => p != null);
      this.opsMatchByCM[c.cm] = {
        coordinador: [...new Set(elegidas.map(x => x.f.o.coordinador).filter(Boolean))].join(' / '),
        peso: pesos.length ? pesos.reduce((a, b) => a + b, 0) : null,
        aaa: elegidas.map(x => x.f.o.aaa).find(Boolean) || '',
        status: elegidas[0].f.o.status,
        nombre: elegidas.map(x => x.f.o.nombre).join(' + '),
        aprox: elegidas.some(x => x.nivel > 0), // el nombre en la hoja no es idéntico
        filas: elegidas.length,
      };
    });
    filas.forEach((f, fi) => { if (!usadas.has(fi)) this.opsFilasSinCruce.push(f.o); });
  },
  // Carga por coordinador: Actual = suma de Peso de las campañas NO terminadas (misma regla
  // que la fórmula de la hoja), Capacidad = la que trae la tabla de la hoja. Se calcula aquí
  // y no se lee el "Actual" de la hoja porque la fórmula de Pablo no excluye las terminadas
  // (le falta el "⚫ " al criterio) y muestra de más.
  _computeCargaCoordinadores() {
    this.cargaCoordinadores = [];
    const ops = this.operaciones;
    if (!ops) return;
    const actual = {}, nombres = {};
    let total = 0;
    ops.rows.forEach(r => {
      if (r.terminada || r.peso == null) return;
      total += r.peso;
      const k = normTexto(r.coordinador);
      actual[k] = (actual[k] || 0) + r.peso;
      if (k && !nombres[k]) nombres[k] = r.coordinador;
    });
    const out = [];
    let capSuma = 0;
    ops.capacidad.filter(c => !c.esTotal).forEach(c => {
      const k = normTexto(c.nombre);
      out.push({ nombre: c.nombre, actual: actual[k] || 0, capacidad: c.capacidad });
      capSuma += c.capacidad;
      delete actual[k];
    });
    Object.keys(actual).forEach(k => { // coordinadores en la lista que no están en la tabla de capacidad
      if (k === '') { if (actual[k] > 0) out.push({ nombre: 'Sin coordinador', actual: actual[k], capacidad: null }); }
      else out.push({ nombre: nombres[k], actual: actual[k], capacidad: null });
    });
    const filaTotal = ops.capacidad.find(c => c.esTotal);
    out.push({ nombre: 'Total Sistema', actual: total, capacidad: filaTotal ? filaTotal.capacidad : capSuma, esTotal: true });
    this.cargaCoordinadores = out;
  },
  loadForecast(forecast, fileMeta) {
    this.forecast = forecast;
    this.metaForecast = fileMeta || null;
    this._applyFuzzyMatchingForecast();
    this._computeActiveSellers();
  },
  loadContactos(records, fileMeta) {
    this.contactos = records;
    this.metaContactos = fileMeta || null;
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
  // Forecast filtrado igual que datos (vendedor, cliente, holding, mes, año) — el Forecast
  // no trae emp/est/loc/cat/tp/cc, así que esos filtros del panel superior no le aplican.
  applyForecastFilters(filters, opts={}) {
    return this.forecast.filter(f => {
      if (!opts.ignoreYear && filters.anio && f.anio != filters.anio) return false;
      if (filters.mes && f.mes != filters.mes) return false;
      if (filters.eje && f.eje !== filters.eje) return false;
      if (filters.cli && f.cli !== filters.cli) return false;
      if (filters.hol && f.hol !== filters.hol) return false;
      return true;
    });
  },
  // Igual que applyForecastFilters pero ya sumado en un arreglo de 12 meses — para las
  // tablas/KPIs que necesitan el total mensual del Forecast respetando los filtros activos.
  forecastMonthlyFilteredArr(filters, anio) {
    const arr = Array(12).fill(0);
    this.applyForecastFilters({ ...filters, anio }, { ignoreYear: false }).forEach(f => {
      if (f.mes >= 1 && f.mes <= 12) arr[f.mes - 1] += f.fcNeto;
    });
    return arr;
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
window.processContactosWorkbook = processContactosWorkbook;
window.processOperacionesValues = processOperacionesValues;
window.CATEGORIAS_MIX = CATEGORIAS_MIX;
