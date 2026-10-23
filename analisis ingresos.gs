/**
 * CLUB · Análisis de ingresos por concepto (para Looker Studio)
 *
 * Genera 3 hojas planas, una fila por registro:
 *   - Ingresos_Detalle  : una fila por CONCEPTO cobrado (explota la columna F de Ingresos)
 *   - Tarifas_Estimadas : precio deducido por concepto y mes (recibos de un solo concepto)
 *   - Deuda_Esperada    : AntiguedaddeDeuda en formato largo (lo que se espera cobrar)
 *
 * Reparto del importe (columna I) entre los conceptos de un recibo:
 *   1. Recibo con un solo concepto  -> se asigna todo el importe.
 *   2. Varios conceptos con tarifa estimada -> proporcional a esas tarifas.
 *   3. Sin ninguna tarifa conocida -> partes iguales.
 *   La suma de lo asignado siempre da igual a la columna I.
 *   Cuando tengas la lista de precios real, se reemplaza el paso de tarifas estimadas.
 *
 * En Ingresos_Detalle, la columna AB repite el texto original de la columna F
 * de Ingresos (todos los conceptos del recibo) para poder controlar el reparto.
 *
 * Columnas J, K, L y M de Ingresos se ignoran a propósito.
 */

const CFG = {
  HOJA_INGRESOS: 'Ingresos',
  HOJA_SOCIOS: 'Socios',
  HOJA_DEUDA: 'AntiguedaddeDeuda',
  OUT_DETALLE: 'Ingresos_Detalle',
  OUT_TARIFAS: 'Tarifas_Estimadas',
  OUT_DEUDA: 'Deuda_Esperada',
  HORA_TRIGGER: 7
};

// Índices de columnas de la hoja Ingresos (A = 0)
const ING = { F_REG: 0, F_COBRO: 1, NRO: 2, CAJA: 3, CLIENTE: 4, CONCEPTOS: 5, VALORES: 6, IMPORTE: 8 };

const MESES = { ene: 0, feb: 1, mar: 2, abr: 3, may: 4, jun: 5, jul: 6, ago: 7, sep: 8, set: 8, oct: 9, nov: 10, dic: 11 };
const RX_PERIODO = /^([A-Za-zÁÉÍÓÚáéíóú]{3})-(\d{2})$/;
const RX_SEPARA_CONCEPTOS = /,\s*(?=[A-ZÁÉÍÓÚÑ])/;

// ───────────────────────── Menú y disparadores ─────────────────────────

function actualizarInformes() {
  generarIngresosDetalle();
  generarDeudaEsperada();
}

function instalarTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'actualizarTodo')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('actualizarTodo').timeBased().everyDays(1).atHour(CFG.HORA_TRIGGER).create();
}

// ───────────────────────── Ingresos por concepto ─────────────────────────

function generarIngresosDetalle() {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName(CFG.HOJA_INGRESOS);
  if (!hoja) throw new Error('No encuentro la hoja ' + CFG.HOJA_INGRESOS);

  const datos = hoja.getDataRange().getValues();
  const socios = cargarSocios_(ss);

  // 1) Leer recibos y separar conceptos
  const recibos = [];
  for (let i = 1; i < datos.length; i++) {
    const r = datos[i];
    const nro = r[ING.NRO];
    const txt = String(r[ING.CONCEPTOS] || '').trim();
    if (!nro && !txt) continue;

    const conceptos = txt
      ? txt.split(RX_SEPARA_CONCEPTOS).map(s => parseConcepto_(s.trim())).filter(c => c.clave)
      : [];
    if (!conceptos.length) conceptos.push({ tipo: 'Sin concepto', detalle: '', modalidad: '', periodo: null, clave: 'Sin concepto' });

    const fCobro = aFecha_(r[ING.F_COBRO]) || aFecha_(r[ING.F_REG]);
    const socio = nroSocio_(r[ING.CLIENTE]);
    recibos.push({
      nro: nro,
      fCobro: fCobro,
      caja: r[ING.CAJA],
      cliente: r[ING.CLIENTE],
      socio: socio,
      medio: medioPago_(r[ING.VALORES]),
      importe: aNumero_(r[ING.IMPORTE]),
      conceptos: conceptos,
      conceptosTxt: txt,            // NUEVO: texto original de la columna F
      infoSocio: socios[socio] || {}
    });
  }

  // 2) Tarifas estimadas: mediana de recibos con un solo concepto,
  //    excluyendo socios con descuento/recargo para no distorsionar el precio.
  const muestras = {}; // clave -> { indiceMes: [importes] }
  recibos.forEach(rc => {
    if (rc.conceptos.length !== 1 || rc.importe <= 0 || rc.infoSocio.desc) return;
    const c = rc.conceptos[0];
    const idx = mesRef_(c, rc);
    if (idx === null) return;
    (muestras[c.clave] = muestras[c.clave] || {});
    (muestras[c.clave][idx] = muestras[c.clave][idx] || []).push(rc.importe);
  });

  const tablaTarifas = {}; // clave -> [{idx, v, n}]
  const filasTarifas = [];
  Object.keys(muestras).sort().forEach(clave => {
    tablaTarifas[clave] = Object.keys(muestras[clave]).map(ix => {
      const lista = muestras[clave][ix];
      const t = { idx: Number(ix), v: mediana_(lista), n: lista.length };
      const d = desdeIndice_(t.idx);
      filasTarifas.push([clave, d, d.getFullYear(), d.getMonth() + 1, t.v, t.n]);
      return t;
    });
  });

  const buscarTarifa = (clave, idx) => {
    const lista = tablaTarifas[clave];
    if (!lista || idx === null) return null;
    let mejor = null;
    lista.forEach(t => {
      if (!mejor || Math.abs(t.idx - idx) < Math.abs(mejor.idx - idx)) mejor = t;
    });
    return mejor;
  };

  // 3) Repartir el importe de cada recibo entre sus conceptos
  const filas = [];
  recibos.forEach(rc => {
    const n = rc.conceptos.length;
    const tarifas = rc.conceptos.map(c => buscarTarifa(c.clave, mesRef_(c, rc)));
    const conocidas = tarifas.filter(Boolean).map(t => t.v);

    let pesos, metodo;
    if (n === 1) {
      pesos = [1];
      metodo = 'Concepto único';
    } else if (!conocidas.length) {
      pesos = tarifas.map(() => 1);
      metodo = 'Partes iguales';
    } else {
      const prom = conocidas.reduce((a, b) => a + b, 0) / conocidas.length;
      pesos = tarifas.map(t => (t ? t.v : prom));
      metodo = conocidas.length === n ? 'Proporcional a tarifa estimada' : 'Proporcional (tarifa parcial)';
    }
    let sumaPesos = pesos.reduce((a, b) => a + b, 0);
    if (!sumaPesos) { pesos = pesos.map(() => 1); sumaPesos = n; metodo = 'Partes iguales'; }

    const fC = rc.fCobro;
    const s = rc.infoSocio;
    let acumulado = 0;

    rc.conceptos.forEach((c, k) => {
      const asignado = k === n - 1
        ? redondear_(rc.importe - acumulado)
        : redondear_(rc.importe * pesos[k] / sumaPesos);
      acumulado += asignado;

      const desfase = (c.periodo && fC) ? mesIndex_(fC) - mesIndex_(c.periodo) : '';
      const situacion = desfase === '' ? 'Sin período'
        : desfase > 0 ? 'Atrasado'
        : desfase < 0 ? 'Adelantado' : 'Al día';

      filas.push([
        rc.nro,                                                   // A
        fC || '',                                                 // B
        fC ? fC.getFullYear() : '',                               // C
        fC ? fC.getMonth() + 1 : '',                              // D
        fC ? new Date(fC.getFullYear(), fC.getMonth(), 1) : '',   // E
        rc.caja,                                                  // F
        rc.socio,                                                 // G
        rc.cliente,                                               // H
        s.div || '',                                              // I
        s.desc || '',                                             // J
        s.perDesc || '',                                          // K
        s.desc ? 'Sí' : 'No',                                     // L
        rc.medio,                                                 // M
        c.tipo,                                                   // N
        c.detalle,                                                // O
        c.modalidad,                                              // P
        c.clave,                                                  // Q
        c.periodo || '',                                          // R
        c.periodo ? c.periodo.getFullYear() : '',                 // S
        c.periodo ? c.periodo.getMonth() + 1 : '',                // T
        desfase,                                                  // U
        situacion,                                                // V
        n,                                                        // W
        rc.importe,                                               // X
        tarifas[k] ? tarifas[k].v : '',                           // Y
        asignado,                                                 // Z
        metodo,                                                   // AA
        rc.conceptosTxt                                           // AB: texto original col F de Ingresos
      ]);
    });
  });

  escribir_(ss, CFG.OUT_DETALLE, [
    'Nº Recibo', 'Fecha Cobro', 'Año Cobro', 'Mes Cobro', 'Período Cobro', 'Caja',
    'Nº Socio', 'Cliente', 'Divisiones', 'Descuento/Recargo', 'Período Descuento', 'Tiene Descuento',
    'Medio de Pago', 'Tipo Concepto', 'Detalle', 'Modalidad / Código', 'Concepto',
    'Período Imputado', 'Año Imputado', 'Mes Imputado', 'Desfase Meses', 'Situación Pago',
    'Cant. Conceptos Recibo', 'Importe Recibo', 'Tarifa Estimada', 'Importe Asignado', 'Método Asignación',
    'Conceptos Original (col F)'
  ], filas, { 2: 'dd/MM/yyyy', 5: 'MM/yyyy', 18: 'MM/yyyy', 24: '#,##0.00', 25: '#,##0.00', 26: '#,##0.00', 28: '@' });

  escribir_(ss, CFG.OUT_TARIFAS, [
    'Concepto', 'Período', 'Año', 'Mes', 'Tarifa Estimada', 'Cant. Recibos'
  ], filasTarifas, { 2: 'MM/yyyy', 5: '#,##0.00' });
}

// ───────────────────────── Deuda esperada ─────────────────────────

function generarDeudaEsperada() {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName(CFG.HOJA_DEUDA);
  if (!hoja) throw new Error('No encuentro la hoja ' + CFG.HOJA_DEUDA);

  const d = hoja.getDataRange().getValues();
  const h = d[0].map(String);
  const socios = cargarSocios_(ss);

  // Columnas con período en el encabezado: "30 Días (Sep-26)"
  const cols = [];
  h.forEach((t, j) => {
    const m = t.match(/\(([^)]+)\)/);
    const p = m ? parsePeriodo_(m[1]) : null;
    if (p) cols.push({ j: j, tramo: t.replace(/\s*\(.*\)/, '').trim(), periodo: p, estado: 'Vencida' });
  });
  if (!cols.length) throw new Error('No reconozco los períodos en los encabezados de ' + CFG.HOJA_DEUDA);

  const idxs = cols.map(c => mesIndex_(c.periodo));
  const jAV = h.findIndex(t => /a vencer/i.test(t));
  const jMas = h.findIndex(t => /\+\s*180/.test(t));
  if (jAV >= 0) cols.push({ j: jAV, tramo: 'A Vencer', periodo: desdeIndice_(Math.max.apply(null, idxs) + 1), estado: 'A vencer' });
  if (jMas >= 0) cols.push({ j: jMas, tramo: '+ 180 Días', periodo: desdeIndice_(Math.min.apply(null, idxs) - 1), estado: 'Vencida' });

  const hoy = new Date();
  const filas = [];
  for (let i = 1; i < d.length; i++) {
    const codigo = String(d[i][0] || '').trim();
    if (!codigo) continue;
    const s = socios[codigo] || {};
    cols.forEach(c => {
      const monto = aNumero_(d[i][c.j]);
      if (!monto) return;
      filas.push([
        codigo, d[i][1], s.div || '', s.desc || '', s.desc ? 'Sí' : 'No',
        c.tramo, c.estado, c.periodo, c.periodo.getFullYear(), c.periodo.getMonth() + 1,
        monto, hoy
      ]);
    });
  }

  escribir_(ss, CFG.OUT_DEUDA, [
    'Nº Socio', 'Socio', 'Divisiones', 'Descuento/Recargo', 'Tiene Descuento',
    'Tramo', 'Estado', 'Período', 'Año', 'Mes', 'Monto', 'Fecha Actualización'
  ], filas, { 8: 'MM/yyyy', 11: '#,##0.00', 12: 'dd/MM/yyyy HH:mm' });
}

// ───────────────────────── Auxiliares ─────────────────────────

/** Socios por número: divisiones (H), descuento/recargo (I) y su período (J). */
function cargarSocios_(ss) {
  const hoja = ss.getSheetByName(CFG.HOJA_SOCIOS);
  const mapa = {};
  if (!hoja) return mapa;
  const d = hoja.getDataRange().getValues();
  const h = d[0].map(t => String(t).toLowerCase());
  const col = (rx, def) => { const j = h.findIndex(t => rx.test(t)); return j >= 0 ? j : def; };
  const cNro = col(/n.?\s*socio/, 0);
  const cDiv = col(/^divisiones/, 7);
  const cDesc = col(/^descuentos y recargos/, 8);
  const cPer = col(/^per[ií]odo en descuentos/, 9);
  for (let i = 1; i < d.length; i++) {
    const nro = String(d[i][cNro] || '').trim();
    if (!nro) continue;
    mapa[nro] = {
      div: d[i][cDiv],
      desc: String(d[i][cDesc] || '').trim(),
      perDesc: String(d[i][cPer] || '').trim()
    };
  }
  return mapa;
}

/** "Actividad | Fútbol Masculino > Jun-26 > Abono Mensual" -> partes */
function parseConcepto_(txt) {
  const tokens = txt.split(/\s*[>|]\s*/).map(t => t.trim()).filter(Boolean);
  let periodo = null;
  const resto = [];
  tokens.forEach(t => {
    const p = parsePeriodo_(t);
    if (p && !periodo) periodo = p; else resto.push(t);
  });
  const tipo = resto[0] || '';
  let detalle = '', modalidad = '';
  if (/^art[ií]culos?$/i.test(tipo)) {          // Artículos > ART-2045 | Matricula
    detalle = resto[resto.length - 1] || '';
    modalidad = resto.length > 2 ? resto.slice(1, -1).join(' > ') : '';
  } else if (/^actividad/i.test(tipo)) {         // Actividad | Fútbol Masculino > Jun-26 > Abono Mensual
    detalle = resto[1] || '';
    modalidad = resto.slice(2).join(' > ');
  } else {                                       // Cuota Social > Jun-26 > Activo
    detalle = resto.slice(1).join(' > ') || tipo;
  }
  return { tipo: tipo, detalle: detalle, modalidad: modalidad, periodo: periodo, clave: resto.join(' > ') };
}

function parsePeriodo_(txt) {
  const m = String(txt).trim().match(RX_PERIODO);
  if (!m) return null;
  const mes = MESES[m[1].toLowerCase()];
  return mes === undefined ? null : new Date(2000 + Number(m[2]), mes, 1);
}

/** Mes de referencia para la tarifa: período imputado o, si no tiene, mes de cobro. */
function mesRef_(c, rc) {
  if (c.periodo) return mesIndex_(c.periodo);
  return rc.fCobro ? mesIndex_(rc.fCobro) : null;
}

function medioPago_(v) {
  const medios = String(v || '')
    .split(RX_SEPARA_CONCEPTOS)
    .map(s => s.split('>')[0].trim())
    .filter(Boolean);
  return Array.from(new Set(medios)).join(' + ');
}

function nroSocio_(cliente) {
  const m = String(cliente || '').match(/\(([^)]+)\)\s*$/);
  return m ? m[1].trim() : '';
}

function aFecha_(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  const p = String(v).trim().split(/[\/\s]/);
  if (p.length < 3) return null;
  const f = new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
  return isNaN(f) ? null : f;
}

function aNumero_(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  const n = Number(String(v).replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function mesIndex_(d) { return d.getFullYear() * 12 + d.getMonth(); }
function desdeIndice_(i) { return new Date(Math.floor(i / 12), i % 12, 1); }
function redondear_(x) { return Math.round(x * 100) / 100; }

function mediana_(arr) {
  const a = arr.slice().sort((x, y) => x - y);
  const m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
}

/** Reescribe una hoja completa. formatos: {nº columna (1 = A): formato} */
function escribir_(ss, nombre, encabezado, filas, formatos) {
  const sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  sh.clearContents();
  sh.getRange(1, 1, 1, encabezado.length).setValues([encabezado]).setFontWeight('bold');
  if (filas.length) {
    sh.getRange(2, 1, filas.length, encabezado.length).setValues(filas);
    Object.keys(formatos || {}).forEach(c => {
      sh.getRange(2, Number(c), filas.length, 1).setNumberFormat(formatos[c]);
    });
  }
  sh.setFrozenRows(1);
}