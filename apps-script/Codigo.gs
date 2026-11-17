/**
 * Seguimiento de Consultas 2026 · Servicio de datos (Google Apps Script)
 * Supervisión Escolar SEMS
 *
 * Guarda y entrega los cortes diarios de la herramienta de seguimiento.
 * Solo recibe conteos agregados por plantel y grupo; nunca respuestas individuales.
 *
 * ANTES DE USAR, edita estas dos líneas:
 */
const ENTORNO = 'PRUEBAS';            // En la hoja de producción escribe: 'PRODUCCION'
const CLAVE = 'ESCRIBE-AQUI-TU-CLAVE'; // Mínimo 12 caracteres. Será la que captures en la herramienta.

/* ------------------------------------------------------------------ */
/*  No es necesario modificar nada debajo de esta línea               */
/* ------------------------------------------------------------------ */
const VERSION = '1.0';
const FECHA_MIN = '2026-11-13';
const FECHA_MAX = '2026-12-19';
const CATALOGO = {
  est: {
    grupos: ['Matutino 1°', 'Matutino 3°', 'Matutino 5°', 'Vespertino 1°', 'Vespertino 3°', 'Vespertino 5°'],
    planteles: ['CBTis 203', 'CETis 95', 'CBTis 29', 'CBTis 133', 'CEB 9/13', 'CBTis 289', 'CBTis 180', 'CEB 9/28',
      'CBTA 128', 'BTED Plantel Texcoco', 'CETis 94', 'CBTis 295', 'CETis 141', 'CETis 165', 'CEB 9/29', 'CETis 97',
      'CEB 9/14', 'CETis 23', 'CETis 36', 'CETis 118', 'CBTA 300']
  },
  doc: {
    grupos: ['Matutino', 'Vespertino', 'Mixto'],
    planteles: ['CBTis 203', 'CETis 95', 'CBTis 29', 'CBTis 133', 'CEB 9/13', 'CBTis 289', 'CBTis 180', 'CEB 9/28',
      'CBTA 128', 'BTED Plantel Texcoco', 'CETis 94', 'CBTis 295', 'CETis 141', 'CETis 165', 'CEB 9/29', 'CETis 97',
      'CEB 9/14', 'CETis 23', 'CETis 36', 'CETis 118', 'CBTA 300']
  }
};
const HOJAS = {
  Cortes: ['id_corte', 'encuesta', 'fecha_corte', 'estado', 'guardado_en', 'archivo', 'huella', 'respuestas_archivo',
    'duplicados', 'id_max', 'version_app', 'datos_json'],
  Conteos: ['id_corte', 'encuesta', 'fecha_corte', 'plantel', 'grupo', 'respuestas'],
  PorRevisar: ['id_corte', 'encuesta', 'fecha_corte', 'tipo', 'plantel_o_escuela', 'detalle', 'respuestas'],
  Bitacora: ['fecha_hora', 'accion', 'encuesta', 'fecha_corte', 'id_corte', 'resultado', 'detalle']
};
const C = { id: 0, encuesta: 1, fecha: 2, estado: 3, guardado: 4, archivo: 5, huella: 6, unicas: 7, dup: 8, maxId: 9, version: 10, datos: 11 };

/** Ejecuta esta función UNA vez después de pegar el código. Crea las hojas y sus encabezados. */
function configurar() {
  if (CLAVE.length < 12 || CLAVE === 'ESCRIBE-AQUI-TU-CLAVE') throw new Error('Primero escribe tu clave (mínimo 12 caracteres) en la línea CLAVE.');
  if (ENTORNO !== 'PRUEBAS' && ENTORNO !== 'PRODUCCION') throw new Error("ENTORNO debe ser 'PRUEBAS' o 'PRODUCCION'.");
  const ss = SpreadsheetApp.getActive();
  Object.keys(HOJAS).forEach(function (n) {
    let sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HOJAS[n].length).setValues([HOJAS[n]]).setFontWeight('bold').setBackground('#002F2A').setFontColor('#E6D194');
      sh.setFrozenRows(1);
    }
    if (!sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) {
      sh.protect().setDescription('Escritura solo desde la herramienta de seguimiento').setWarningOnly(true);
    }
  });
  ['Hoja 1', 'Sheet1', 'Hoja1'].forEach(function (n) {
    const h = ss.getSheetByName(n);
    if (h && h.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(h);
  });
  bitacora('configurar', '', '', '', 'ok', 'Entorno ' + ENTORNO + ', versión ' + VERSION);
  Logger.log('Listo. Entorno: ' + ENTORNO + '. Ahora implementa como aplicación web.');
}

/* ----------------------------- Entradas web ----------------------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.accion === 'ping' || !p.accion) return salida({ ok: true, entorno: ENTORNO, version: VERSION });
  return salida({ ok: false, codigo: 'metodo', mensaje: 'Usa POST para esta acción.' });
}
function doPost(e) {
  let p;
  try { p = JSON.parse(e.postData.contents); } catch (err) { return salida({ ok: false, codigo: 'json', mensaje: 'Solicitud con formato inválido.' }); }
  return salida(manejar(p));
}
function manejar(p) {
  try {
    if (p.accion === 'ping') return { ok: true, entorno: ENTORNO, version: VERSION };
    if (p.clave !== CLAVE) return { ok: false, codigo: 'clave', mensaje: 'Clave incorrecta.' };
    if (p.accion === 'leer') return { ok: true, entorno: ENTORNO, datos: leer() };
    if (p.accion === 'guardar') return conBloqueo(function () { return guardar(p.corte, !!p.forzar); });
    if (p.accion === 'eliminar') return conBloqueo(function () { return eliminar(p.encuesta, p.fecha); });
    return { ok: false, codigo: 'accion', mensaje: 'Acción desconocida.' };
  } catch (err) {
    return { ok: false, codigo: 'error', mensaje: String((err && err.message) || err) };
  }
}
function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function conBloqueo(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, codigo: 'ocupado', mensaje: 'Otro guardado está en curso. Intenta de nuevo en unos segundos.' };
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ----------------------------- Lectura ----------------------------- */
function filasCortes() {
  const sh = hoja('Cortes'), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, HOJAS.Cortes.length).getValues();
}
function vigentes(enc) {
  const porFecha = {};
  filasCortes().forEach(function (r, i) {
    if (r[C.estado] !== 'vigente' || (enc && r[C.encuesta] !== enc)) return;
    const k = r[C.encuesta] + '|' + txtFecha(r[C.fecha]);
    if (!porFecha[k] || String(r[C.guardado]) > String(porFecha[k].r[C.guardado])) porFecha[k] = { r: r, fila: i + 2 };
  });
  return Object.keys(porFecha).map(function (k) { return porFecha[k]; });
}
function leer() {
  const out = { est: [], doc: [] };
  vigentes().forEach(function (v) {
    const r = v.r;
    let d = null;
    try { d = JSON.parse(r[C.datos]); } catch (e) { return; }
    out[r[C.encuesta]].push({ id: r[C.id], fecha: txtFecha(r[C.fecha]), guardado: String(r[C.guardado]), datos: d });
  });
  out.est.sort(porFechaAsc); out.doc.sort(porFechaAsc);
  return out;
}

/* ----------------------------- Guardado ----------------------------- */
function guardar(c, forzar) {
  const err = validar(c);
  if (err) { bitacora('guardar', c && c.encuesta, c && c.fecha, '', 'rechazado', err); return { ok: false, codigo: 'invalido', mensaje: err }; }
  const vs = vigentes(c.encuesta);
  if (!forzar) {
    let ant = null, sig = null;
    vs.forEach(function (v) {
      const f = txtFecha(v.r[C.fecha]);
      if (f < c.fecha && (!ant || f > txtFecha(ant.r[C.fecha]))) ant = v;
      if (f > c.fecha && (!sig || f < txtFecha(sig.r[C.fecha]))) sig = v;
    });
    if (ant && Number(ant.r[C.unicas]) > c.unicas) return aviso(c, 'menor_que_anterior',
      'Este archivo tiene ' + c.unicas + ' respuestas y el corte anterior (' + txtFecha(ant.r[C.fecha]) + ') tenía ' + ant.r[C.unicas] + '. Puede ser el archivo equivocado.');
    if (sig && Number(sig.r[C.unicas]) < c.unicas) return aviso(c, 'mayor_que_siguiente',
      'Este archivo tiene más respuestas (' + c.unicas + ') que el corte posterior del ' + txtFecha(sig.r[C.fecha]) + ' (' + sig.r[C.unicas] + '). Revisa la fecha del corte.');
    if (c.huella) {
      const rep = vs.filter(function (v) { return v.r[C.huella] === c.huella && txtFecha(v.r[C.fecha]) !== c.fecha; })[0];
      if (rep) return aviso(c, 'archivo_repetido', 'Este mismo archivo ya se guardó como corte del ' + txtFecha(rep.r[C.fecha]) + '.');
    }
  }
  const ahora = new Date(), tz = Session.getScriptTimeZone();
  const id = c.encuesta + '-' + c.fecha + '-' + Utilities.formatDate(ahora, tz, 'HHmmss');
  const guardado = Utilities.formatDate(ahora, tz, "yyyy-MM-dd'T'HH:mm:ss");
  const grupos = CATALOGO[c.encuesta].grupos;

  // 1) Conteos y casos por revisar (solo informativos; la herramienta lee el corte desde la hoja Cortes)
  const cont = [];
  Object.keys(c.g).forEach(function (p) { c.g[p].forEach(function (n, j) { if (n > 0) cont.push([id, c.encuesta, c.fecha, p, grupos[j], n]); }); });
  agregar('Conteos', cont);
  const rev = [];
  Object.keys(c.inc || {}).forEach(function (p) {
    Object.keys(c.inc[p]).forEach(function (k) {
      const det = k === 'e' ? 'Edad menor a 20 años (posible estudiante)' : k === 'x' ? 'Sin turno o semestre' : grupos[+k] + ' (grupo que el plantel no tiene)';
      rev.push([id, c.encuesta, c.fecha, 'Plantel del catálogo', p, det, c.inc[p][k]]);
    });
  });
  Object.keys(c.otras || {}).forEach(function (o) { rev.push([id, c.encuesta, c.fecha, 'Otra escuela', o, 'Fuera del catálogo', c.otras[o]]); });
  if (c.sinPlantel) rev.push([id, c.encuesta, c.fecha, 'Sin plantel', '', 'No señaló plantel', c.sinPlantel]);
  agregar('PorRevisar', rev);

  // 2) El corte nuevo queda vigente
  agregar('Cortes', [[id, c.encuesta, c.fecha, 'vigente', guardado, c.archivo || '', c.huella || '', c.unicas, c.dup || 0, c.maxId || 0, c.version || '', JSON.stringify(c)]]);

  // 3) El corte anterior de la misma fecha pasa a "reemplazado" (se conserva para auditoría)
  const sh = hoja('Cortes');
  let reemplazos = 0;
  vs.forEach(function (v) { if (txtFecha(v.r[C.fecha]) === c.fecha) { sh.getRange(v.fila, C.estado + 1).setValue('reemplazado'); reemplazos++; } });

  bitacora('guardar', c.encuesta, c.fecha, id, 'ok', c.unicas + ' respuestas' + (reemplazos ? '; reemplaza corte anterior del mismo día' : '') + (forzar ? '; guardado con confirmación' : ''));
  return { ok: true, id: id, reemplazo: reemplazos > 0, entorno: ENTORNO };
}
function aviso(c, codigo, mensaje) {
  bitacora('guardar', c.encuesta, c.fecha, '', 'pendiente de confirmar', mensaje);
  return { ok: false, codigo: codigo, confirmable: true, mensaje: mensaje };
}
function validar(c) {
  if (!c || typeof c !== 'object') return 'No llegó ningún corte.';
  const cat = CATALOGO[c.encuesta];
  if (!cat) return 'Encuesta desconocida.';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(c.fecha || '') || c.fecha < FECHA_MIN || c.fecha > FECHA_MAX) return 'La fecha del corte debe estar entre ' + FECHA_MIN + ' y ' + FECHA_MAX + '.';
  if (!entero(c.unicas) || !entero(c.sinPlantel || 0) || !entero(c.dup || 0)) return 'Totales con formato inválido.';
  if (!c.g || typeof c.g !== 'object') return 'Faltan los conteos por plantel.';
  let suma = c.sinPlantel || 0;
  for (const p in c.g) {
    if (cat.planteles.indexOf(p) < 0) return 'Plantel fuera del catálogo: ' + p;
    const a = c.g[p];
    if (!Array.isArray(a) || a.length !== cat.grupos.length || !a.every(entero)) return 'Conteos inválidos en ' + p + '.';
    suma += a.reduce(function (x, y) { return x + y; }, 0);
  }
  for (const p in (c.inc || {})) {
    if (cat.planteles.indexOf(p) < 0) return 'Plantel fuera del catálogo: ' + p;
    for (const k in c.inc[p]) {
      if (!(k === 'x' || k === 'e' || (/^\d+$/.test(k) && +k < cat.grupos.length))) return 'Categoría por revisar inválida en ' + p + '.';
      if (!entero(c.inc[p][k])) return 'Conteo por revisar inválido en ' + p + '.';
      suma += c.inc[p][k];
    }
  }
  for (const o in (c.otras || {})) { if (!entero(c.otras[o])) return 'Conteo inválido de otra escuela.'; suma += c.otras[o]; }
  if (suma !== c.unicas) return 'Los conteos no cuadran con el total del archivo (' + suma + ' contra ' + c.unicas + '). No se guardó nada.';
  return null;
}

/* ----------------------------- Eliminación ----------------------------- */
function eliminar(enc, fecha) {
  if (!CATALOGO[enc] || !/^\d{4}-\d{2}-\d{2}$/.test(fecha || '')) return { ok: false, codigo: 'invalido', mensaje: 'Datos incompletos.' };
  const sh = hoja('Cortes');
  let n = 0;
  vigentes(enc).forEach(function (v) { if (txtFecha(v.r[C.fecha]) === fecha) { sh.getRange(v.fila, C.estado + 1).setValue('eliminado'); n++; } });
  bitacora('eliminar', enc, fecha, '', n ? 'ok' : 'sin cambios', n ? 'Corte marcado como eliminado' : 'No había corte vigente en esa fecha');
  return { ok: true, eliminados: n };
}

/* ----------------------------- Utilidades ----------------------------- */
function hoja(n) {
  const sh = SpreadsheetApp.getActive().getSheetByName(n);
  if (!sh) throw new Error('Falta la hoja ' + n + '. Ejecuta configurar().');
  return sh;
}
function agregar(n, filas) {
  if (!filas.length) return;
  const sh = hoja(n);
  sh.getRange(sh.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
}
function bitacora(accion, enc, fecha, id, resultado, detalle) {
  try {
    agregar('Bitacora', [[Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'), accion, enc || '', fecha || '', id || '', resultado, detalle || '']]);
  } catch (e) { /* la bitácora nunca debe impedir la operación */ }
}
function entero(n) { return typeof n === 'number' && isFinite(n) && n >= 0 && Math.floor(n) === n; }
function txtFecha(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v); }
function porFechaAsc(a, b) { return a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0; }

/* ----------------------------- Autoprueba (solo PRUEBAS) ----------------------------- */
/** Ejecútala en la hoja de PRUEBAS para comprobar que todo funciona. Escribe cortes ficticios. */
function autoprueba() {
  if (ENTORNO !== 'PRUEBAS') throw new Error('La autoprueba solo se ejecuta en el entorno de PRUEBAS.');
  function corte(fecha, n, huella) {
    return { encuesta: 'doc', fecha: fecha, archivo: 'autoprueba.xlsx', huella: huella, unicas: n, dup: 0, maxId: n, sinPlantel: 0,
      g: { 'CBTis 133': [n - 2, 1, 0] }, inc: { 'CBTis 133': { e: 1 } }, otras: {}, version: 'autoprueba' };
  }
  const r = [];
  r.push(['1. Guardar corte del 13 nov', guardar(corte('2026-11-13', 10, 'a1'), false).ok === true]);
  r.push(['2. Guardar otra vez el 13 nov (reemplaza)', guardar(corte('2026-11-13', 12, 'a2'), false).reemplazo === true]);
  r.push(['3. Corte del 14 nov con menos respuestas (debe pedir confirmación)', guardar(corte('2026-11-14', 5, 'a3'), false).codigo === 'menor_que_anterior']);
  r.push(['4. Mismo archivo con otra fecha (debe pedir confirmación)', guardar(corte('2026-11-14', 15, 'a2'), false).codigo === 'archivo_repetido']);
  const malo = corte('2026-11-14', 15, 'a4'); malo.unicas = 99;
  r.push(['5. Conteos que no cuadran (debe rechazar)', guardar(malo, false).codigo === 'invalido']);
  const ajeno = corte('2026-11-14', 15, 'a5'); ajeno.g = { 'CETis 999': [13, 1, 0] };
  r.push(['6. Plantel fuera del catálogo (debe rechazar)', guardar(ajeno, false).codigo === 'invalido']);
  const l = leer().doc.filter(function (x) { return x.fecha === '2026-11-13'; });
  r.push(['7. Solo hay un corte vigente del 13 nov y es el de 12 respuestas', l.length === 1 && l[0].datos.unicas === 12]);
  r.push(['8. Eliminar el corte del 13 nov', eliminar('doc', '2026-11-13').eliminados === 1]);
  r.forEach(function (x) { Logger.log((x[1] ? 'CORRECTO   ' : 'FALLÓ      ') + x[0]); });
  Logger.log(r.every(function (x) { return x[1]; }) ? 'Autoprueba completa: todo funciona.' : 'Hay pruebas fallidas: avísale a Claude.');
}
