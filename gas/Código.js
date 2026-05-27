// ══════════════════════════════════════════════════════════════════════════
//  SM Academia — Google Apps Script  (pega TODO este archivo en Apps Script)
//
//  Despliegue:
//    Implementaciones → Nueva implementación → Web app
//    Ejecutar como:  Yo (la cuenta del propietario)
//    Acceso:         Cualquiera (Anyone, even anonymous)
//    → Copia la URL /exec resultante y pégala en la app de admin
// ══════════════════════════════════════════════════════════════════════════

// ─── doGet ─ todas las peticiones de la admin-app (JSONP) ─────────────────
function doGet(e) {
  const action   = (e.parameter.action   || '').trim();
  const callback = (e.parameter.callback || '').trim();
  const sheetP   = e.parameter.sheet ? decodeURIComponent(e.parameter.sheet) : '';
  const idP      = e.parameter.id || '';
  let   rowP     = {};
  try { rowP = JSON.parse(e.parameter.row || '{}'); } catch (_) {}

  let result;
  try {
    switch (action) {
      case 'PING':                  result = ping();                          break;
      case 'GET_ALL':               result = getAllData();                     break;
      case 'GET_PAGOS':             result = getPagosData();                   break;
      case 'GET_SESIONES':          result = getSesionesData();                break;
      case 'GET_REGISTRO_HORARIO':  result = getRegistroHorario();             break;
      case 'GET_ASISTENCIAS_EXTRA': result = getAsistenciasExtra();            break;
      case 'GET_SHEET':             result = getSheetData(sheetP);             break;
      case 'INSPECT_SESIONES':      result = inspectSesiones();                break;
      case 'APPEND':                result = appendRow(sheetP, rowP);          break;
      case 'UPDATE':                result = updateRow(sheetP, idP, rowP);     break;
      case 'NUEVA_INSCRIPCION':     result = nuevaInscripcion(rowP);           break;
      default:
        result = { ok: false, error: 'Acción desconocida: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ─── doPost ─ peticiones desde Make.com (JSON body) ───────────────────────
function doPost(e) {
  let result;
  try {
    const body   = JSON.parse(e.postData.contents || '{}');
    const action = (body.action || '').trim();
    switch (action) {
      case 'NUEVA_INSCRIPCION': result = nuevaInscripcion(body);             break;
      case 'APPEND':            result = appendRow(body.sheet, body.row || {}); break;
      default:
        result = { ok: false, error: 'Acción POST desconocida: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// ══════════════════════════════════════════════════════════════════════════
//  HELPERS INTERNOS
// ══════════════════════════════════════════════════════════════════════════

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/** Primera hoja cuyo nombre contiene `kw` (insensible a mayúsculas). */
function sheetByKw_(kw) {
  const u = kw.toUpperCase();
  return ss_().getSheets().find(s => s.getName().toUpperCase().includes(u)) || null;
}

/** Convierte toda la hoja en array de objetos {cabecera: valor}. */
function sheetToObjects_(name) {
  const sh = ss_().getSheetByName(name) || sheetByKw_(name);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const hdrs = vals[0];
  return vals.slice(1)
    .map(row => {
      const obj = {};
      hdrs.forEach((h, i) => { if (String(h).trim()) obj[String(h).trim()] = row[i]; });
      return obj;
    })
    .filter(r => Object.values(r).some(v => v !== '' && v !== null && v !== undefined));
}

/**
 * Normaliza las claves de un objeto a snake_case ASCII minúsculas.
 * Úsalo para PAGOS, donde la app espera claves como id_pago, clave_alumno, etc.
 */
function normalizeKeys_(obj) {
  const norm = k => k
    .toLowerCase().trim()
    .replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i').replace(/ó/g,'o').replace(/ú/g,'u').replace(/ñ/g,'n')
    .replace(/[\s\-]+/g,'_')
    .replace(/[^a-z0-9_]/g,'_')
    .replace(/_+/g,'_')
    .replace(/^_|_$/g,'');
  const out = {};
  Object.entries(obj).forEach(([k, v]) => { out[norm(k)] = v; });
  return out;
}

/** Devuelve el siguiente ID numérico de una hoja (máx existente + 1). */
function nextNumericId_(sheetName, idColName) {
  const sh = ss_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return 1;
  const data = sh.getDataRange().getValues();
  const hdrs = data[0];
  const ci   = idColName ? hdrs.findIndex(h => String(h).trim() === idColName) : 0;
  const col  = ci >= 0 ? ci : 0;
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const n = parseInt(data[i][col]) || 0;
    if (n > max) max = n;
  }
  return max + 1;
}

/** Siguiente ID de INSCRIPCIONES con formato INSC-001, INSC-002 … */
function nextInscripcionId_() {
  const sh = ss_().getSheetByName('INSCRIPCIONES');
  if (!sh || sh.getLastRow() < 2) return 'INSC-001';
  const data = sh.getDataRange().getValues();
  const hdrs = data[0];
  const ci   = hdrs.findIndex(h => String(h).includes('ID_INSCRIPCI'));
  if (ci < 0) return 'INSC-' + String(sh.getLastRow()).padStart(3, '0');
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const n = parseInt(String(data[i][ci] || '').replace(/\D/g, '')) || 0;
    if (n > max) max = n;
  }
  return 'INSC-' + String(max + 1).padStart(3, '0');
}


// ══════════════════════════════════════════════════════════════════════════
//  ACCIONES DE LECTURA
// ══════════════════════════════════════════════════════════════════════════

function ping() {
  return { ok: true, ts: new Date().toISOString() };
}

function getAllData() {
  return {
    ok: true,
    data: {
      clientes:         sheetToObjects_('CLIENTES'),
      alumnos:          sheetToObjects_('ALUMNOS'),
      inscripciones:    sheetToObjects_('INSCRIPCIONES'),
      gastos:           sheetToObjects_('GASTOS'),
      personal:         sheetToObjects_('PERSONAL'),
      'eventos pitusa': sheetToObjects_('EVENTOS PITUSA'),
    }
  };
}

function getPagosData() {
  const raw   = sheetToObjects_('PAGOS');
  const pagos = raw
    .map(r => normalizeKeys_(r))
    .filter(p => String(p.id_pago || '').trim() !== '');
  return { ok: true, pagos, total: pagos.length, ts: new Date().toISOString() };
}

function getSesionesData() {
  const sesiones = sheetToObjects_('SESIONES');
  return { ok: true, sesiones };
}

function getRegistroHorario() {
  const registros = sheetToObjects_('REGISTRO HORARIO');
  return { ok: true, registros };
}

function getAsistenciasExtra() {
  const asistencias = sheetToObjects_('ASISTENCIAS_EXTRA');
  return { ok: true, asistencias };
}

function getSheetData(sheetName) {
  if (!sheetName) return { ok: false, error: 'Parámetro "sheet" requerido' };
  const rows = sheetToObjects_(sheetName);
  return { ok: true, rows, sheet: sheetName };
}

function inspectSesiones() {
  const hojas  = ss_().getSheets().map(s => s.getName());
  const sesSh  = sheetByKw_('SESIONES');
  const info   = {
    ok:            true,
    todasLasHojas: hojas,
    hojaSesiones:  sesSh ? sesSh.getName() : null,
    filas:         sesSh ? sesSh.getLastRow()    : 0,
    columnas:      sesSh ? sesSh.getLastColumn() : 0,
  };
  if (sesSh && sesSh.getLastRow() > 0) {
    info.cabeceras = sesSh.getRange(1, 1, 1, sesSh.getLastColumn()).getValues()[0];
    if (sesSh.getLastRow() > 1) {
      info.primeraFila = sesSh.getRange(2, 1, 1, sesSh.getLastColumn()).getValues()[0];
    }
  }
  return info;
}


// ══════════════════════════════════════════════════════════════════════════
//  ACCIONES DE ESCRITURA
// ══════════════════════════════════════════════════════════════════════════

/**
 * Añade una fila al final de `sheetName` mapeando por nombre de columna.
 * rowObj: objeto {NombreColumna: valor} o array de valores.
 */
function appendRow(sheetName, rowObj) {
  if (!sheetName) return { ok: false, error: 'Parámetro "sheet" requerido' };
  const sh = ss_().getSheetByName(sheetName) || sheetByKw_(sheetName);
  if (!sh) return { ok: false, error: 'Hoja no encontrada: ' + sheetName };

  if (typeof rowObj === 'string') {
    try { rowObj = JSON.parse(rowObj); } catch(_) { return { ok: false, error: 'row JSON inválido' }; }
  }
  if (Array.isArray(rowObj)) {
    sh.appendRow(rowObj);
    return { ok: true };
  }

  const hdrs   = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  const newRow = hdrs.map(h => {
    const key = String(h).trim();
    return rowObj[key] !== undefined ? rowObj[key] : '';
  });
  sh.appendRow(newRow);
  return { ok: true };
}

/**
 * Actualiza la fila de `sheetName` cuya primera columna coincide con `id`.
 * Sobrescribe solo las claves presentes en rowObj; el resto queda intacto.
 */
function updateRow(sheetName, id, rowObj) {
  if (!sheetName || id === undefined || id === '') {
    return { ok: false, error: 'Parámetros "sheet" e "id" requeridos' };
  }
  const sh = ss_().getSheetByName(sheetName) || sheetByKw_(sheetName);
  if (!sh) return { ok: false, error: 'Hoja no encontrada: ' + sheetName };

  if (typeof rowObj === 'string') {
    try { rowObj = JSON.parse(rowObj); } catch(_) { return { ok: false, error: 'row JSON inválido' }; }
  }

  const data = sh.getDataRange().getValues();
  const hdrs = data[0];
  const idStr = String(id).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === idStr) {
      const updated = hdrs.map((h, ci) => {
        const key = String(h).trim();
        return rowObj[key] !== undefined ? rowObj[key] : data[i][ci];
      });
      sh.getRange(i + 1, 1, 1, updated.length).setValues([updated]);
      return { ok: true, rowIndex: i + 1 };
    }
  }
  return { ok: false, error: 'No se encontró id=' + id + ' en ' + sheetName };
}


// ══════════════════════════════════════════════════════════════════════════
//  NUEVA INSCRIPCIÓN  (admin-app + Make.com webhook)
// ══════════════════════════════════════════════════════════════════════════

/**
 * Registra un nuevo alumno en CLIENTES + ALUMNOS + INSCRIPCIONES.
 *
 * Acepta dos formatos en `body`:
 *   • Plano (admin-app):  { nombre_alumno, apellidos_alumno, email, ... }
 *   • Labels Jetpack (Make.com): { "Nombre del jugador/a", "Correo electrónico", ... }
 *
 * Devuelve { ok: true, inscrId: 'INSC-XXX' }
 */
function nuevaInscripcion(body) {
  // Helper: prueba primero clave GAS, luego label original del formulario
  const v = (gasKey, label) => String(body[gasKey] || body[label] || '').trim();

  const nomAlu        = v('nombre_alumno',     'Nombre del jugador/a');
  const apeAlu        = v('apellidos_alumno',  'Apellidos del jugador/a');
  const fechaNac      = v('fecha_nac',         'Fecha nacimiento (AAAA-MM-DD)');
  const colegio       = v('colegio',           'Colegio');
  const observaciones = v('observaciones',     'Observaciones');
  const nomTut        = v('nombre_tutor',      'Nombre del tutor/a');
  const apeTut        = v('apellidos_tutor',   'Apellidos del tutor/a');
  const email         = v('email',             'Correo electrónico').toLowerCase();
  const telefono      = v('telefono',          'Teléfono');
  const servicio      = v('servicio',          'Servicio').toUpperCase();
  const modalidad     = v('modalidad',         'Modalidad').toUpperCase();
  const prefCentro    = v('preferencia_centro','Preferencias');
  const dias          = v('dias',              'Días preferidos');
  const autoriza      = v('autoriza_imagen',   'Autorizo uso de imagen') || 'Sí';

  if (!nomAlu || !email) {
    return { ok: false, error: 'Campos obligatorios: nombre_alumno y email' };
  }

  const shCli = ss_().getSheetByName('CLIENTES');
  const shAlu = ss_().getSheetByName('ALUMNOS');
  const shIns = ss_().getSheetByName('INSCRIPCIONES');
  if (!shCli || !shAlu || !shIns) {
    return { ok: false, error: 'Hojas CLIENTES, ALUMNOS o INSCRIPCIONES no encontradas' };
  }

  const tz    = Session.getScriptTimeZone();
  const hoy   = new Date();
  const fecha = Utilities.formatDate(hoy, tz, 'yyyy-MM-dd');
  const mes   = Utilities.formatDate(hoy, tz, 'MMMM yyyy');

  // ── 1. CLIENTES — evitar duplicados por email ─────────────────────────
  let clienteId;
  const cliData    = shCli.getDataRange().getValues();
  const cliHdrs    = cliData[0];
  const emailIdx   = cliHdrs.findIndex(h => String(h).toUpperCase().includes('EMAIL'));
  const idCliIdx   = cliHdrs.findIndex(h => String(h).trim() === 'ID_CLIENTE');

  let yaExiste = false;
  for (let i = 1; i < cliData.length; i++) {
    if (String(cliData[i][emailIdx] || '').toLowerCase().trim() === email) {
      clienteId  = parseInt(cliData[i][idCliIdx >= 0 ? idCliIdx : 0]) || i;
      yaExiste   = true;
      break;
    }
  }

  if (!yaExiste) {
    clienteId             = nextNumericId_('CLIENTES', 'ID_CLIENTE');
    const nomTutFull      = (nomTut + ' ' + apeTut).trim().toUpperCase();
    const claveCliente    = clienteId + '|' + nomTutFull;
    appendRow('CLIENTES', {
      'ID_CLIENTE':    clienteId,
      'NOMBRE':        nomTut,
      'APELLIDOS':     apeTut,
      'TELÉFONO 1':    telefono,
      'EMAIL':         email,
      'CLAVE CLIENTE': claveCliente,
    });
  }

  // ── 2. ALUMNOS ────────────────────────────────────────────────────────
  const alumnoId    = nextNumericId_('ALUMNOS', 'ID_ALUMNO');
  const nomAluFull  = (nomAlu + ' ' + apeAlu).trim().toUpperCase();
  const claveAlu    = alumnoId + '|' + nomAluFull;
  const claveCli    = clienteId + '|' + (nomTut + ' ' + apeTut).trim().toUpperCase();

  appendRow('ALUMNOS', {
    'ID_ALUMNO':           alumnoId,
    'NOMBRE':              nomAlu,
    'APELLIDOS':           apeAlu,
    'ACTIVO':              'SÍ',
    'FECHA NACIMIENTO':    fechaNac,
    'CLAVE_CLIENTE':       claveCli,
    'CLAVE ALUMNO':        claveAlu,
    'CESIÓN IMAGEN':       autoriza.toUpperCase().startsWith('S') ? 'SÍ' : 'NO',
    'OBSERVACIONES':       observaciones,
    'TOTAL INSCRIPCIONES': 0,
  });

  // ── 3. INSCRIPCIONES ──────────────────────────────────────────────────
  const inscrId  = nextInscripcionId_();
  const claveIns = inscrId + '|' + nomAluFull;

  appendRow('INSCRIPCIONES', {
    'ID_INSCRIPCIÓN':    inscrId,
    'MES':               mes,
    'FECHA':             fecha,
    'CLAVE CLIENTE':     claveCli,
    'CLAVE ALUMNO':      claveAlu,
    'SERVICIO':          servicio,
    'MODALIDAD':         modalidad,
    'IGIC':              3,
    'CLAVE INSCRIPCIÓN': claveIns,
    'PRECIO':            '',
    'CENTRO':            prefCentro,
    'GRUPO_DIA':         dias,
    'ACTIVA':            'SÍ',
  });

  return { ok: true, inscrId };
}
