// ══════════════════════════════════════════════════════════════════════════
//  SM Academia — Google Apps Script  (pega TODO este archivo en Apps Script)
//
//  Despliegue:
//    Implementaciones → Nueva implementación → Web app
//    Ejecutar como:  Yo (la cuenta del propietario)
//    Acceso:         Cualquiera (Anyone, even anonymous)
//    → Copia la URL /exec resultante y pégala en la app de admin
//
//  Script Properties requeridas (para VeriFactu AEAT):
//    VF_PRIVATE_KEY_PEM  — clave privada RSA en formato PEM (opcional)
//    VF_CERT_B64         — certificado X.509 en Base64 sin cabeceras (opcional)
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
      case 'SEND_ALERTS_IMPAGOS':   result = sendAlertsImpagos(rowP);          break;
      case 'SUBMIT_VERIFACTU':      result = submitVerifactu(rowP);            break;
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
      case 'NUEVA_INSCRIPCION':     result = nuevaInscripcion(body);                 break;
      case 'NUEVA_PREINSCRIPCION':  result = nuevaPreinscripcion(body);             break;
      case 'APPEND':                result = appendRow(body.sheet, body.row || {}); break;
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
      clientes:          sheetToObjects_('CLIENTES'),
      alumnos:           sheetToObjects_('ALUMNOS'),
      inscripciones:     sheetToObjects_('INSCRIPCIONES'),
      gastos:            sheetToObjects_('GASTOS'),
      personal:          sheetToObjects_('PERSONAL'),
      preinscripciones:  sheetToObjects_('PREINSCRIPCIONES'),
      actividades:       sheetToObjects_('ACTIVIDADES'),
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
//  NUEVA INSCRIPCIÓN  (admin-app + Make.com mailhook)
// ══════════════════════════════════════════════════════════════════════════

/**
 * Registra un nuevo alumno en CLIENTES + ALUMNOS + INSCRIPCIONES.
 *
 * Acepta tres formatos en `body`:
 *   • Plano (admin-app):  { nombre_alumno, apellidos_alumno, email, ... }
 *   • Labels Jetpack (Make.com): { "Nombre del jugador/a", "Correo electrónico", ... }
 *   • Email body (Make.com mailhook): { raw_text: "Campo: Valor\nCampo2: Valor2\n..." }
 *
 * Devuelve { ok: true, inscrId: 'INSC-XXX' }
 */
function nuevaInscripcion(body) {
  // ── [A] Parsear email body de Jetpack (Make.com mailhook) ─────────────
  // Make.com envía el cuerpo del email como raw_text / body_text / text
  const rawEmail = body.raw_text || body.body_text || body.body_plain || body.text || '';
  if (rawEmail) {
    rawEmail.split(/\r?\n/).forEach(function(line) {
      const colon = line.indexOf(':');
      if (colon > 0) {
        const k = line.substring(0, colon).trim();
        const v = line.substring(colon + 1).trim();
        if (k && v && body[k] === undefined) body[k] = v;
      }
    });
  }

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
    'RGPD_FIRMADO':      fecha,
  });

  return { ok: true, inscrId };
}


// ══════════════════════════════════════════════════════════════════════════
//  NUEVA PREINSCRIPCIÓN  (formulario de interés familias — Escaleritas)
// ══════════════════════════════════════════════════════════════════════════

/**
 * Registra una preinscripción de interés en la hoja PREINSCRIPCIONES.
 *
 * Columnas esperadas en la hoja (créalas en ese orden):
 *   ID_PREINSCRIPCION | FECHA | ESTADO | CURSO_ACADEMICO | CENTRO
 *   NOMBRE_ALUMNO | APELLIDOS_ALUMNO | CURSO_ALUMNO | FECHA_NAC | SALUD | OBSERVACIONES
 *   NOMBRE_TUTOR1 | APELLIDOS_TUTOR1 | RELACION_TUTOR1 | NIF_TUTOR1 | TELEFONO_TUTOR1 | EMAIL_TUTOR1
 *   NOMBRE_TUTOR2 | APELLIDOS_TUTOR2 | RELACION_TUTOR2 | NIF_TUTOR2 | TELEFONO_TUTOR2 | EMAIL_TUTOR2
 *   AUTORIZA_IMAGEN | ACEPTA_DATOS | ACEPTA_DATOS_SALUD | SOCIO_AMPA | IBAN
 *   ACTIVIDADES | DIAS_DISPONIBLES
 */
function nuevaPreinscripcion(body) {
  const v = (k1, k2) => String(body[k1] || (k2 ? body[k2] : '') || '').trim();

  const nomAlu      = v('nombre_alumno',      'Nombre del alumno/a');
  const apeAlu      = v('apellidos_alumno',   'Apellidos del alumno/a');
  if (!nomAlu || !apeAlu) {
    return { ok: false, error: 'Campos obligatorios: nombre_alumno y apellidos_alumno' };
  }

  const sh = ss_().getSheetByName('PREINSCRIPCIONES');
  if (!sh) return { ok: false, error: 'Hoja PREINSCRIPCIONES no encontrada. Créala con las cabeceras indicadas.' };

  const tz    = Session.getScriptTimeZone();
  const fecha = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // Calcular siguiente ID
  const lastRow = sh.getLastRow();
  let nextId = 'PREINSC-001';
  if (lastRow >= 2) {
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    let max = 0;
    ids.forEach(function(id) {
      const n = parseInt(String(id).replace(/\D/g, '')) || 0;
      if (n > max) max = n;
    });
    nextId = 'PREINSC-' + String(max + 1).padStart(3, '0');
  }

  // Actividades puede llegar como array (checkboxes) o string CSV
  let actividades = body.actividades || body['Actividades de interés'] || '';
  if (Array.isArray(actividades)) actividades = actividades.join(', ');

  appendRow('PREINSCRIPCIONES', {
    'ID_PREINSCRIPCION': nextId,
    'FECHA':             fecha,
    'ESTADO':            'RECIBIDA',
    'CURSO_ACADEMICO':   v('curso_academico')   || '2026-2027',
    'CENTRO':            v('centro')             || 'ESCALERITAS',
    'NOMBRE_ALUMNO':     nomAlu,
    'APELLIDOS_ALUMNO':  apeAlu,
    'CURSO_ALUMNO':      v('curso_alumno',       'Curso escolar'),
    'FECHA_NAC':         v('fecha_nac',          'Fecha de nacimiento'),
    'SALUD':             v('salud',              'Enfermedades/alergias'),
    'OBSERVACIONES':     v('observaciones',      'Observaciones'),
    'NOMBRE_TUTOR1':     v('nombre_tutor1',      'Nombre tutor 1'),
    'APELLIDOS_TUTOR1':  v('apellidos_tutor1',   'Apellidos tutor 1'),
    'RELACION_TUTOR1':   v('relacion_tutor1',    'Relación tutor 1'),
    'NIF_TUTOR1':        v('nif_tutor1',         'NIF/DNI tutor 1'),
    'TELEFONO_TUTOR1':   v('telefono_tutor1',    'Teléfono tutor 1'),
    'EMAIL_TUTOR1':      v('email_tutor1',       'Email tutor 1').toLowerCase(),
    'NOMBRE_TUTOR2':     v('nombre_tutor2',      'Nombre tutor 2'),
    'APELLIDOS_TUTOR2':  v('apellidos_tutor2',   'Apellidos tutor 2'),
    'RELACION_TUTOR2':   v('relacion_tutor2',    'Relación tutor 2'),
    'NIF_TUTOR2':        v('nif_tutor2',         'NIF/DNI tutor 2'),
    'TELEFONO_TUTOR2':   v('telefono_tutor2',    'Teléfono tutor 2'),
    'EMAIL_TUTOR2':      v('email_tutor2',       'Email tutor 2').toLowerCase(),
    'AUTORIZA_IMAGEN':   v('autoriza_imagen',    'Cesión de imágenes').toUpperCase().startsWith('S') ? 'SÍ' : 'NO',
    'ACEPTA_DATOS':      v('acepta_datos',       'Política de datos').toUpperCase().startsWith('S') ? 'SÍ' : 'NO',
    'ACEPTA_DATOS_SALUD':v('acepta_datos_salud', 'Consentimiento datos salud').toUpperCase().startsWith('S') ? 'SÍ' : 'NO',
    'SOCIO_AMPA':        v('socio_ampa',         '¿Desea hacerse socio del AMPA?').toUpperCase().startsWith('S') ? 'SÍ' : 'NO',
    'IBAN':              v('iban',               'IBAN para domiciliación'),
    'ACTIVIDADES':       String(actividades),
    'DIAS_DISPONIBLES':  v('dias_disponibles',   'Días disponibles'),
  });

  return { ok: true, preinscId: nextId };
}


// ══════════════════════════════════════════════════════════════════════════
//  [D] ALERTAS DE IMPAGOS — envía email a familias con pago pendiente
// ══════════════════════════════════════════════════════════════════════════

/**
 * Envía emails de aviso a familias con pagos pendientes en el mes indicado.
 * params: { mes: 'mayo', anyo: 2026 }
 * Devuelve { ok: true, enviados: N, errores: [...], total: N }
 */
function sendAlertsImpagos(params) {
  const mes  = String(params.mes || '').toLowerCase().trim();
  const anyo = parseInt(params.anyo) || new Date().getFullYear();
  if (!mes) return { ok: false, error: 'Parámetro "mes" requerido' };

  const pagos         = sheetToObjects_('PAGOS').map(r => normalizeKeys_(r));
  const clientes      = sheetToObjects_('CLIENTES');
  const inscripciones = sheetToObjects_('INSCRIPCIONES');

  if (!clientes.length || !inscripciones.length) {
    return { ok: false, error: 'Hojas CLIENTES o INSCRIPCIONES vacías' };
  }

  // Alumnos que SÍ tienen PAGADO en el mes
  const pagadosIds = new Set(
    pagos.filter(p => {
      const pMes  = String(p.mes || '').toLowerCase();
      const pAnyo = p.fecha ? new Date(String(p.fecha)).getFullYear() : anyo;
      return pMes === mes && pAnyo === anyo &&
             String(p.estado || '').toUpperCase() === 'PAGADO';
    }).map(p => String(p.id_alumno || p.alumnoid || p.alumno_id || ''))
  );

  // Inscripciones activas que deben pagar (tecnificación / recogida, modalidad mensual)
  const SERV_OK = ['TECNIF', 'RECOGIDA'];
  const MOD_NO  = ['SUELTA', 'PUNTUAL', 'ESPORAD'];
  const MOD_OK  = ['SEMANA', 'MENSUAL'];

  const inscDeuda = inscripciones.filter(i => {
    const serv   = String(i.SERVICIO   || i.servicio   || '').toUpperCase();
    const mod    = String(i.MODALIDAD  || i.modalidad  || '').toUpperCase();
    const activa = String(i.ACTIVA     || i.activa     || '').toUpperCase();
    if (activa !== 'SÍ' && activa !== 'SI') return false;
    if (!SERV_OK.some(s => serv.includes(s))) return false;
    if (MOD_NO.some(m  => mod.includes(m)))   return false;
    if (!MOD_OK.some(m => mod.includes(m)))   return false;
    return true;
  });

  // Quedarnos solo con los que NO han pagado
  const sinPagar = inscDeuda.filter(i => {
    const aId = String(i.ID_ALUMNO || i['ID ALUMNO'] || '');
    return !pagadosIds.has(aId);
  });

  const mesLabel = mes.charAt(0).toUpperCase() + mes.slice(1) + ' ' + anyo;
  let enviados   = 0;
  const errores  = [];

  sinPagar.forEach(function(i) {
    const cliKey = String(i['CLAVE CLIENTE'] || i.clave_cliente || '');
    const nomAlu = (String(i['CLAVE ALUMNO'] || i.clave_alumno || '').split('|')[1] || 'el/la alumno/a').trim();
    const servicio = String(i.SERVICIO || i.servicio || '');

    // Buscar cliente por clave
    const cli = clientes.find(function(c) {
      return String(c['CLAVE CLIENTE'] || '').trim() === cliKey;
    });
    if (!cli) return;

    const emailDest = String(cli.EMAIL || cli['EMAIL TUTOR'] || '').toLowerCase().trim();
    if (!emailDest || emailDest.indexOf('@') < 0) return;

    const nomTutor = String(cli.NOMBRE || '').split(' ')[0] || 'familia';

    try {
      MailApp.sendEmail({
        to:        emailDest,
        subject:   'SM Academia · Pago pendiente ' + mesLabel,
        body:
          'Hola ' + nomTutor + ',\n\n' +
          'Te escribimos desde SM Academia porque en nuestro sistema no consta el pago de ' +
          mesLabel + ' correspondiente a ' + nomAlu + ' (' + servicio + ').\n\n' +
          'Si ya has realizado el pago, indícanos la fecha y el método para actualizar nuestros registros. ' +
          'Si tienes alguna duda, estamos a tu disposición.\n\n' +
          'Un saludo,\nSamy Martín\nSM Academia\nhttps://academiasmfutbol.com',
        htmlBody:
          '<p>Hola <strong>' + nomTutor + '</strong>,</p>' +
          '<p>Te escribimos desde <strong>SM Academia</strong> porque en nuestro sistema no consta el pago de ' +
          '<strong>' + mesLabel + '</strong> correspondiente a <strong>' + nomAlu + '</strong> ' +
          '(<em>' + servicio + '</em>).</p>' +
          '<p>Si ya has realizado el pago, por favor indícanos la fecha y el método para actualizar nuestros registros. ' +
          'Si tienes alguna duda, estamos a tu disposición.</p>' +
          '<p>Un saludo,<br><strong>Samy Martín</strong><br>SM Academia<br>' +
          '<a href="https://academiasmfutbol.com">academiasmfutbol.com</a></p>',
      });
      enviados++;
    } catch (e) {
      errores.push(emailDest + ': ' + e.message);
    }
  });

  return { ok: true, enviados: enviados, errores: errores, total: sinPagar.length };
}


// ══════════════════════════════════════════════════════════════════════════
//  [C] VERIFACTU AEAT — envío de facturas al sistema de la AEAT
//
//  Requisito previo: en Script Properties de Apps Script configura:
//    VF_PRIVATE_KEY_PEM  → clave privada RSA de tu certificado (formato PEM)
//    VF_CERT_B64         → certificado X.509 en Base64 (sin cabeceras PEM)
//    VF_ENTORNO          → 'TEST' | 'PROD' (por defecto TEST)
//
//  Endpoint AEAT test: https://prewww1.aeat.es/wlpl/TIKE-CONT/ws/SistemaFacturacion/VerifactuSOAP
//  Endpoint AEAT prod: https://www1.aeat.es/wlpl/TIKE-CONT/ws/SistemaFacturacion/VerifactuSOAP
// ══════════════════════════════════════════════════════════════════════════

function submitVerifactu(params) {
  // params: { factura: {...}, xmlBody: '...', entorno: 'TEST'|'PROD' }
  const props       = PropertiesService.getScriptProperties();
  const privateKey  = props.getProperty('VF_PRIVATE_KEY_PEM') || '';
  const certB64     = props.getProperty('VF_CERT_B64')        || '';
  const entorno     = (params.entorno || props.getProperty('VF_ENTORNO') || 'TEST').toUpperCase();

  const ENDPOINT_TEST = 'https://prewww1.aeat.es/wlpl/TIKE-CONT/ws/SistemaFacturacion/VerifactuSOAP';
  const ENDPOINT_PROD = 'https://www1.aeat.es/wlpl/TIKE-CONT/ws/SistemaFacturacion/VerifactuSOAP';
  const endpoint      = entorno === 'PROD' ? ENDPOINT_PROD : ENDPOINT_TEST;

  // Guardar en Sheets con estado PENDIENTE_ENVIO aunque falle el envío
  const factura = params.factura || {};
  if (factura['Id Factura']) {
    try {
      updateRow('FACTURAS', factura['Id Factura'], {
        'Estado VeriFactu': privateKey ? 'ENVIANDO' : 'CERT_NO_CONFIGURADO',
      });
    } catch(_) {}
  }

  if (!privateKey || !certB64) {
    return {
      ok:    false,
      error: 'Certificado no configurado. Añade VF_PRIVATE_KEY_PEM y VF_CERT_B64 en Script Properties del GAS.',
      ayuda: 'Abre Apps Script → Proyecto → Configuración del proyecto → Propiedades de script'
    };
  }

  // Construir XML VeriFactu según esquema AEAT 1.0
  const xmlSoap = buildVerifactuSoap_(params.xmlBody || buildVerifactuXml_(factura), certB64, privateKey);
  if (!xmlSoap.ok) return xmlSoap;

  try {
    const response = UrlFetchApp.fetch(endpoint, {
      method:      'post',
      contentType: 'text/xml; charset=UTF-8',
      headers: {
        'SOAPAction': 'https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/tikeV1/cont/ws/SistemaFacturacion/RegFactuSistemaFacturacion',
      },
      payload:          xmlSoap.xml,
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    const respBody   = response.getContentText();

    // Parsear respuesta AEAT
    const estadoAEAT = respBody.includes('<CSV>') ? 'ACEPTADA' :
                       respBody.includes('KO')    ? 'RECHAZADA' : 'DESCONOCIDO';

    // Actualizar Sheets con resultado
    if (factura['Id Factura']) {
      try {
        updateRow('FACTURAS', factura['Id Factura'], {
          'Estado VeriFactu': estadoAEAT,
          'CSV AEAT':         (respBody.match(/<CSV>([^<]+)<\/CSV>/) || ['',''])[1],
          'Fecha envío AEAT': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
        });
      } catch(_) {}
    }

    return {
      ok:          statusCode === 200,
      statusCode:  statusCode,
      estado:      estadoAEAT,
      entorno:     entorno,
      responseXml: respBody.substring(0, 2000), // truncar para no saturar la respuesta
    };

  } catch (e) {
    return { ok: false, error: 'Error de red: ' + e.message };
  }
}

/** Construye el XML de la factura según esquema VeriFactu 1.0 */
function buildVerifactuXml_(f) {
  const nif    = f['NIF Emisor']     || '45766626F';
  const nombre = f['Nombre emisor']  || 'Antonio Samuel Martín Rivera';
  const numFac = f['Id Factura']     || '';
  const fecha  = f['Fecha expedición'] || '';
  const hora   = f['Hora expedición']  || '00:00:00';
  const total  = parseFloat(f['Total factura'] || 0).toFixed(2);
  const hash   = f['Hash VeriFactu'] || '';

  return (
    '<sfc:SuministroLRFacturasEmitidas ' +
    'xmlns:sfc="https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/tikeV1/cont/ws/SistemaFacturacion" ' +
    'xmlns:sf="https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/tikeV1/cont/xml/catalogos">' +
      '<sfc:Cabecera>' +
        '<sfc:Obligado>' +
          '<sf:NombreRazon>' + xmlEscape_(nombre) + '</sf:NombreRazon>' +
          '<sf:NIF>' + xmlEscape_(nif) + '</sf:NIF>' +
        '</sfc:Obligado>' +
      '</sfc:Cabecera>' +
      '<sfc:RegistroFactura>' +
        '<sfc:RegistroFacturacion>' +
          '<sfc:IDVersion>1.0</sfc:IDVersion>' +
          '<sfc:IDFactura>' +
            '<sfc:IDEmisorFactura>' + xmlEscape_(nif) + '</sfc:IDEmisorFactura>' +
            '<sfc:NumSerieFactura>' + xmlEscape_(numFac) + '</sfc:NumSerieFactura>' +
            '<sfc:FechaExpedicionFactura>' + xmlEscape_(fecha) + '</sfc:FechaExpedicionFactura>' +
          '</sfc:IDFactura>' +
          '<sfc:NombreRazonEmisor>' + xmlEscape_(nombre) + '</sfc:NombreRazonEmisor>' +
          '<sfc:TipoFactura>F1</sfc:TipoFactura>' +
          '<sfc:DescripcionOperacion>' + xmlEscape_(f['Descripción operación'] || '') + '</sfc:DescripcionOperacion>' +
          '<sfc:ImporteTotal>' + total + '</sfc:ImporteTotal>' +
          '<sfc:FechaHoraHuella>' + xmlEscape_(fecha + 'T' + hora) + '</sfc:FechaHoraHuella>' +
          '<sfc:Huella>' + xmlEscape_(hash) + '</sfc:Huella>' +
        '</sfc:RegistroFacturacion>' +
      '</sfc:RegistroFactura>' +
    '</sfc:SuministroLRFacturasEmitidas>'
  );
}

/** Envuelve el XML en un SOAP envelope firmado (WS-Security XMLDSig RSA-SHA256) */
function buildVerifactuSoap_(bodyXml, certB64, privateKeyPem) {
  try {
    // ID del body para la referencia en la firma
    const bodyId = 'Body-' + Utilities.getUuid().replace(/-/g,'').substring(0,16);

    const bodyElem =
      '<soapenv:Body xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" wsu:Id="' + bodyId + '" ' +
      'xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">' +
      bodyXml +
      '</soapenv:Body>';

    // Digest SHA-256 del body (canonicalización simple — sin C14N completo)
    const bodyDigest = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bodyElem, Utilities.Charset.UTF_8)
    );

    const signedInfo =
      '<ds:SignedInfo xmlns:ds="http://www.w3.org/2000/09/xmldsig#">' +
        '<ds:CanonicalizationMethod Algorithm="http://www.w3.org/2001/10/xml-exc-c14n#"/>' +
        '<ds:SignatureMethod Algorithm="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"/>' +
        '<ds:Reference URI="#' + bodyId + '">' +
          '<ds:Transforms>' +
            '<ds:Transform Algorithm="http://www.w3.org/2001/10/xml-exc-c14n#"/>' +
          '</ds:Transforms>' +
          '<ds:DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>' +
          '<ds:DigestValue>' + bodyDigest + '</ds:DigestValue>' +
        '</ds:Reference>' +
      '</ds:SignedInfo>';

    // Firmar SignedInfo con la clave privada RSA
    const signatureBytes = Utilities.computeRsaSha256Signature(signedInfo, privateKeyPem);
    const signatureB64   = Utilities.base64Encode(signatureBytes);

    const certId = 'Cert-' + Utilities.getUuid().replace(/-/g,'').substring(0,16);

    const header =
      '<soapenv:Header>' +
        '<wsse:Security xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" ' +
        'xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">' +
          '<wsse:BinarySecurityToken ' +
            'EncodingType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-soap-message-security-1.0#Base64Binary" ' +
            'ValueType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-x509-token-profile-1.0#X509v3" ' +
            'wsu:Id="' + certId + '">' + certB64 +
          '</wsse:BinarySecurityToken>' +
          '<ds:Signature xmlns:ds="http://www.w3.org/2000/09/xmldsig#">' +
            signedInfo +
            '<ds:SignatureValue>' + signatureB64 + '</ds:SignatureValue>' +
            '<ds:KeyInfo>' +
              '<wsse:SecurityTokenReference>' +
                '<wsse:Reference URI="#' + certId + '" ' +
                'ValueType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-x509-token-profile-1.0#X509v3"/>' +
              '</wsse:SecurityTokenReference>' +
            '</ds:KeyInfo>' +
          '</ds:Signature>' +
        '</wsse:Security>' +
      '</soapenv:Header>';

    const envelope =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<soapenv:Envelope ' +
      'xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" ' +
      'xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">' +
      header + bodyElem +
      '</soapenv:Envelope>';

    return { ok: true, xml: envelope };
  } catch(e) {
    return { ok: false, error: 'Error al construir SOAP: ' + e.message };
  }
}

function xmlEscape_(s) {
  return String(s || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}
