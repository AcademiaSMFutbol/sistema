// ═══════════════════════════════════════════════════════════════════════════
//  SM ACADEMIA — Google Apps Script v2.3
//  Sheets ID: 1oHJIUoyR3V5N0iweWnMagZic_8YkWtVV6CsshXloX4k
// ═══════════════════════════════════════════════════════════════════════════

var SPREADSHEET_ID = '1oHJIUoyR3V5N0iweWnMagZic_8YkWtVV6CsshXloX4k';

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ── Peticiones GET (JSONP desde la app) ───────────────────────────────────
function doGet(e) {
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : 'callback';
  var action   = (e && e.parameter && e.parameter.action)   ? e.parameter.action   : '';

  try {
    var result;
    switch (action) {

      case 'PING':
        result = { ok: true, ts: new Date().toISOString() };
        break;

      case 'GET_ALL':
        result = getAllData();
        break;

      case 'APPEND':
        result = appendRow(
          e.parameter.sheet,
          JSON.parse(e.parameter.row || '{}')
        );
        break;

      case 'UPDATE':
        result = updateRow(
          e.parameter.sheet,
          e.parameter.id,
          JSON.parse(e.parameter.row || '{}')
        );
        break;

      case 'DELETE':
        result = deleteRow(e.parameter.sheet, e.parameter.id);
        break;

      default:
        result = { ok: false, error: 'Acción desconocida: ' + action };
    }

    return jsonp(callback, result);

  } catch (err) {
    return jsonp(callback, { ok: false, error: err.toString() });
  }
}

// ── Peticiones POST (formulario público inscripcion.html) ─────────────────
function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var result = nuevaPreinscripcion(data);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Helper JSONP ──────────────────────────────────────────────────────────
function jsonp(callback, data) {
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(data) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ═══════════════════════════════════════════════════════════════════════════
//  LEER TODAS LAS HOJAS
// ═══════════════════════════════════════════════════════════════════════════
function getAllData() {
  var book = ss_();
  return {
    ok:                      true,
    clientes:                sheetToJSON(book, 'CLIENTES'),
    alumnos:                 sheetToJSON(book, 'ALUMNOS'),
    actividades:             sheetToJSON(book, 'ACTIVIDADES'),
    preinscripciones:        sheetToJSON(book, 'PREINSCRIPCIONES'),
    inscripciones:           sheetToJSON(book, 'INSCRIPCIONES'),
    pagos:                   sheetToJSON(book, 'PAGOS'),
    facturas:                sheetToJSON(book, 'FACTURAS'),
    personal:                sheetToJSON(book, 'PERSONAL'),
    ausencias:               sheetToJSON(book, 'AUSENCIAS'),
    registroHorario:         sheetToJSON(book, 'REGISTRO_HORARIO'),
    sesiones:                sheetToJSON(book, 'SESIONES'),
    asistenciasFutbol:       sheetToJSON(book, 'ASISTENCIAS_FUTBOL'),
    asistenciasExtra:        sheetToJSON(book, 'ASISTENCIAS_EXTRA'),
    nominas:                 sheetToJSON(book, 'NOMINAS'),
    gastos:                  sheetToJSON(book, 'GASTOS'),
    movimientosBancarios:    sheetToJSON(book, 'MOVIMIENTOS_BANCARIOS'),
    ampaSocios:              sheetToJSON(book, 'AMPA_SOCIOS'),
    ampaAlumnos:             sheetToJSON(book, 'AMPA_ALUMNOS'),
    encuestasSatisfaccion:   sheetToJSON(book, 'ENCUESTAS_SATISFACCION'),
  };
}

// Convierte una hoja en array de objetos — si la hoja no existe devuelve []
function sheetToJSON(book, sheetName) {
  var sheet = book.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('AVISO: hoja no encontrada → ' + sheetName);
    return [];
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var data    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var rows    = [];

  for (var i = 1; i < data.length; i++) {
    var row     = data[i];
    var isEmpty = row.every(function(c) { return c === '' || c === null || c === undefined; });
    if (isEmpty) continue;

    var obj = {};
    headers.forEach(function(h, j) {
      var v = row[j];
      if (v instanceof Date) {
        // Columnas de hora → serializar como HH:mm (no perder la hora)
        if (h.indexOf('HORA') !== -1) {
          obj[h] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
        } else {
          obj[h] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      } else {
        obj[h] = (v !== null && v !== undefined) ? v : '';
      }
    });
    rows.push(obj);
  }
  return rows;
}

// ═══════════════════════════════════════════════════════════════════════════
//  AÑADIR FILA (APPEND)
// ═══════════════════════════════════════════════════════════════════════════
function appendRow(sheetName, rowObj) {
  var book  = ss_();
  var sheet = book.getSheetByName(sheetName);
  if (!sheet) return { ok: false, error: 'Hoja no encontrada: ' + sheetName };

  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { ok: false, error: 'La hoja ' + sheetName + ' no tiene cabeceras' };

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h).trim(); });

  // Generar ID automático si la columna ID está vacía
  var idCol = headers[0];
  if (idCol && !rowObj[idCol]) {
    rowObj[idCol] = nextId_(sheet, sheetName, headers);
  }

  var newRow = headers.map(function(h) {
    var v = (rowObj[h] !== undefined && rowObj[h] !== null) ? rowObj[h] : '';
    // Forzar texto en columnas HORA para que Sheets no las convierta a timestamp
    if (h.indexOf('HORA') !== -1 && v !== '') return String(v);
    return v;
  });

  sheet.appendRow(newRow);
  return { ok: true, id: rowObj[idCol] || '' };
}

// ═══════════════════════════════════════════════════════════════════════════
//  ACTUALIZAR FILA (UPDATE)
// ═══════════════════════════════════════════════════════════════════════════
function updateRow(sheetName, id, rowObj) {
  var book  = ss_();
  var sheet = book.getSheetByName(sheetName);
  if (!sheet) return { ok: false, error: 'Hoja no encontrada: ' + sheetName };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { ok: false, error: 'Hoja vacía' };

  var data    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      var updatedRow = headers.map(function(h, j) {
        if (rowObj[h] !== undefined) {
          // Forzar texto en columnas HORA también al editar
          if (h.indexOf('HORA') !== -1 && rowObj[h] !== '') return String(rowObj[h]);
          return rowObj[h];
        }
        return data[i][j];
      });
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([updatedRow]);
      return { ok: true, id: id };
    }
  }
  return { ok: false, error: 'Fila no encontrada con ID: ' + id };
}

// ═══════════════════════════════════════════════════════════════════════════
//  ELIMINAR FILA (DELETE)
// ═══════════════════════════════════════════════════════════════════════════
function deleteRow(sheetName, id) {
  var book  = ss_();
  var sheet = book.getSheetByName(sheetName);
  if (!sheet) return { ok: false, error: 'Hoja no encontrada: ' + sheetName };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: false, error: 'Hoja vacía' };

  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Fila no encontrada con ID: ' + id };
}

// ═══════════════════════════════════════════════════════════════════════════
//  GENERACIÓN DE IDs AUTOMÁTICOS
// ═══════════════════════════════════════════════════════════════════════════
function nextId_(sheet, sheetName, headers) {
  var PREFIX_MAP = {
    'PREINSCRIPCIONES':       'PREINSC-',
    'INSCRIPCIONES':          'INSC-',
    'FACTURAS':               'FAC-',
    'AMPA_SOCIOS':            'SOC-',
    'AMPA_ALUMNOS':           'ALU-',
    'AUSENCIAS':              'AUS-',
    'REGISTRO_HORARIO':       'RH-',
    'ASISTENCIAS_FUTBOL':     'AF-',
    'ASISTENCIAS_EXTRA':      'AE-',
    'SESIONES':               'SES-',
    'NOMINAS':                'NOM-',
    'ENCUESTAS_SATISFACCION': 'ENC-',
  };

  var prefix  = PREFIX_MAP[sheetName];
  var lastRow = sheet.getLastRow();

  if (prefix) {
    var num = 1;
    if (lastRow > 1) {
      var vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
        .map(function(r) { return String(r[0]); })
        .filter(function(v) { return v.indexOf(prefix) === 0; })
        .map(function(v) { return parseInt(v.replace(prefix, ''), 10) || 0; });
      if (vals.length) num = Math.max.apply(null, vals) + 1;
    }
    return prefix + (num < 10 ? '00' + num : num < 100 ? '0' + num : String(num));
  } else {
    if (lastRow < 2) return 1;
    var nums = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .map(function(r) { return parseInt(r[0], 10) || 0; });
    return nums.length ? Math.max.apply(null, nums) + 1 : 1;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
//  NORMALIZACIÓN DE TEXTO
// ═══════════════════════════════════════════════════════════════════════════
function stripAccents_(str) {
  return String(str || '')
    .toUpperCase()
    .replace(/Á/g, 'A').replace(/É/g, 'E').replace(/Í/g, 'I')
    .replace(/Ó/g, 'O').replace(/Ú/g, 'U').replace(/Ü/g, 'U')
    .replace(/Ñ/g, 'N').replace(/Ç/g, 'C');
}

// ═══════════════════════════════════════════════════════════════════════════
//  PREINSCRIPCIÓN PÚBLICA (desde inscripcion.html)
// ═══════════════════════════════════════════════════════════════════════════
function nuevaPreinscripcion(data) {
  var book  = ss_();
  var sheet = book.getSheetByName('PREINSCRIPCIONES');
  if (!sheet) return { ok: false, error: 'Hoja PREINSCRIPCIONES no encontrada' };

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h).trim(); });

  var existing   = sheetToJSON(book, 'PREINSCRIPCIONES');
  var nombreNorm = stripAccents_(data.NOMBRE_ALUMNO);
  var emailNorm  = String(data.EMAIL_TUTOR || '').toLowerCase().trim();

  var dup = existing.some(function(r) {
    return stripAccents_(r.NOMBRE_ALUMNO)           === nombreNorm &&
           String(r.EMAIL_TUTOR || '').toLowerCase() === emailNorm &&
           String(r.ESTADO || '').toUpperCase()      !== 'RECHAZADA';
  });

  if (dup) return { ok: false, error: 'Ya existe una preinscripción para este alumno.' };

  var id = nextId_(sheet, 'PREINSCRIPCIONES', headers);

  var rowObj = {
    ID:               id,
    FECHA_SOLICITUD:  Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    NOMBRE_ALUMNO:    String(data.NOMBRE_ALUMNO   || '').toUpperCase().trim(),
    APELLIDOS_ALUMNO: String(data.APELLIDOS_ALUMNO|| '').toUpperCase().trim(),
    FECHA_NAC:        data.FECHA_NAC              || '',
    NOMBRE_TUTOR:     String(data.NOMBRE_TUTOR    || '').toUpperCase().trim(),
    APELLIDOS_TUTOR:  String(data.APELLIDOS_TUTOR || '').toUpperCase().trim(),
    EMAIL_TUTOR:      emailNorm,
    'TELÉFONO_TUTOR': data.TELEFONO_TUTOR || data['TELÉFONO_TUTOR'] || '',
    ACTIVIDAD_INTERES:data.ACTIVIDAD_INTERES      || '',
    LINEA:            data.LINEA                  || 'EXTRAESCOLARES',
    ESTADO:           'RECIBIDA',
    NOTAS:            data.NOTAS                  || '',
  };

  var newRow = headers.map(function(h) {
    return rowObj[h] !== undefined ? rowObj[h] : '';
  });
  sheet.appendRow(newRow);

  return { ok: true, id: id };
}

// ═══════════════════════════════════════════════════════════════════════════
//  ALERTAS DE IMPAGOS (ejecutar manualmente o con trigger semanal)
// ═══════════════════════════════════════════════════════════════════════════
function sendAlertsImpagos() {
  var book     = ss_();
  var pagos    = sheetToJSON(book, 'PAGOS');
  var alumnos  = sheetToJSON(book, 'ALUMNOS');
  var clientes = sheetToJSON(book, 'CLIENTES');
  var enviados = 0;

  pagos
    .filter(function(p) { return String(p.ESTADO || '').toUpperCase() === 'PENDIENTE'; })
    .forEach(function(p) {
      var alumno  = alumnos.find(function(a) { return String(a.ID_ALUMNO) === String(p.ID_ALUMNO); });
      var cliente = alumno ? clientes.find(function(c) { return String(c.ID_CLIENTE) === String(alumno.ID_CLIENTE); }) : null;
      if (cliente && cliente.EMAIL) {
        MailApp.sendEmail({
          to:      cliente.EMAIL,
          subject: 'SM Academia — Recordatorio de pago pendiente',
          body:    'Estimado/a ' + cliente.NOMBRE + ',\n\n' +
                   'Tiene un pago pendiente:\n' +
                   '  Concepto: ' + (p.CONCEPTO || '—') + '\n' +
                   '  Mes:      ' + (p.MES || '—') + '\n' +
                   '  Importe:  ' + (p.TOTAL_FACTURADO || '—') + ' €\n\n' +
                   'Contacte con nosotros para regularizarlo.\n\nSM Academia',
        });
        enviados++;
      }
    });

  return { ok: true, alertas: enviados };
}

// ═══════════════════════════════════════════════════════════════════════════
//  FUNCIÓN DE DIAGNÓSTICO — ejecutar desde el editor para ver errores
// ═══════════════════════════════════════════════════════════════════════════
function diagnostico() {
  try {
    var book   = ss_();
    var sheets = book.getSheets().map(function(s) { return s.getName(); });
    Logger.log('Hojas encontradas: ' + JSON.stringify(sheets));

    var esperadas = [
      'CLIENTES','ALUMNOS','ACTIVIDADES','PREINSCRIPCIONES','INSCRIPCIONES',
      'PAGOS','FACTURAS','PERSONAL','AUSENCIAS','REGISTRO_HORARIO',
      'SESIONES','ASISTENCIAS_FUTBOL','ASISTENCIAS_EXTRA','NOMINAS',
      'GASTOS','MOVIMIENTOS_BANCARIOS','AMPA_SOCIOS','AMPA_ALUMNOS',
      'ENCUESTAS_SATISFACCION',
    ];

    var faltantes = esperadas.filter(function(n) { return sheets.indexOf(n) === -1; });
    if (faltantes.length) {
      Logger.log('⚠ HOJAS FALTANTES: ' + JSON.stringify(faltantes));
    } else {
      Logger.log('✓ Todas las hojas encontradas');
    }

    var result = getAllData();
    Logger.log('GET_ALL ok=' + result.ok);
    Logger.log('Clientes: '  + result.clientes.length);
    Logger.log('Personal: '  + result.personal.length);
    Logger.log('Encuestas: ' + result.encuestasSatisfaccion.length);

  } catch(err) {
    Logger.log('ERROR: ' + err.toString());
  }
}
