/**
 * ============================================================================
 * CASCADAS HOTEL — ENCUESTA DE SATISFACCION
 * ============================================================================
 * Un solo link para todos los huespedes (?panel ausente) y un panel de staff
 * (?panel=1) para generar/copiar el link y ver los informes con graficos.
 *
 * Primer uso: ejecuta crearBaseDeDatos() UNA vez desde el editor de Apps
 * Script. Crea el Spreadsheet con las hojas necesarias y guarda su ID en
 * PropertiesService, igual que en Concierge. No hay que pegar ningun ID a mano.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// PUNTO DE ENTRADA WEB
// ---------------------------------------------------------------------------
function doGet(e) {
  var esPanel = e && e.parameter && (e.parameter.panel === '1' || e.parameter.panel === 'true');
  var plantilla = HtmlService.createTemplateFromFile(esPanel ? 'Index' : 'Encuesta');
  return plantilla.evaluate()
    .setTitle(esPanel ? 'Panel de Encuestas - Cascadas Hotel' : 'Encuesta de Satisfaccion - Cascadas Hotel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------------
// SETUP (ejecutar UNA vez desde el editor)
// ---------------------------------------------------------------------------
function crearBaseDeDatos() {
  var idExistente = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (idExistente) {
    Logger.log('Ya existe una base de datos con ID: ' + idExistente);
    Logger.log('Si deseas recrearla, borra la propiedad SPREADSHEET_ID primero.');
    return idExistente;
  }

  var ss = SpreadsheetApp.create('Cascadas Hotel - Encuestas');
  var id = ss.getId();

  _crearHojaRespuestas(ss);
  _crearHojaUsuarios(ss);
  _crearHojaConfiguracion(ss);
  _eliminarHojaPorDefecto(ss);

  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);

  Logger.log('Base de datos creada con exito.');
  Logger.log('SPREADSHEET_ID: ' + id);
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('Ahora agrega tu email en la hoja "Usuarios" (columna Activo = TRUE) para poder entrar al panel.');
  return id;
}

function _crearHoja(ss, nombre, encabezados, filas) {
  var hoja = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  hoja.clear();
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  var rangoEnc = hoja.getRange(1, 1, 1, encabezados.length);
  rangoEnc.setFontWeight('bold');
  rangoEnc.setBackground('#414143');
  rangoEnc.setFontColor('#FFFFFF');
  hoja.setFrozenRows(1);
  if (filas && filas.length > 0) {
    hoja.getRange(2, 1, filas.length, encabezados.length).setValues(filas);
  }
  hoja.autoResizeColumns(1, encabezados.length);
  return hoja;
}

function _crearHojaRespuestas(ss) {
  _crearHoja(ss, 'Respuestas', [
    'ID', 'Timestamp', 'Nombre', 'Habitacion',
    'PuntuacionGeneral', 'Limpieza', 'Atencion', 'Instalaciones', 'Comentario'
  ], []);
}

function _crearHojaUsuarios(ss) {
  _crearHoja(ss, 'Usuarios', ['Email', 'Nombre', 'Activo'], [
    ['admin@cascadashotel.cl', 'Administracion', 'TRUE']
  ]);
}

function _crearHojaConfiguracion(ss) {
  _crearHoja(ss, 'Configuracion', ['Clave', 'Valor', 'Descripcion'], [
    ['HOTEL_NOMBRE', 'Cascadas Hotel', 'Nombre mostrado en la encuesta y el panel'],
    ['MENSAJE_ENCUESTA',
      'Gracias por hospedarte en Cascadas Hotel. Nos encantaria conocer tu opinion, ¿nos regalas 1 minuto para responder esta breve encuesta?',
      'Mensaje opcional que se copia junto al link para enviar por WhatsApp/mail']
  ]);
}

function _eliminarHojaPorDefecto(ss) {
  var hojaDefault = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hojaDefault && ss.getSheets().length > 1) ss.deleteSheet(hojaDefault);
}

// ---------------------------------------------------------------------------
// ACCESO AL SPREADSHEET
// ---------------------------------------------------------------------------
var _ssCache = null;
function _ss() {
  if (_ssCache) return _ssCache;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('No se encontro SPREADSHEET_ID. Ejecuta crearBaseDeDatos() primero.');
  _ssCache = SpreadsheetApp.openById(id);
  return _ssCache;
}
function _hoja(nombre) {
  var h = _ss().getSheetByName(nombre);
  if (!h) throw new Error('Hoja no encontrada: ' + nombre);
  return h;
}
function _leerHojaComoObjetos(nombreHoja) {
  var hoja = _hoja(nombreHoja);
  var datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return [];
  var encabezados = datos[0];
  var out = [];
  for (var i = 1; i < datos.length; i++) {
    var obj = {};
    for (var c = 0; c < encabezados.length; c++) obj[encabezados[c]] = datos[i][c];
    obj._fila = i + 1;
    out.push(obj);
  }
  return out;
}
function _aBooleano(v) {
  if (v === true) return true;
  if (typeof v === 'string') return v.toUpperCase() === 'TRUE';
  if (typeof v === 'number') return v === 1;
  return false;
}
function generarID() { return Utilities.getUuid(); }

/** Devuelve el valor de una clave de Configuracion (o '' si no existe). */
function _obtenerConfigValor(clave) {
  var filas = _leerHojaComoObjetos('Configuracion');
  for (var i = 0; i < filas.length; i++) {
    if (filas[i].Clave === clave) return filas[i].Valor;
  }
  return '';
}

// ---------------------------------------------------------------------------
// DATOS PUBLICOS (encuesta - sin login)
// ---------------------------------------------------------------------------

/** Datos basicos para pintar la portada de la encuesta (nombre del hotel). */
function obtenerDatosPublicos() {
  return {
    hotelNombre: _obtenerConfigValor('HOTEL_NOMBRE') || 'Cascadas Hotel'
  };
}

/**
 * Guarda una respuesta de encuesta. Valida server-side que las 4
 * puntuaciones vengan entre 1 y 10 y que el nombre no venga vacio.
 * @param {Object} datos {nombre, habitacion, general, limpieza, atencion, instalaciones, comentario}
 * @return {Object} {success, mensaje}
 */
function enviarRespuestaEncuesta(datos) {
  try {
    datos = datos || {};
    var nombre = String(datos.nombre || '').trim();
    if (!nombre) return { success: false, mensaje: 'Falta tu nombre.' };

    function puntuacion(v) {
      var n = Math.round(Number(v));
      return (n >= 1 && n <= 10) ? n : null;
    }
    var general = puntuacion(datos.general);
    var limpieza = puntuacion(datos.limpieza);
    var atencion = puntuacion(datos.atencion);
    var instalaciones = puntuacion(datos.instalaciones);
    if (general === null || limpieza === null || atencion === null || instalaciones === null) {
      return { success: false, mensaje: 'Faltan puntuaciones por seleccionar.' };
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var hoja = _hoja('Respuestas');
      hoja.appendRow([
        generarID(), new Date(), nombre, String(datos.habitacion || '').trim(),
        general, limpieza, atencion, instalaciones,
        String(datos.comentario || '').trim()
      ]);
    } finally {
      lock.releaseLock();
    }
    return { success: true, mensaje: '¡Gracias por tu tiempo!' };
  } catch (err) {
    return { success: false, mensaje: 'Error: ' + err.message };
  }
}

// ---------------------------------------------------------------------------
// STAFF: LOGIN
// ---------------------------------------------------------------------------
function autenticarStaff(email) {
  if (!email) return { success: false, mensaje: 'Ingresa tu email.' };
  var usuarios = _leerHojaComoObjetos('Usuarios');
  for (var i = 0; i < usuarios.length; i++) {
    if (String(usuarios[i].Email).toLowerCase() === String(email).toLowerCase() && _aBooleano(usuarios[i].Activo)) {
      return { success: true, usuario: { Email: usuarios[i].Email, Nombre: usuarios[i].Nombre }, mensaje: 'Bienvenido ' + (usuarios[i].Nombre || '') };
    }
  }
  return { success: false, mensaje: 'Usuario no encontrado o inactivo.' };
}
function _emailAutorizado(email) {
  if (!email) return false;
  var usuarios = _leerHojaComoObjetos('Usuarios');
  for (var i = 0; i < usuarios.length; i++) {
    if (String(usuarios[i].Email).toLowerCase() === String(email).toLowerCase() && _aBooleano(usuarios[i].Activo)) return true;
  }
  return false;
}

// ---------------------------------------------------------------------------
// STAFF: ENLACE DE LA ENCUESTA
// ---------------------------------------------------------------------------
/**
 * Devuelve el link base (sin ?panel) para enviar a los huespedes, mas el
 * mensaje opcional configurado. Un solo link sirve para todos los huespedes.
 */
function obtenerDatosEnlace(email) {
  if (!_emailAutorizado(email)) return { success: false, mensaje: 'No tienes acceso.' };
  var base = '';
  try { base = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  return {
    success: true,
    link: base,
    mensaje: _obtenerConfigValor('MENSAJE_ENCUESTA') || ''
  };
}
function guardarMensajeEncuesta(email, mensaje) {
  if (!_emailAutorizado(email)) return { success: false, mensaje: 'No tienes acceso.' };
  var hoja = _hoja('Configuracion');
  var filas = _leerHojaComoObjetos('Configuracion');
  for (var i = 0; i < filas.length; i++) {
    if (filas[i].Clave === 'MENSAJE_ENCUESTA') {
      hoja.getRange(filas[i]._fila, 2).setValue(mensaje || '');
      return { success: true };
    }
  }
  hoja.appendRow(['MENSAJE_ENCUESTA', mensaje || '', '']);
  return { success: true };
}

// ---------------------------------------------------------------------------
// STAFF: INFORMES
// ---------------------------------------------------------------------------
function numero_(v) {
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

/**
 * Devuelve KPIs, distribuciones y la tabla de respuestas para el rango de
 * fechas indicado (o todo el historico si no se pasan fechas).
 * @param {string} email
 * @param {Object} filtros {fechaDesde, fechaHasta} formato YYYY-MM-DD
 */
function obtenerDatosReporte(email, filtros) {
  if (!_emailAutorizado(email)) return { success: false, mensaje: 'No tienes acceso.' };
  filtros = filtros || {};
  try {
    var tz = Session.getScriptTimeZone();
    var filas = _leerHojaComoObjetos('Respuestas');

    var desde = filtros.fechaDesde ? String(filtros.fechaDesde) : null;
    var hasta = filtros.fechaHasta ? String(filtros.fechaHasta) : null;

    var filtradas = filas.filter(function (r) {
      var fechaObj = r.Timestamp instanceof Date ? r.Timestamp : new Date(r.Timestamp);
      if (isNaN(fechaObj.getTime())) return false;
      var fechaStr = Utilities.formatDate(fechaObj, tz, 'yyyy-MM-dd');
      if (desde && fechaStr < desde) return false;
      if (hasta && fechaStr > hasta) return false;
      r._fechaObj = fechaObj;
      r._fechaStr = fechaStr;
      return true;
    });

    var total = filtradas.length;
    var sumaGeneral = 0, sumaLimpieza = 0, sumaAtencion = 0, sumaInstalaciones = 0;
    var promotores = 0, neutros = 0, detractores = 0;
    var distribucion = {}; for (var d = 1; d <= 10; d++) distribucion[d] = 0;
    var porDiaMap = {};

    filtradas.forEach(function (r) {
      var g = numero_(r.PuntuacionGeneral), l = numero_(r.Limpieza), a = numero_(r.Atencion), ins = numero_(r.Instalaciones);
      sumaGeneral += g; sumaLimpieza += l; sumaAtencion += a; sumaInstalaciones += ins;
      if (distribucion[g] !== undefined) distribucion[g]++;
      if (g >= 9) promotores++; else if (g >= 7) neutros++; else detractores++;

      if (!porDiaMap[r._fechaStr]) porDiaMap[r._fechaStr] = { suma: 0, cantidad: 0 };
      porDiaMap[r._fechaStr].suma += g;
      porDiaMap[r._fechaStr].cantidad++;
    });

    var promedio = function (s) { return total > 0 ? Math.round((s / total) * 10) / 10 : 0; };
    var pct = function (n) { return total > 0 ? Math.round((n / total) * 1000) / 10 : 0; };
    var nps = total > 0 ? Math.round(pct(promotores) - pct(detractores)) : 0;

    var porDia = Object.keys(porDiaMap).sort().map(function (f) {
      var o = porDiaMap[f];
      return [f, Math.round((o.suma / o.cantidad) * 10) / 10, o.cantidad];
    });

    var distribucionArr = [];
    for (var k = 1; k <= 10; k++) distribucionArr.push([k, distribucion[k]]);

    var tabla = filtradas
      .sort(function (a, b) { return b._fechaObj.getTime() - a._fechaObj.getTime(); })
      .slice(0, 300) // limite razonable para no sobrecargar el navegador
      .map(function (r) {
        return {
          fecha: Utilities.formatDate(r._fechaObj, tz, 'dd/MM/yyyy HH:mm'),
          nombre: r.Nombre || '',
          habitacion: r.Habitacion || '',
          general: numero_(r.PuntuacionGeneral),
          limpieza: numero_(r.Limpieza),
          atencion: numero_(r.Atencion),
          instalaciones: numero_(r.Instalaciones),
          comentario: r.Comentario || ''
        };
      });

    return {
      success: true,
      totalRespuestas: total,
      promedioGeneral: promedio(sumaGeneral),
      promedioLimpieza: promedio(sumaLimpieza),
      promedioAtencion: promedio(sumaAtencion),
      promedioInstalaciones: promedio(sumaInstalaciones),
      promotores: promotores, neutros: neutros, detractores: detractores,
      pctPromotores: pct(promotores), pctNeutros: pct(neutros), pctDetractores: pct(detractores),
      nps: nps,
      distribucionGeneral: distribucionArr,
      porDia: porDia,
      tabla: tabla
    };
  } catch (err) {
    return { success: false, mensaje: 'Error: ' + err.message };
  }
}
