/**
 * ============================================================================
 * CASCADAS HOTEL — ENCUESTA DE SATISFACCION
 * ============================================================================
 * Un solo link para todos los huespedes (?panel ausente) y un panel de staff
 * (?panel=1) para generar/copiar el link, editar las preguntas y ver los
 * informes con graficos.
 *
 * Las preguntas de la encuesta son configurables desde el panel (hoja
 * "Preguntas"): se pueden agregar, editar, eliminar y reordenar, eligiendo
 * el tipo de respuesta (puntuacion 1-10, escala de satisfaccion, alternativas,
 * si/no o texto libre). El informe se arma dinamicamente segun esos tipos.
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
  _crearHojaPreguntas(ss);
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
    'ID', 'Timestamp', 'Idioma', 'Nombre', 'Habitacion', 'RespuestasJSON'
  ], []);
}

/**
 * Preguntas de la encuesta. Tipo puede ser:
 *  - 'puntuacion'  : escala 1-10 (si EsNPS=TRUE, alimenta el NPS y el KPI principal)
 *  - 'escala'      : Muy bueno / Bueno / Regular / Malo / Muy malo / No aplica
 *  - 'alternativas': opciones a eleccion (definidas en la columna Opciones, separadas por "|")
 *  - 'sino'        : Si / No
 *  - 'texto'       : texto libre
 * Solo puede haber UNA pregunta con EsNPS=TRUE (se controla al guardar).
 */
function _crearHojaPreguntas(ss) {
  _crearHoja(ss, 'Preguntas', ['ID', 'Orden', 'Texto', 'Tipo', 'Opciones', 'EsNPS', 'Requerida', 'Activa'], [
    [generarID(), 1, '¿Qué tan probable es que recomiende Cascadas Hotel a un amigo o colega?', 'puntuacion', '', 'TRUE', 'TRUE', 'TRUE'],
    [generarID(), 2, 'Experiencia de reserva', 'escala', '', 'FALSE', 'TRUE', 'TRUE'],
    [generarID(), 3, 'Check-in / Check-out', 'escala', '', 'FALSE', 'TRUE', 'TRUE'],
    [generarID(), 4, 'Cordialidad del personal', 'escala', '', 'FALSE', 'TRUE', 'TRUE'],
    [generarID(), 5, 'Limpieza general', 'escala', '', 'FALSE', 'TRUE', 'TRUE'],
    [generarID(), 6, 'Comodidad de la habitación', 'escala', '', 'FALSE', 'TRUE', 'TRUE'],
    [generarID(), 7, 'Desayuno', 'escala', '', 'FALSE', 'TRUE', 'TRUE'],
    [generarID(), 8, 'Servicio de Almuerzo/Cena', 'escala', '', 'FALSE', 'TRUE', 'TRUE'],
    [generarID(), 9, 'Instalaciones e infraestructura', 'escala', '', 'FALSE', 'TRUE', 'TRUE'],
    [generarID(), 10, '¿Cómo conoció Cascadas Hotel?', 'alternativas', 'Internet|Redes Sociales|Agencia de viajes|Recomendación|Otro', 'FALSE', 'TRUE', 'TRUE'],
    [generarID(), 11, '¿Cuál fue su motivo de visita?', 'alternativas', 'Turismo|Trabajo|Otro', 'FALSE', 'TRUE', 'TRUE'],
    [generarID(), 12, '¿Desea dejarnos alguna sugerencia o comentario?', 'texto', '', 'FALSE', 'FALSE', 'TRUE'],
    [generarID(), 13, '¿Desearía indicarnos el nombre de alguna persona que se haya caracterizado por su cordialidad?', 'texto', '', 'FALSE', 'FALSE', 'TRUE']
  ]);
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
function _indiceColumna(hoja, nombre) {
  var enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  return enc.indexOf(nombre);
}
function generarID() { return Utilities.getUuid(); }
function numero_(v) {
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

/** Devuelve el valor de una clave de Configuracion (o '' si no existe). */
function _obtenerConfigValor(clave) {
  var filas = _leerHojaComoObjetos('Configuracion');
  for (var i = 0; i < filas.length; i++) {
    if (filas[i].Clave === clave) return filas[i].Valor;
  }
  return '';
}

/**
 * Migracion auto-reparable (mismo patron que Concierge): si la base de datos
 * fue creada antes de que existieran las preguntas configurables, crea la
 * hoja "Preguntas" con las preguntas por defecto y renueva "Respuestas" al
 * nuevo esquema (columna RespuestasJSON). No hay perdida de datos reales
 * porque esto solo corre quando el esquema aun es el antiguo.
 */
function _asegurarEsquema() {
  var ss = _ss();
  if (!ss.getSheetByName('Preguntas')) _crearHojaPreguntas(ss);
  var hojaResp = ss.getSheetByName('Respuestas');
  if (!hojaResp) { _crearHojaRespuestas(ss); return; }
  var encabezados = hojaResp.getLastColumn() > 0
    ? hojaResp.getRange(1, 1, 1, hojaResp.getLastColumn()).getValues()[0] : [];
  if (encabezados.indexOf('RespuestasJSON') === -1) _crearHojaRespuestas(ss);
}

// ---------------------------------------------------------------------------
// DATOS PUBLICOS (encuesta - sin login)
// ---------------------------------------------------------------------------

/** Datos para pintar la encuesta: nombre del hotel + preguntas activas. */
function obtenerDatosPublicos() {
  _asegurarEsquema();
  return {
    hotelNombre: _obtenerConfigValor('HOTEL_NOMBRE') || 'Cascadas Hotel',
    preguntas: obtenerPreguntasActivas()
  };
}

/** Preguntas activas, ordenadas, con solo los datos que necesita el formulario publico. */
function obtenerPreguntasActivas() {
  _asegurarEsquema();
  var filas = _leerHojaComoObjetos('Preguntas').filter(function (p) { return _aBooleano(p.Activa); });
  filas.sort(function (a, b) { return numero_(a.Orden) - numero_(b.Orden); });
  return filas.map(function (p) {
    return {
      id: p.ID, texto: p.Texto, tipo: p.Tipo,
      opciones: p.Opciones ? String(p.Opciones).split('|') : [],
      esNPS: _aBooleano(p.EsNPS), requerida: _aBooleano(p.Requerida)
    };
  });
}

var NIVELES_ESCALA = ['Muy bueno', 'Bueno', 'Regular', 'Malo', 'Muy malo', 'No aplica'];

/**
 * Guarda una respuesta de encuesta. Valida server-side segun el tipo de cada
 * pregunta activa (independiente de lo que venga del cliente).
 * @param {Object} datos {nombre, habitacion, idioma, respuestas: {preguntaID: valor}}
 * @return {Object} {success, mensaje}
 */
function enviarRespuestaEncuesta(datos) {
  try {
    _asegurarEsquema();
    datos = datos || {};
    var nombre = String(datos.nombre || '').trim();
    if (!nombre) return { success: false, mensaje: 'Falta tu nombre.' };

    var preguntas = _leerHojaComoObjetos('Preguntas').filter(function (p) { return _aBooleano(p.Activa); });
    var entrada = datos.respuestas || {};
    var salida = {};

    for (var i = 0; i < preguntas.length; i++) {
      var p = preguntas[i];
      var val = entrada[p.ID];
      var requerida = _aBooleano(p.Requerida);

      if (p.Tipo === 'puntuacion') {
        var n = Math.round(Number(val));
        if (isNaN(n) || n < 1 || n > 10) {
          if (requerida) return { success: false, mensaje: 'Falta responder: "' + p.Texto + '".' };
          continue;
        }
        salida[p.ID] = n;
      } else if (p.Tipo === 'escala') {
        if (NIVELES_ESCALA.indexOf(val) === -1) {
          if (requerida) return { success: false, mensaje: 'Falta responder: "' + p.Texto + '".' };
          continue;
        }
        salida[p.ID] = val;
      } else if (p.Tipo === 'alternativas') {
        var opciones = p.Opciones ? String(p.Opciones).split('|') : [];
        if (opciones.indexOf(val) === -1) {
          if (requerida) return { success: false, mensaje: 'Falta responder: "' + p.Texto + '".' };
          continue;
        }
        salida[p.ID] = val;
      } else if (p.Tipo === 'sino') {
        if (val !== 'Si' && val !== 'No') {
          if (requerida) return { success: false, mensaje: 'Falta responder: "' + p.Texto + '".' };
          continue;
        }
        salida[p.ID] = val;
      } else if (p.Tipo === 'texto') {
        var texto = String(val || '').trim();
        if (!texto) {
          if (requerida) return { success: false, mensaje: 'Falta responder: "' + p.Texto + '".' };
          continue;
        }
        salida[p.ID] = texto;
      }
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var hoja = _hoja('Respuestas');
      hoja.appendRow([
        generarID(), new Date(), String(datos.idioma || 'es'), nombre,
        String(datos.habitacion || '').trim(), JSON.stringify(salida)
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
/** URL de la hoja de calculo (Google Sheets) donde vive toda la base de datos. */
function obtenerUrlHojaCalculo(email) {
  if (!_emailAutorizado(email)) return { success: false, mensaje: 'No tienes acceso.' };
  _asegurarEsquema();
  return { success: true, url: _ss().getUrl() };
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
// STAFF: PREGUNTAS (editor)
// ---------------------------------------------------------------------------
var TIPOS_PREGUNTA = ['puntuacion', 'escala', 'alternativas', 'sino', 'texto'];

/** Todas las preguntas (incl. inactivas) para el editor del panel. */
function obtenerPreguntas(email) {
  if (!_emailAutorizado(email)) return { success: false, mensaje: 'No tienes acceso.' };
  _asegurarEsquema();
  var filas = _leerHojaComoObjetos('Preguntas');
  filas.sort(function (a, b) { return numero_(a.Orden) - numero_(b.Orden); });
  return {
    success: true,
    preguntas: filas.map(function (p) {
      return {
        id: p.ID, orden: numero_(p.Orden), texto: p.Texto, tipo: p.Tipo,
        opciones: p.Opciones ? String(p.Opciones).split('|') : [],
        esNPS: _aBooleano(p.EsNPS), requerida: _aBooleano(p.Requerida), activa: _aBooleano(p.Activa)
      };
    })
  };
}

/**
 * Crea o actualiza una pregunta (si datos.id viene, actualiza; si no, crea).
 * Solo puede existir UNA pregunta con esNPS=true: si se marca una nueva,
 * las demas se desmarcan automaticamente.
 */
function guardarPregunta(datos, email) {
  if (!_emailAutorizado(email)) return { success: false, mensaje: 'No tienes acceso.' };
  _asegurarEsquema();
  datos = datos || {};
  var texto = String(datos.texto || '').trim();
  if (!texto) return { success: false, mensaje: 'Falta el texto de la pregunta.' };
  if (TIPOS_PREGUNTA.indexOf(datos.tipo) === -1) return { success: false, mensaje: 'Tipo de pregunta invalido.' };

  var opciones = '';
  if (datos.tipo === 'alternativas') {
    var lista = (datos.opciones || []).map(function (o) { return String(o).trim(); }).filter(function (o) { return o; });
    if (lista.length < 2) return { success: false, mensaje: 'Agrega al menos 2 opciones.' };
    opciones = lista.join('|');
  }
  var esNPS = datos.tipo === 'puntuacion' && !!datos.esNPS;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var hoja = _hoja('Preguntas');
    var filas = _leerHojaComoObjetos('Preguntas');
    var colEsNPS = _indiceColumna(hoja, 'EsNPS') + 1;

    // Solo puede haber una pregunta NPS: si esta se marca, se desmarcan las demas.
    if (esNPS) {
      filas.forEach(function (p) {
        if (_aBooleano(p.EsNPS) && p.ID !== datos.id) hoja.getRange(p._fila, colEsNPS).setValue('FALSE');
      });
    }

    if (datos.id) {
      var existente = filas.filter(function (p) { return p.ID === datos.id; })[0];
      if (!existente) return { success: false, mensaje: 'Pregunta no encontrada.' };
      var fila = existente._fila;
      hoja.getRange(fila, _indiceColumna(hoja, 'Texto') + 1).setValue(texto);
      hoja.getRange(fila, _indiceColumna(hoja, 'Tipo') + 1).setValue(datos.tipo);
      hoja.getRange(fila, _indiceColumna(hoja, 'Opciones') + 1).setValue(opciones);
      hoja.getRange(fila, colEsNPS).setValue(esNPS ? 'TRUE' : 'FALSE');
      hoja.getRange(fila, _indiceColumna(hoja, 'Requerida') + 1).setValue(datos.requerida === false ? 'FALSE' : 'TRUE');
      if (datos.activa !== undefined) {
        hoja.getRange(fila, _indiceColumna(hoja, 'Activa') + 1).setValue(datos.activa === false ? 'FALSE' : 'TRUE');
      }
      return { success: true, mensaje: 'Pregunta actualizada.' };
    }

    var maxOrden = filas.reduce(function (m, p) { return Math.max(m, numero_(p.Orden)); }, 0);
    hoja.appendRow([generarID(), maxOrden + 1, texto, datos.tipo, opciones,
      esNPS ? 'TRUE' : 'FALSE', datos.requerida === false ? 'FALSE' : 'TRUE', 'TRUE']);
    return { success: true, mensaje: 'Pregunta agregada.' };
  } finally {
    lock.releaseLock();
  }
}

function eliminarPregunta(id, email) {
  if (!_emailAutorizado(email)) return { success: false, mensaje: 'No tienes acceso.' };
  var hoja = _hoja('Preguntas');
  var filas = _leerHojaComoObjetos('Preguntas');
  var fila = filas.filter(function (p) { return p.ID === id; })[0];
  if (!fila) return { success: false, mensaje: 'Pregunta no encontrada.' };
  hoja.deleteRow(fila._fila);
  return { success: true, mensaje: 'Pregunta eliminada.' };
}

/** Reordena segun el arreglo de IDs recibido (el orden del arreglo = el nuevo Orden). */
function reordenarPreguntas(idsEnOrden, email) {
  if (!_emailAutorizado(email)) return { success: false, mensaje: 'No tienes acceso.' };
  var hoja = _hoja('Preguntas');
  var filas = _leerHojaComoObjetos('Preguntas');
  var colOrden = _indiceColumna(hoja, 'Orden') + 1;
  (idsEnOrden || []).forEach(function (id, i) {
    var fila = filas.filter(function (p) { return p.ID === id; })[0];
    if (fila) hoja.getRange(fila._fila, colOrden).setValue(i + 1);
  });
  return { success: true };
}

// ---------------------------------------------------------------------------
// STAFF: INFORMES
// ---------------------------------------------------------------------------

/**
 * Devuelve KPIs, distribuciones y la tabla de respuestas para el rango de
 * fechas indicado (o todo el historico si no se pasan fechas). El informe se
 * arma dinamicamente segun las preguntas configuradas y sus tipos.
 * @param {string} email
 * @param {Object} filtros {fechaDesde, fechaHasta} formato YYYY-MM-DD
 */
function obtenerDatosReporte(email, filtros) {
  if (!_emailAutorizado(email)) return { success: false, mensaje: 'No tienes acceso.' };
  _asegurarEsquema();
  filtros = filtros || {};
  try {
    var tz = Session.getScriptTimeZone();
    var preguntas = _leerHojaComoObjetos('Preguntas');
    preguntas.sort(function (a, b) { return numero_(a.Orden) - numero_(b.Orden); });
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
      try { r._respuestas = JSON.parse(r.RespuestasJSON || '{}'); } catch (e) { r._respuestas = {}; }
      return true;
    });

    var total = filtradas.length;

    // La pregunta "cabecera" del informe: la marcada EsNPS, o si no hay
    // ninguna, la primera de tipo puntuacion (asi el informe sigue teniendo
    // un KPI principal aunque el hotel no haya marcado ninguna como NPS).
    var preguntaNPS = preguntas.filter(function (p) { return _aBooleano(p.EsNPS); })[0] || null;
    var preguntaGeneral = preguntaNPS || preguntas.filter(function (p) { return p.Tipo === 'puntuacion'; })[0] || null;

    var promedioGeneral = 0, nps = null, promotores = 0, neutros = 0, detractores = 0;
    var pctPromotores = 0, pctNeutros = 0, pctDetractores = 0;
    var distribucionGeneral = [];
    var porDia = [];

    if (preguntaGeneral) {
      var suma = 0, cant = 0;
      var distribucion = {}; for (var d = 1; d <= 10; d++) distribucion[d] = 0;
      var porDiaMap = {};
      filtradas.forEach(function (r) {
        if (!r._respuestas.hasOwnProperty(preguntaGeneral.ID)) return;
        var v = numero_(r._respuestas[preguntaGeneral.ID]);
        suma += v; cant++;
        if (distribucion[v] !== undefined) distribucion[v]++;
        if (v >= 9) promotores++; else if (v >= 7) neutros++; else detractores++;
        if (!porDiaMap[r._fechaStr]) porDiaMap[r._fechaStr] = { suma: 0, cantidad: 0 };
        porDiaMap[r._fechaStr].suma += v;
        porDiaMap[r._fechaStr].cantidad++;
      });
      promedioGeneral = cant > 0 ? Math.round((suma / cant) * 10) / 10 : 0;
      var pct = function (n) { return cant > 0 ? Math.round((n / cant) * 1000) / 10 : 0; };
      pctPromotores = pct(promotores); pctNeutros = pct(neutros); pctDetractores = pct(detractores);
      if (preguntaNPS) nps = Math.round(pctPromotores - pctDetractores);
      for (var k = 1; k <= 10; k++) distribucionGeneral.push([k, distribucion[k]]);
      porDia = Object.keys(porDiaMap).sort().map(function (f) {
        var o = porDiaMap[f];
        return [f, Math.round((o.suma / o.cantidad) * 10) / 10, o.cantidad];
      });
    }

    // Preguntas de puntuacion secundarias (todas menos la "cabecera").
    var puntuaciones = preguntas
      .filter(function (p) { return p.Tipo === 'puntuacion' && (!preguntaGeneral || p.ID !== preguntaGeneral.ID); })
      .map(function (p) {
        var s = 0, c = 0;
        filtradas.forEach(function (r) {
          if (!r._respuestas.hasOwnProperty(p.ID)) return;
          s += numero_(r._respuestas[p.ID]); c++;
        });
        return { id: p.ID, texto: p.Texto, promedio: c > 0 ? Math.round((s / c) * 10) / 10 : 0, respuestas: c };
      });

    // Preguntas de escala -> matriz de satisfaccion.
    var escalas = preguntas.filter(function (p) { return p.Tipo === 'escala'; }).map(function (p) {
      var conteo = {}; NIVELES_ESCALA.forEach(function (n) { conteo[n] = 0; });
      var c = 0;
      filtradas.forEach(function (r) {
        var v = r._respuestas[p.ID];
        if (conteo.hasOwnProperty(v)) { conteo[v]++; c++; }
      });
      return {
        id: p.ID, texto: p.Texto, respuestas: c,
        distribucion: NIVELES_ESCALA.map(function (n) {
          return { nivel: n, cantidad: conteo[n], pct: c > 0 ? Math.round((conteo[n] / c) * 1000) / 10 : 0 };
        })
      };
    });

    // Preguntas de alternativas -> donas.
    var alternativas = preguntas.filter(function (p) { return p.Tipo === 'alternativas'; }).map(function (p) {
      var opciones = p.Opciones ? String(p.Opciones).split('|') : [];
      var conteo = {}; opciones.forEach(function (o) { conteo[o] = 0; });
      var c = 0;
      filtradas.forEach(function (r) {
        var v = r._respuestas[p.ID];
        if (conteo.hasOwnProperty(v)) { conteo[v]++; c++; }
      });
      return {
        id: p.ID, texto: p.Texto, respuestas: c,
        distribucion: opciones.map(function (o) {
          return { opcion: o, cantidad: conteo[o], pct: c > 0 ? Math.round((conteo[o] / c) * 1000) / 10 : 0 };
        })
      };
    });

    // Preguntas si/no.
    var siNo = preguntas.filter(function (p) { return p.Tipo === 'sino'; }).map(function (p) {
      var si = 0, no = 0;
      filtradas.forEach(function (r) {
        var v = r._respuestas[p.ID];
        if (v === 'Si') si++; else if (v === 'No') no++;
      });
      var c = si + no;
      return {
        id: p.ID, texto: p.Texto, respuestas: c,
        pctSi: c > 0 ? Math.round((si / c) * 1000) / 10 : 0,
        pctNo: c > 0 ? Math.round((no / c) * 1000) / 10 : 0
      };
    });

    // Preguntas de texto libre -> lista de comentarios (mas recientes primero).
    var ordenadasReciente = filtradas.slice().sort(function (a, b) { return b._fechaObj.getTime() - a._fechaObj.getTime(); });
    var textos = preguntas.filter(function (p) { return p.Tipo === 'texto'; }).map(function (p) {
      var items = [];
      ordenadasReciente.forEach(function (r) {
        var v = r._respuestas[p.ID];
        if (!v) return;
        items.push({
          fecha: Utilities.formatDate(r._fechaObj, tz, 'dd/MM/yyyy'),
          habitacion: r.Habitacion || '', nombre: r.Nombre || '', texto: v,
          puntuacionGeneral: preguntaGeneral && r._respuestas.hasOwnProperty(preguntaGeneral.ID)
            ? numero_(r._respuestas[preguntaGeneral.ID]) : null
        });
      });
      return { id: p.ID, texto: p.Texto, respuestas: items.length, items: items.slice(0, 100) };
    });

    // Tabla de respuestas individuales (los comentarios de texto se combinan en una columna).
    var tabla = ordenadasReciente.slice(0, 300).map(function (r) {
      var comentarios = textos
        .map(function (t) { return r._respuestas[t.id]; })
        .filter(function (v) { return v; })
        .join(' · ');
      return {
        fecha: Utilities.formatDate(r._fechaObj, tz, 'dd/MM/yyyy HH:mm'),
        nombre: r.Nombre || '',
        habitacion: r.Habitacion || '',
        idioma: r.Idioma || '',
        puntuacionGeneral: preguntaGeneral && r._respuestas.hasOwnProperty(preguntaGeneral.ID)
          ? numero_(r._respuestas[preguntaGeneral.ID]) : null,
        comentario: comentarios
      };
    });

    return {
      success: true,
      totalRespuestas: total,
      preguntaGeneralTexto: preguntaGeneral ? preguntaGeneral.Texto : '',
      esNPS: !!preguntaNPS,
      promedioGeneral: promedioGeneral,
      nps: nps,
      promotores: promotores, neutros: neutros, detractores: detractores,
      pctPromotores: pctPromotores, pctNeutros: pctNeutros, pctDetractores: pctDetractores,
      distribucionGeneral: distribucionGeneral,
      porDia: porDia,
      puntuaciones: puntuaciones,
      escalas: escalas,
      alternativas: alternativas,
      siNo: siNo,
      textos: textos,
      tabla: tabla
    };
  } catch (err) {
    return { success: false, mensaje: 'Error: ' + err.message };
  }
}
