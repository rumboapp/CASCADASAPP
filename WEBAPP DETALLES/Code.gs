/**
 * ============================================================================
 * CASCADAS HOTEL — DETALLES CASCADAS
 * ============================================================================
 * Generador de tarjetas de atencion (aniversario, cumpleanos, bienvenida,
 * estadia de regalo, gift card de noches, welcome drink, etc.) y cupones
 * canjeables. Cada tarjeta/cupon se arma y se descarga como PDF en el
 * navegador, sin pasar por planilla.
 *
 * La UNICA excepcion es la Gift Card: cada vez que se genera una, sus datos
 * (codigo, noches, vigencia, para/de, fecha) se guardan en una planilla para
 * poder llevar registro de las que se han emitido. La planilla se crea sola
 * la primera vez que se guarda una Gift Card (no hace falta configurar nada
 * a mano): el ID queda guardado en las Propiedades del Script.
 * ============================================================================
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Detalles Cascadas — Cascadas Hotel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------------
// PLANILLA DE GIFT CARDS EMITIDAS (se crea sola la primera vez)
// ---------------------------------------------------------------------------
var GIFTCARDS_HOJA = 'GiftCards';
var GIFTCARDS_ENCABEZADOS = [
  'Codigo', 'FechaEmision', 'Noches', 'VigenciaMeses', 'MesesValidos', 'Inclusiones',
  'Para', 'De', 'Idioma', 'Mensaje', 'FechaDesde', 'FechaHasta', 'EstadoReserva'
];

function _ssDetalles() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('DETALLES_SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* si la borraron, se crea otra abajo */ }
  }
  var ss = SpreadsheetApp.create('Cascadas Hotel - Detalles Cascadas (Gift Cards)');
  props.setProperty('DETALLES_SPREADSHEET_ID', ss.getId());
  return ss;
}

function _hojaGiftCards() {
  var ss = _ssDetalles();
  var hoja = ss.getSheetByName(GIFTCARDS_HOJA);
  if (!hoja) {
    hoja = ss.insertSheet(GIFTCARDS_HOJA);
    hoja.getRange(1, 1, 1, GIFTCARDS_ENCABEZADOS.length).setValues([GIFTCARDS_ENCABEZADOS]);
    var rangoEnc = hoja.getRange(1, 1, 1, GIFTCARDS_ENCABEZADOS.length);
    rangoEnc.setFontWeight('bold'); rangoEnc.setBackground('#414143'); rangoEnc.setFontColor('#FFFFFF');
    hoja.setFrozenRows(1);
    hoja.autoResizeColumns(1, GIFTCARDS_ENCABEZADOS.length);
    var hojaDefault = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
    if (hojaDefault && ss.getSheets().length > 1) ss.deleteSheet(hojaDefault);
  }
  _asegurarColumnasGiftCards(hoja);
  return hoja;
}

/* La planilla puede venir de una version anterior con menos columnas: las que
   falten se agregan al final sin tocar lo ya registrado. */
function _asegurarColumnasGiftCards(hoja) {
  var ancho = Math.max(hoja.getLastColumn(), 1);
  var actuales = hoja.getRange(1, 1, 1, ancho).getValues()[0].map(function (c) { return String(c).trim(); });
  var faltan = GIFTCARDS_ENCABEZADOS.filter(function (h) { return actuales.indexOf(h) === -1; });
  if (!faltan.length) return actuales;
  hoja.getRange(1, actuales.length + 1, 1, faltan.length).setValues([faltan])
      .setFontWeight('bold').setBackground('#414143').setFontColor('#FFFFFF');
  return actuales.concat(faltan);
}

/**
 * Guarda una Gift Card recien emitida. Se llama desde el navegador justo
 * despues de generar el PDF. `datos` = {codigo, noches, vigencia,
 * mesesValidos, inclusiones, destinatario, remitente, idioma, mensaje}.
 */
function guardarGiftCard(datos) {
  var hoja = _hojaGiftCards();
  var valores = {
    'Codigo': datos.codigo || '', 'FechaEmision': new Date(), 'Noches': datos.noches || '',
    'VigenciaMeses': datos.vigencia || '', 'MesesValidos': datos.mesesValidos || '',
    'Inclusiones': datos.inclusiones || '', 'Para': datos.destinatario || '', 'De': datos.remitente || '',
    'Idioma': datos.idioma || '', 'Mensaje': datos.mensaje || '',
    'FechaDesde': datos.fechaDesde || '', 'FechaHasta': datos.fechaHasta || '',
    'EstadoReserva': datos.estadoReserva || ''
  };
  // Se escribe por nombre de columna: si la planilla tenía otro orden o
  // columnas extra, cada dato igual cae donde corresponde.
  var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0]
                        .map(function (c) { return String(c).trim(); });
  hoja.appendRow(encabezados.map(function (h) {
    return Object.prototype.hasOwnProperty.call(valores, h) ? valores[h] : '';
  }));
  return { ok: true };
}

/**
 * Devuelve las Gift Cards emitidas, mas recientes primero (hasta `limite`).
 */
function listarGiftCards(limite) {
  limite = limite || 50;
  var hoja = _hojaGiftCards();
  var datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return [];
  var encabezados = datos[0];
  var filas = datos.slice(1).map(function (fila) {
    var obj = {};
    for (var c = 0; c < encabezados.length; c++) obj[encabezados[c]] = fila[c];
    if (obj.FechaEmision instanceof Date) obj.FechaEmision = Utilities.formatDate(obj.FechaEmision, Session.getScriptTimeZone() || 'America/Santiago', 'dd/MM/yyyy HH:mm');
    return obj;
  });
  filas.reverse();
  return filas.slice(0, limite);
}
