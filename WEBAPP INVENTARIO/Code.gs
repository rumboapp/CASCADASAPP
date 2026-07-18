/**
 * ============================================================================
 * CASCADAS HOTEL — INVENTARIO DE BAR/RESTAURANT
 * ============================================================================
 * Reemplaza la ficha de inventario en papel (Word de ~6 paginas) que
 * imprimian cada 2 dias. Sin login: cualquiera con el link entra y edita
 * (no hay datos sensibles de huespedes en esta app).
 *
 * Un solo link (sin ?movil) abre el panel completo (Inventario/QR/Historial).
 * El mismo link + ?movil=1 abre la vista de conteo movil (la que se abre al
 * escanear el QR), pensada para caminar por el bar sumando/restando con el
 * celular.
 *
 * Primer uso: ejecuta crearBaseDeDatos() UNA vez desde el editor de Apps
 * Script. Crea el Spreadsheet, guarda su ID en PropertiesService (igual
 * patron que las demas apps) y siembra el inventario con los 122 productos
 * de la ficha original, en cantidad 0.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// PUNTO DE ENTRADA WEB
// ---------------------------------------------------------------------------
function doGet(e) {
  var esMovil = e && e.parameter && (e.parameter.movil === '1' || e.parameter.movil === 'true');
  var plantilla = HtmlService.createTemplateFromFile(esMovil ? 'Movil' : 'Index');
  return plantilla.evaluate()
    .setTitle(esMovil ? 'Conteo de Inventario - Cascadas Hotel' : 'Inventario Bar/Restaurant - Cascadas Hotel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------------
// SETUP (ejecutar UNA vez desde el editor)
// ---------------------------------------------------------------------------
var CATEGORIAS_SEED = [
    ['Bebida', 'Coca cola'],
    ['Bebida', 'Coca cola sin azúcar'],
    ['Bebida', 'Coca cola light'],
    ['Bebida', 'Sprite'],
    ['Bebida', 'Sprite sin azúcar'],
    ['Bebida', 'Fanta'],
    ['Bebida', 'Schweppes Ginger ale 0'],
    ['Bebida', 'Schweppes Ginger ale'],
    ['Bebida', 'Schweppes Tónica rose'],
    ['Bebida', 'Canada dry zero'],
    ['Bebida', 'Schweppes Agua tónica sin azúcar'],
    ['Bebida', 'Ginger ale canada dry'],
    ['Bebida', 'Red bull'],
    ['Bebida', 'Red bull 0'],
    ['Bebida', 'Red bull red'],
    ['Bebida', 'Red bull yellow'],
    ['Bebida', 'Bless'],
    ['Bebida', 'Fentimans Tónica water'],
    ['Bebida', 'Fentimans ginger beer'],
    ['Agua', 'Con gas'],
    ['Agua', 'Sin gas'],
    ['Cerveza', 'Koiwe'],
    ['Cerveza', 'Lenga'],
    ['Cerveza', 'Arrayan'],
    ['Cerveza', 'Mañio'],
    ['Cerveza', 'Quila'],
    ['Cerveza', 'Tepu'],
    ['Cerveza', 'Michay'],
    ['Cerveza', 'Budweiser'],
    ['Cerveza', 'Austral lager'],
    ['Cerveza', 'Tropical stout'],
    ['Cerveza', 'Tropera strong'],
    ['Cerveza', 'Tropera brown'],
    ['Cerveza', 'Tropera blonde'],
    ['Cerveza', 'Tropera blanche'],
    ['Cerveza', 'Austral calafate'],
    ['Cerveza', 'Heineken sin alcohol'],
    ['Cerveza', 'Kustman Torobayo'],
    ['Destilados', 'Pisco tres erres'],
    ['Destilados', 'Alto del Carmen 40'],
    ['Destilados', 'Alto del Carmen 35'],
    ['Destilados', 'Bauzá aniversario'],
    ['Destilados', 'Horno quemado'],
    ['Destilados', 'Pisco espíritu de los andes'],
    ['Destilados', 'Bauzá blanco'],
    ['Destilados', 'Bauzá blanco especial'],
    ['Destilados', 'Buchanan\'s de luxe'],
    ['Destilados', 'Johnnie Walker Black'],
    ['Destilados', 'Johnnie Walker Red'],
    ['Destilados', 'Chivas Regal 18'],
    ['Destilados', 'Chivas Regal 15'],
    ['Destilados', 'Chivas Regal 12'],
    ['Destilados', 'Jack Daniel\'s Black'],
    ['Destilados', 'Jack Daniel\'s Honey'],
    ['Destilados', 'Jack Daniel\'s Apple'],
    ['Destilados', 'Jack Daniel\'s Fire'],
    ['Destilados', 'Drambuie'],
    ['Destilados', 'Gin Andes bloom'],
    ['Destilados', 'Beefeater'],
    ['Destilados', 'Gin cítrico'],
    ['Destilados', 'Gin floral'],
    ['Destilados', 'Goodwinds'],
    ['Destilados', 'Hendrick\'s'],
    ['Destilados', 'Bombay Sapphire'],
    ['Destilados', 'Grey Goose'],
    ['Destilados', 'Smirnoff'],
    ['Destilados', 'Absolut'],
    ['Destilados', 'Havana Club Especial'],
    ['Destilados', 'Havana Club 7'],
    ['Destilados', 'Bacardí'],
    ['Destilados', 'Malibu'],
    ['Destilados', 'Don Julio'],
    ['Destilados', 'Olmeca Reposado'],
    ['Destilados', 'Olmeca Silver'],
    ['Destilados', 'Sangría'],
    ['Destilados', 'Jägermeister'],
    ['Destilados', 'Ramazzotti Rosado'],
    ['Destilados', 'Ramazzotti Violetto'],
    ['Destilados', 'Martini'],
    ['Destilados', 'Aperol'],
    ['Destilados', 'Vermouth'],
    ['Destilados', 'Rossard Spritz'],
    ['Destilados', 'Carpano'],
    ['Destilados', 'Fernet'],
    ['Destilados', 'Granadina'],
    ['Destilados', 'Mitjans triple sec'],
    ['Destilados', 'Mitjans licor café'],
    ['Destilados', 'Mitjans cassis'],
    ['Destilados', 'Amaretto'],
    ['Destilados', 'Cachaca'],
    ['Destilados', 'Pipeño'],
    ['Tinto', 'Presumido'],
    ['Tinto', 'C.D reserva merlot'],
    ['Tinto', 'Casa patronales cabernet'],
    ['Tinto', 'Isabel'],
    ['Tinto', 'C.D Cabernet sauvignon'],
    ['Tinto', 'C.D Carmenere'],
    ['Tinto', 'P.C limited edition carmenere'],
    ['Tinto', 'Veraz'],
    ['Tinto', 'Caballo loco'],
    ['Tinto', 'Milla cala'],
    ['Tinto', 'Casa donoso ensamblaje'],
    ['Tinto', 'Villard pinot noir'],
    ['Tinto', 'Amplus'],
    ['Tinto', 'Pc limited edition syrah'],
    ['Tinto', 'Tanagra'],
    ['Tinto', 'Pc gran reserva'],
    ['Tinto', 'Toro de piedra'],
    ['Tinto', 'Santa ema merlot'],
    ['Tinto', 'Coyam'],
    ['Tinto', 'Amayna'],
    ['Tinto', 'Atrevido'],
    ['Tinto', 'Piu belle'],
    ['Tinto', 'Alboroto'],
    ['Tinto', 'Valiente'],
    ['Blanco', 'Villard chardonnay'],
    ['Blanco', 'Casa donoso Chardonnay'],
    ['Blanco', 'Casa donoso Sauvignon blanc'],
    ['Espumante', 'Casa bauza gala'],
    ['Espumante', 'Casa donoso brut'],
    ['Espumante', 'Idolatría brut'],
    ['Espumante', 'SBX brut']
];

function crearBaseDeDatos() {
  var idExistente = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (idExistente) {
    Logger.log('Ya existe una base de datos con ID: ' + idExistente);
    Logger.log('Si deseas recrearla, borra la propiedad SPREADSHEET_ID primero.');
    return idExistente;
  }

  var ss = SpreadsheetApp.create('Cascadas Hotel - Inventario Bar');
  var id = ss.getId();

  _crearHojaInventario(ss);
  _crearHojaHistorial(ss);
  _eliminarHojaPorDefecto(ss);

  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);

  Logger.log('Base de datos creada con exito.');
  Logger.log('SPREADSHEET_ID: ' + id);
  Logger.log('URL: ' + ss.getUrl());
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

function _crearHojaInventario(ss) {
  var filas = CATEGORIAS_SEED.map(function (par, i) {
    return [generarID(), par[0], par[1], 'un.', 0, 0, 0, '', '', i + 1];
  });
  _crearHoja(ss, 'Inventario', [
    'ID', 'Categoria', 'Producto', 'Unidad', 'Cantidad', 'Minimo', 'Ideal', 'Ubicacion', 'Detalle', 'Orden'
  ], filas);
}

function _crearHojaHistorial(ss) {
  _crearHoja(ss, 'Historial', [
    'SnapshotID', 'Fecha', 'ItemID', 'Categoria', 'Producto', 'Unidad', 'Cantidad', 'Minimo', 'Ideal'
  ], []);
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
function _indiceColumna(hoja, nombre) {
  var enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  return enc.indexOf(nombre);
}
function generarID() { return Utilities.getUuid(); }
function numero_(v) {
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

// ---------------------------------------------------------------------------
// INVENTARIO
// ---------------------------------------------------------------------------

/** Devuelve todo el inventario ordenado por categoria/orden. */
function obtenerInventario() {
  var filas = _leerHojaComoObjetos('Inventario');
  filas.sort(function (a, b) {
    if (a.Categoria !== b.Categoria) return a.Categoria < b.Categoria ? -1 : 1;
    return numero_(a.Orden) - numero_(b.Orden);
  });
  return filas.map(function (r) {
    return {
      id: r.ID, categoria: r.Categoria, producto: r.Producto, unidad: r.Unidad || 'un.',
      cantidad: numero_(r.Cantidad), minimo: numero_(r.Minimo), ideal: numero_(r.Ideal),
      ubicacion: r.Ubicacion || '', detalle: r.Detalle || '', orden: numero_(r.Orden)
    };
  });
}

/**
 * Crea o actualiza un item. Si datos.id viene, actualiza; si no, crea uno
 * nuevo al final de su categoria.
 * @param {Object} datos {id, categoria, producto, unidad, cantidad, minimo, ideal, ubicacion, detalle}
 */
function guardarItem(datos) {
  datos = datos || {};
  var categoria = String(datos.categoria || '').trim();
  var producto = String(datos.producto || '').trim();
  if (!categoria) return { success: false, mensaje: 'Falta la categoria.' };
  if (!producto) return { success: false, mensaje: 'Falta el nombre del producto.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var hoja = _hoja('Inventario');
    var filas = _leerHojaComoObjetos('Inventario');

    var set = function (fila, col, val) {
      var i = _indiceColumna(hoja, col);
      if (i !== -1) hoja.getRange(fila, i + 1).setValue(val);
    };

    if (datos.id) {
      var existente = filas.filter(function (r) { return r.ID === datos.id; })[0];
      if (!existente) return { success: false, mensaje: 'Item no encontrado.' };
      var fila = existente._fila;
      set(fila, 'Categoria', categoria);
      set(fila, 'Producto', producto);
      set(fila, 'Unidad', String(datos.unidad || 'un.').trim() || 'un.');
      if (datos.cantidad !== undefined) set(fila, 'Cantidad', numero_(datos.cantidad));
      set(fila, 'Minimo', numero_(datos.minimo));
      set(fila, 'Ideal', numero_(datos.ideal));
      set(fila, 'Ubicacion', String(datos.ubicacion || ''));
      set(fila, 'Detalle', String(datos.detalle || ''));
      return { success: true, id: datos.id, mensaje: 'Item actualizado.' };
    }

    var maxOrden = filas.reduce(function (m, r) { return Math.max(m, numero_(r.Orden)); }, 0);
    var id = generarID();
    hoja.appendRow([
      id, categoria, producto, String(datos.unidad || 'un.').trim() || 'un.',
      numero_(datos.cantidad), numero_(datos.minimo), numero_(datos.ideal),
      String(datos.ubicacion || ''), String(datos.detalle || ''), maxOrden + 1
    ]);
    return { success: true, id: id, mensaje: 'Item agregado.' };
  } finally {
    lock.releaseLock();
  }
}

function eliminarItem(id) {
  var hoja = _hoja('Inventario');
  var fila = _leerHojaComoObjetos('Inventario').filter(function (r) { return r.ID === id; })[0];
  if (!fila) return { success: false, mensaje: 'Item no encontrado.' };
  hoja.deleteRow(fila._fila);
  return { success: true, mensaje: 'Item eliminado.' };
}

/**
 * Suma (o resta, con delta negativo) una cantidad al stock actual de un
 * item, de forma atomica (para los +/- del conteo movil, donde varios
 * celulares pueden estar tocando al mismo tiempo).
 * @param {string} id
 * @param {number} delta
 * @return {Object} {success, cantidad, mensaje}
 */
function incrementarCantidad(id, delta) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var hoja = _hoja('Inventario');
    var fila = _leerHojaComoObjetos('Inventario').filter(function (r) { return r.ID === id; })[0];
    if (!fila) return { success: false, mensaje: 'Item no encontrado.' };
    var nueva = Math.max(0, numero_(fila.Cantidad) + numero_(delta));
    var colCant = _indiceColumna(hoja, 'Cantidad');
    hoja.getRange(fila._fila, colCant + 1).setValue(nueva);
    return { success: true, cantidad: nueva };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Guarda una foto del inventario actual en Historial y pone todas las
 * cantidades en 0. Se usa cuando el restaurant quiere empezar un conteo
 * nuevo desde cero.
 */
function borrarTodoElStock() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var hojaInv = _hoja('Inventario');
    var filas = _leerHojaComoObjetos('Inventario');
    if (!filas.length) return { success: false, mensaje: 'No hay inventario para reiniciar.' };

    var snapshotID = generarID();
    var fecha = new Date();
    var hojaHist = _hoja('Historial');
    var filasHist = filas.map(function (r) {
      return [snapshotID, fecha, r.ID, r.Categoria, r.Producto, r.Unidad || 'un.',
        numero_(r.Cantidad), numero_(r.Minimo), numero_(r.Ideal)];
    });
    hojaHist.getRange(hojaHist.getLastRow() + 1, 1, filasHist.length, 9).setValues(filasHist);

    var colCant = _indiceColumna(hojaInv, 'Cantidad');
    var ceros = filas.map(function () { return [0]; });
    hojaInv.getRange(2, colCant + 1, ceros.length, 1).setValues(ceros);

    return { success: true, mensaje: 'Stock reiniciado. Se guardo una copia en el historial.' };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// HISTORIAL / REPORTES
// ---------------------------------------------------------------------------

/** Lista de snapshots pasados (fecha + cantidad de items), mas reciente primero. */
function obtenerListaHistorial() {
  var filas = _leerHojaComoObjetos('Historial');
  var porSnapshot = {};
  var orden = [];
  filas.forEach(function (r) {
    if (!porSnapshot[r.SnapshotID]) {
      porSnapshot[r.SnapshotID] = { snapshotId: r.SnapshotID, fecha: _fechaHoraTexto(r.Fecha), fechaObj: r.Fecha, items: 0, totalUnidades: 0 };
      orden.push(r.SnapshotID);
    }
    porSnapshot[r.SnapshotID].items++;
    porSnapshot[r.SnapshotID].totalUnidades += numero_(r.Cantidad);
  });
  return orden.map(function (id) { return porSnapshot[id]; })
    .sort(function (a, b) { return new Date(b.fechaObj) - new Date(a.fechaObj); })
    .map(function (s) { delete s.fechaObj; return s; });
}

/** Detalle completo de un snapshot especifico (para expandirlo en la UI). */
function obtenerDetalleSnapshot(snapshotId) {
  return _leerHojaComoObjetos('Historial')
    .filter(function (r) { return r.SnapshotID === snapshotId; })
    .map(function (r) {
      return { categoria: r.Categoria, producto: r.Producto, unidad: r.Unidad || 'un.', cantidad: numero_(r.Cantidad) };
    })
    .sort(function (a, b) { return a.categoria !== b.categoria ? (a.categoria < b.categoria ? -1 : 1) : 0; });
}

/**
 * Ranking de consumo estimado por producto: para cada par de snapshots
 * consecutivos (el mas antiguo primero, terminando en el stock actual), si
 * la cantidad bajo, esa baja se cuenta como "consumido". No se registran
 * entradas de mercaderia, asi que es una aproximacion (si llega stock nuevo
 * entre medio, se subestima el consumo real), pero sirve para ver de un
 * vistazo que se esta moviendo mas.
 * @return {Array} [{producto, categoria, unidad, consumoEstimado}] ordenado desc.
 */
function obtenerConsumoEstimado() {
  var hist = _leerHojaComoObjetos('Historial');
  var actual = _leerHojaComoObjetos('Inventario');

  // Agrupa historial por snapshot, ordenado cronologicamente.
  var porSnapshot = {};
  hist.forEach(function (r) {
    if (!porSnapshot[r.SnapshotID]) porSnapshot[r.SnapshotID] = { fecha: r.Fecha, items: {} };
    porSnapshot[r.SnapshotID].items[r.ItemID] = {
      cantidad: numero_(r.Cantidad), producto: r.Producto, categoria: r.Categoria, unidad: r.Unidad || 'un.'
    };
  });
  var snapshotsOrdenados = Object.keys(porSnapshot)
    .map(function (id) { return porSnapshot[id]; })
    .sort(function (a, b) { return new Date(a.fecha) - new Date(b.fecha); });

  // Secuencia: snapshot1 -> snapshot2 -> ... -> stock actual.
  var actualPorId = {};
  actual.forEach(function (r) {
    actualPorId[r.ID] = { cantidad: numero_(r.Cantidad), producto: r.Producto, categoria: r.Categoria, unidad: r.Unidad || 'un.' };
  });
  var secuencia = snapshotsOrdenados.concat([{ items: actualPorId }]);

  var consumoPorItem = {}; // itemId -> {producto, categoria, unidad, total}
  for (var i = 0; i < secuencia.length - 1; i++) {
    var actualEtapa = secuencia[i].items;
    var siguienteEtapa = secuencia[i + 1].items;
    Object.keys(actualEtapa).forEach(function (itemId) {
      var antes = actualEtapa[itemId];
      var despues = siguienteEtapa[itemId];
      if (!despues) return;
      var baja = antes.cantidad - despues.cantidad;
      if (baja > 0) {
        if (!consumoPorItem[itemId]) {
          consumoPorItem[itemId] = { producto: despues.producto, categoria: despues.categoria, unidad: despues.unidad, total: 0 };
        }
        consumoPorItem[itemId].total += baja;
      }
    });
  }

  return Object.keys(consumoPorItem)
    .map(function (id) { return consumoPorItem[id]; })
    .filter(function (x) { return x.total > 0; })
    .sort(function (a, b) { return b.total - a.total; });
}

function _fechaHoraTexto(fecha) {
  if (!fecha) return '';
  var d = fecha instanceof Date ? fecha : new Date(fecha);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

// ---------------------------------------------------------------------------
// LINK MOVIL (para pestaña QR)
// ---------------------------------------------------------------------------
function obtenerUrlMovil() {
  var base = '';
  try { base = ScriptApp.getService().getUrl() || ''; } catch (err) {}
  return base ? (base + '?movil=1') : '';
}
