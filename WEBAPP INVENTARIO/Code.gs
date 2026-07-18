/**
 * ============================================================================
 * CASCADAS HOTEL — APP DE COCINA/RESTAURANT
 * ============================================================================
 * Dos herramientas en una sola app (mismo link de siempre, no cambia el QR
 * ya impreso):
 *  - Inventario de bar: reemplaza la ficha en papel que imprimian cada 2
 *    dias. Sin login: cualquiera con el link entra y edita.
 *  - Sugerencia del Chef: reemplaza el proceso de escribir el menu especial
 *    en un papel y pasarlo a recepcion para pasarlo a Word/Canva a mano.
 *    El chef (o quien reciba el papelito) lo carga en el panel y se genera
 *    un PDF elegante con los colores del hotel, listo para imprimir.
 *
 * Un solo link (sin ?movil) abre el panel completo (Inventario/Sugerencia
 * del Chef/QR/Historial). El mismo link + ?movil=1 abre la vista de conteo
 * movil (la que se abre al escanear el QR), pensada para caminar por el bar
 * sumando/restando con el celular.
 *
 * Primer uso: ejecuta crearBaseDeDatos() UNA vez desde el editor de Apps
 * Script. Crea el Spreadsheet, guarda su ID en PropertiesService (igual
 * patron que las demas apps) y siembra el inventario con los 122 productos
 * de la ficha original, en cantidad 0. Si ya tenias la base de datos de
 * Inventario creada, volver a ejecutar esta funcion es seguro: solo agrega
 * la hoja de Sugerencia del Chef que falte, sin tocar el inventario.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// PUNTO DE ENTRADA WEB
// ---------------------------------------------------------------------------
function doGet(e) {
  var esMovil = e && e.parameter && (e.parameter.movil === '1' || e.parameter.movil === 'true');
  var plantilla = HtmlService.createTemplateFromFile(esMovil ? 'Movil' : 'Index');
  return plantilla.evaluate()
    .setTitle(esMovil ? 'Conteo de Inventario - Cascadas Hotel' : 'Cocina y Restaurant - Cascadas Hotel')
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
  var props = PropertiesService.getScriptProperties();
  var idExistente = props.getProperty('SPREADSHEET_ID');
  if (idExistente) {
    // Instalacion existente: solo agrega hojas nuevas que falten (por ejemplo
    // al actualizar de una version que no tenia Sugerencia del Chef), sin
    // tocar los datos que ya existen.
    var ssExistente = SpreadsheetApp.openById(idExistente);
    if (!ssExistente.getSheetByName('MenuHistorial')) {
      _crearHojaMenuChef(ssExistente);
      Logger.log('Se agrego la hoja MenuHistorial (Sugerencia del Chef) a la base de datos existente.');
    }
    Logger.log('Ya existe una base de datos con ID: ' + idExistente);
    return idExistente;
  }

  var ss = SpreadsheetApp.create('Cascadas Hotel - Cocina y Restaurant');
  var id = ss.getId();

  _crearHojaInventario(ss);
  _crearHojaHistorial(ss);
  _crearHojaMenuChef(ss);
  _eliminarHojaPorDefecto(ss);

  props.setProperty('SPREADSHEET_ID', id);

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
    return [generarID(), par[0], par[1], 'un.', 0, 0, 0, '', i + 1];
  });
  _crearHoja(ss, 'Inventario', [
    'ID', 'Categoria', 'Producto', 'Unidad', 'Cantidad', 'Minimo', 'Ideal', 'Detalle', 'Orden'
  ], filas);
}

function _crearHojaHistorial(ss) {
  _crearHoja(ss, 'Historial', [
    'SnapshotID', 'Fecha', 'ItemID', 'Categoria', 'Producto', 'Unidad', 'Cantidad', 'Minimo', 'Ideal'
  ], []);
}

function _crearHojaMenuChef(ss) {
  _crearHoja(ss, 'MenuHistorial', [
    'SnapshotID', 'FechaGenerado', 'Titulo', 'FechaCena', 'PlatoOrden', 'Tiempo', 'Nombre', 'Descripcion'
  ], []);
}

// La hoja de Sugerencia del Chef se agrego despues del lanzamiento inicial
// de Inventario; esto la crea sola si alguien no volvio a correr
// crearBaseDeDatos() tras actualizar el codigo, en vez de tirar error. Ademas
// repone el encabezado si la hoja quedo sin el.
var _MENU_ENCABEZADOS = ['SnapshotID', 'FechaGenerado', 'Titulo', 'FechaCena', 'PlatoOrden', 'Tiempo', 'Nombre', 'Descripcion'];
function _asegurarHojaMenuChef() {
  var ss = _ss();
  var hoja = ss.getSheetByName('MenuHistorial');
  if (!hoja) { _crearHojaMenuChef(ss); return; }
  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, _MENU_ENCABEZADOS.length).setValues([_MENU_ENCABEZADOS]);
  }
}

function _eliminarHojaPorDefecto(ss) {
  var hojaDefault = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hojaDefault && ss.getSheets().length > 1) ss.deleteSheet(hojaDefault);
}

// ---------------------------------------------------------------------------
// ACCESO AL SPREADSHEET
// ---------------------------------------------------------------------------
// Sin cache: cada llamada abre la planilla de nuevo. Un cache en variable de
// script (var a nivel de archivo) puede quedar apuntando a un estado viejo
// si Apps Script reutiliza el mismo contenedor entre ejecuciones distintas,
// y eso puede hacer que una lectura no vea una escritura muy reciente.
function _ss() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('No se encontro SPREADSHEET_ID. Ejecuta crearBaseDeDatos() primero.');
  return SpreadsheetApp.openById(id);
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
// Orden de categorias tal como aparecian en la ficha original (para que los
// vinos queden agrupados al final: Tinto, Espumante, Blanco). Las categorias
// que no esten en esta lista (agregadas a mano) se ordenan alfabeticamente
// al final de todas.
var ORDEN_CATEGORIAS = ['Bebida', 'Agua', 'Cerveza', 'Destilados', 'Tinto', 'Espumante', 'Blanco'];
function _indiceOrdenCategoria(categoria) {
  var i = ORDEN_CATEGORIAS.indexOf(categoria);
  return i === -1 ? ORDEN_CATEGORIAS.length : i;
}

function obtenerInventario() {
  var filas = _leerHojaComoObjetos('Inventario');
  filas.sort(function (a, b) {
    if (a.Categoria !== b.Categoria) {
      var ia = _indiceOrdenCategoria(a.Categoria), ib = _indiceOrdenCategoria(b.Categoria);
      if (ia !== ib) return ia - ib;
      return a.Categoria < b.Categoria ? -1 : 1;
    }
    return numero_(a.Orden) - numero_(b.Orden);
  });
  return filas.map(function (r) {
    return {
      id: r.ID, categoria: r.Categoria, producto: r.Producto, unidad: r.Unidad || 'un.',
      cantidad: numero_(r.Cantidad), minimo: numero_(r.Minimo), ideal: numero_(r.Ideal),
      detalle: r.Detalle || '', orden: numero_(r.Orden)
    };
  });
}

/**
 * Crea o actualiza un item. Si datos.id viene, actualiza; si no, crea uno
 * nuevo al final de su categoria.
 * @param {Object} datos {id, categoria, producto, unidad, cantidad, minimo, ideal, detalle}
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
      set(fila, 'Detalle', String(datos.detalle || ''));
      return { success: true, id: datos.id, mensaje: 'Item actualizado.' };
    }

    var maxOrden = filas.reduce(function (m, r) { return Math.max(m, numero_(r.Orden)); }, 0);
    var id = generarID();
    // Arma la fila por nombre de columna (no por posicion): en instalaciones
    // creadas antes de quitar "Ubicacion" el orden de columnas es distinto,
    // y esto evita que los valores queden desalineados.
    var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var valoresPorNombre = {
      ID: id, Categoria: categoria, Producto: producto,
      Unidad: String(datos.unidad || 'un.').trim() || 'un.',
      Cantidad: numero_(datos.cantidad), Minimo: numero_(datos.minimo), Ideal: numero_(datos.ideal),
      Detalle: String(datos.detalle || ''), Orden: maxOrden + 1
    };
    var fila = encabezados.map(function (h) { return valoresPorNombre.hasOwnProperty(h) ? valoresPorNombre[h] : ''; });
    hoja.appendRow(fila);
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

// ---------------------------------------------------------------------------
// SUGERENCIA DEL CHEF
// ---------------------------------------------------------------------------

/**
 * Guarda una copia del menu en el historial. Se llama cada vez que se genera
 * el PDF, para que recepcion pueda volver a imprimir una cena pasada.
 * @param {Object} datos {titulo, fechaCena, platos: [{tiempo, nombre, descripcion}]}
 */
function guardarMenuChef(datos) {
  datos = datos || {};
  var platos = (datos.platos || []).filter(function (p) { return p && String(p.nombre || '').trim(); });
  if (!platos.length) return { success: false, mensaje: 'Agrega al menos un plato con nombre.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    _asegurarHojaMenuChef();
    var hoja = _hoja('MenuHistorial');
    var snapshotId = generarID();
    var fechaGenerado = new Date();
    var titulo = String(datos.titulo || '').trim();
    var fechaCena = String(datos.fechaCena || '').trim();

    var filas = platos.map(function (p, i) {
      return [
        snapshotId, fechaGenerado, titulo, fechaCena, i + 1,
        String(p.tiempo || '').trim(), String(p.nombre || '').trim(), String(p.descripcion || '').trim()
      ];
    });
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
    SpreadsheetApp.flush(); // asegura que quede persistido antes de responder

    return {
      success: true, snapshotId: snapshotId, mensaje: 'Menu guardado en el historial.',
      fecha: _fechaHoraTexto(fechaGenerado)
    };
  } finally {
    lock.releaseLock();
  }
}

/** Lista de menus guardados, mas reciente primero. */
function obtenerHistorialMenus() {
  _asegurarHojaMenuChef();
  SpreadsheetApp.flush(); // fuerza a ver cualquier escritura reciente antes de leer
  var filas = _leerHojaComoObjetos('MenuHistorial');
  var porSnapshot = {};
  var orden = [];
  filas.forEach(function (r) {
    var id = String(r.SnapshotID || '').trim();
    if (!id) return; // fila sin id (vacia/corrupta): se ignora, no rompe el resto
    if (!porSnapshot[id]) {
      porSnapshot[id] = {
        snapshotId: id, fecha: _fechaHoraTexto(r.FechaGenerado),
        fechaOrden: r.FechaGenerado instanceof Date ? r.FechaGenerado.getTime() : 0,
        titulo: r.Titulo || '', fechaCena: r.FechaCena || '', cantidadPlatos: 0
      };
      orden.push(id);
    }
    porSnapshot[id].cantidadPlatos++;
  });
  return orden.map(function (id) { return porSnapshot[id]; })
    .sort(function (a, b) { return b.fechaOrden - a.fechaOrden; });
}

/** Detalle completo (platos ordenados) de un menu guardado. */
function obtenerDetalleMenu(snapshotId) {
  _asegurarHojaMenuChef();
  var filas = _leerHojaComoObjetos('MenuHistorial').filter(function (r) { return r.SnapshotID === snapshotId; });
  if (!filas.length) return null;
  filas.sort(function (a, b) { return numero_(a.PlatoOrden) - numero_(b.PlatoOrden); });
  return {
    titulo: filas[0].Titulo || '', fechaCena: filas[0].FechaCena || '',
    fecha: _fechaHoraTexto(filas[0].FechaGenerado),
    platos: filas.map(function (r) { return { tiempo: r.Tiempo || '', nombre: r.Nombre || '', descripcion: r.Descripcion || '' }; })
  };
}

/**
 * Diagnostico del historial de menus: dice si la hoja existe, cuantas filas
 * tiene, y que devuelve exactamente obtenerHistorialMenus() (la MISMA
 * funcion que usa la app, llamada aca mismo) para que sea imposible que el
 * diagnostico y la lista real digan cosas distintas. Si esta funcion no
 * existe al llamarla desde la web, es señal de que el Code.gs desplegado
 * esta desactualizado.
 */
function diagnosticoMenus() {
  var info = { ok: true, version: 'menus-3' };
  try {
    var ss = _ss();
    info.spreadsheetId = ss.getId();
    info.spreadsheetUrl = ss.getUrl();
    var hoja = ss.getSheetByName('MenuHistorial');
    info.hojaExiste = !!hoja;
    if (hoja) {
      info.ultimaFila = hoja.getLastRow();
      info.registros = _leerHojaComoObjetos('MenuHistorial').length;
      var lista = obtenerHistorialMenus(); // la funcion real, no una copia paralela
      info.menusDistintos = lista.length;
      info.listaReal = lista;
    }
  } catch (e) {
    info.ok = false;
    info.error = String(e && e.message ? e.message : e);
  }
  return info;
}
