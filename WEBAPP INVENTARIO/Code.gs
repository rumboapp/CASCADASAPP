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
 * La planilla es siempre "Cascadas Hotel - Inventario Bar" (SPREADSHEET_ID_FIJO
 * mas abajo, el ID esta fijo en el codigo, no hay que configurar nada).
 * Ejecuta crearBaseDeDatos() UNA vez desde el editor de Apps Script para
 * asegurar que esten todas las hojas (Inventario/Historial): es segura de
 * re-correr, solo agrega lo que falte sin tocar datos existentes.
 *
 * Sugerencia del Chef no guarda historial: cada PDF se genera al momento a
 * partir de lo que hay cargado en el formulario, sin pasar por la planilla.
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

// Segura de re-correr: usa siempre la planilla fija (SPREADSHEET_ID_FIJO) y
// solo agrega las hojas que falten, sin tocar los datos que ya existen.
function crearBaseDeDatos() {
  var ss = _ss();
  if (!ss.getSheetByName('Inventario')) _crearHojaInventario(ss);
  if (!ss.getSheetByName('Historial')) _crearHojaHistorial(ss);
  if (!ss.getSheetByName('MenusChef')) _asegurarHojaMenusChef();
  Logger.log('Listo. Planilla: ' + ss.getUrl());
  return ss.getId();
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

function _eliminarHojaPorDefecto(ss) {
  var hojaDefault = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hojaDefault && ss.getSheets().length > 1) ss.deleteSheet(hojaDefault);
}

// ---------------------------------------------------------------------------
// ACCESO AL SPREADSHEET
// ---------------------------------------------------------------------------
// ID fijo de la planilla real (Cascadas Hotel - Inventario Bar), para que no
// haya ninguna duda de a cual planilla se lee/escribe. Si en algun momento
// se necesita otra, se cambia aca.
var SPREADSHEET_ID_FIJO = '1jJ4sxTASvUACS6b1nX7m3F8UCd0J286H-5DHEUnyQQ0';
function _ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID_FIJO);
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
// MENUS GUARDADOS (Sugerencia del Chef)
// ---------------------------------------------------------------------------
// Los platos se guardan como JSON en una sola celda. IMPORTANTE: todo lo que
// se devuelve al cliente son strings/numeros planos; si se devuelve un objeto
// Date, google.script.run falla EN SILENCIO (no llama al success ni al
// failure handler) y el historial aparece vacio sin ningun error.

/** Crea la hoja MenusChef si no existe (auto-reparacion, sin migracion manual). */
function _asegurarHojaMenusChef() {
  var ss = _ss();
  var hoja = ss.getSheetByName('MenusChef');
  if (!hoja) {
    hoja = _crearHoja(ss, 'MenusChef', ['ID', 'Guardado', 'FechaCena', 'Titulo', 'PlatosJSON'], []);
  }
  return hoja;
}

/** Fecha de la cena como texto "yyyy-MM-dd" (Sheets puede convertirla a Date). */
function _textoFechaCena(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v ? String(v) : '';
}

/**
 * Guarda (o actualiza, si viene id) un menu del chef.
 * datos = { id?, titulo, fechaCena, platos: [{tiempo, nombre, descripcion}] }
 */
function guardarMenuChef(datos) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var hoja = _asegurarHojaMenusChef();
    var platosJSON = JSON.stringify((datos && datos.platos) || []);
    var titulo = String((datos && datos.titulo) || '');
    var fechaCena = String((datos && datos.fechaCena) || '');
    if (datos && datos.id) {
      var fila = _leerHojaComoObjetos('MenusChef').filter(function (r) {
        return String(r.ID) === String(datos.id);
      })[0];
      if (fila) {
        hoja.getRange(fila._fila, 1, 1, 5).setValues([[String(datos.id), new Date(), fechaCena, titulo, platosJSON]]);
        return { success: true, id: String(datos.id) };
      }
    }
    var id = generarID();
    hoja.appendRow([id, new Date(), fechaCena, titulo, platosJSON]);
    return { success: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

/** Lista los menus guardados, del mas reciente al mas antiguo. */
function obtenerMenusChef() {
  _asegurarHojaMenusChef();
  var filas = _leerHojaComoObjetos('MenusChef');
  filas.sort(function (a, b) {
    var da = a.Guardado instanceof Date ? a.Guardado.getTime() : 0;
    var db = b.Guardado instanceof Date ? b.Guardado.getTime() : 0;
    return db - da;
  });
  return filas.map(function (r) {
    return {
      id: String(r.ID),
      guardado: _fechaHoraTexto(r.Guardado),
      fechaCena: _textoFechaCena(r.FechaCena),
      titulo: r.Titulo ? String(r.Titulo) : '',
      platosJSON: r.PlatosJSON ? String(r.PlatosJSON) : '[]'
    };
  });
}

/** Elimina un menu guardado. */
function eliminarMenuChef(id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var hoja = _asegurarHojaMenusChef();
    var fila = _leerHojaComoObjetos('MenusChef').filter(function (r) {
      return String(r.ID) === String(id);
    })[0];
    if (!fila) return { success: false, mensaje: 'Menu no encontrado.' };
    hoja.deleteRow(fila._fila);
    return { success: true, mensaje: 'Menu eliminado.' };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// LINK MOVIL (para pestaña QR)
// ---------------------------------------------------------------------------
function obtenerUrlMovil() {
  var base = '';
  try { base = ScriptApp.getService().getUrl() || ''; } catch (err) {}
  if (!base) return '';
  // getUrl() puede devolver la URL con un fragmento (#...) o un query previo.
  // Si se pega "?movil=1" tal cual y ya habia un "#", el parametro queda
  // DENTRO del fragmento y el servidor nunca lo recibe: doGet no detecta
  // movil y termina abriendo el panel de escritorio. Por eso limpiamos
  // cualquier ?/# y normalizamos a /exec (a veces getUrl da la /dev de editor)
  // antes de agregar el parametro, para que el QR abra siempre el conteo movil.
  base = base.split('#')[0].split('?')[0].replace(/\/dev$/, '/exec');
  return base + '?movil=1';
}
