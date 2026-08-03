/**
 * ============================================================================
 * CASCADAS CONCIERGE — ACTO 2: BACKEND (Code.gs)
 * ============================================================================
 * Backend, APIs, logica de negocio, seguridad y validaciones server-side.
 *
 * Todas las funciones expuestas al frontend validan permisos y datos en el
 * servidor. Nada de horarios/capacidades/precios esta hardcodeado: todo se
 * lee de las hojas Configuracion y Servicios.
 *
 * El ID del Spreadsheet se guarda en ScriptProperties por Setup.gs.
 * ============================================================================
 */

/**
 * FUNCION DE DIAGNOSTICO. Ejecutala desde el editor de Apps Script y revisa el
 * "Registro de ejecucion". No forma parte de la app; sirve solo para depurar.
 */
function diagnostico() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  Logger.log('SPREADSHEET_ID guardado: ' + id);
  if (!id) {
    Logger.log('ERROR: no hay SPREADSHEET_ID. Ejecuta crearBaseDeDatos() primero.');
    return;
  }
  var ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (e) {
    Logger.log('ERROR abriendo el Spreadsheet: ' + e.message);
    return;
  }
  Logger.log('Nombre del Spreadsheet: ' + ss.getName());
  Logger.log('Hojas encontradas: ' + ss.getSheets().map(function (s) { return s.getName(); }).join(', '));

  var serv = ss.getSheetByName('Servicios');
  if (!serv) {
    Logger.log('ERROR: no existe la hoja "Servicios".');
  } else {
    Logger.log('Servicios -> ultima fila: ' + serv.getLastRow() + ', ultima columna: ' + serv.getLastColumn());
    Logger.log('Servicios contenido: ' + JSON.stringify(serv.getDataRange().getValues()));
  }

  var conf = ss.getSheetByName('Configuracion');
  Logger.log('Configuracion -> ultima fila: ' + (conf ? conf.getLastRow() : 'NO EXISTE'));

  try {
    Logger.log('obtenerServiciosActivos() devuelve: ' + JSON.stringify(obtenerServiciosActivos()));
  } catch (e) {
    Logger.log('obtenerServiciosActivos() ERROR: ' + e.message);
  }
}

/**
 * MIGRACION. Fusiona Almuerzo y Cena en un solo servicio "Almuerzo / Cena"
 * (S003, 13:00-22:00) y desactiva S004. Ademas crea la clave de configuracion
 * RESERVA_PASO_MINUTOS=30 para ofrecer horarios cada media hora.
 * Ejecutala UNA vez desde el editor.
 */
function migrarAlmuerzoCena() {
  var ss = _ss();
  var hoja = ss.getSheetByName('Servicios');
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] === 'S003') {
      hoja.getRange(i + 1, 2).setValue('Almuerzo / Cena');
      hoja.getRange(i + 1, 9).setNumberFormat('@').setValue('13:00');
      hoja.getRange(i + 1, 10).setNumberFormat('@').setValue('22:00');
      hoja.getRange(i + 1, 11).setValue('Restaurant: almuerzo y cena');
    }
    if (datos[i][0] === 'S004') {
      hoja.getRange(i + 1, 8).setValue('FALSE'); // columna Activo
      hoja.getRange(i + 1, 11).setValue('Fusionado en Almuerzo / Cena (S003)');
    }
  }

  // Crea RESERVA_PASO_MINUTOS si no existe todavia.
  var conf = ss.getSheetByName('Configuracion');
  var claves = conf.getRange(1, 1, conf.getLastRow(), 1).getValues().map(function (f) { return f[0]; });
  if (claves.indexOf('RESERVA_PASO_MINUTOS') === -1) {
    conf.appendRow(['RESERVA_PASO_MINUTOS', 30, 'Cada cuantos minutos se ofrecen horarios de reserva']);
  }

  _invalidarCaches('Servicios');
  _invalidarCaches('Configuracion');
  Logger.log('Migracion lista: S003 = "Almuerzo / Cena" (13:00-22:00), S004 desactivada, paso de reserva 30 min.');
}

/**
 * MIGRACION V3. Ajustes puntuales: desayuno a la habitacion requiere aprobacion,
 * agrega la regla de dias maximos para huespedes, y asegura Cena (S004) inactiva.
 * Ejecutala UNA vez desde el editor. Es idempotente.
 */
function migrarV3() {
  var ss = _ss();
  var serv = ss.getSheetByName('Servicios');
  var datos = serv.getDataRange().getValues();
  var colAprob = _indiceColumna(serv, 'RequiereAprobacion') + 1;
  var colActivo = _indiceColumna(serv, 'Activo') + 1;
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] === 'S002') serv.getRange(i + 1, colAprob).setValue('TRUE');
    if (datos[i][0] === 'S004') serv.getRange(i + 1, colActivo).setValue('FALSE');
  }

  // Config: dias maximos de reserva para huespedes.
  var conf = ss.getSheetByName('Configuracion');
  var claves = conf.getRange(1, 1, conf.getLastRow(), 1).getValues().map(function (f) { return f[0]; });
  if (claves.indexOf('MAX_DIAS_RESERVA_HUESPED') === -1) {
    conf.appendRow(['MAX_DIAS_RESERVA_HUESPED', 3, 'Dias maximos hacia adelante que un huesped puede reservar (recepcion sin limite)']);
  }

  _invalidarCaches('Servicios');
  _invalidarCaches('Configuracion');
  Logger.log('Migracion V3 lista: desayuno habitacion requiere aprobacion, regla de 3 dias, Cena inactiva.');
}

/**
 * MIGRACION V5. Agrega columnas para: entrega del pedido (Restaurant/Habitacion)
 * y el link del grupo de WhatsApp por habitacion. Ejecutala UNA vez.
 */
function migrarV5() {
  var ss = _ss();
  _asegurarColumna(ss.getSheetByName('Pedidos'), 'Entrega', 'Restaurant');
  _asegurarColumna(ss.getSheetByName('Habitaciones'), 'WhatsAppLink', '');
  // Config: interruptor para habilitar/deshabilitar el pedido a la habitacion.
  var conf = ss.getSheetByName('Configuracion');
  var claves = conf.getRange(1, 1, conf.getLastRow(), 1).getValues().map(function (f) { return f[0]; });
  if (claves.indexOf('PEDIDO_HABITACION_ACTIVO') === -1) {
    conf.appendRow(['PEDIDO_HABITACION_ACTIVO', 'TRUE', 'Habilita/deshabilita la opcion de pedir a la habitacion']);
  }
  _invalidarCaches('Configuracion');
  Logger.log('Migracion V5 lista: columna Entrega en Pedidos, WhatsAppLink en Habitaciones y config PEDIDO_HABITACION_ACTIVO.');
  Logger.log('Pega los links de los grupos de WhatsApp en la columna WhatsAppLink de la hoja Habitaciones.');
}

/**
 * MIGRACION V6 (BILINGUE). Agrega columnas en ingles a Servicios, Categorias y
 * Productos, y las rellena con la traduccion de fabrica (solo donde esten
 * vacias, sin pisar ediciones manuales). Ejecutala UNA vez desde el editor.
 * Los productos que son nombres propios/marcas (cervezas, gins, vinos, etc.)
 * no llevan traduccion: en ese caso la app muestra el nombre original.
 */
function migrarV6() {
  // Garantiza tambien lo de V5 (es idempotente): columna Entrega, WhatsAppLink
  // y el interruptor PEDIDO_HABITACION_ACTIVO. Correr V6 deja TODO al dia.
  migrarV5();
  aplicarTraduccionesEN();
  Logger.log('Migracion V6 lista: columnas NombreEN/DescripcionEN y traducciones de fabrica aplicadas.');
}

/**
 * MIGRACION V7. Agrega el recargo del pedido a la habitacion y el interruptor
 * de prepedidos. Ejecuta tambien V5 y V6 (todas idempotentes): correr V7
 * deja TODO al dia. Ejecutala UNA vez desde el editor.
 */
function migrarV7() {
  migrarV6();
  var conf = _ss().getSheetByName('Configuracion');
  var claves = conf.getRange(1, 1, conf.getLastRow(), 1).getValues().map(function (f) { return f[0]; });
  if (claves.indexOf('PEDIDO_HABITACION_RECARGO') === -1) {
    conf.appendRow(['PEDIDO_HABITACION_RECARGO', 0, 'Valor que se suma al total del pedido cuando la entrega es a la habitacion']);
  }
  if (claves.indexOf('PREPEDIDO_ACTIVO') === -1) {
    conf.appendRow(['PREPEDIDO_ACTIVO', 'TRUE', 'Habilita/deshabilita los pedidos anticipados del restaurant']);
  }
  _invalidarCaches('Configuracion');
  Logger.log('Migracion V7 lista: PEDIDO_HABITACION_RECARGO y PREPEDIDO_ACTIVO en Configuracion.');
  Logger.log('Define el valor del recargo en Configuracion (parte en 0 = sin recargo).');
}

/**
 * Diccionario de traducciones de fabrica (espanol -> ingles), por ID.
 * Solo se traduce lo que cambia en ingles; los nombres propios/marcas se omiten
 * a proposito para que la app muestre el nombre original.
 */
var TRADUCCIONES = {
  servicios: {
    S001: { nombre: 'Buffet Breakfast', desc: 'House buffet' },
    S002: { nombre: 'Room Breakfast', desc: 'Continental breakfast served in your room. Not a la carte.' },
    S003: { nombre: 'Lunch / Dinner', desc: 'Restaurant: lunch and dinner' },
    S004: { nombre: 'Dinner', desc: 'Merged into Lunch / Dinner (S003)' },
    S005: { nombre: 'Hot Tub', desc: 'Outdoor wood-fired hot tub' },
    S006: { nombre: 'Bikes', desc: '27.5-inch' },
    S007: { nombre: 'Massages', desc: 'Massage session' }
  },
  categorias: {
    CAT01: { nombre: 'Starters' }, CAT02: { nombre: 'Main Courses' },
    CAT03: { nombre: 'Fish & Seafood' }, CAT04: { nombre: 'Meats' },
    CAT05: { nombre: 'Salads' }, CAT06: { nombre: 'Pasta & Sauces' },
    CAT07: { nombre: 'For Kids' }, CAT08: { nombre: 'Desserts' },
    CAT09: { nombre: 'Coffee' }, CAT10: { nombre: 'Sandwiches' },
    CAT11: { nombre: 'Sharing Boards' }, CAT12: { nombre: 'Beers' },
    CAT13: { nombre: 'Gin' }, CAT14: { nombre: 'Vodka' }, CAT15: { nombre: 'Whisky' },
    CAT16: { nombre: 'Tequila' }, CAT17: { nombre: 'Pisco' }, CAT18: { nombre: 'Cocktails' },
    CAT19: { nombre: 'Wines' }, CAT20: { nombre: 'Non-Alcoholic' }, CAT21: { nombre: 'Bar (General)' }
  },
  productos: {
    P001: { nombre: 'Old-style beef tartare', desc: 'Chopped beef fillet mixed with mayonnaise and wholegrain mustard, lemon juice, capers, gherkins and onion. Served with house bread.' },
    P002: { nombre: 'Octopus carpaccio with olive sauce', desc: 'Octopus slices with olive sauce, crostini and avocado mousse.' },
    P003: { nombre: 'Salmon tiradito with acevichada sauce', desc: 'Salmon cuts bathed in passion-fruit acevichada sauce, finished with pico de gallo and sweet-potato mousse.' },
    P004: { nombre: 'Mixed ceviche', desc: "Tuna, shrimp and octopus marinated in tiger's milk and lemon, with onion and cilantro." },
    P005: { nombre: 'Soup of the day', desc: 'Mushroom | Squash | Asparagus | Onion soup | Ajiaco | Others of the day.' },
    P006: { nombre: 'Mushroom & heart-of-palm ceviche', desc: "Mushrooms and hearts of palm marinated in tiger's milk, with fried sweet potato and avocado mousse." },
    P007: { nombre: 'Smoked salmon oriental salad', desc: 'Green mix, smoked salmon, asparagus, avocado, confit tomatoes, crispy cream-cheese balls and jalapeno.' },
    P008: { nombre: 'Chicken Caesar salad', desc: 'Green leaf mix, grilled chicken, crispy bacon, croutons, parmesan flakes, Caesar dressing.' },
    P009: { nombre: 'Shrimp Caesar salad', desc: 'Green leaf mix, grilled salmon, crispy bacon, croutons, parmesan flakes, Caesar dressing.' },
    P010: { nombre: 'Las Cascadas garden salad', desc: 'Green leaf mix, hearts of palm, confit tomatoes, avocado, artichokes, crostini, olives and cilantro vinaigrette.' },
    P011: { nombre: 'Honey grilled salmon with pesto mote risotto and avocado mousse', desc: 'Salmon marinated in lemon and honey, grilled, served with pesto mote risotto and avocado mousse.' },
    P012: { nombre: 'Fried Southern hake with pico de gallo and merquen puree', desc: 'Crispy fried hake with homemade potato puree and a touch of smoked merquen, with pico de gallo sauce.' },
    P013: { nombre: 'Baked conger eel with shrimp sauce and wok vegetables', desc: 'Baked conger eel bathed in shrimp sauce, served with seasonal vegetables sauteed in the wok.' },
    P014: { nombre: 'Chilote crab pie', desc: 'Creamy crab pie gratinated with parmesan: a classic of the south.' },
    P015: { nombre: 'Beef fillet with sauteed asparagus and portobellos with pears', desc: 'Beef medallion wrapped in bacon and bathed in gorgonzola sauce, with asparagus, portobellos and pears sauteed in fine-herb butter.' },
    P016: { nombre: 'Skirt steak with chimichurri and rosemary native potatoes', desc: 'Grilled skirt steak marinated with house chimichurri, with potatoes sauteed in garlic and rosemary.' },
    P017: { nombre: 'Short ribs in Cabernet Sauvignon sauce with corn pudding', desc: 'Slow-cooked short ribs in red wine, bathed in their sauce and topped with sweet-corn pastelera.' },
    P018: { nombre: 'Sirloin with creamy quinoa', desc: 'Grilled sirloin over creamy quinoa with seasonal vegetables.' },
    P019: { nombre: 'Spaghetti, fettuccine or gnocchi with sauce of choice', desc: 'Bolognese | Alfredo | Chicken & pesto | Four cheeses' },
    P020: { nombre: 'Smoked salmon cannelloni', desc: 'Homemade pasta filled with smoked salmon, mushrooms, onion and local cheese. Gratinated in rose sauce and parmesan.' },
    P021: { nombre: 'Squash risotto', desc: 'Traditional rice slowly cooked in seasonal vegetable stock and white wine, finished with squash puree.' },
    P022: { nombre: 'Lucuma suspiro', desc: 'Classic Peruvian dessert with dulce de leche cream and the creamy flavor of lucuma.' },
    P023: { nombre: 'Murta panna cotta', desc: 'Traditional homemade preparation with the flavor of this seasonal southern berry.' },
    P024: { nombre: 'House cheesecake (varies by availability)', desc: 'Baked cheesecake with a biscuit base and sauce.' },
    P025: { nombre: 'Brownie with ice cream', desc: 'Homemade brownie with two scoops of ice cream.' },
    P026: { nombre: 'Grilled chicken 150g', desc: 'Served with fries, rice or puree.' },
    P027: { nombre: 'Grilled salmon 150g', desc: 'Served with fries, rice or puree.' },
    P028: { nombre: 'Mini burger 170g', desc: 'Served with fries, rice or puree.' },
    P029: { nombre: 'Homemade nuggets (5 pcs)', desc: 'Served with fries, rice or puree.' },
    P031: { nombre: 'Double espresso', desc: '' },
    P034: { nombre: 'Tea & infusions', desc: '' },
    P035: { nombre: 'Lake sandwich', desc: 'Artisan ciabatta, cream cheese, arugula, avocado, olive and smoked salmon.' },
    P036: { nombre: 'Barros Luco sandwich', desc: 'Beef or chicken steak with cheese.' },
    P037: { nombre: 'Italiano sandwich', desc: 'Beef or chicken steak, tomato, lettuce, mayonnaise, avocado.' },
    P038: { nombre: 'Cascadas burger', desc: 'Potato bun, burger, bacon, cheese, red onion, lettuce, tomato, pickle, golf sauce.' },
    P039: { nombre: 'Calbuco sandwich', desc: 'Ciabatta, black-garlic spread, serrano ham, arugula, roasted pepper.' },
    P040: { nombre: 'Osorno board', desc: 'Fries, sauteed chicken, caramelized onion, mushrooms, spicy sauce.' },
    P041: { nombre: 'Cascadas board', desc: 'Charcuterie, cheeses, fruit and nuts.' },
    P042: { nombre: 'Traditional pisco sour', desc: '' },
    P046: { nombre: 'Cathedral pisco sour', desc: '' },
    P061: { nombre: 'Sangria', desc: '' },
    P066: { nombre: 'Mineral water', desc: '' },
    P067: { nombre: 'Soft drinks', desc: '' },
    P068: { nombre: 'Juices', desc: '' },
    P069: { nombre: 'Classic lemonade', desc: '' },
    P070: { nombre: 'Special lemonade', desc: '' }
  }
};

/**
 * Asegura las columnas EN y rellena las traducciones de fabrica (solo celdas
 * vacias). La usan tanto crearBaseDeDatos() como migrarV6().
 */
function aplicarTraduccionesEN(ssOpt) {
  var ss = ssOpt || _ss();
  var serv = ss.getSheetByName('Servicios');
  _asegurarColumna(serv, 'NombreEN', '');
  _asegurarColumna(serv, 'DescripcionEN', '');
  _rellenarEN(serv, TRADUCCIONES.servicios, [['NombreEN', 'nombre'], ['DescripcionEN', 'desc']]);

  var cat = ss.getSheetByName('CategoriasProducto');
  _asegurarColumna(cat, 'NombreEN', '');
  _rellenarEN(cat, TRADUCCIONES.categorias, [['NombreEN', 'nombre']]);

  var prod = ss.getSheetByName('Productos');
  _asegurarColumna(prod, 'NombreEN', '');
  _asegurarColumna(prod, 'DescripcionEN', '');
  _rellenarEN(prod, TRADUCCIONES.productos, [['NombreEN', 'nombre'], ['DescripcionEN', 'desc']]);

  _invalidarCaches('Servicios');
  _invalidarCaches('CategoriasProducto');
  _invalidarCaches('Productos');
}

/**
 * Rellena columnas EN de una hoja a partir de un mapa {ID: {campo: valor}}.
 * Solo escribe si la celda destino esta vacia (no pisa ediciones manuales).
 * Escritura por lote (una llamada por columna) para no agotar el tiempo.
 * @param {Sheet} hoja
 * @param {Object} mapa   { ID: {nombre:'', desc:''} }
 * @param {Array} specs   [[nombreColumna, campoEnMapa], ...]
 */
function _rellenarEN(hoja, mapa, specs) {
  var ultima = hoja.getLastRow();
  if (ultima < 2) return;
  var idCol = _indiceColumna(hoja, 'ID');
  var ids = hoja.getRange(2, idCol + 1, ultima - 1, 1).getValues();
  specs.forEach(function (spec) {
    var colIdx = _indiceColumna(hoja, spec[0]);
    if (colIdx === -1) return;
    var rango = hoja.getRange(2, colIdx + 1, ultima - 1, 1);
    var actuales = rango.getValues();
    var cambio = false;
    for (var i = 0; i < ids.length; i++) {
      var entry = mapa[ids[i][0]];
      // 'desc' puede ser cadena vacia intencional: se aplica igual para "fijar" vacio.
      if (entry && entry[spec[1]] !== undefined && !String(actuales[i][0]).trim()) {
        actuales[i][0] = entry[spec[1]];
        cambio = true;
      }
    }
    if (cambio) rango.setValues(actuales);
  });
}

/**
 * MIGRACION V4. Agrega la columna Color a Servicios (si falta) y asigna un
 * color distintivo a cada servicio base. Ejecutala UNA vez desde el editor.
 */
function migrarColoresServicios() {
  var ss = _ss();
  var serv = ss.getSheetByName('Servicios');
  _asegurarColumna(serv, 'Color', '');
  var col = _indiceColumna(serv, 'Color') + 1;
  // Colores distintos entre si, sin dorados ni mostazas (estilo limpio).
  var colores = {
    S001: '#6E7F3E', // Desayuno Buffet - oliva
    S002: '#A6524B', // Desayuno Habitacion - ladrillo
    S003: '#9C4A63', // Almuerzo/Cena - vino
    S004: '#556070', // Cena (inactiva) - pizarra
    S005: '#2E8A86', // Tinaja - teal
    S006: '#3E6FA3', // Bicicletas - azul
    S007: '#7E4EA3'  // Masajes - purpura
  };
  var datos = serv.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var id = datos[i][0];
    // Reescribe los servicios base (aunque ya tuvieran color) para dejarlos limpios.
    if (colores[id]) serv.getRange(i + 1, col).setValue(colores[id]);
  }
  _invalidarCaches('Servicios');
  Logger.log('Migracion V4 lista: colores distintivos (sin dorados) asignados a los servicios.');
}

/**
 * MIGRACION. Agrega a la hoja Servicios las columnas Icono, Variantes y
 * UsoExclusivo (si faltan), y a Reservas la columna Variante. Rellena valores
 * por defecto para los servicios existentes y define masajes con dos variantes.
 * Ejecutala UNA vez desde el editor.
 */
function migrarServiciosVariantes() {
  var ss = _ss();

  // --- Hoja Servicios: nuevas columnas ---
  var serv = ss.getSheetByName('Servicios');
  _asegurarColumna(serv, 'Icono', '');
  _asegurarColumna(serv, 'Variantes', '');
  _asegurarColumna(serv, 'UsoExclusivo', 'FALSE');

  var colIcono = _indiceColumna(serv, 'Icono') + 1;
  var colVar = _indiceColumna(serv, 'Variantes') + 1;
  var colExcl = _indiceColumna(serv, 'UsoExclusivo') + 1;

  // Iconos y exclusividad por ID para los servicios base.
  var iconos = {
    S001: 'fa-bowl-food', S002: 'fa-mug-saucer', S003: 'fa-utensils',
    S004: 'fa-wine-glass', S005: 'fa-hot-tub-person', S006: 'fa-bicycle', S007: 'fa-spa'
  };
  var exclusivos = { S005: true, S007: true };

  var datos = serv.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var id = datos[i][0];
    // Icono: solo si esta vacio (no pisar personalizaciones).
    if (iconos[id] && !datos[i][colIcono - 1]) serv.getRange(i + 1, colIcono).setValue(iconos[id]);
    // Exclusividad segun mapa (default FALSE ya quedo por _asegurarColumna).
    if (exclusivos[id]) serv.getRange(i + 1, colExcl).setValue('TRUE');
    // Masajes: dos variantes.
    if (id === 'S007' && !datos[i][colVar - 1]) {
      serv.getRange(i + 1, colVar).setNumberFormat('@').setValue('Relajacion:60000|Descontracturante:65000');
      serv.getRange(i + 1, _indiceColumna(serv, 'Descripcion') + 1).setValue('Sesion de masaje');
    }
  }

  // --- Hoja Reservas: columna Variante ---
  _asegurarColumna(ss.getSheetByName('Reservas'), 'Variante', '');

  _invalidarCaches('Servicios');
  Logger.log('Migracion de variantes lista: columnas Icono/Variantes/UsoExclusivo en Servicios, Variante en Reservas. Masajes con 2 variantes.');
}

/**
 * Agrega una columna al final de una hoja si aun no existe, con un valor por
 * defecto para las filas de datos existentes.
 * @param {Sheet} hoja
 * @param {string} nombre
 * @param {*} valorDefecto
 */
function _asegurarColumna(hoja, nombre, valorDefecto) {
  var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  if (encabezados.indexOf(nombre) !== -1) return; // ya existe
  var col = hoja.getLastColumn() + 1;
  hoja.getRange(1, col).setValue(nombre);
  var nFilas = hoja.getLastRow() - 1;
  if (nFilas > 0 && valorDefecto !== '') {
    var valores = [];
    for (var i = 0; i < nFilas; i++) valores.push([valorDefecto]);
    hoja.getRange(2, col, nFilas, 1).setValues(valores);
  }
}

/**
 * Auto-reparacion: asegura que la hoja Productos tenga la columna
 * RequierePuntoCoccion (TRUE por defecto para no cambiar el comportamiento
 * actual). Se llama al inicio de las funciones que leen o escriben Productos
 * relacionadas con el prepedido/carta, para que la columna exista siempre
 * antes de leerla o escribirla, sin necesidad de una migracion manual aparte.
 */
function _asegurarColumnaPuntoCoccion() {
  _asegurarColumna(_hoja(HOJAS.PRODUCTOS), 'RequierePuntoCoccion', 'TRUE');
}

/**
 * Auto-reparacion: asegura que la hoja Usuarios tenga la columna Password
 * (vacia por defecto = sin contrasena, mantiene el comportamiento actual).
 * Se llama antes de leer/escribir Usuarios relacionado con el login.
 */
function _asegurarColumnaPasswordUsuarios() {
  _asegurarColumna(_hoja(HOJAS.USUARIOS), 'Password', '');
}

/**
 * Auto-reparacion: asegura que la hoja Reservas tenga la columna NotasInternas
 * (notas que SOLO ve el personal en el Centro de Operaciones; el huesped nunca
 * las ve). Vacia por defecto, no cambia el comportamiento de reservas viejas.
 */
function _asegurarColumnaNotasInternas() {
  _asegurarColumna(_hoja(HOJAS.RESERVAS), 'NotasInternas', '');
}

/**
 * Auto-reparacion: asegura que la hoja Servicios tenga la columna Visible
 * (si el servicio se muestra en la lista que ve el huesped). TRUE por
 * defecto para no ocultar de golpe servicios ya existentes; un servicio con
 * Activo=TRUE y Visible=FALSE sigue disponible para que recepcion reserve
 * desde el Centro de Operaciones, pero no aparece en la app del huesped
 * (ej: promos exclusivas de Instagram).
 */
function _asegurarColumnaVisibleServicios() {
  _asegurarColumna(_hoja(HOJAS.SERVICIOS), 'Visible', 'TRUE');
}

/**
 * Auto-reparacion: asegura que la hoja Servicios tenga la columna
 * AnticipoMinimoHoras (regla opcional, activable por servicio: "solo se
 * puede reservar con N horas de anticipacion"). 0 = sin restriccion (valor
 * por defecto para no afectar servicios existentes). La primera vez que se
 * crea la columna, deja Masajes (S007) en 12 horas, que es la regla que
 * pidio el hotel para ese servicio; el resto queda en 0 (sin restriccion)
 * y se puede activar desde el editor de servicios cuando haga falta.
 */
function _asegurarColumnaAnticipoServicios() {
  var hoja = _hoja(HOJAS.SERVICIOS);
  var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var yaExistia = encabezados.indexOf('AnticipoMinimoHoras') !== -1;
  _asegurarColumna(hoja, 'AnticipoMinimoHoras', 0);
  if (!yaExistia) {
    var datos = hoja.getDataRange().getValues();
    var col = _indiceColumna(hoja, 'AnticipoMinimoHoras') + 1;
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] === 'S007') hoja.getRange(i + 1, col).setValue(12);
    }
  }
}

/**
 * Auto-reparacion: asegura que la hoja Servicios tenga la columna
 * PermiteParticipantes (activa por servicio la opcion de ir sumando una lista
 * de personas a la reserva, ej. una salida de trekking). FALSE por defecto:
 * solo los servicios donde el hotel lo active mostraran el boton de
 * participantes. No afecta a los servicios existentes.
 */
function _asegurarColumnaParticipantesServicios() {
  _asegurarColumna(_hoja(HOJAS.SERVICIOS), 'PermiteParticipantes', 'FALSE');
}

/**
 * Auto-reparacion: asegura la columna AvisoReserva en Servicios. Es un texto
 * opcional (ej. politica de cancelacion) que, si esta lleno, se le muestra al
 * huesped como pop-up "De acuerdo" antes de confirmar la reserva de ese
 * servicio. Vacio por defecto = sin aviso. Lo usan Masajes/Tinaja, etc.
 */
function _asegurarColumnaAvisoServicios() {
  var hoja = _hoja(HOJAS.SERVICIOS);
  _asegurarColumna(hoja, 'AvisoReserva', '');
  _asegurarColumna(hoja, 'AvisoReservaEN', '');
}

/**
 * Escribe un valor en una columna por nombre, solo si la columna existe.
 * Evita romper hojas que aun no tengan la columna bilingue.
 */
function _setSiExisteCol(hoja, fila, nombreColumna, valor) {
  var idx = _indiceColumna(hoja, nombreColumna);
  if (idx !== -1) hoja.getRange(fila, idx + 1).setValue(valor);
}

/**
 * MIGRACION. Reescribe los horarios de Servicios y Configuracion como TEXTO
 * "HH:MM", porque Google Sheets los convirtio a valores de hora (fechas) y eso
 * rompe el motor de horarios. Ejecutala UNA vez desde el editor.
 */
function repararHorarios() {
  _normalizarHorariosComoTexto(_ss());
  Logger.log('Horarios reparados como texto. Revisa la hoja Servicios (columnas I y J).');
}

/**
 * Fuerza formato de texto y reescribe los horarios correctos en un Spreadsheet.
 * Se usa tanto en la migracion como al final de crearBaseDeDatos().
 * @param {Spreadsheet} ss
 */
function _normalizarHorariosComoTexto(ss) {
  // Servicios: columnas HorarioInicio (9) y HorarioFin (10).
  var serv = ss.getSheetByName('Servicios');
  if (serv) {
    var horariosServicios = {
      'S001': ['08:30', '10:30'], 'S002': ['08:30', '10:30'], 'S003': ['13:00', '22:00'],
      'S004': ['19:00', '22:00'], 'S005': ['10:00', '22:00'], 'S006': ['10:00', '18:00'],
      'S007': ['10:00', '20:00']
    };
    var datosServ = serv.getDataRange().getValues();
    for (var i = 1; i < datosServ.length; i++) {
      var idServ = datosServ[i][0];
      if (horariosServicios[idServ]) {
        serv.getRange(i + 1, 9).setNumberFormat('@').setValue(horariosServicios[idServ][0]);
        serv.getRange(i + 1, 10).setNumberFormat('@').setValue(horariosServicios[idServ][1]);
      }
    }
  }

  // Configuracion: los valores de horario tambien como texto.
  var conf = ss.getSheetByName('Configuracion');
  if (conf) {
    var horariosConfig = {
      'DESAYUNO_HORARIO_INICIO': '08:30', 'DESAYUNO_HORARIO_FIN': '10:30',
      'ALMUERZO_HORARIO_INICIO': '13:00', 'ALMUERZO_HORARIO_FIN': '22:00',
      'CENA_HORARIO_INICIO': '19:00', 'CENA_HORARIO_FIN': '22:00',
      'TINAJA_HORARIO_INICIO': '10:00', 'TINAJA_HORARIO_FIN': '22:00'
    };
    var datosConf = conf.getDataRange().getValues();
    for (var j = 1; j < datosConf.length; j++) {
      var clave = datosConf[j][0];
      if (horariosConfig[clave]) {
        conf.getRange(j + 1, 2).setNumberFormat('@').setValue(horariosConfig[clave]);
      }
    }
  }
}

// ---------------------------------------------------------------------------
// CONSTANTES DE NOMBRES DE HOJAS (unica fuente de verdad)
// ---------------------------------------------------------------------------
var HOJAS = {
  CONFIGURACION: 'Configuracion',
  HABITACIONES: 'Habitaciones',
  SERVICIOS: 'Servicios',
  CATEGORIAS: 'CategoriasProducto',
  PRODUCTOS: 'Productos',
  RESERVAS: 'Reservas',
  PEDIDOS: 'Pedidos',
  DETALLE_PEDIDOS: 'DetallePedidos',
  USUARIOS: 'Usuarios',
  EVENTOS_BLOQUEOS: 'EventosBloqueos',
  PARTICIPANTES: 'Participantes',
  NOTIFICACIONES: 'Notificaciones',
  DISPONIBILIDAD_PERSONAL: 'DisponibilidadPersonal',
  LOG: 'LogActividad',
  HISTORIAL: 'HistorialPedidos'
};

// Estados permitidos del sistema (unica fuente de verdad).
var ESTADOS = {
  DISPONIBLE: 'Disponible',
  SOLICITADA: 'Solicitada',
  PENDIENTE: 'Pendiente aprobacion',
  CONFIRMADA: 'Confirmada',
  EN_CURSO: 'En curso',
  FINALIZADA: 'Finalizada',
  CANCELADA_HUESPED: 'Cancelada huesped',
  CANCELADA_HOTEL: 'Cancelada hotel',
  BLOQUEADA: 'Bloqueada',
  NO_ASISTIO: 'No asistio'
};

// ===========================================================================
// 6.1. FUNCIONES DE RENDERIZADO
// ===========================================================================

/**
 * Punto de entrada web. Si viene ?hab=XXX renderiza vista huesped inyectando
 * la habitacion validada. Si no, renderiza el selector de rol/acceso.
 * @param {Object} e Evento de peticion GET.
 * @return {HtmlOutput}
 */
function doGet(e) {
  var plantilla = HtmlService.createTemplateFromFile('Index');

  // Valor por defecto: sin habitacion (mostrara selector de acceso).
  plantilla.HABITACION_QR = '';

  if (e && e.parameter && e.parameter.hab) {
    var numero = String(e.parameter.hab).trim();
    // Validacion server-side: la habitacion debe existir.
    if (_habitacionExiste(numero)) {
      plantilla.HABITACION_QR = numero;
    }
  }

  return plantilla.evaluate()
    .setTitle(_obtenerConfigValor('HOTEL_NOMBRE') || 'Cascadas Concierge')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper para incluir archivos HTML parciales (si se usaran).
 * @param {string} filename
 * @return {string} Contenido evaluado.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===========================================================================
// UTILIDADES INTERNAS DE ACCESO A HOJAS
// ===========================================================================

/** Cache del Spreadsheet durante la ejecucion (abrirlo es costoso). */
var _ssCache = null;

/** Devuelve el Spreadsheet activo segun el ID guardado por Setup.gs. */
function _ss() {
  if (_ssCache) return _ssCache;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('No se encontro SPREADSHEET_ID. Ejecuta crearBaseDeDatos() primero.');
  }
  _ssCache = SpreadsheetApp.openById(id);
  return _ssCache;
}

/**
 * Convierte un valor de hora (texto o Date de Sheets) a texto "HH:MM".
 * IMPORTANTE: google.script.run NO puede transferir objetos Date al navegador;
 * ademas Sheets convierte "13:00" en hora automaticamente. Todo lo que salga
 * al cliente debe pasar por aqui.
 */
function _horaATexto(valor) {
  if (valor === null || valor === undefined || valor === '') return '';
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, _ss().getSpreadsheetTimeZone(), 'HH:mm');
  }
  return String(valor);
}

/** Convierte un Date a texto "yyyy-MM-dd HH:mm:ss" transferible al cliente. */
function _fechaHoraTexto(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, _ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return (valor === null || valor === undefined) ? '' : String(valor);
}

/** Devuelve una hoja por nombre. */
function _hoja(nombre) {
  var hoja = _ss().getSheetByName(nombre);
  if (!hoja) throw new Error('Hoja no encontrada: ' + nombre);
  return hoja;
}

/**
 * Cache de lecturas dentro de una misma ejecucion. Solo para hojas "estaticas"
 * (catalogos que no cambian durante una peticion): evita releer la planilla
 * varias veces por llamada, que es lo mas lento de Apps Script.
 */
var _lecturaCache = {};
var HOJAS_CACHEABLES = {
  'Configuracion': true, 'Servicios': true, 'CategoriasProducto': true,
  'Productos': true, 'Habitaciones': true, 'Usuarios': true
};

/**
 * Lee una hoja completa como array de objetos {columna: valor}.
 * @param {string} nombreHoja
 * @return {Array<Object>}
 */
function _leerHojaComoObjetos(nombreHoja) {
  if (HOJAS_CACHEABLES[nombreHoja] && _lecturaCache[nombreHoja]) {
    return _lecturaCache[nombreHoja];
  }
  var hoja = _hoja(nombreHoja);
  var datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return [];
  var encabezados = datos[0];
  var objetos = [];
  for (var i = 1; i < datos.length; i++) {
    var obj = {};
    for (var c = 0; c < encabezados.length; c++) {
      obj[encabezados[c]] = datos[i][c];
    }
    obj._fila = i + 1; // numero de fila real en la hoja (1-indexed)
    objetos.push(obj);
  }
  if (HOJAS_CACHEABLES[nombreHoja]) _lecturaCache[nombreHoja] = objetos;
  return objetos;
}

// ---------------------------------------------------------------------------
// CACHE ENTRE PETICIONES (CacheService)
// Los catalogos (config, servicios, carta) casi no cambian: se sirven desde
// cache por unos minutos y se invalidan al editar productos/categorias/config.
// ---------------------------------------------------------------------------
var CACHE_TTL_SEGUNDOS = 120;

// La hoja Notificaciones crece con cada reserva/pedido y el polling la lee
// entera cada 20s. Se purgan las mas antiguas para que la lectura siga siendo
// rapida sin importar cuantos meses lleve operando la app.
var DIAS_RETENER_NOTIFICACIONES = 30;

/** Lee un objeto JSON del cache de script (o null). */
function _cacheGet(clave) {
  try {
    var v = CacheService.getScriptCache().get(clave);
    return v ? JSON.parse(v) : null;
  } catch (e) { return null; }
}

/** Guarda un objeto JSON en el cache de script. */
function _cachePut(clave, obj, ttl) {
  try {
    CacheService.getScriptCache().put(clave, JSON.stringify(obj), ttl || CACHE_TTL_SEGUNDOS);
  } catch (e) { /* si el objeto es muy grande, simplemente no se cachea */ }
}

/** Invalida los caches derivados y la lectura en memoria de una hoja. */
function _invalidarCaches(nombreHoja) {
  if (nombreHoja) _lecturaCache[nombreHoja] = null;
  try {
    CacheService.getScriptCache().removeAll(['datosIniciales', 'cartaCompleta']);
  } catch (e) { /* nunca romper el flujo por el cache */ }
}

/** Convierte un valor de celda a booleano real (soporta "TRUE"/true/1). */
function _aBooleano(valor) {
  if (valor === true) return true;
  if (typeof valor === 'string') return valor.toUpperCase() === 'TRUE';
  if (typeof valor === 'number') return valor === 1;
  return false;
}

/** Devuelve el indice (0-based) de una columna por su encabezado. */
function _indiceColumna(hoja, nombreColumna) {
  var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  return encabezados.indexOf(nombreColumna);
}

// ===========================================================================
// 6.2. FUNCIONES DE DATOS
// ===========================================================================

/**
 * Devuelve en UNA sola llamada todo lo necesario para arrancar la app:
 * configuracion, servicios activos y categorias. Cada llamada de
 * google.script.run cuesta ~1-2s, asi que agruparlas acelera mucho la carga.
 * Se sirve desde CacheService cuando es posible.
 * @return {Object} {config, servicios, categorias, categoriasTodas}
 */
function obtenerDatosIniciales() {
  var cacheado = _cacheGet('datosIniciales');
  if (cacheado) return cacheado;
  var datos = {
    config: obtenerConfiguracionCompleta(),
    servicios: obtenerServiciosActivos(),
    categorias: obtenerCategoriasMenu(),
    categoriasTodas: obtenerCategoriasTodas()
  };
  _cachePut('datosIniciales', datos);
  return datos;
}

/**
 * Devuelve la carta completa (categorias visibles + productos agrupados por
 * categoria) en UNA sola llamada, cacheada. Evita una llamada por categoria.
 * @return {Object} {categorias, productosPorCategoria}
 */
function obtenerCartaCompleta() {
  var cacheado = _cacheGet('cartaCompleta');
  if (cacheado) return cacheado;
  _asegurarColumnaPuntoCoccion();
  var categorias = obtenerCategoriasMenu();
  var productos = _leerHojaComoObjetos(HOJAS.PRODUCTOS)
    .filter(function (p) { return _aBooleano(p.Disponible) && _aBooleano(p.Visible); })
    .map(_normalizarProducto)
    .sort(function (a, b) { return a.Orden - b.Orden; });
  var porCategoria = {};
  productos.forEach(function (p) {
    if (!porCategoria[p.CategoriaID]) porCategoria[p.CategoriaID] = [];
    porCategoria[p.CategoriaID].push(p);
  });
  var datos = { categorias: categorias, productosPorCategoria: porCategoria };
  _cachePut('cartaCompleta', datos);
  return datos;
}

/**
 * Devuelve toda la configuracion como objeto {clave: valor}.
 * @return {Object}
 */
function obtenerConfiguracionCompleta() {
  _asegurarClavesPush(); // crea las claves de avisos al telefono si faltan
  var filas = _leerHojaComoObjetos(HOJAS.CONFIGURACION);
  var config = {};
  filas.forEach(function (f) {
    // Los Date no son transferibles al navegador: se convierten a "HH:MM".
    config[f.Clave] = (f.Valor instanceof Date) ? _horaATexto(f.Valor) : f.Valor;
  });
  return config;
}

/** Helper interno: valor de una clave de configuracion. */
function _obtenerConfigValor(clave) {
  var filas = _leerHojaComoObjetos(HOJAS.CONFIGURACION);
  for (var i = 0; i < filas.length; i++) {
    if (filas[i].Clave === clave) return filas[i].Valor;
  }
  return null;
}

/**
 * True si el pedido a la habitacion esta habilitado en Configuracion.
 * Por defecto TRUE (si la clave no existe todavia).
 */
function _pedidoHabitacionActivo() {
  var v = _obtenerConfigValor('PEDIDO_HABITACION_ACTIVO');
  if (v === null || v === undefined || v === '') return true;
  return String(v).toUpperCase() !== 'FALSE';
}

/**
 * True si el restaurant esta aceptando pedidos anticipados (prepedidos).
 * Por defecto TRUE (si la clave no existe todavia).
 */
function _prepedidoActivo() {
  var v = _obtenerConfigValor('PREPEDIDO_ACTIVO');
  if (v === null || v === undefined || v === '') return true;
  return String(v).toUpperCase() !== 'FALSE';
}

/** Recargo configurable que se suma al pedido entregado a la habitacion. */
function _recargoHabitacion() {
  return Number(_obtenerConfigValor('PEDIDO_HABITACION_RECARGO')) || 0;
}

/**
 * Devuelve los servicios activos (Activo=TRUE).
 * @return {Array<Object>}
 */
function obtenerServiciosActivos() {
  _asegurarColumnaVisibleServicios();
  _asegurarColumnaAnticipoServicios();
  _asegurarColumnaParticipantesServicios();
  _asegurarColumnaAvisoServicios();
  return _leerHojaComoObjetos(HOJAS.SERVICIOS).filter(function (s) {
    return _aBooleano(s.Activo);
  }).map(_normalizarServicio);
}

/** Normaliza un servicio a tipos utiles para el frontend. */
function _normalizarServicio(s) {
  return {
    ID: s.ID,
    Nombre: s.Nombre,
    NombreEN: s.NombreEN ? String(s.NombreEN) : '',
    Categoria: s.Categoria,
    DuracionMinutos: Number(s.DuracionMinutos) || 0,
    Capacidad: Number(s.Capacidad) || 0,
    CostoBase: Number(s.CostoBase) || 0,
    RequiereAprobacion: _aBooleano(s.RequiereAprobacion),
    Activo: _aBooleano(s.Activo),
    HorarioInicio: _horaATexto(s.HorarioInicio),
    HorarioFin: _horaATexto(s.HorarioFin),
    Descripcion: s.Descripcion,
    PermitePrepedido: _aBooleano(s.PermitePrepedido),
    DescripcionEN: s.DescripcionEN ? String(s.DescripcionEN) : '',
    EsIncluible: _aBooleano(s.EsIncluible),
    // Visible para el huesped en la lista de servicios. Columna nueva: si
    // esta vacia (instalaciones viejas antes de la auto-reparacion) se
    // asume TRUE para no ocultar nada de golpe.
    Visible: s.Visible === '' || s.Visible === undefined ? true : _aBooleano(s.Visible),
    Icono: s.Icono ? String(s.Icono) : '',
    Variantes: _parsearVariantes(s.Variantes),
    UsoExclusivo: _aBooleano(s.UsoExclusivo),
    Color: s.Color ? String(s.Color) : '',
    AnticipoMinimoHoras: Number(s.AnticipoMinimoHoras) || 0,
    PermiteParticipantes: _aBooleano(s.PermiteParticipantes),
    AvisoReserva: s.AvisoReserva ? String(s.AvisoReserva) : '',
    AvisoReservaEN: s.AvisoReservaEN ? String(s.AvisoReservaEN) : ''
  };
}

/**
 * Convierte el texto de variantes en un array [{nombre, precio,
 * duracionMinutos, visible}]. Formato "Nombre:precio:duracion:visible",
 * con duracion y visible opcionales (compatible con el formato viejo
 * "Nombre:precio" que ya existe en servicios como Masajes):
 *  - duracion vacia o 0 = usa la duracion normal del servicio.
 *  - visible vacio = TRUE (no oculta nada de golpe en datos viejos).
 * Devuelve [] si no hay variantes.
 * @param {string} texto
 * @return {Array<Object>}
 */
function _parsearVariantes(texto) {
  if (!texto) return [];
  return String(texto).split('|').map(function (par) {
    var partes = par.split(':');
    if (partes.length < 2) return null;
    var nombre = partes[0].trim();
    if (!nombre) return null;
    var precio = Number(partes[1]) || 0;
    var duracionMinutos = partes[2] !== undefined && partes[2] !== '' ? (Number(partes[2]) || 0) : 0;
    var visible = partes[3] !== undefined && partes[3] !== '' ? String(partes[3]).toUpperCase() !== 'FALSE' : true;
    return { nombre: nombre, precio: precio, duracionMinutos: duracionMinutos, visible: visible };
  }).filter(function (v) { return v; });
}

/** Convierte un array de variantes de vuelta a texto "Nombre:precio:duracion:visible|...". */
function _serializarVariantes(variantes) {
  if (!variantes || !variantes.length) return '';
  return variantes.map(function (v) {
    var nombre = String(v.nombre).replace(/[|:]/g, ' ').trim();
    var precio = Number(v.precio) || 0;
    var duracion = Number(v.duracionMinutos) || 0;
    var visible = v.visible === false ? 'FALSE' : 'TRUE';
    return nombre + ':' + precio + ':' + duracion + ':' + visible;
  }).join('|');
}

/** Devuelve un servicio normalizado por ID (o null). */
function _obtenerServicio(servicioID) {
  _asegurarColumnaVisibleServicios();
  _asegurarColumnaAnticipoServicios();
  _asegurarColumnaParticipantesServicios();
  _asegurarColumnaAvisoServicios();
  var filas = _leerHojaComoObjetos(HOJAS.SERVICIOS);
  for (var i = 0; i < filas.length; i++) {
    if (filas[i].ID === servicioID) return _normalizarServicio(filas[i]);
  }
  return null;
}

/**
 * Devuelve todas las habitaciones.
 * @return {Array<Object>}
 */
function obtenerHabitaciones() {
  return _leerHojaComoObjetos(HOJAS.HABITACIONES).map(function (h) {
    return {
      Numero: String(h.Numero),
      Estado: h.Estado,
      Tipo: h.Tipo,
      Capacidad: Number(h.Capacidad) || 0,
      Notas: h.Notas
    };
  });
}

/** Comprueba si una habitacion existe. */
function _habitacionExiste(numero) {
  var habs = _leerHojaComoObjetos(HOJAS.HABITACIONES);
  for (var i = 0; i < habs.length; i++) {
    if (String(habs[i].Numero) === String(numero)) return true;
  }
  return false;
}

/**
 * Devuelve el link del grupo de WhatsApp de una habitacion (o '' si no tiene).
 * @param {string} numero
 * @return {string}
 */
function obtenerWhatsappHabitacion(numero) {
  var habs = _leerHojaComoObjetos(HOJAS.HABITACIONES);
  for (var i = 0; i < habs.length; i++) {
    if (String(habs[i].Numero) === String(numero)) {
      return habs[i].WhatsAppLink ? String(habs[i].WhatsAppLink) : '';
    }
  }
  return '';
}

/**
 * Devuelve la URL del Spreadsheet (base de datos) para abrirlo desde la app.
 * @return {string}
 */
function obtenerUrlSheet() {
  try { return _ss().getUrl(); } catch (e) { return ''; }
}

/**
 * Asegura que existan en Configuracion las claves del portal de reservas:
 * PORTAL_BASE_URL (URL /exec del web app, se puede pegar a mano) y PORTAL_MENSAJE
 * (texto opcional que acompana el enlace).
 */
function _asegurarConfigPortal() {
  var hoja = _hoja(HOJAS.CONFIGURACION);
  var claves = hoja.getRange(1, 1, hoja.getLastRow(), 1).getValues().map(function (f) { return f[0]; });
  var cambios = false;
  if (claves.indexOf('PORTAL_BASE_URL') === -1) {
    hoja.appendRow(['PORTAL_BASE_URL', '', 'URL base del portal (termina en /exec). Si se deja vacio se intenta detectar sola.']);
    cambios = true;
  }
  if (claves.indexOf('PORTAL_MENSAJE') === -1) {
    hoja.appendRow(['PORTAL_MENSAJE',
      'Este es nuestro portal de reservas del hotel. Desde aqui puede reservar desayuno, almuerzo/cena, tinaja, bicicletas y masajes durante su estadia:',
      'Mensaje opcional que acompana el enlace de la habitacion']);
    cambios = true;
  }
  if (cambios) _invalidarCaches('Configuracion');
}

/**
 * Devuelve la URL base del portal: la fijada a mano en Configuracion
 * (PORTAL_BASE_URL) o, si esta vacia, la URL del web app desplegado.
 * @return {string}
 */
function obtenerPortalBaseUrl() {
  var manual = _obtenerConfigValor('PORTAL_BASE_URL');
  if (manual && String(manual).trim()) return String(manual).trim();
  try {
    var u = ScriptApp.getService().getUrl();
    if (u) return u;
  } catch (e) {}
  return '';
}

/**
 * Devuelve todo lo necesario para la seccion "Enlaces por habitacion" de la
 * pestana de configuracion: URL base, mensaje opcional y lista de habitaciones.
 * @return {Object} {base, mensaje, habitaciones:[Numero]}
 */
function obtenerDatosPortal() {
  _asegurarConfigPortal();
  return {
    base: obtenerPortalBaseUrl(),
    mensaje: _obtenerConfigValor('PORTAL_MENSAJE') || '',
    habitaciones: obtenerHabitaciones().map(function (h) { return String(h.Numero); })
  };
}

/**
 * Devuelve las categorias de menu visibles, ordenadas.
 * @return {Array<Object>}
 */
function obtenerCategoriasMenu() {
  return _leerHojaComoObjetos(HOJAS.CATEGORIAS)
    .filter(function (c) { return _aBooleano(c.Visible); })
    .map(_normalizarCategoria)
    .sort(function (a, b) { return a.Orden - b.Orden; });
}

/** Normaliza una categoria. */
function _normalizarCategoria(c) {
  return {
    ID: c.ID,
    Nombre: c.Nombre,
    NombreEN: c.NombreEN ? String(c.NombreEN) : '',
    Orden: Number(c.Orden) || 0,
    Visible: _aBooleano(c.Visible),
    IconoFontAwesome: c.IconoFontAwesome,
    Color: c.Color
  };
}

/**
 * Devuelve todas las categorias (incluidas no visibles) para gestion.
 * @return {Array<Object>}
 */
function obtenerCategoriasTodas() {
  return _leerHojaComoObjetos(HOJAS.CATEGORIAS)
    .map(_normalizarCategoria)
    .sort(function (a, b) { return a.Orden - b.Orden; });
}

/**
 * Devuelve productos de una categoria disponibles y visibles, ordenados.
 * @param {string} categoriaID
 * @return {Array<Object>}
 */
function obtenerProductosPorCategoria(categoriaID) {
  return _leerHojaComoObjetos(HOJAS.PRODUCTOS)
    .filter(function (p) {
      return p.CategoriaID === categoriaID && _aBooleano(p.Disponible) && _aBooleano(p.Visible);
    })
    .map(_normalizarProducto)
    .sort(function (a, b) { return a.Orden - b.Orden; });
}

/** Normaliza un producto (sin ImagenURL). */
function _normalizarProducto(p) {
  return {
    ID: p.ID,
    CategoriaID: p.CategoriaID,
    Nombre: p.Nombre,
    NombreEN: p.NombreEN ? String(p.NombreEN) : '',
    Descripcion: p.Descripcion,
    DescripcionEN: p.DescripcionEN ? String(p.DescripcionEN) : '',
    Precio: Number(p.Precio) || 0,
    Disponible: _aBooleano(p.Disponible),
    Visible: _aBooleano(p.Visible),
    Orden: Number(p.Orden) || 0,
    Etiquetas: p.Etiquetas ? String(p.Etiquetas) : '',
    TiempoPreparacionMin: Number(p.TiempoPreparacionMin) || 0,
    EsMenuDelDia: _aBooleano(p.EsMenuDelDia),
    // Por defecto TRUE (columna nueva, self-heal via _asegurarColumnaPuntoCoccion).
    // Solo tiene efecto en la carta para platos de Carnes; en el resto no se usa.
    RequierePuntoCoccion: p.RequierePuntoCoccion === '' || p.RequierePuntoCoccion === undefined
      ? true : _aBooleano(p.RequierePuntoCoccion),
    FechaModificacion: _fechaHoraTexto(p.FechaModificacion),
    ModificadoPor: p.ModificadoPor
  };
}

/**
 * Devuelve un producto completo mas su categoria.
 * @param {string} productoID
 * @return {Object|null}
 */
function obtenerProductoCompleto(productoID) {
  var productos = _leerHojaComoObjetos(HOJAS.PRODUCTOS);
  for (var i = 0; i < productos.length; i++) {
    if (productos[i].ID === productoID) {
      var prod = _normalizarProducto(productos[i]);
      prod.Categoria = _obtenerCategoriaPorID(prod.CategoriaID);
      return prod;
    }
  }
  return null;
}

/** Devuelve una categoria normalizada por ID. */
function _obtenerCategoriaPorID(categoriaID) {
  var cats = _leerHojaComoObjetos(HOJAS.CATEGORIAS);
  for (var i = 0; i < cats.length; i++) {
    if (cats[i].ID === categoriaID) return _normalizarCategoria(cats[i]);
  }
  return null;
}

// ===========================================================================
// UTILIDADES DE TIEMPO
// ===========================================================================

/** Convierte "HH:MM" a minutos desde medianoche. */
function _horaAMinutos(hora) {
  var texto = _horaATexto(hora); // tolera celdas convertidas a Date por Sheets
  if (!texto) return 0;
  var partes = texto.split(':');
  return (parseInt(partes[0], 10) || 0) * 60 + (parseInt(partes[1], 10) || 0);
}

/**
 * Convierte minutos desde medianoche a "HH:MM". Si da 1440 (24:00, ej. una
 * reserva de 22:00 + 2h) lo deja en 23:59: al escribir "24:00" en una celda,
 * Sheets lo interpreta como el dia siguiente a las 00:00 y luego, al leerlo
 * de vuelta, esa hora "00:00" hace que _finalizarReservasVencidas() la de
 * por terminada de inmediato (0 minutos siempre es "menor" que la hora
 * actual). 23:59 evita esa confusion sin cambiar la duracion real.
 */
function _minutosAHora(minutos) {
  if (minutos >= 1440) minutos = 1439;
  var h = Math.floor(minutos / 60);
  var m = minutos % 60;
  return _pad2(h) + ':' + _pad2(m);
}

/** Rellena a 2 digitos. */
function _pad2(n) {
  return (n < 10 ? '0' : '') + n;
}

/** Formatea una fecha (Date o string) a "YYYY-MM-DD". */
function _fechaISO(fecha) {
  if (fecha instanceof Date) {
    return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(fecha).substring(0, 10);
}

/** Devuelve true si la fecha "YYYY-MM-DD" es hoy o futura. */
function _esFechaFuturaOHoy(fechaISO) {
  var hoy = _fechaISO(new Date());
  return fechaISO >= hoy;
}

// ===========================================================================
// 6.3. MOTOR DE RESERVAS UNIFICADO
// ===========================================================================

/**
 * Consulta bloques horarios disponibles para un servicio/fecha/personas.
 * Considera capacidad, reservas existentes, bloqueos y duracion. Si el
 * servicio tiene variantes con duracion propia (ej. Tinaja "1 hora"), pasa
 * el nombre de la variante para que los bloques usen esa duracion en vez
 * de la del servicio; el chequeo de conflictos sigue siendo contra TODAS
 * las reservas del servicio (misma tina/recurso), sin importar la variante.
 * @param {string} servicioID
 * @param {string} fecha "YYYY-MM-DD"
 * @param {number} personas
 * @param {string} [variante]
 * @return {Object} {success, bloques:[{hora, disponibles, capacidad}], mensaje}
 */
function consultarDisponibilidad(servicioID, fecha, personas, variante) {
  try {
    var servicio = _obtenerServicio(servicioID);
    if (!servicio || !servicio.Activo) {
      return { success: false, bloques: [], mensaje: 'Servicio no disponible.' };
    }
    personas = Number(personas) || 1;
    fecha = _fechaISO(fecha);

    var inicio = _horaAMinutos(servicio.HorarioInicio);
    var fin = _horaAMinutos(servicio.HorarioFin);
    var duracion = _duracionServicio(servicio, variante) || 30;

    // Reservas activas del dia para este servicio.
    var reservas = _reservasActivasDelDia(servicioID, fecha);
    // Bloqueos que afectan a este servicio en la fecha.
    var bloqueos = _bloqueosDelDia(servicioID, fecha);

    // Paso entre horarios ofrecidos: cada media hora por defecto (configurable).
    var paso = Number(_obtenerConfigValor('RESERVA_PASO_MINUTOS')) || 30;
    if (paso > duracion) paso = duracion;
    var exclusivo = _esServicioExclusivo(servicio);

    var bloques = [];
    // Genera bloques desde inicio hasta que quepa la duracion completa.
    for (var t = inicio; t + duracion <= fin; t += paso) {
      var horaBloque = _minutosAHora(t);

      // Si el bloque esta dentro de un bloqueo, se omite.
      if (_bloqueSolapaBloqueos(t, t + duracion, bloqueos)) continue;

      // Servicios exclusivos (Tinaja/Masajes): una sola sesion a la vez,
      // cualquier reserva solapada bloquea el horario.
      if (exclusivo && _hayReservaSolapada(reservas, t, t + duracion)) continue;

      // Ocupacion: personas ya reservadas que solapan este bloque.
      var ocupadas = _personasOcupadasEnBloque(reservas, t, t + duracion);
      var disponibles = servicio.Capacidad - ocupadas;

      if (disponibles >= personas) {
        bloques.push({
          hora: horaBloque,
          disponibles: disponibles,
          capacidad: servicio.Capacidad
        });
      }
    }

    // Si la fecha es hoy, descarta bloques ya pasados.
    if (fecha === _fechaISO(new Date())) {
      var ahora = _minutosAhoraLocal();
      bloques = bloques.filter(function (b) {
        return _horaAMinutos(b.hora) > ahora;
      });
    }

    return { success: true, bloques: bloques, mensaje: '' };
  } catch (err) {
    return { success: false, bloques: [], mensaje: 'Error: ' + err.message };
  }
}

/** Minutos transcurridos hoy segun zona horaria del script. */
function _minutosAhoraLocal() {
  var ahora = new Date();
  var hhmm = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'HH:mm');
  return _horaAMinutos(hhmm);
}

/** Dias de diferencia entre dos fechas "YYYY-MM-DD" (hastaISO - desdeISO). */
function _diasEntreISO(desdeISO, hastaISO) {
  var d1 = new Date(desdeISO + 'T00:00:00Z');
  var d2 = new Date(hastaISO + 'T00:00:00Z');
  return Math.round((d2.getTime() - d1.getTime()) / 86400000);
}

/**
 * Horas (puede ser fraccion o negativo) entre ahora y una fecha+hora de
 * reserva, segun la zona horaria del script. Se usa para la regla de
 * "anticipo minimo" por servicio (ej. Masajes: 12 horas).
 */
function _horasHastaReserva(fechaISO, horaInicio) {
  var hoy = _fechaISO(new Date());
  var diffDias = _diasEntreISO(hoy, fechaISO);
  var minutosReserva = _horaAMinutos(horaInicio);
  var ahoraMin = _minutosAhoraLocal();
  return diffDias * 24 + (minutosReserva - ahoraMin) / 60;
}

/** Reservas activas (no canceladas) del dia para un servicio. */
function _reservasActivasDelDia(servicioID, fecha) {
  return _leerHojaComoObjetos(HOJAS.RESERVAS).filter(function (r) {
    return r.ServicioID === servicioID &&
      _fechaISO(r.Fecha) === fecha &&
      !_esEstadoCancelado(r.Estado);
  });
}

/** True si un estado es cancelado (huesped u hotel). */
function _esEstadoCancelado(estado) {
  return estado === ESTADOS.CANCELADA_HUESPED || estado === ESTADOS.CANCELADA_HOTEL;
}

/**
 * Servicios de uso exclusivo: una sola sesion a la vez (una reserva bloquea el
 * horario completo). Ahora es un flag explicito por servicio (UsoExclusivo);
 * como respaldo, Bienestar sigue siendo exclusivo si el flag no existe.
 */
function _esServicioExclusivo(servicio) {
  if (servicio.UsoExclusivo === true || servicio.UsoExclusivo === false) return servicio.UsoExclusivo;
  return servicio.Categoria === 'Bienestar';
}

/** True si alguna reserva activa solapa el intervalo [ini,fin). */
function _hayReservaSolapada(reservas, ini, fin) {
  for (var i = 0; i < reservas.length; i++) {
    var rIni = _horaAMinutos(reservas[i].HoraInicio);
    var rFin = _horaAMinutos(reservas[i].HoraFin);
    if (rIni < fin && rFin > ini) return true;
  }
  return false;
}

/** Suma de personas de reservas que solapan un intervalo [ini,fin). */
function _personasOcupadasEnBloque(reservas, ini, fin) {
  var total = 0;
  reservas.forEach(function (r) {
    var rIni = _horaAMinutos(r.HoraInicio);
    var rFin = _horaAMinutos(r.HoraFin);
    if (rIni < fin && rFin > ini) { // hay solapamiento
      total += Number(r.Personas) || 1;
    }
  });
  return total;
}

/**
 * Decide si un bloqueo/evento afecta a un servicio dado, segun su campo
 * ServicioID (que ahora admite varios alcances):
 *   - vacio          -> bloquea TODOS los servicios (compatibilidad + "todos").
 *   - 'NINGUNO'      -> es solo un evento, no bloquea ningun servicio.
 *   - 'S001,S003...' -> bloquea SOLO esos servicios (lista separada por comas).
 * @param {string} bServicioID Valor crudo de la columna ServicioID del bloqueo.
 * @param {string} servicioID  Servicio contra el que se evalua.
 * @return {boolean}
 */
function _bloqueoAplicaAServicio(bServicioID, servicioID) {
  var raw = bServicioID === null || bServicioID === undefined ? '' : String(bServicioID).trim();
  if (raw === '') return true;          // todos
  if (raw === 'NINGUNO') return false;  // evento que no bloquea nada
  return raw.split(',').map(function (s) { return s.trim(); }).indexOf(servicioID) !== -1;
}

/** Bloqueos que aplican a un servicio (o a todos) en una fecha. */
function _bloqueosDelDia(servicioID, fecha) {
  return _leerHojaComoObjetos(HOJAS.EVENTOS_BLOQUEOS).filter(function (b) {
    var aplicaServicio = _bloqueoAplicaAServicio(b.ServicioID, servicioID);
    var desde = _fechaISO(b.FechaInicio);
    var hasta = _fechaISO(b.FechaFin || b.FechaInicio);
    return aplicaServicio && fecha >= desde && fecha <= hasta;
  });
}

/** True si el intervalo solapa algun bloqueo horario. */
function _bloqueSolapaBloqueos(ini, fin, bloqueos) {
  for (var i = 0; i < bloqueos.length; i++) {
    var b = bloqueos[i];
    var bIni = b.HoraInicio ? _horaAMinutos(b.HoraInicio) : 0;
    var bFin = b.HoraFin ? _horaAMinutos(b.HoraFin) : 24 * 60;
    if (bIni < fin && bFin > ini) return true;
  }
  return false;
}

/**
 * Crea una reserva aplicando todas las validaciones server-side.
 * @param {Object} datos {habitacion, servicioID, fecha, horaInicio, personas, notas, solicitadoPor}
 * @return {Object} {success, id, estado, mensaje}
 */
function crearReserva(datos) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // evita condiciones de carrera en la capacidad

    // Validaciones basicas.
    if (!datos || !datos.habitacion || !datos.servicioID || !datos.fecha || !datos.horaInicio) {
      return { success: false, mensaje: 'Faltan datos obligatorios.' };
    }

    // Es personal validado? (cualquier rol interno puede agendar desde el panel).
    var esStaff = datos.emailStaff && _validarRolPermitido(datos.emailStaff, ['RECEPCION', 'ADMINISTRADOR', 'COCINA', 'RESTAURANT']);
    // Clientes externos (pasante / grupo) solo los agenda el personal; no exigen habitacion real.
    var esExterno = esStaff && (datos.tipoCliente === 'pasante' || datos.tipoCliente === 'grupo');

    if (!esExterno && !_habitacionExiste(datos.habitacion)) {
      return { success: false, mensaje: 'La habitacion no existe.' };
    }
    var servicio = _obtenerServicio(datos.servicioID);
    if (!servicio || !servicio.Activo) {
      return { success: false, mensaje: 'El servicio no existe o no esta activo.' };
    }

    // (El desayuno a la habitacion se habilita/deshabilita con el interruptor
    //  "Activo" del servicio S002 en Servicios y precios; ya se valido arriba.)

    // Variante: si el servicio tiene variantes, se exige una valida.
    var variante = '';
    if (servicio.Variantes.length) {
      var elegida = _buscarVariante(servicio.Variantes, datos.variante);
      if (!elegida) return { success: false, mensaje: 'Debes elegir una opcion del servicio.' };
      variante = elegida.nombre;
    }

    var fecha = _fechaISO(datos.fecha);
    if (!_esFechaFuturaOHoy(fecha)) {
      return { success: false, mensaje: 'La fecha debe ser hoy o futura.' };
    }

    // Regla: los huespedes solo reservan hasta N dias hacia adelante.
    // El personal (recepcion/admin) no tiene ese limite.
    if (!esStaff) {
      var maxDias = Number(_obtenerConfigValor('MAX_DIAS_RESERVA_HUESPED'));
      if (!maxDias && maxDias !== 0) maxDias = 3;
      var limite = new Date();
      limite.setDate(limite.getDate() + maxDias);
      if (fecha > _fechaISO(limite)) {
        return { success: false, mensaje: 'Solo puedes reservar hasta ' + maxDias + ' dias desde hoy. Para fechas mas lejanas, consulta en recepcion.' };
      }
    }

    // Regla configurable por servicio: reservar con un minimo de horas de
    // anticipacion (ej. Masajes = 12h). Se activa/desactiva por servicio
    // (AnticipoMinimoHoras = 0 significa sin restriccion). No aplica al
    // personal, que puede agendar en el momento desde el Centro de Operaciones.
    if (!esStaff && servicio.AnticipoMinimoHoras > 0) {
      var horasFaltantes = _horasHastaReserva(fecha, datos.horaInicio);
      if (horasFaltantes < servicio.AnticipoMinimoHoras) {
        return { success: false, mensaje: 'Este servicio requiere reservar con al menos ' + servicio.AnticipoMinimoHoras + ' horas de anticipacion.' };
      }
    }

    var personas = Number(datos.personas) || 1;
    var iniMin = _horaAMinutos(datos.horaInicio);

    // Un pedido a la habitacion es una entrega puntual, no ocupa una mesa: no
    // consume capacidad del restaurant ni dura el bloque de una mesa sentada.
    // Por eso, cuando la entrega es a la habitacion, solo se valida que la
    // hora este dentro del horario del restaurant; la capacidad y la duracion
    // de mesa no aplican.
    var esEntregaHabitacion = String(datos.entrega || '').toLowerCase() === 'habitacion';
    var finMin = esEntregaHabitacion ? iniMin : iniMin + _duracionServicio(servicio, variante);

    // Horario dentro del rango del servicio.
    var fueraDeRango = esEntregaHabitacion
      ? (iniMin < _horaAMinutos(servicio.HorarioInicio) || iniMin >= _horaAMinutos(servicio.HorarioFin))
      : (iniMin < _horaAMinutos(servicio.HorarioInicio) || finMin > _horaAMinutos(servicio.HorarioFin));
    if (fueraDeRango) {
      return { success: false, mensaje: 'El horario esta fuera del rango permitido.' };
    }

    // No debe caer dentro de un bloqueo (aplica tambien al pedido a la pieza:
    // un bloqueo significa que el restaurant no esta operando).
    var bloqueos = _bloqueosDelDia(datos.servicioID, fecha);
    if (_bloqueSolapaBloqueos(iniMin, Math.max(finMin, iniMin + 1), bloqueos)) {
      return { success: false, mensaje: 'El horario esta bloqueado por el hotel.' };
    }

    // Capacidad y exclusividad: solo para reservas que ocupan mesa. El pedido a
    // la habitacion se las salta.
    if (!esEntregaHabitacion) {
      var reservas = _reservasActivasDelDia(datos.servicioID, fecha);
      var ocupadas = _personasOcupadasEnBloque(reservas, iniMin, finMin);
      if (ocupadas + personas > servicio.Capacidad) {
        return { success: false, mensaje: 'No hay capacidad disponible en ese horario.' };
      }
      if (_esServicioExclusivo(servicio) && _hayReservaSolapada(reservas, iniMin, finMin)) {
        return { success: false, mensaje: 'Ese horario ya esta tomado.' };
      }
    }

    // Determina el estado inicial segun reglas de negocio.
    var estado = _estadoInicialSegunServicio(servicio);

    // Si la crea personal validado (recepcion en persona), queda confirmada
    // de inmediato, sin importar el servicio (incluso los que requieren
    // aprobacion): recepcion la esta creando presencialmente, ya esta OK.
    if (esStaff) {
      estado = ESTADOS.CONFIRMADA;
    }

    // Inserta la fila.
    _asegurarColumnaNotasInternas();
    var id = generarID();
    var horaFin = _minutosAHora(finMin);
    var hojaReservas = _hoja(HOJAS.RESERVAS);
    hojaReservas.appendRow([
      id, new Date(), String(datos.habitacion), datos.servicioID, fecha,
      datos.horaInicio, horaFin, personas, estado,
      datos.solicitadoPor || 'Huesped', datos.notas || '', '', '', 'FALSE'
    ]);
    var filaNueva = hojaReservas.getLastRow();
    // Guarda la variante elegida (columna agregada por migracion, al final).
    if (variante) {
      var colVariante = _indiceColumna(hojaReservas, 'Variante');
      if (colVariante !== -1) hojaReservas.getRange(filaNueva, colVariante + 1).setValue(variante);
    }
    // Nota interna (solo personal): la escribe recepcion, el huesped no la ve.
    if (datos.notasInternas) {
      var colNI = _indiceColumna(hojaReservas, 'NotasInternas');
      if (colNI !== -1) hojaReservas.getRange(filaNueva, colNI + 1).setValue(datos.notasInternas);
    }

    var nombreCompleto = servicio.Nombre + (variante ? ' (' + variante + ')' : '');
    registrarLog('Crear reserva', nombreCompleto + ' ' + fecha + ' ' + datos.horaInicio, datos.habitacion);

    // Genera notificacion para el staff. Mensaje estructurado (el cliente lo
    // muestra como pop-up grande): RES|habitacion|servicio|fecha|hora|estado
    var mensajeNotif = 'RES|' + datos.habitacion + '|' + nombreCompleto + '|' + fecha + '|' + datos.horaInicio + '|' + estado;
    _crearNotificacion('reserva', mensajeNotif, 'RECEPCION', datos.habitacion, datos.servicioID, fecha);

    return { success: true, id: id, estado: estado, mensaje: 'Reserva creada correctamente.' };
  } catch (err) {
    return { success: false, mensaje: 'Error: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/** Busca una variante por nombre (insensible a mayusculas). */
function _buscarVariante(variantes, nombre) {
  if (!nombre) return null;
  for (var i = 0; i < variantes.length; i++) {
    if (variantes[i].nombre.toLowerCase() === String(nombre).toLowerCase()) return variantes[i];
  }
  return null;
}

/**
 * Duracion efectiva en minutos para una reserva: la de la variante elegida
 * si tiene una propia (ej. Tinaja "1 hora - Promo" = 60), o si no la
 * duracion normal del servicio (ej. Tinaja normal = 120). Esto es lo que
 * hace que, aunque compartan el mismo servicio (mismo pool de reservas,
 * sin riesgo de doble reserva), la version corta libere el horario antes.
 * @param {Object} servicio Normalizado (de _obtenerServicio).
 * @param {string} nombreVariante
 * @return {number}
 */
function _duracionServicio(servicio, nombreVariante) {
  var variante = _buscarVariante(servicio.Variantes || [], nombreVariante);
  if (variante && variante.duracionMinutos > 0) return variante.duracionMinutos;
  return servicio.DuracionMinutos;
}

/**
 * Determina el estado inicial de una reserva.
 * Servicios que requieren aprobacion (Tinaja, Bicicletas, Masajes) nacen
 * "Pendiente aprobacion". El resto nace "Solicitada".
 */
function _estadoInicialSegunServicio(servicio) {
  if (servicio.RequiereAprobacion) return ESTADOS.PENDIENTE;
  return ESTADOS.SOLICITADA;
}

/**
 * Modifica fecha/hora/personas/notas de una reserva.
 * No permite modificar si esta En curso, Finalizada o Cancelada.
 * @param {string} id
 * @param {Object} datos {fecha, horaInicio, personas, notas}
 * @return {Object} {success, mensaje}
 */
function modificarReserva(id, datos) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var hoja = _hoja(HOJAS.RESERVAS);
    var reserva = _buscarReservaPorID(id);
    if (!reserva) return { success: false, mensaje: 'Reserva no encontrada.' };

    if ([ESTADOS.EN_CURSO, ESTADOS.FINALIZADA, ESTADOS.CANCELADA_HUESPED,
         ESTADOS.CANCELADA_HOTEL].indexOf(reserva.Estado) !== -1) {
      return { success: false, mensaje: 'No se puede modificar una reserva ' + reserva.Estado + '.' };
    }

    var servicio = _obtenerServicio(reserva.ServicioID);
    var fecha = datos.fecha ? _fechaISO(datos.fecha) : _fechaISO(reserva.Fecha);
    var horaInicio = datos.horaInicio || reserva.HoraInicio;
    var personas = datos.personas ? Number(datos.personas) : Number(reserva.Personas);

    if (!_esFechaFuturaOHoy(fecha)) {
      return { success: false, mensaje: 'La fecha debe ser hoy o futura.' };
    }

    var iniMin = _horaAMinutos(horaInicio);
    var finMin = iniMin + _duracionServicio(servicio, reserva.Variante);
    if (iniMin < _horaAMinutos(servicio.HorarioInicio) ||
        finMin > _horaAMinutos(servicio.HorarioFin)) {
      return { success: false, mensaje: 'El horario esta fuera del rango permitido.' };
    }

    // Revalida capacidad excluyendo esta misma reserva.
    var reservas = _reservasActivasDelDia(reserva.ServicioID, fecha).filter(function (r) {
      return r.ID !== id;
    });
    var ocupadas = _personasOcupadasEnBloque(reservas, iniMin, finMin);
    if (ocupadas + personas > servicio.Capacidad) {
      return { success: false, mensaje: 'No hay capacidad disponible en ese horario.' };
    }

    var fila = reserva._fila;
    hoja.getRange(fila, _indiceColumna(hoja, 'Fecha') + 1).setValue(fecha);
    hoja.getRange(fila, _indiceColumna(hoja, 'HoraInicio') + 1).setValue(horaInicio);
    hoja.getRange(fila, _indiceColumna(hoja, 'HoraFin') + 1).setValue(_minutosAHora(finMin));
    hoja.getRange(fila, _indiceColumna(hoja, 'Personas') + 1).setValue(personas);
    if (datos.notas !== undefined) {
      hoja.getRange(fila, _indiceColumna(hoja, 'Notas') + 1).setValue(datos.notas);
    }
    if (datos.notasInternas !== undefined) {
      _asegurarColumnaNotasInternas();
      var colNI = _indiceColumna(hoja, 'NotasInternas');
      if (colNI !== -1) hoja.getRange(fila, colNI + 1).setValue(datos.notasInternas);
    }

    registrarLog('Modificar reserva', id, reserva.Habitacion);
    return { success: true, mensaje: 'Reserva modificada.' };
  } catch (err) {
    return { success: false, mensaje: 'Error: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cancela una reserva. Libera capacidad automaticamente (queda excluida por
 * su estado en los calculos de ocupacion).
 * @param {string} id
 * @param {string} motivo
 * @param {boolean} esHotel Si true, "Cancelada hotel"; si no, "Cancelada huesped".
 * @return {Object} {success, mensaje}
 */
function cancelarReserva(id, motivo, esHotel) {
  var hoja = _hoja(HOJAS.RESERVAS);
  var reserva = _buscarReservaPorID(id);
  if (!reserva) return { success: false, mensaje: 'Reserva no encontrada.' };

  if (_esEstadoCancelado(reserva.Estado)) {
    return { success: false, mensaje: 'La reserva ya estaba cancelada.' };
  }
  if ([ESTADOS.EN_CURSO, ESTADOS.FINALIZADA].indexOf(reserva.Estado) !== -1) {
    return { success: false, mensaje: 'No se puede cancelar una reserva ' + reserva.Estado + '.' };
  }

  var nuevoEstado = esHotel ? ESTADOS.CANCELADA_HOTEL : ESTADOS.CANCELADA_HUESPED;
  var fila = reserva._fila;
  hoja.getRange(fila, _indiceColumna(hoja, 'Estado') + 1).setValue(nuevoEstado);
  hoja.getRange(fila, _indiceColumna(hoja, 'MotivoCancelacion') + 1).setValue(motivo || '');

  // Si tenia un pedido asociado, marcarlo Cancelado para que no quede "fantasma".
  var pedido = _buscarPedidoPorReserva(id);
  if (pedido) {
    var hojaPed = _hoja(HOJAS.PEDIDOS);
    hojaPed.getRange(pedido._fila, _indiceColumna(hojaPed, 'Estado') + 1).setValue('Cancelado');
  }

  registrarLog('Cancelar reserva', id + ' -> ' + nuevoEstado, reserva.Habitacion);
  return { success: true, mensaje: 'Reserva cancelada.' };
}

/**
 * Finaliza automaticamente las reservas cuyo horario ya paso. Asi, cuando una
 * habitacion rota de huesped, las solicitudes antiguas desaparecen solas de
 * la vista del nuevo huesped sin necesidad de "check-out" manual.
 * Con throttle: corre a lo mas una vez cada 5 minutos.
 */
function _finalizarReservasVencidas() {
  try {
    var cache = CacheService.getScriptCache();
    if (cache.get('finalizacionReciente')) return;
    cache.put('finalizacionReciente', '1', 300);
  } catch (e) { /* sin cache, igual continua */ }

  try {
    var hoja = _hoja(HOJAS.RESERVAS);
    var datos = _leerHojaComoObjetos(HOJAS.RESERVAS);
    var colEstado = _indiceColumna(hoja, 'Estado') + 1;
    var hoy = _fechaISO(new Date());
    var ahora = _minutosAhoraLocal();
    var activos = [ESTADOS.SOLICITADA, ESTADOS.PENDIENTE, ESTADOS.CONFIRMADA, ESTADOS.EN_CURSO];

    datos.forEach(function (r) {
      if (activos.indexOf(r.Estado) === -1) return;
      var fecha = _fechaISO(r.Fecha);
      var vencida = fecha < hoy || (fecha === hoy && _horaAMinutos(r.HoraFin) < ahora);
      if (vencida) {
        hoja.getRange(r._fila, colEstado).setValue(ESTADOS.FINALIZADA);
      }
    });
  } catch (e) { /* la limpieza nunca debe romper una lectura */ }
}

/** Busca una reserva por ID (devuelve objeto con _fila o null). */
function _buscarReservaPorID(id) {
  var reservas = _leerHojaComoObjetos(HOJAS.RESERVAS);
  for (var i = 0; i < reservas.length; i++) {
    if (reservas[i].ID === id) return reservas[i];
  }
  return null;
}

/**
 * Cambia el estado de una reserva (uso de staff: confirmar, en curso, etc.).
 * @param {string} id
 * @param {string} nuevoEstado
 * @return {Object} {success, mensaje}
 */
function cambiarEstadoReserva(id, nuevoEstado) {
  // Valida que el estado sea uno permitido.
  var permitidos = [];
  for (var k in ESTADOS) permitidos.push(ESTADOS[k]);
  if (permitidos.indexOf(nuevoEstado) === -1) {
    return { success: false, mensaje: 'Estado no valido.' };
  }
  var hoja = _hoja(HOJAS.RESERVAS);
  var reserva = _buscarReservaPorID(id);
  if (!reserva) return { success: false, mensaje: 'Reserva no encontrada.' };
  hoja.getRange(reserva._fila, _indiceColumna(hoja, 'Estado') + 1).setValue(nuevoEstado);
  registrarLog('Cambiar estado reserva', id + ' -> ' + nuevoEstado, reserva.Habitacion);
  return { success: true, mensaje: 'Estado actualizado.' };
}

/**
 * Devuelve reservas filtradas y enriquecidas con nombre de servicio.
 * @param {Object} filtros {fechaDesde, fechaHasta, habitacion, servicioID, estados}
 * @return {Array<Object>}
 */
function obtenerReservas(filtros) {
  filtros = filtros || {};
  if (filtros.incluirInternas) _asegurarColumnaNotasInternas();
  _finalizarReservasVencidas(); // limpieza automatica de reservas pasadas
  return _construirReservas(filtros, _leerHojaComoObjetos(HOJAS.RESERVAS));
}

/**
 * Filtra/enriquece un conjunto de reservas YA leido. Se separa de
 * obtenerReservas para poder reutilizar una unica lectura de la hoja Reservas
 * cuando varios calculos la necesitan en la misma ejecucion (ej. el centro de
 * operaciones), evitando releer la planilla (lo mas lento de Apps Script).
 * @param {Object} filtros
 * @param {Array<Object>} todasReservas Reservas ya leidas de la hoja.
 * @return {Array<Object>}
 */
function _construirReservas(filtros, todasReservas) {
  filtros = filtros || {};
  var servicios = {};
  _leerHojaComoObjetos(HOJAS.SERVICIOS).forEach(function (s) { servicios[s.ID] = s.Nombre; });
  var hoy = _fechaISO(new Date());
  var estadosCerrados = [ESTADOS.FINALIZADA, ESTADOS.CANCELADA_HUESPED,
    ESTADOS.CANCELADA_HOTEL, ESTADOS.NO_ASISTIO];

  return todasReservas.filter(function (r) {
    var fecha = _fechaISO(r.Fecha);
    if (filtros.fechaDesde && fecha < _fechaISO(filtros.fechaDesde)) return false;
    if (filtros.fechaHasta && fecha > _fechaISO(filtros.fechaHasta)) return false;
    if (filtros.habitacion && String(r.Habitacion) !== String(filtros.habitacion)) return false;
    if (filtros.servicioID && r.ServicioID !== filtros.servicioID) return false;
    if (filtros.estados && filtros.estados.length && filtros.estados.indexOf(r.Estado) === -1) return false;
    // soloActivas: lo que ve el huesped. Oculta reservas cerradas y pasadas,
    // para que un huesped nuevo en la misma habitacion no vea lo del anterior.
    if (filtros.soloActivas) {
      if (estadosCerrados.indexOf(r.Estado) !== -1) return false;
      if (fecha < hoy) return false;
    }
    return true;
  }).map(function (r) {
    var obj = {
      ID: r.ID,
      Timestamp: _fechaHoraTexto(r.Timestamp),
      Habitacion: String(r.Habitacion),
      ServicioID: r.ServicioID,
      ServicioNombre: servicios[r.ServicioID] || r.ServicioID,
      Fecha: _fechaISO(r.Fecha),
      HoraInicio: _horaATexto(r.HoraInicio),
      HoraFin: _horaATexto(r.HoraFin),
      Personas: Number(r.Personas) || 0,
      Estado: r.Estado,
      SolicitadoPor: r.SolicitadoPor,
      Notas: r.Notas,
      Variante: r.Variante ? String(r.Variante) : '',
      PrepedidoID: r.PrepedidoID,
      MotivoCancelacion: r.MotivoCancelacion,
      EsEvento: _aBooleano(r.EsEvento)
    };
    // Las notas internas SOLO se incluyen cuando lo pide el panel de staff
    // (filtros.incluirInternas). Las llamadas del huesped nunca lo piden, asi
    // que nunca las recibe.
    if (filtros.incluirInternas) obj.NotasInternas = r.NotasInternas ? String(r.NotasInternas) : '';
    return obj;
  }).sort(function (a, b) {
    return (a.Fecha + a.HoraInicio) < (b.Fecha + b.HoraInicio) ? -1 : 1;
  });
}

// ===========================================================================
// 6.4. GESTION DE PEDIDOS (PREPEDIDO)
// ===========================================================================

/**
 * Cuando se agrega o modifica un prepedido sobre una reserva YA CONFIRMADA,
 * la reserva vuelve a estado "Solicitada" para que recepcion/cocina la revise
 * de nuevo (por si no hay disponibilidad del plato pedido). Notifica a recepcion.
 * No hace nada si la reserva no estaba confirmada (o ya paso).
 * @param {Object} reserva Objeto de reserva (con _fila y Estado).
 * @return {boolean} true si se reabrio la confirmacion.
 */
function _reabrirConfirmacionPorPedido(reserva) {
  if (!reserva || reserva.Estado !== ESTADOS.CONFIRMADA) return false;
  var hoja = _hoja(HOJAS.RESERVAS);
  hoja.getRange(reserva._fila, _indiceColumna(hoja, 'Estado') + 1).setValue(ESTADOS.SOLICITADA);
  reserva.Estado = ESTADOS.SOLICITADA;
  var serv = _obtenerServicio(reserva.ServicioID);
  var nombreServicio = serv ? serv.Nombre : 'Restaurant';
  registrarLog('Reabrir confirmacion por pedido', reserva.ID + ' (' + nombreServicio + ')', reserva.Habitacion);
  // Notificacion con formato de reserva para que salga como pop-up en recepcion.
  var msg = 'RES|' + reserva.Habitacion + '|' + nombreServicio + '|' +
            _fechaISO(reserva.Fecha) + '|' + _horaATexto(reserva.HoraInicio) + '|' + ESTADOS.SOLICITADA;
  _crearNotificacion('reserva', msg, 'RECEPCION', reserva.Habitacion, reserva.ServicioID, _fechaISO(reserva.Fecha));
  return true;
}

/**
 * Crea un pedido asociado a una reserva.
 * @param {string} reservaID
 * @param {Array} items [{productoID, cantidad, notas}]
 * @param {string} notas Notas generales del pedido.
 * @param {string} entrega
 * @param {string} staffEmail Si viene de RECEPCION/COCINA/RESTAURANT/ADMINISTRADOR,
 *   se omite la ventana de tiempo del prepedido (ellos pueden operar hasta el ultimo minuto).
 * @return {Object} {success, pedidoID, total, mensaje}
 */
function crearPedido(reservaID, items, notas, entrega, staffEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var reserva = _buscarReservaPorID(reservaID);
    if (!reserva) return { success: false, mensaje: 'Reserva no encontrada.' };
    if (!items || !items.length) return { success: false, mensaje: 'El pedido esta vacio.' };

    // Si el restaurant no esta aceptando prepedidos, se rechaza de entrada.
    if (!_prepedidoActivo()) {
      return { success: false, mensaje: 'El restaurant no esta recibiendo pedidos anticipados hoy.' };
    }

    // Verifica ventana de tiempo de prepedido (el staff opera sin esta restriccion).
    var esStaff = staffEmail && _validarRolPermitido(staffEmail, ['RECEPCION', 'COCINA', 'RESTAURANT', 'ADMINISTRADOR']);
    if (!esStaff) {
      var validacionTiempo = _validarVentanaPrepedido(reserva);
      if (!validacionTiempo.ok) return { success: false, mensaje: validacionTiempo.mensaje };
    }

    // Entrega: 'Habitacion' o 'Restaurant' (por defecto).
    entrega = (entrega === 'Habitacion') ? 'Habitacion' : 'Restaurant';
    // Si el pedido a la habitacion esta deshabilitado, se rechaza por seguridad.
    if (entrega === 'Habitacion' && !_pedidoHabitacionActivo()) {
      return { success: false, mensaje: 'El pedido a la habitacion no esta disponible por ahora.' };
    }

    // Si ya existe un pedido para la reserva, redirige a modificar.
    var existente = _buscarPedidoPorReserva(reservaID);
    if (existente) {
      return modificarPedido(existente.ID, items, notas, entrega, staffEmail);
    }

    // Calcula subtotales y total leyendo precios desde la hoja Productos.
    var calculo = _calcularItems(items);
    if (!calculo.ok) return { success: false, mensaje: calculo.mensaje };

    // Recargo por entrega a la habitacion (configurable, se suma al total).
    var recargo = (entrega === 'Habitacion') ? _recargoHabitacion() : 0;
    var totalFinal = calculo.total + recargo;

    var pedidoID = generarID();
    var hojaPed = _hoja(HOJAS.PEDIDOS);
    var filaPed = [pedidoID, reservaID, String(reserva.Habitacion), 'Registrado',
      totalFinal, new Date(), notas || ''];
    // Rellena hasta el numero real de columnas y setea Entrega por indice.
    var colEntrega = _indiceColumna(hojaPed, 'Entrega');
    while (filaPed.length < hojaPed.getLastColumn()) filaPed.push('');
    if (colEntrega !== -1) filaPed[colEntrega] = entrega;
    hojaPed.appendRow(filaPed);

    // Inserta todos los detalles en UNA sola escritura (mucho mas rapido
    // que un appendRow por plato).
    _insertarDetallesPedido(pedidoID, calculo.detalles);

    // Vincula el pedido a la reserva.
    var hojaReservas = _hoja(HOJAS.RESERVAS);
    hojaReservas.getRange(reserva._fila, _indiceColumna(hojaReservas, 'PrepedidoID') + 1).setValue(pedidoID);

    // Si la reserva ya estaba confirmada, se reabre para volver a confirmarla.
    var reabierta = _reabrirConfirmacionPorPedido(reserva);

    registrarLog('Crear pedido', pedidoID + ' total ' + totalFinal +
      (recargo ? ' (incluye recargo habitacion ' + recargo + ')' : ''), reserva.Habitacion);
    // 'TODOS' en vez de un rol fijo: asi la notificacion siempre llega a quien
    // atienda el restaurant, sin depender del nombre exacto del rol en la hoja
    // Usuarios (que puede cambiar, como paso de COCINA a RESTAURANT).
    _crearNotificacion('pedido', 'Nuevo prepedido Hab ' + reserva.Habitacion, 'TODOS',
      reserva.Habitacion, reserva.ServicioID, _fechaISO(reserva.Fecha));

    return { success: true, pedidoID: pedidoID, total: totalFinal, reabierta: reabierta,
      mensaje: reabierta ? 'Pedido registrado. La reserva vuelve a estado Solicitada para reconfirmar.' : 'Pedido registrado.' };
  } catch (err) {
    return { success: false, mensaje: 'Error: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Modifica los items de un pedido existente. Solo si falta mas del limite
 * configurado para la hora de la reserva (el staff RECEPCION/COCINA/RESTAURANT/
 * ADMINISTRADOR queda exento de esta ventana).
 * @param {string} pedidoID
 * @param {Array} items [{productoID, cantidad, notas}]
 * @param {string} notas
 * @param {string} entrega
 * @param {string} staffEmail
 * @return {Object} {success, total, mensaje}
 */
function modificarPedido(pedidoID, items, notas, entrega, staffEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var pedido = _buscarPedidoPorID(pedidoID);
    if (!pedido) return { success: false, mensaje: 'Pedido no encontrado.' };
    var reserva = _buscarReservaPorID(pedido.ReservaID);
    if (!reserva) return { success: false, mensaje: 'Reserva asociada no encontrada.' };

    // Si el restaurant no esta aceptando prepedidos, tampoco se aceptan cambios.
    if (!_prepedidoActivo()) {
      return { success: false, mensaje: 'El restaurant no esta recibiendo pedidos anticipados hoy.' };
    }

    var esStaff = staffEmail && _validarRolPermitido(staffEmail, ['RECEPCION', 'COCINA', 'RESTAURANT', 'ADMINISTRADOR']);
    if (!esStaff) {
      var validacionTiempo = _validarVentanaPrepedido(reserva);
      if (!validacionTiempo.ok) return { success: false, mensaje: validacionTiempo.mensaje };
    }

    if (!items || !items.length) return { success: false, mensaje: 'El pedido no puede quedar vacio.' };

    // La entrega se valida ANTES de tocar los detalles: si se rechaza aqui,
    // el pedido queda exactamente como estaba (sin estados a medias).
    var entregaFinal;
    if (entrega !== undefined) {
      entregaFinal = (entrega === 'Habitacion') ? 'Habitacion' : 'Restaurant';
      if (entregaFinal === 'Habitacion' && !_pedidoHabitacionActivo()) {
        return { success: false, mensaje: 'El pedido a la habitacion no esta disponible por ahora.' };
      }
    }
    // Entrega con la que quedara el pedido (la nueva, o la que ya tenia).
    var entregaEfectiva = (entregaFinal !== undefined)
      ? entregaFinal
      : ((pedido.Entrega === 'Habitacion') ? 'Habitacion' : 'Restaurant');

    var calculo = _calcularItems(items);
    if (!calculo.ok) return { success: false, mensaje: calculo.mensaje };

    // Recargo por entrega a la habitacion (configurable, se suma al total).
    var recargo = (entregaEfectiva === 'Habitacion') ? _recargoHabitacion() : 0;
    var totalFinal = calculo.total + recargo;

    // Borra los detalles anteriores del pedido.
    _borrarDetallesDePedido(pedidoID);

    // Inserta los nuevos detalles en una sola escritura.
    _insertarDetallesPedido(pedidoID, calculo.detalles);

    // Actualiza el total, notas y entrega del pedido.
    var hojaPedidos = _hoja(HOJAS.PEDIDOS);
    hojaPedidos.getRange(pedido._fila, _indiceColumna(hojaPedidos, 'Total') + 1).setValue(totalFinal);
    if (notas !== undefined) {
      hojaPedidos.getRange(pedido._fila, _indiceColumna(hojaPedidos, 'Notas') + 1).setValue(notas || '');
    }
    if (entregaFinal !== undefined) {
      var colEnt = _indiceColumna(hojaPedidos, 'Entrega');
      if (colEnt !== -1) hojaPedidos.getRange(pedido._fila, colEnt + 1).setValue(entregaFinal);
    }

    // Si la reserva ya estaba confirmada, se reabre para volver a confirmarla.
    var reabierta = _reabrirConfirmacionPorPedido(reserva);

    registrarLog('Modificar pedido', pedidoID + ' total ' + totalFinal +
      (recargo ? ' (incluye recargo habitacion ' + recargo + ')' : ''), pedido.Habitacion);
    return { success: true, total: totalFinal, reabierta: reabierta,
      mensaje: reabierta ? 'Pedido actualizado. La reserva vuelve a estado Solicitada para reconfirmar.' : 'Pedido actualizado.' };
  } catch (err) {
    return { success: false, mensaje: 'Error: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Valida que la reserva permita (o siga permitiendo) prepedido segun la
 * ventana de PREPEDIDO_MINUTOS_LIMITE.
 * @return {Object} {ok, mensaje}
 */
function _validarVentanaPrepedido(reserva) {
  var limite = Number(_obtenerConfigValor('PREPEDIDO_MINUTOS_LIMITE')) || 30;
  var fecha = _fechaISO(reserva.Fecha);
  var hoy = _fechaISO(new Date());

  // Si la reserva es en fecha futura, siempre se puede.
  if (fecha > hoy) return { ok: true, mensaje: '' };
  // Si ya paso el dia, no se puede.
  if (fecha < hoy) return { ok: false, mensaje: 'La reserva ya paso.' };

  // Mismo dia: comparar contra la hora limite.
  var horaReserva = _horaAMinutos(reserva.HoraInicio);
  var ahora = _minutosAhoraLocal();
  if (ahora + limite > horaReserva) {
    return { ok: false, mensaje: 'El pedido solo puede modificarse hasta ' + limite +
      ' min antes de la reserva.' };
  }
  return { ok: true, mensaje: '' };
}

/** Calcula detalles/subtotales/total leyendo precios reales de Productos. */
function _calcularItems(items) {
  var productos = {};
  _leerHojaComoObjetos(HOJAS.PRODUCTOS).forEach(function (p) {
    productos[p.ID] = _normalizarProducto(p);
  });

  var detalles = [];
  var total = 0;
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var prod = productos[it.productoID];
    if (!prod) return { ok: false, mensaje: 'Producto no encontrado: ' + it.productoID };
    if (!prod.Disponible) return { ok: false, mensaje: 'Producto agotado: ' + prod.Nombre };
    var cantidad = Number(it.cantidad) || 0;
    if (cantidad <= 0) continue;
    var subtotal = prod.Precio * cantidad;
    total += subtotal;
    detalles.push({
      productoID: prod.ID,
      cantidad: cantidad,
      precioUnitario: prod.Precio,
      subtotal: subtotal,
      notas: it.notas || ''
    });
  }
  if (!detalles.length) return { ok: false, mensaje: 'No hay items validos en el pedido.' };
  return { ok: true, detalles: detalles, total: total };
}

/** Busca un pedido por ID. */
function _buscarPedidoPorID(pedidoID) {
  var pedidos = _leerHojaComoObjetos(HOJAS.PEDIDOS);
  for (var i = 0; i < pedidos.length; i++) {
    if (pedidos[i].ID === pedidoID) return pedidos[i];
  }
  return null;
}

/** Busca un pedido por reserva. */
function _buscarPedidoPorReserva(reservaID) {
  var pedidos = _leerHojaComoObjetos(HOJAS.PEDIDOS);
  for (var i = 0; i < pedidos.length; i++) {
    if (pedidos[i].ReservaID === reservaID) return pedidos[i];
  }
  return null;
}

/**
 * Inserta los detalles de un pedido en UNA sola escritura por lote.
 * (Un appendRow por plato costaba ~200-300ms cada uno.)
 * @param {string} pedidoID
 * @param {Array} detalles [{productoID, cantidad, precioUnitario, subtotal, notas}]
 */
function _insertarDetallesPedido(pedidoID, detalles) {
  if (!detalles || !detalles.length) return;
  var hoja = _hoja(HOJAS.DETALLE_PEDIDOS);
  var filas = detalles.map(function (d) {
    return [generarID(), pedidoID, d.productoID, d.cantidad, d.precioUnitario, d.subtotal, d.notas || ''];
  });
  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
}

/** Borra fisicamente los detalles de un pedido (para reemplazo). */
function _borrarDetallesDePedido(pedidoID) {
  var hoja = _hoja(HOJAS.DETALLE_PEDIDOS);
  var datos = hoja.getDataRange().getValues();
  // Recorre de abajo hacia arriba para borrar filas con seguridad.
  for (var i = datos.length - 1; i >= 1; i--) {
    if (datos[i][1] === pedidoID) { // columna PedidoID (indice 1)
      hoja.deleteRow(i + 1);
    }
  }
}

/**
 * Devuelve el pedido de una reserva con sus detalles enriquecidos.
 * @param {string} reservaID
 * @return {Object|null} {pedido, detalles:[{...producto}]}
 */
function obtenerPedidosPorReserva(reservaID) {
  var pedido = _buscarPedidoPorReserva(reservaID);
  if (!pedido) return null;

  _asegurarColumnaPuntoCoccion();
  var productos = {};
  _leerHojaComoObjetos(HOJAS.PRODUCTOS).forEach(function (p) {
    productos[p.ID] = _normalizarProducto(p);
  });

  var detalles = _leerHojaComoObjetos(HOJAS.DETALLE_PEDIDOS)
    .filter(function (d) { return d.PedidoID === pedido.ID; })
    .map(function (d) {
      var prod = productos[d.ProductoID] || {};
      return {
        ProductoID: d.ProductoID,
        Nombre: prod.Nombre || d.ProductoID,
        CategoriaID: prod.CategoriaID || '',
        RequierePuntoCoccion: prod.RequierePuntoCoccion !== false,
        Cantidad: Number(d.Cantidad) || 0,
        PrecioUnitario: Number(d.PrecioUnitario) || 0,
        Subtotal: Number(d.Subtotal) || 0,
        Notas: d.Notas
      };
    });

  return {
    pedido: {
      ID: pedido.ID,
      ReservaID: pedido.ReservaID,
      Habitacion: String(pedido.Habitacion),
      Estado: pedido.Estado,
      Total: Number(pedido.Total) || 0,
      Timestamp: _fechaHoraTexto(pedido.Timestamp),
      Notas: pedido.Notas,
      Entrega: pedido.Entrega || 'Restaurant'
    },
    detalles: detalles
  };
}

// ===========================================================================
// 6.5. CENTRO DE OPERACIONES Y ALERTAS
// ===========================================================================

/**
 * Devuelve el timeline del dia agrupado por hora.
 * @param {string} fecha "YYYY-MM-DD"
 * @return {Object} {fecha, bloques:[{hora, items:[...]}], alertas:[...]}
 */
function obtenerCentroOperaciones(fecha, reservasPre) {
  fecha = _fechaISO(fecha || new Date());
  // Si el llamador ya finalizo vencidas y leyo las reservas (endpoint
  // combinado), no repetimos ese trabajo.
  if (!reservasPre) _finalizarReservasVencidas();

  var servicios = {};
  _leerHojaComoObjetos(HOJAS.SERVICIOS).forEach(function (s) {
    servicios[s.ID] = _normalizarServicio(s);
  });

  // Lee TODAS las reservas una sola vez y reutilizalas (optimizacion).
  var todasReservas = reservasPre || _leerHojaComoObjetos(HOJAS.RESERVAS);
  var reservas = todasReservas.filter(function (r) {
    return _fechaISO(r.Fecha) === fecha && !_esEstadoCancelado(r.Estado);
  });

  // Agrupa por hora de inicio (normalizada a texto "HH:MM").
  var mapaHoras = {};
  reservas.forEach(function (r) {
    var hora = _horaATexto(r.HoraInicio);
    if (!mapaHoras[hora]) mapaHoras[hora] = [];
    var serv = servicios[r.ServicioID] || {};
    mapaHoras[hora].push({
      reservaID: r.ID,
      servicioID: r.ServicioID,
      servicioNombre: serv.Nombre || r.ServicioID,
      habitacion: String(r.Habitacion),
      personas: Number(r.Personas) || 0,
      capacidad: serv.Capacidad || 0,
      estado: r.Estado,
      horaInicio: hora,
      notas: r.Notas ? String(r.Notas) : '',
      notasInternas: r.NotasInternas ? String(r.NotasInternas) : '',
      variante: r.Variante ? String(r.Variante) : '',
      tienePedido: !!r.PrepedidoID
    });
  });

  // Ordena bloques por hora.
  var horas = Object.keys(mapaHoras).sort(function (a, b) {
    return _horaAMinutos(a) - _horaAMinutos(b);
  });
  var bloques = horas.map(function (h) {
    // Calcula ocupacion por servicio dentro del bloque.
    var ocupacionPorServicio = {};
    mapaHoras[h].forEach(function (it) {
      ocupacionPorServicio[it.servicioID] = (ocupacionPorServicio[it.servicioID] || 0) + it.personas;
    });
    return {
      hora: h,
      items: mapaHoras[h],
      ocupacionPorServicio: ocupacionPorServicio
    };
  });

  // Genera alertas de capacidad casi llena (reutiliza las reservas del dia).
  var alertas = _generarAlertasCapacidad(reservas, servicios);

  // Pedidos del dia incluidos: se muestran DENTRO de cada reserva. Reutiliza
  // las reservas ya leidas para no volver a leer la hoja.
  var pedidos = obtenerPedidosDelDia(fecha, todasReservas);

  return { fecha: fecha, bloques: bloques, alertas: alertas, pedidos: pedidos };
}

/**
 * Endpoint COMBINADO del Centro de Operaciones: en UNA sola llamada devuelve
 * el centro del dia, los productos agotados ("86") y las reservas por venir.
 * Antes el frontend hacia 3 llamadas en paralelo (3 viajes al servidor de
 * ~1-2s cada uno) que ademas leian Reservas y Servicios por duplicado; ahora
 * es 1 viaje y cada hoja se lee una sola vez (Reservas se comparte, y los
 * catalogos se sirven del cache de lectura por-ejecucion). Esto tambien aligera
 * el refresco automatico que corre cada 20s mientras el staff esta en la vista.
 * @param {string} fecha
 * @return {Object} {centro, agotados, reservasProximas}
 */
function obtenerDatosOperaciones(fecha) {
  fecha = _fechaISO(fecha || new Date());
  _asegurarColumnaNotasInternas(); // la columna debe existir antes de leer
  _finalizarReservasVencidas(); // una sola vez para todo el endpoint
  var todasReservas = _leerHojaComoObjetos(HOJAS.RESERVAS);

  var estadosProximas = [ESTADOS.SOLICITADA, ESTADOS.PENDIENTE, ESTADOS.CONFIRMADA];
  // Rango de eventos "por venir": desde hoy hasta un ano hacia adelante. Es
  // independiente del dia que este mirando el staff, para que el widget
  // "Reservas por venir" muestre siempre los eventos futuros.
  var hoy = _fechaISO(new Date());
  var lejano = new Date();
  lejano.setFullYear(lejano.getFullYear() + 1);
  return {
    centro: obtenerCentroOperaciones(fecha, todasReservas),
    agotados: obtenerProductosAgotados(),
    reservasProximas: _construirReservas({ estados: estadosProximas }, todasReservas),
    // Bloqueos/eventos que caen en el dia seleccionado, para las burbujas de "hoy".
    eventos: obtenerBloqueos(fecha, fecha, ''),
    // Eventos futuros (desde hoy) para el widget "Reservas por venir", sin
    // importar que dia se este mirando en el Centro de Operaciones.
    eventosProximos: obtenerBloqueos(hoy, _fechaISO(lejano), ''),
    // Participantes por reserva/evento, para mostrarlos dentro de cada burbuja
    // sin una llamada extra por cada una.
    participantes: _mapaParticipantes()
  };
}

/**
 * Genera alertas de cupos. Redaccion neutra (sin genero): usa "Sin cupos para X"
 * en vez de "completo/completa".
 * @param {Array} reservasDelDia Reservas ya filtradas del dia.
 * @param {Object} servicios Mapa servicioID -> servicio normalizado.
 */
function _generarAlertasCapacidad(reservasDelDia, servicios) {
  var alertas = [];
  var mapa = {};
  reservasDelDia.forEach(function (r) {
    var clave = r.ServicioID + '|' + _horaATexto(r.HoraInicio);
    mapa[clave] = (mapa[clave] || 0) + (Number(r.Personas) || 0);
  });
  Object.keys(mapa).forEach(function (clave) {
    var partes = clave.split('|');
    var serv = servicios[partes[0]];
    if (!serv) return;
    // Servicios de uso exclusivo (ej. Tinaja, Masajes) se reservan enteros:
    // una sola reserva ocupa todo el horario sin importar cuantas personas
    // asistan (2 personas en una tinaja de 4 igual la dejan tomada por
    // completo). Ahi no existe un "quedan X cupos" parcial: o esta libre o
    // esta ocupada.
    if (_esServicioExclusivo(serv)) {
      if (mapa[clave] > 0) {
        alertas.push('Sin cupos para ' + serv.Nombre + ' a las ' + partes[1] + ' (uso exclusivo).');
      }
      return;
    }
    var restante = serv.Capacidad - mapa[clave];
    if (restante <= 2 && restante > 0) {
      alertas.push('Quedan ' + restante + ' cupos para ' + serv.Nombre + ' a las ' + partes[1] + '.');
    } else if (restante <= 0) {
      alertas.push('Sin cupos para ' + serv.Nombre + ' a las ' + partes[1] + '.');
    }
  });
  return alertas;
}

/**
 * Devuelve notificaciones para un rol/habitacion. Genera automaticamente las
 * alertas de capacidad faltantes del dia si se consulta.
 * @param {string} rol
 * @param {string} habitacion
 * @param {boolean} soloNoLeidas
 * @return {Array<Object>}
 */
function obtenerNotificaciones(rol, habitacion, soloNoLeidas) {
  // Lectura liviana de UNA hoja: apta para consultarse por polling cada 20s.
  // (Las alertas de capacidad ya se muestran en el Centro de Operaciones.)
  _purgarNotificacionesAntiguas(); // mantiene la hoja chica (auto-limpieza)
  return _leerHojaComoObjetos(HOJAS.NOTIFICACIONES).filter(function (n) {
    if (soloNoLeidas && _aBooleano(n.Leida)) return false;
    if (rol && !_esNotifParaRol(n, rol)) return false;
    if (habitacion && n.Habitacion && String(n.Habitacion) !== String(habitacion)) return false;
    return true;
  }).map(function (n) {
    return {
      ID: n.ID,
      Timestamp: _fechaHoraTexto(n.Timestamp),
      Tipo: n.Tipo,
      Mensaje: n.Mensaje,
      DestinatarioRol: n.DestinatarioRol,
      Habitacion: String(n.Habitacion || ''),
      ServicioID: n.ServicioID,
      Leida: _aBooleano(n.Leida),
      FechaReferencia: _fechaISO(n.FechaReferencia)
    };
  }).sort(function (a, b) { return a.Timestamp < b.Timestamp ? 1 : -1; });
}

/**
 * Decide si una notificacion corresponde a un rol.
 * El ADMINISTRADOR ve todas las notificaciones de todos los roles.
 */
function _esNotifParaRol(n, rol) {
  if (rol === 'ADMINISTRADOR') return true;
  return !n.DestinatarioRol || n.DestinatarioRol === 'TODOS' || n.DestinatarioRol === rol;
}

/** Inserta una notificacion. */
function _crearNotificacion(tipo, mensaje, destinatarioRol, habitacion, servicioID, fechaReferencia) {
  _hoja(HOJAS.NOTIFICACIONES).appendRow([
    generarID(), new Date(), tipo, mensaje, destinatarioRol || 'TODOS',
    habitacion || '', servicioID || '', 'FALSE', _fechaISO(fechaReferencia || new Date())
  ]);
  // Aviso REAL al telefono (suena aunque la app este cerrada). Ver bloque
  // "PUSH EXTERNO" mas abajo. Nunca puede romper la reserva: va en try/catch.
  _enviarPushExterno(tipo, mensaje, destinatarioRol, habitacion, servicioID);
}

// ===========================================================================
// PUSH EXTERNO (avisos reales al telefono, con la app cerrada)
// ---------------------------------------------------------------------------
// Por que existe: la webapp de Apps Script se sirve dentro de un iframe
// sandbox de googleusercontent.com. Ahi el navegador NO permite registrar un
// Service Worker ni usar la Web Push API, asi que un aviso "nativo" desde la
// pagina es imposible: si el telefono se bloquea o el navegador se minimiza,
// el JavaScript se suspende y el polling deja de correr.
//
// La solucion es invertir quien avisa: en vez de que el telefono pregunte,
// es el SERVIDOR (este script) el que empuja el aviso hacia una app que si
// tiene push nativo. Se soportan dos canales, ambos gratis:
//   - Telegram : se crea un bot y un grupo con el personal del restaurant.
//   - ntfy.sh  : app dedicada a avisos, sin cuenta ni registro.
// Se puede usar uno, el otro, o los dos a la vez.
//
// Todo esto corre en el servidor, o sea funciona con el telefono bloqueado,
// la app cerrada y el navegador sin abrir.
// ===========================================================================

/** Claves de Configuracion que usa el push, con su valor por defecto. */
var CLAVES_PUSH = [
  ['PUSH_ACTIVO', 'FALSE', 'Envia los avisos al telefono aunque la app este cerrada (Telegram / ntfy)'],
  ['PUSH_ROLES', 'TODOS', 'Que avisos se envian al telefono: TODOS, GASTRONOMIA o NO_GASTRONOMIA'],
  ['PUSH_TELEGRAM_TOKEN', '', 'Token del bot de Telegram (te lo da @BotFather)'],
  ['PUSH_TELEGRAM_CHAT', '', 'ID del chat o grupo de Telegram. Varios separados por coma'],
  ['PUSH_NTFY_TOPIC', '', 'Nombre del canal en ntfy.sh (usa algo largo y dificil de adivinar)']
];

/**
 * Crea en la hoja Configuracion las claves de push que falten. Se llama al
 * leer la configuracion, asi la app se auto-actualiza sin tocar el Sheet.
 */
function _asegurarClavesPush() {
  try {
    var hoja = _hoja(HOJAS.CONFIGURACION);
    var existentes = {};
    _leerHojaComoObjetos(HOJAS.CONFIGURACION).forEach(function (f) { existentes[f.Clave] = true; });
    var faltantes = CLAVES_PUSH.filter(function (c) { return !existentes[c[0]]; });
    if (!faltantes.length) return;
    faltantes.forEach(function (c) { hoja.appendRow(c); });
    _invalidarCaches(HOJAS.CONFIGURACION);
  } catch (e) {
    /* si no se pueden crear, el push simplemente queda apagado */
  }
}

/**
 * Envia el aviso a los canales configurados. Silencioso ante cualquier fallo:
 * un problema de red jamas debe impedir que se guarde una reserva.
 */
function _enviarPushExterno(tipo, mensaje, destinatarioRol, habitacion, servicioID) {
  try {
    if (String(_obtenerConfigValor('PUSH_ACTIVO') || '').toUpperCase() !== 'TRUE') return;
    if (!_avisoPasaFiltro(_obtenerConfigValor('PUSH_ROLES'), tipo, servicioID)) return;

    var aviso = _formatearMensajePush(tipo, mensaje, habitacion);
    _pushTelegram(aviso.titulo, aviso.cuerpo);
    _pushNtfy(aviso.titulo, aviso.cuerpo);
  } catch (e) {
    /* nunca romper el flujo de reservas/pedidos */
  }
}

/**
 * Decide si un aviso pasa el filtro elegido en Configuracion.
 *
 * OJO: NO se puede filtrar por DestinatarioRol. Todas las reservas se generan
 * con rol RECEPCION sin importar el servicio (ver crearReserva), asi que
 * filtrar por rol dejaba fuera absolutamente todas las reservas. Se filtra por
 * la categoria real del servicio, que es lo que una persona entiende cuando
 * dice "solo los del restaurant".
 *
 * @param {string} filtro Valor de PUSH_ROLES
 * @param {string} tipo   'reserva' | 'pedido' | ...
 * @param {string} servicioID
 * @return {boolean}
 */
function _avisoPasaFiltro(filtro, tipo, servicioID) {
  var f = String(filtro || 'TODOS').toUpperCase().trim();

  // Valores de la primera version, cuando el filtro era por rol. Se traducen
  // para que nadie quede sin avisos despues de actualizar.
  if (f === 'RESTAURANT' || f === 'COCINA' || f === 'RESTAURANT,COCINA') f = 'GASTRONOMIA';
  if (f === 'RECEPCION' || f === '') f = 'TODOS';

  if (f === 'TODOS') return true;

  var esComida = (String(tipo || '').toLowerCase() === 'pedido') || _servicioEsGastronomia(servicioID);
  if (f === 'GASTRONOMIA') return esComida;
  if (f === 'NO_GASTRONOMIA') return !esComida;
  return true; // filtro desconocido: mejor avisar de mas que de menos
}

/** True si el servicio pertenece a la categoria Gastronomia. */
function _servicioEsGastronomia(servicioID) {
  if (!servicioID) return false;
  try {
    var servicios = _leerHojaComoObjetos(HOJAS.SERVICIOS);
    for (var i = 0; i < servicios.length; i++) {
      if (String(servicios[i].ID) === String(servicioID)) {
        return String(servicios[i].Categoria || '').toLowerCase().indexOf('gastronom') === 0;
      }
    }
  } catch (e) { /* si no se puede leer, no se bloquea el aviso */ }
  return false;
}

/**
 * Convierte el mensaje interno en algo legible en la pantalla del telefono.
 * Entiende el formato estructurado "RES|hab|servicio|fecha|hora|estado".
 */
function _formatearMensajePush(tipo, mensaje, habitacion) {
  var m = String(mensaje || '');

  if (m.indexOf('RES|') === 0) {
    var p = m.split('|');
    var hab = p[1] || habitacion || '';
    return {
      titulo: 'Nueva reserva' + (hab ? '  ·  Hab ' + hab : ''),
      cuerpo: (p[2] || 'Servicio') +
        '\n' + _fechaLegiblePush(p[3]) + (p[4] ? '  ·  ' + p[4] + ' hrs' : '') +
        (p[5] ? '\nEstado: ' + p[5] : '')
    };
  }

  var titulos = {
    pedido: 'Nuevo pedido',
    reserva: 'Nueva reserva',
    cancelacion: 'Reserva cancelada',
    aviso: 'Aviso'
  };
  // El titulo ya lleva la habitacion: se saca del cuerpo para no repetirla
  // ("Nuevo pedido · Hab 204" / "Nuevo prepedido Hab 204").
  var cuerpo = m;
  if (habitacion) {
    cuerpo = cuerpo.replace(new RegExp('\\s*Hab\\.?\\s*' + habitacion + '\\s*$', 'i'), '').trim();
  }
  return {
    titulo: (titulos[String(tipo || '').toLowerCase()] || 'Cascadas Concierge') +
            (habitacion ? '  ·  Hab ' + habitacion : ''),
    cuerpo: cuerpo || m || 'Tienes un aviso nuevo en el Concierge.'
  };
}

/** "2026-08-01" -> "sabado 1 de agosto". Si no puede, devuelve el original. */
function _fechaLegiblePush(iso) {
  try {
    if (!iso) return '';
    var d = new Date(String(iso) + 'T12:00:00');
    if (isNaN(d.getTime())) return String(iso);
    var dias = ['domingo', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado'];
    var meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
                 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    return dias[d.getDay()] + ' ' + d.getDate() + ' de ' + meses[d.getMonth()];
  } catch (e) { return String(iso || ''); }
}

/** Escapa lo minimo para el parse_mode HTML de Telegram. */
function _escaparHtmlTelegram(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Envia por Telegram. Acepta varios chat id separados por coma. */
function _pushTelegram(titulo, cuerpo) {
  var token = String(_obtenerConfigValor('PUSH_TELEGRAM_TOKEN') || '').trim();
  var chats = String(_obtenerConfigValor('PUSH_TELEGRAM_CHAT') || '').trim();
  if (!token || !chats) return;

  var texto = '<b>' + _escaparHtmlTelegram(titulo) + '</b>\n' + _escaparHtmlTelegram(cuerpo);
  chats.split(',').forEach(function (chat) {
    chat = String(chat).trim();
    if (!chat) return;
    try {
      UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        payload: JSON.stringify({
          chat_id: chat, text: texto, parse_mode: 'HTML', disable_web_page_preview: true
        })
      });
    } catch (e) { /* un chat caido no puede frenar a los demas */ }
  });
}

/**
 * Envia por ntfy.sh. Se usa el endpoint JSON (y no las cabeceras) para que
 * los acentos del titulo viajen sin problemas de codificacion.
 */
function _pushNtfy(titulo, cuerpo) {
  var topic = String(_obtenerConfigValor('PUSH_NTFY_TOPIC') || '').trim();
  if (!topic) return;
  try {
    UrlFetchApp.fetch('https://ntfy.sh/', {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        topic: topic, title: titulo, message: cuerpo, priority: 4, tags: ['bell']
      })
    });
  } catch (e) { /* silencioso */ }
}

/**
 * Manda un aviso de prueba a los canales configurados. Se llama desde el
 * boton "Enviar aviso de prueba" en Configuracion.
 * @param {string} email
 * @return {Object} {success, mensaje}
 */
function enviarPushDePrueba(email) {
  if (!_validarRolPermitido(email, ['RECEPCION', 'ADMINISTRADOR', 'COCINA', 'RESTAURANT'])) {
    return { success: false, mensaje: 'No tienes permisos para probar las notificaciones.' };
  }
  if (String(_obtenerConfigValor('PUSH_ACTIVO') || '').toUpperCase() !== 'TRUE') {
    return { success: false, mensaje: 'Primero activa "Avisar al telefono" y guarda los datos del canal.' };
  }
  var token = String(_obtenerConfigValor('PUSH_TELEGRAM_TOKEN') || '').trim();
  var chats = String(_obtenerConfigValor('PUSH_TELEGRAM_CHAT') || '').trim();
  var topic = String(_obtenerConfigValor('PUSH_NTFY_TOPIC') || '').trim();
  if ((!token || !chats) && !topic) {
    return { success: false, mensaje: 'Falta configurar Telegram (token + chat) o el canal de ntfy.' };
  }

  var canales = [];
  if (token && chats) {
    var r = _probarTelegram(token, chats);
    if (!r.ok) return { success: false, mensaje: 'Telegram: ' + r.detalle };
    canales.push('Telegram');
  }
  if (topic) {
    _pushNtfy('Prueba de aviso', 'Si ves esto en tu telefono, los avisos del Concierge estan funcionando.');
    canales.push('ntfy');
  }

  // La prueba se salta el filtro a proposito (comprueba el canal). Si el filtro
  // esta acotado hay que decirlo, o el "funciono" da una falsa seguridad.
  var f = _avisoPasaFiltro(_obtenerConfigValor('PUSH_ROLES'), 'reserva', null) &&
          _avisoPasaFiltro(_obtenerConfigValor('PUSH_ROLES'), 'pedido', null);
  var nota = f ? '' : ' Ojo: tienes un filtro activo, asi que no todas las reservas van a avisar.';
  return { success: true, mensaje: 'Aviso de prueba enviado por ' + canales.join(' y ') + '.' + nota };
}

/**
 * Marca de version del codigo de avisos. Sirve para saber, desde la app, si
 * la implementacion publicada es la ultima o quedo una antigua.
 */
var VERSION_PUSH = '2026-08-03-c';

/**
 * Radiografia completa de los avisos: que hay guardado, que responde cada
 * canal y si una reserva concreta pasaria el filtro. Devuelve texto plano
 * para poder leerlo o mandarlo por pantallazo.
 * @param {string} email
 * @return {Object} {success, texto}
 */
function diagnosticoPush(email) {
  if (!_validarRolPermitido(email, ['RECEPCION', 'ADMINISTRADOR', 'COCINA', 'RESTAURANT'])) {
    return { success: false, texto: 'No tienes permisos.' };
  }
  var L = [];
  L.push('VERSION DEL CODIGO: ' + VERSION_PUSH);
  L.push('(si esto no dice 2026-08-03-c, quedo publicada una version antigua:');
  L.push(' hay que volver a publicar en Implementar > Administrar implementaciones)');
  L.push('');

  var activo = String(_obtenerConfigValor('PUSH_ACTIVO') || '').toUpperCase();
  L.push('AVISOS ACTIVOS: ' + (activo || '(vacio)') + (activo === 'TRUE' ? '' : '   <-- APAGADO'));

  var filtro = String(_obtenerConfigValor('PUSH_ROLES') || 'TODOS');
  L.push('FILTRO: ' + filtro);
  var srv = _leerHojaComoObjetos(HOJAS.SERVICIOS) || [];
  var unaGastro = null, unaOtra = null;
  srv.forEach(function (s) {
    var esG = String(s.Categoria || '').toLowerCase().indexOf('gastronom') === 0;
    if (esG && !unaGastro) unaGastro = s;
    if (!esG && !unaOtra) unaOtra = s;
  });
  if (unaGastro) {
    L.push('  ' + unaGastro.Nombre + ': ' +
      (_avisoPasaFiltro(filtro, 'reserva', unaGastro.ID) ? 'SI avisa' : 'NO avisa (filtro)'));
  }
  if (unaOtra) {
    L.push('  ' + unaOtra.Nombre + ': ' +
      (_avisoPasaFiltro(filtro, 'reserva', unaOtra.ID) ? 'SI avisa' : 'NO avisa (filtro)'));
  }
  L.push('');

  // ---- ntfy ----
  var topic = String(_obtenerConfigValor('PUSH_NTFY_TOPIC') || '');
  L.push('NTFY');
  if (!topic.trim()) {
    L.push('  Sin canal configurado.');
  } else {
    L.push('  Canal guardado: "' + topic + '"');
    L.push('  Largo: ' + topic.length + ' caracteres' +
      (topic !== topic.trim() ? '   <-- TIENE ESPACIOS AL PRINCIPIO O AL FINAL' : ''));
    L.push('  El telefono debe estar suscrito a ESTE nombre, identico.');
    try {
      var rn = UrlFetchApp.fetch('https://ntfy.sh/', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({
          topic: topic.trim(), title: 'Diagnostico', message: 'Mensaje de diagnostico del Concierge.',
          priority: 4, tags: ['bell']
        })
      });
      var cod = rn.getResponseCode();
      L.push('  Respuesta del servidor: HTTP ' + cod + (cod === 200 ? '  (enviado OK)' : '  <-- ERROR'));
      L.push('  Detalle: ' + String(rn.getContentText() || '').slice(0, 300));
      if (cod === 200) {
        L.push('  El servidor SI envio. Si no llego al telefono, el problema esta');
        L.push('  en el telefono: canal mal escrito, permisos, o ahorro de bateria.');
      }
    } catch (e) {
      L.push('  No se pudo conectar: ' + e.message);
    }
  }
  L.push('');

  // ---- Telegram ----
  var token = String(_obtenerConfigValor('PUSH_TELEGRAM_TOKEN') || '').trim();
  var chats = String(_obtenerConfigValor('PUSH_TELEGRAM_CHAT') || '').trim();
  L.push('TELEGRAM');
  if (!token || !chats) {
    L.push('  Sin configurar (token o chat vacio).');
  } else {
    L.push('  Chat: ' + chats);
    var r = _probarTelegram(token, chats);
    L.push('  ' + (r.ok ? 'Enviado OK' : 'ERROR: ' + r.detalle));
  }
  L.push('');
  L.push('TRUCO: abre en el navegador  https://ntfy.sh/' + topic.trim());
  L.push('Ahi ves los mensajes que SI estan llegando al canal, en vivo.');
  L.push('Si los ves ahi pero no en el telefono, el problema es del telefono.');

  return { success: true, texto: L.join('\n') };
}

/** Envia la prueba por Telegram devolviendo el error real si algo falla. */
function _probarTelegram(token, chats) {
  var texto = '<b>Prueba de aviso</b>\nSi ves esto en tu telefono, los avisos del Concierge estan funcionando.';
  var ultimo = 'sin respuesta';
  var alguno = false;
  chats.split(',').forEach(function (chat) {
    chat = String(chat).trim();
    if (!chat) return;
    try {
      var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ chat_id: chat, text: texto, parse_mode: 'HTML' })
      });
      var cuerpo = JSON.parse(res.getContentText() || '{}');
      if (cuerpo.ok) alguno = true;
      else ultimo = cuerpo.description || res.getContentText();
    } catch (e) { ultimo = e.message; }
  });
  return alguno ? { ok: true } : { ok: false, detalle: ultimo };
}

/**
 * Ayuda para encontrar el ID del chat: escribe cualquier mensaje en el grupo
 * (o al bot) y despues ejecuta esto. Devuelve los chats que vieron al bot.
 * @param {string} email
 * @return {Object} {success, mensaje, chats}
 */
function detectarChatsTelegram(email) {
  if (!_validarRolPermitido(email, ['RECEPCION', 'ADMINISTRADOR', 'COCINA', 'RESTAURANT'])) {
    return { success: false, mensaje: 'No tienes permisos.', chats: [] };
  }
  var token = String(_obtenerConfigValor('PUSH_TELEGRAM_TOKEN') || '').trim();
  if (!token) return { success: false, mensaje: 'Primero pega el token del bot.', chats: [] };
  try {
    var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates',
                                { muteHttpExceptions: true });
    var cuerpo = JSON.parse(res.getContentText() || '{}');
    if (!cuerpo.ok) {
      return { success: false, mensaje: 'Telegram respondio: ' + (cuerpo.description || 'error'), chats: [] };
    }
    var vistos = {}, chats = [];
    (cuerpo.result || []).forEach(function (u) {
      var c = (u.message && u.message.chat) || (u.channel_post && u.channel_post.chat);
      if (!c || vistos[c.id]) return;
      vistos[c.id] = true;
      chats.push({ id: String(c.id), nombre: c.title || c.first_name || c.username || ('Chat ' + c.id) });
    });
    if (!chats.length) {
      return { success: false, chats: [],
        mensaje: 'No vi ningun chat. Escribe cualquier mensaje en el grupo (o al bot) y vuelve a intentar.' };
    }
    return { success: true, mensaje: 'Encontre ' + chats.length + ' chat(s).', chats: chats };
  } catch (e) {
    return { success: false, mensaje: 'No pude consultar Telegram: ' + e.message, chats: [] };
  }
}

/**
 * Borra las notificaciones con mas de DIAS_RETENER_NOTIFICACIONES dias, para
 * que la hoja no crezca sin limite y el polling cada 20s siga siendo rapido.
 * Corre como mucho una vez por hora (guardada por cache) y reescribe la hoja
 * en UNA sola operacion (clear + setValues) en lugar de borrar fila por fila.
 * Nunca rompe el flujo: cualquier error se traga.
 */
function _purgarNotificacionesAntiguas() {
  try {
    var cache = CacheService.getScriptCache();
    if (cache.get('purgaNotifReciente')) return;
    cache.put('purgaNotifReciente', '1', 3600); // a lo mas 1 vez por hora
  } catch (e) { return; }

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(3000)) return; // si otra ejecucion la esta haciendo, se salta
    var hoja = _hoja(HOJAS.NOTIFICACIONES);
    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return;
    var encabezados = datos[0];
    var colTs = encabezados.indexOf('Timestamp');
    if (colTs === -1) return;

    var limite = new Date();
    limite.setDate(limite.getDate() - DIAS_RETENER_NOTIFICACIONES);

    var sobreviven = [encabezados];
    for (var i = 1; i < datos.length; i++) {
      var ts = datos[i][colTs];
      var fecha = ts instanceof Date ? ts : new Date(ts);
      // Si la fecha es invalida, se conserva (mejor no borrar por las dudas).
      if (isNaN(fecha.getTime()) || fecha >= limite) sobreviven.push(datos[i]);
    }
    if (sobreviven.length === datos.length) return; // nada que purgar

    hoja.clearContents();
    hoja.getRange(1, 1, sobreviven.length, encabezados.length).setValues(sobreviven);
    _lecturaCache[HOJAS.NOTIFICACIONES] = null; // por si acaso (no es cacheable, pero limpio)
  } catch (e) {
    /* la purga nunca debe romper una lectura de notificaciones */
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Marca una notificacion como leida.
 * @param {string} id
 * @return {Object} {success}
 */
function marcarNotificacionLeida(id) {
  var hoja = _hoja(HOJAS.NOTIFICACIONES);
  var datos = _leerHojaComoObjetos(HOJAS.NOTIFICACIONES);
  for (var i = 0; i < datos.length; i++) {
    if (datos[i].ID === id) {
      hoja.getRange(datos[i]._fila, _indiceColumna(hoja, 'Leida') + 1).setValue('TRUE');
      return { success: true };
    }
  }
  return { success: false, mensaje: 'Notificacion no encontrada.' };
}

/**
 * Marca como leidas todas las notificaciones de un rol.
 * @param {string} rol
 * @return {Object} {success, marcadas}
 */
function marcarTodasNotificacionesLeidas(rol) {
  var hoja = _hoja(HOJAS.NOTIFICACIONES);
  var datos = _leerHojaComoObjetos(HOJAS.NOTIFICACIONES);
  var col = _indiceColumna(hoja, 'Leida') + 1;
  var marcadas = 0;
  datos.forEach(function (d) {
    var paraRol = _esNotifParaRol(d, rol);
    if (paraRol && !_aBooleano(d.Leida)) {
      hoja.getRange(d._fila, col).setValue('TRUE');
      marcadas++;
    }
  });
  return { success: true, marcadas: marcadas };
}

// ===========================================================================
// 6.6. BLOQUEOS Y EVENTOS
// ===========================================================================

/** Auto-reparacion: asegura la columna Notas en EventosBloqueos (nota o
 *  descripcion libre del evento/bloqueo). Vacia por defecto. */
function _asegurarColumnaNotasEventos() {
  _asegurarColumna(_hoja(HOJAS.EVENTOS_BLOQUEOS), 'Notas', '');
}

/**
 * Devuelve bloqueos en un rango que afecten a un servicio (o todos).
 * @param {string} fechaInicio
 * @param {string} fechaFin
 * @param {string} servicioID Opcional.
 * @return {Array<Object>}
 */
function obtenerBloqueos(fechaInicio, fechaFin, servicioID) {
  _asegurarColumnaNotasEventos();
  var desde = _fechaISO(fechaInicio);
  var hasta = _fechaISO(fechaFin);
  return _leerHojaComoObjetos(HOJAS.EVENTOS_BLOQUEOS).filter(function (b) {
    var bDesde = _fechaISO(b.FechaInicio);
    var bHasta = _fechaISO(b.FechaFin || b.FechaInicio);
    // Solapamiento de rangos.
    if (bHasta < desde || bDesde > hasta) return false;
    if (servicioID && !_bloqueoAplicaAServicio(b.ServicioID, servicioID)) return false;
    return true;
  }).map(function (b) {
    return {
      ID: b.ID,
      Tipo: b.Tipo,
      ServicioID: b.ServicioID ? String(b.ServicioID) : '',
      FechaInicio: _fechaISO(b.FechaInicio),
      FechaFin: _fechaISO(b.FechaFin || b.FechaInicio),
      HoraInicio: _horaATexto(b.HoraInicio),
      HoraFin: _horaATexto(b.HoraFin),
      Motivo: b.Motivo,
      Notas: b.Notas ? String(b.Notas) : '',
      CreadoPor: b.CreadoPor,
      Timestamp: _fechaHoraTexto(b.Timestamp)
    };
  });
}

/**
 * Crea un bloqueo/evento. Solo roles RECEPCION o ADMINISTRADOR.
 * @param {Object} datos {tipo, servicioID, fechaInicio, fechaFin, horaInicio, horaFin, motivo, email}
 * @return {Object} {success, id, mensaje}
 */
function crearBloqueo(datos) {
  if (!_validarRolPermitido(datos.email, ['RECEPCION', 'ADMINISTRADOR'])) {
    return { success: false, mensaje: 'No tienes permisos para crear bloqueos.' };
  }
  if (!datos.fechaInicio) return { success: false, mensaje: 'Falta la fecha de inicio.' };

  _asegurarColumnaNotasEventos();
  var hoja = _hoja(HOJAS.EVENTOS_BLOQUEOS);
  var id = generarID();
  hoja.appendRow([
    id, datos.tipo || 'Bloqueo', datos.servicioID || '', _fechaISO(datos.fechaInicio),
    _fechaISO(datos.fechaFin || datos.fechaInicio), datos.horaInicio || '', datos.horaFin || '',
    datos.motivo || '', datos.email || '', new Date()
  ]);
  // La nota va por nombre de columna (la columna Notas se agrego despues del
  // orden posicional original de la hoja).
  if (datos.notas) {
    var colNotas = _indiceColumna(hoja, 'Notas');
    if (colNotas !== -1) hoja.getRange(hoja.getLastRow(), colNotas + 1).setValue(datos.notas);
  }
  registrarLog('Crear bloqueo', datos.tipo + ' ' + (datos.servicioID || 'TODOS'), '');
  return { success: true, id: id, mensaje: 'Bloqueo creado.' };
}

/**
 * Edita un bloqueo/evento existente. Solo RECEPCION o ADMINISTRADOR.
 * @param {Object} datos {id, tipo, servicioID, fechaInicio, fechaFin, horaInicio, horaFin, motivo, email}
 */
function editarBloqueo(datos) {
  if (!_validarRolPermitido(datos.email, ['RECEPCION', 'ADMINISTRADOR'])) {
    return { success: false, mensaje: 'No tienes permisos para editar bloqueos.' };
  }
  if (!datos.id) return { success: false, mensaje: 'Falta el bloqueo a editar.' };
  if (!datos.fechaInicio) return { success: false, mensaje: 'Falta la fecha de inicio.' };

  _asegurarColumnaNotasEventos();
  var hoja = _hoja(HOJAS.EVENTOS_BLOQUEOS);
  var filas = _leerHojaComoObjetos(HOJAS.EVENTOS_BLOQUEOS);
  var fila = filas.filter(function (b) { return b.ID === datos.id; })[0];
  if (!fila) return { success: false, mensaje: 'Bloqueo no encontrado.' };

  var set = function (col, val) {
    var i = _indiceColumna(hoja, col);
    if (i !== -1) hoja.getRange(fila._fila, i + 1).setValue(val);
  };
  set('Tipo', datos.tipo || 'Bloqueo');
  set('ServicioID', datos.servicioID || '');
  set('FechaInicio', _fechaISO(datos.fechaInicio));
  set('FechaFin', _fechaISO(datos.fechaFin || datos.fechaInicio));
  set('HoraInicio', datos.horaInicio || '');
  set('HoraFin', datos.horaFin || '');
  set('Motivo', datos.motivo || '');
  set('Notas', datos.notas || '');
  registrarLog('Editar bloqueo', datos.tipo + ' ' + (datos.servicioID || 'TODOS'), '');
  return { success: true, mensaje: 'Bloqueo actualizado.' };
}

/** Elimina un bloqueo/evento. Solo RECEPCION o ADMINISTRADOR. */
function eliminarBloqueo(id, email) {
  if (!_validarRolPermitido(email, ['RECEPCION', 'ADMINISTRADOR'])) {
    return { success: false, mensaje: 'No tienes permisos para eliminar bloqueos.' };
  }
  var hoja = _hoja(HOJAS.EVENTOS_BLOQUEOS);
  var fila = _leerHojaComoObjetos(HOJAS.EVENTOS_BLOQUEOS).filter(function (b) { return b.ID === id; })[0];
  if (!fila) return { success: false, mensaje: 'Bloqueo no encontrado.' };
  hoja.deleteRow(fila._fila);
  registrarLog('Eliminar bloqueo', id, '');
  return { success: true, mensaje: 'Bloqueo eliminado.' };
}

// ===========================================================================
// 6.6.b PARTICIPANTES (reservas de servicios y eventos)
// ===========================================================================
// Lista libre de personas que se van sumando a una reserva de servicio o a un
// evento (ej. una salida de trekking agendada a cierta hora, donde se anota
// gente de a poco). Solo lo maneja el personal, el huesped no la ve ni la
// edita. tipo: 'reserva' (Reservas.ID) o 'evento' (EventosBloqueos.ID).

/** Auto-reparacion: crea la hoja Participantes si todavia no existe, y le
 *  asegura la columna Notas (agregada despues de la primera version). */
function _asegurarHojaParticipantes() {
  var ss = _ss();
  var hoja = ss.getSheetByName(HOJAS.PARTICIPANTES);
  if (!hoja) {
    hoja = ss.insertSheet(HOJAS.PARTICIPANTES);
    hoja.getRange(1, 1, 1, 9).setValues([[
      'ID', 'Tipo', 'RefID', 'Nombre', 'Telefono', 'Habitacion', 'Notas', 'AgregadoPor', 'Timestamp'
    ]]);
    hoja.setFrozenRows(1);
    return;
  }
  _asegurarColumna(hoja, 'Notas', '');
}

/**
 * Devuelve los participantes de una reserva o evento. Solo personal.
 * @param {string} tipo 'reserva' o 'evento'.
 * @param {string} refID ID de la reserva o del evento/bloqueo.
 * @param {string} email
 * @return {Array<Object>}
 */
function obtenerParticipantes(tipo, refID, email) {
  if (!_validarRolPermitido(email, ['RECEPCION', 'ADMINISTRADOR', 'COCINA', 'RESTAURANT'])) return [];
  _asegurarHojaParticipantes();
  return _leerHojaComoObjetos(HOJAS.PARTICIPANTES).filter(function (p) {
    return p.Tipo === tipo && p.RefID === refID;
  }).map(_normalizarParticipante);
}

/** Normaliza una fila de participante a objeto transferible al cliente. */
function _normalizarParticipante(p) {
  return {
    ID: p.ID,
    Nombre: p.Nombre ? String(p.Nombre) : '',
    Telefono: p.Telefono ? String(p.Telefono) : '',
    Habitacion: p.Habitacion ? String(p.Habitacion) : '',
    Notas: p.Notas ? String(p.Notas) : ''
  };
}

/**
 * Mapa de TODOS los participantes agrupados por RefID (reserva o evento). Se usa
 * para que el Centro de Operaciones muestre la lista dentro de cada burbuja sin
 * una llamada por reserva. Clave = RefID. NO lee la hoja si aun no existe.
 * @return {Object} { refID: [ {ID, Nombre, Telefono, Habitacion, Notas} ] }
 */
function _mapaParticipantes() {
  var ss = _ss();
  if (!ss.getSheetByName(HOJAS.PARTICIPANTES)) return {};
  var mapa = {};
  _leerHojaComoObjetos(HOJAS.PARTICIPANTES).forEach(function (p) {
    if (!p.RefID) return;
    if (!mapa[p.RefID]) mapa[p.RefID] = [];
    mapa[p.RefID].push(_normalizarParticipante(p));
  });
  return mapa;
}

/**
 * Agrega un participante a una reserva o evento. Solo personal.
 * @param {Object} datos {tipo, refID, nombre, telefono, habitacion, notas, email}
 * @return {Object} {success, mensaje}
 */
function agregarParticipante(datos) {
  if (!_validarRolPermitido(datos.email, ['RECEPCION', 'ADMINISTRADOR', 'COCINA', 'RESTAURANT'])) {
    return { success: false, mensaje: 'No tienes permisos para agregar participantes.' };
  }
  if (!datos.tipo || !datos.refID) return { success: false, mensaje: 'Falta la reserva o el evento.' };
  if (!datos.nombre) return { success: false, mensaje: 'Falta el nombre del participante.' };
  _asegurarHojaParticipantes();
  var hoja = _hoja(HOJAS.PARTICIPANTES);
  var id = generarID();
  // Escribe respetando el orden real de columnas (la hoja puede haber sido
  // creada antes de agregar "Notas", ya reparada por _asegurarHojaParticipantes).
  var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var valores = {
    ID: id, Tipo: datos.tipo, RefID: datos.refID, Nombre: datos.nombre,
    Telefono: datos.telefono || '', Habitacion: datos.habitacion || '',
    Notas: datos.notas || '', AgregadoPor: datos.email || '', Timestamp: new Date()
  };
  var fila = encabezados.map(function (col) { return valores[col] !== undefined ? valores[col] : ''; });
  hoja.appendRow(fila);
  registrarLog('Agregar participante', datos.nombre, datos.habitacion || '');
  return { success: true, mensaje: 'Participante agregado.' };
}

/**
 * Edita un participante existente. Solo personal.
 * @param {Object} datos {id, nombre, telefono, habitacion, notas, email}
 * @return {Object} {success, mensaje}
 */
function editarParticipante(datos) {
  if (!_validarRolPermitido(datos.email, ['RECEPCION', 'ADMINISTRADOR', 'COCINA', 'RESTAURANT'])) {
    return { success: false, mensaje: 'No tienes permisos para editar participantes.' };
  }
  if (!datos.id) return { success: false, mensaje: 'Falta el participante a editar.' };
  if (!datos.nombre) return { success: false, mensaje: 'Falta el nombre del participante.' };
  _asegurarHojaParticipantes();
  var hoja = _hoja(HOJAS.PARTICIPANTES);
  var fila = _leerHojaComoObjetos(HOJAS.PARTICIPANTES).filter(function (p) { return p.ID === datos.id; })[0];
  if (!fila) return { success: false, mensaje: 'Participante no encontrado.' };
  var set = function (col, val) {
    var i = _indiceColumna(hoja, col);
    if (i !== -1) hoja.getRange(fila._fila, i + 1).setValue(val);
  };
  set('Nombre', datos.nombre);
  set('Telefono', datos.telefono || '');
  set('Habitacion', datos.habitacion || '');
  set('Notas', datos.notas || '');
  registrarLog('Editar participante', datos.nombre, datos.habitacion || '');
  return { success: true, mensaje: 'Participante actualizado.' };
}

/** Elimina un participante. Solo personal. */
function eliminarParticipante(participanteID, email) {
  if (!_validarRolPermitido(email, ['RECEPCION', 'ADMINISTRADOR', 'COCINA', 'RESTAURANT'])) {
    return { success: false, mensaje: 'No tienes permisos para quitar participantes.' };
  }
  _asegurarHojaParticipantes();
  var hoja = _hoja(HOJAS.PARTICIPANTES);
  var fila = _leerHojaComoObjetos(HOJAS.PARTICIPANTES).filter(function (p) { return p.ID === participanteID; })[0];
  if (!fila) return { success: false, mensaje: 'Participante no encontrado.' };
  hoja.deleteRow(fila._fila);
  registrarLog('Quitar participante', participanteID, '');
  return { success: true, mensaje: 'Participante eliminado.' };
}

// ===========================================================================
// 6.7. CRUD DE PRODUCTOS (GESTION DE CARTA)
// ===========================================================================

/**
 * Crea o actualiza un producto. Rol COCINA, RESTAURANT o ADMINISTRADOR. NO incluye imagen.
 * @param {Object} datos {id, categoriaID, nombre, descripcion, precio, disponible,
 *   visible, orden, etiquetas, tiempoPreparacionMin, esMenuDelDia, email}
 * @return {Object} {success, id, mensaje}
 */
function guardarProducto(datos) {
  if (!_validarRolPermitido(datos.email, ['COCINA', 'RESTAURANT', 'ADMINISTRADOR'])) {
    return { success: false, mensaje: 'No tienes permisos para editar la carta.' };
  }
  if (!datos.nombre || !datos.categoriaID) {
    return { success: false, mensaje: 'Nombre y categoria son obligatorios.' };
  }

  _asegurarColumnaPuntoCoccion();
  var hoja = _hoja(HOJAS.PRODUCTOS);
  var ahora = new Date();
  var modificadoPor = datos.email || 'sistema';

  // Booleanos como texto para consistencia con la hoja.
  var disponible = datos.disponible ? 'TRUE' : 'FALSE';
  var visible = datos.visible ? 'TRUE' : 'FALSE';
  var esMenu = datos.esMenuDelDia ? 'TRUE' : 'FALSE';
  // Por defecto TRUE (solo se guarda FALSE si el formulario lo desmarca a proposito).
  var requierePunto = datos.requierePuntoCoccion === false ? 'FALSE' : 'TRUE';
  var precio = Number(datos.precio) || 0;
  var orden = Number(datos.orden) || 0;
  var tiempo = Number(datos.tiempoPreparacionMin) || 0;

  if (datos.id) {
    // Actualiza producto existente.
    var existente = null;
    var productos = _leerHojaComoObjetos(HOJAS.PRODUCTOS);
    for (var i = 0; i < productos.length; i++) {
      if (productos[i].ID === datos.id) { existente = productos[i]; break; }
    }
    if (!existente) return { success: false, mensaje: 'Producto no encontrado.' };
    var fila = existente._fila;
    // Escritura por LOTE: lee la fila una vez, aplica los cambios por nombre de
    // columna y escribe una sola vez (antes eran ~14 escrituras individuales).
    var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var rango = hoja.getRange(fila, 1, 1, encabezados.length);
    var valores = rango.getValues()[0];
    var cambios = {
      CategoriaID: datos.categoriaID, Nombre: datos.nombre,
      Descripcion: datos.descripcion || '', NombreEN: datos.nombreEN || '',
      DescripcionEN: datos.descripcionEN || '', Precio: precio,
      Disponible: disponible, Visible: visible, Orden: orden,
      Etiquetas: datos.etiquetas || '', TiempoPreparacionMin: tiempo,
      EsMenuDelDia: esMenu, RequierePuntoCoccion: requierePunto,
      FechaModificacion: ahora, ModificadoPor: modificadoPor
    };
    Object.keys(cambios).forEach(function (k) {
      var idx = encabezados.indexOf(k);
      if (idx !== -1) valores[idx] = cambios[k];
    });
    rango.setValues([valores]);
    _invalidarCaches(HOJAS.PRODUCTOS);
    registrarLog('Editar producto', datos.id + ' ' + datos.nombre, '');
    return { success: true, id: datos.id, mensaje: 'Producto actualizado.' };
  } else {
    // Crea producto nuevo: arma la fila completa segun los encabezados reales
    // (incluye NombreEN/DescripcionEN si existen) y la inserta en una escritura.
    var nuevoID = _generarProximoProductoID();
    var encabezadosN = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var nuevo = {
      ID: nuevoID, CategoriaID: datos.categoriaID, Nombre: datos.nombre,
      Descripcion: datos.descripcion || '', NombreEN: datos.nombreEN || '',
      DescripcionEN: datos.descripcionEN || '', Precio: precio,
      Disponible: disponible, Visible: visible, Orden: orden,
      Etiquetas: datos.etiquetas || '', TiempoPreparacionMin: tiempo,
      EsMenuDelDia: esMenu, RequierePuntoCoccion: requierePunto,
      FechaModificacion: ahora, ModificadoPor: modificadoPor
    };
    hoja.appendRow(encabezadosN.map(function (h) { return (h in nuevo) ? nuevo[h] : ''; }));
    _invalidarCaches(HOJAS.PRODUCTOS);
    registrarLog('Crear producto', nuevoID + ' ' + datos.nombre, '');
    return { success: true, id: nuevoID, mensaje: 'Producto creado.' };
  }
}

/** Genera el proximo ID de producto tipo P###. */
function _generarProximoProductoID() {
  var productos = _leerHojaComoObjetos(HOJAS.PRODUCTOS);
  var max = 0;
  productos.forEach(function (p) {
    var m = String(p.ID).match(/^P(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return 'P' + _pad3(max + 1);
}

/** Rellena a 3 digitos. */
function _pad3(n) {
  if (n < 10) return '00' + n;
  if (n < 100) return '0' + n;
  return String(n);
}

/**
 * Soft delete de producto: Visible=FALSE y Disponible=FALSE.
 * @param {string} productoID
 * @param {string} email
 * @return {Object} {success, mensaje}
 */
function eliminarProducto(productoID, email) {
  if (!_validarRolPermitido(email, ['COCINA', 'RESTAURANT', 'ADMINISTRADOR'])) {
    return { success: false, mensaje: 'No tienes permisos.' };
  }
  var hoja = _hoja(HOJAS.PRODUCTOS);
  var productos = _leerHojaComoObjetos(HOJAS.PRODUCTOS);
  for (var i = 0; i < productos.length; i++) {
    if (productos[i].ID === productoID) {
      var fila = productos[i]._fila;
      hoja.getRange(fila, _indiceColumna(hoja, 'Visible') + 1).setValue('FALSE');
      hoja.getRange(fila, _indiceColumna(hoja, 'Disponible') + 1).setValue('FALSE');
      hoja.getRange(fila, _indiceColumna(hoja, 'FechaModificacion') + 1).setValue(new Date());
      hoja.getRange(fila, _indiceColumna(hoja, 'ModificadoPor') + 1).setValue(email || 'sistema');
      _invalidarCaches(HOJAS.PRODUCTOS);
      registrarLog('Eliminar producto (soft)', productoID, '');
      return { success: true, mensaje: 'Producto ocultado.' };
    }
  }
  return { success: false, mensaje: 'Producto no encontrado.' };
}

/**
 * Devuelve un producto para poblar el formulario de edicion.
 * @param {string} productoID
 * @return {Object|null}
 */
function obtenerProductoParaEditar(productoID) {
  _asegurarColumnaPuntoCoccion();
  var productos = _leerHojaComoObjetos(HOJAS.PRODUCTOS);
  for (var i = 0; i < productos.length; i++) {
    if (productos[i].ID === productoID) return _normalizarProducto(productos[i]);
  }
  return null;
}

/**
 * Cambia solo el campo Disponible (toggle "Agotado hoy").
 * @param {string} productoID
 * @param {boolean} nuevoEstado
 * @param {string} email
 * @return {Object} {success, mensaje}
 */
function toggleDisponibleProducto(productoID, nuevoEstado, email) {
  if (!_validarRolPermitido(email, ['COCINA', 'RESTAURANT', 'ADMINISTRADOR'])) {
    return { success: false, mensaje: 'No tienes permisos.' };
  }
  var hoja = _hoja(HOJAS.PRODUCTOS);
  var productos = _leerHojaComoObjetos(HOJAS.PRODUCTOS);
  for (var i = 0; i < productos.length; i++) {
    if (productos[i].ID === productoID) {
      var fila = productos[i]._fila;
      hoja.getRange(fila, _indiceColumna(hoja, 'Disponible') + 1).setValue(nuevoEstado ? 'TRUE' : 'FALSE');
      hoja.getRange(fila, _indiceColumna(hoja, 'FechaModificacion') + 1).setValue(new Date());
      hoja.getRange(fila, _indiceColumna(hoja, 'ModificadoPor') + 1).setValue(email || 'sistema');
      _invalidarCaches(HOJAS.PRODUCTOS);
      registrarLog('Toggle disponible', productoID + ' -> ' + nuevoEstado, '');
      return { success: true, mensaje: 'Disponibilidad actualizada.' };
    }
  }
  return { success: false, mensaje: 'Producto no encontrado.' };
}

/**
 * Devuelve todos los productos (incluidos no visibles) para gestion.
 * @return {Array<Object>}
 */
function obtenerProductosGestion() {
  _asegurarColumnaPuntoCoccion();
  return _leerHojaComoObjetos(HOJAS.PRODUCTOS)
    .map(_normalizarProducto)
    .sort(function (a, b) { return a.Orden - b.Orden; });
}

/**
 * Devuelve los productos "86" (agotados hoy): siguen visibles en la carta pero
 * Disponible=FALSE. Se usa en el widget del Centro de Operaciones para que
 * los garzones vean rapido que no hay en el restaurant en este momento.
 * @return {Array<Object>} [{ID, Nombre, NombreEN, CategoriaID, CategoriaNombre}]
 */
function obtenerProductosAgotados() {
  var categorias = {};
  _leerHojaComoObjetos(HOJAS.CATEGORIAS).forEach(function (c) { categorias[c.ID] = c.Nombre; });

  return _leerHojaComoObjetos(HOJAS.PRODUCTOS)
    .filter(function (p) { return !_aBooleano(p.Disponible) && _aBooleano(p.Visible); })
    .map(function (p) {
      return {
        ID: p.ID,
        Nombre: p.Nombre,
        NombreEN: p.NombreEN ? String(p.NombreEN) : '',
        CategoriaID: p.CategoriaID,
        CategoriaNombre: categorias[p.CategoriaID] || 'Otros',
        Orden: Number(p.Orden) || 0
      };
    })
    .sort(function (a, b) {
      if (a.CategoriaNombre !== b.CategoriaNombre) return a.CategoriaNombre < b.CategoriaNombre ? -1 : 1;
      return a.Orden - b.Orden;
    });
}

// ===========================================================================
// 6.8. CRUD DE CATEGORIAS (SOLO ADMIN)
// ===========================================================================

/**
 * Crea o actualiza una categoria. Rol ADMINISTRADOR.
 * @param {Object} datos {id, nombre, orden, visible, iconoFontAwesome, color, email}
 * @return {Object} {success, id, mensaje}
 */
function guardarCategoria(datos) {
  if (!_validarRolPermitido(datos.email, ['ADMINISTRADOR'])) {
    return { success: false, mensaje: 'Solo el administrador puede gestionar categorias.' };
  }
  if (!datos.nombre) return { success: false, mensaje: 'El nombre es obligatorio.' };

  var hoja = _hoja(HOJAS.CATEGORIAS);
  var visible = datos.visible ? 'TRUE' : 'FALSE';
  var orden = Number(datos.orden) || 0;

  if (datos.id) {
    var cats = _leerHojaComoObjetos(HOJAS.CATEGORIAS);
    for (var i = 0; i < cats.length; i++) {
      if (cats[i].ID === datos.id) {
        var fila = cats[i]._fila;
        hoja.getRange(fila, _indiceColumna(hoja, 'Nombre') + 1).setValue(datos.nombre);
        _setSiExisteCol(hoja, fila, 'NombreEN', datos.nombreEN || '');
        hoja.getRange(fila, _indiceColumna(hoja, 'Orden') + 1).setValue(orden);
        hoja.getRange(fila, _indiceColumna(hoja, 'Visible') + 1).setValue(visible);
        hoja.getRange(fila, _indiceColumna(hoja, 'IconoFontAwesome') + 1).setValue(datos.iconoFontAwesome || 'fa-utensils');
        hoja.getRange(fila, _indiceColumna(hoja, 'Color') + 1).setValue(datos.color || '#414143');
        _invalidarCaches(HOJAS.CATEGORIAS);
        registrarLog('Editar categoria', datos.id, '');
        return { success: true, id: datos.id, mensaje: 'Categoria actualizada.' };
      }
    }
    return { success: false, mensaje: 'Categoria no encontrada.' };
  } else {
    var nuevoID = _generarProximaCategoriaID();
    hoja.appendRow([nuevoID, datos.nombre, orden, visible,
      datos.iconoFontAwesome || 'fa-utensils', datos.color || '#414143']);
    _setSiExisteCol(hoja, hoja.getLastRow(), 'NombreEN', datos.nombreEN || '');
    _invalidarCaches(HOJAS.CATEGORIAS);
    registrarLog('Crear categoria', nuevoID + ' ' + datos.nombre, '');
    return { success: true, id: nuevoID, mensaje: 'Categoria creada.' };
  }
}

/** Genera el proximo ID de categoria tipo CAT##. */
function _generarProximaCategoriaID() {
  var cats = _leerHojaComoObjetos(HOJAS.CATEGORIAS);
  var max = 0;
  cats.forEach(function (c) {
    var m = String(c.ID).match(/^CAT(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  var n = max + 1;
  return 'CAT' + (n < 10 ? '0' + n : String(n));
}

/**
 * Soft delete de categoria: Visible=FALSE. Rol ADMINISTRADOR.
 * @param {string} categoriaID
 * @param {string} email
 * @return {Object} {success, mensaje}
 */
function eliminarCategoria(categoriaID, email) {
  if (!_validarRolPermitido(email, ['ADMINISTRADOR'])) {
    return { success: false, mensaje: 'Solo el administrador puede eliminar categorias.' };
  }
  var hoja = _hoja(HOJAS.CATEGORIAS);
  var cats = _leerHojaComoObjetos(HOJAS.CATEGORIAS);
  for (var i = 0; i < cats.length; i++) {
    if (cats[i].ID === categoriaID) {
      hoja.getRange(cats[i]._fila, _indiceColumna(hoja, 'Visible') + 1).setValue('FALSE');
      _invalidarCaches(HOJAS.CATEGORIAS);
      registrarLog('Eliminar categoria (soft)', categoriaID, '');
      return { success: true, mensaje: 'Categoria ocultada.' };
    }
  }
  return { success: false, mensaje: 'Categoria no encontrada.' };
}

// ===========================================================================
// 6.9. HISTORIAL DE PEDIDOS
// ===========================================================================

/**
 * Busca pedidos historicos enriquecidos. Registra la consulta para auditoria.
 * NO genera comprobantes. Solo consulta interna.
 * @param {Object} filtros {habitacion, fechaDesde, fechaHasta, productoID, servicioID, usuarioConsulta}
 * @return {Array<Object>}
 */
function buscarHistorialPedidos(filtros) {
  filtros = filtros || {};

  var productos = {};
  _leerHojaComoObjetos(HOJAS.PRODUCTOS).forEach(function (p) {
    productos[p.ID] = _normalizarProducto(p);
  });
  var servicios = {};
  _leerHojaComoObjetos(HOJAS.SERVICIOS).forEach(function (s) {
    servicios[s.ID] = _normalizarServicio(s);
  });
  var reservas = {};
  _leerHojaComoObjetos(HOJAS.RESERVAS).forEach(function (r) {
    reservas[r.ID] = r;
  });

  // Agrupa detalles por pedido.
  var detallesPorPedido = {};
  _leerHojaComoObjetos(HOJAS.DETALLE_PEDIDOS).forEach(function (d) {
    if (!detallesPorPedido[d.PedidoID]) detallesPorPedido[d.PedidoID] = [];
    detallesPorPedido[d.PedidoID].push(d);
  });

  var resultados = _leerHojaComoObjetos(HOJAS.PEDIDOS).map(function (ped) {
    var reserva = reservas[ped.ReservaID] || {};
    var servicioID = reserva.ServicioID || '';
    var fecha = reserva.Fecha ? _fechaISO(reserva.Fecha) : _fechaISO(ped.Timestamp);
    var detalles = (detallesPorPedido[ped.ID] || []).map(function (d) {
      var prod = productos[d.ProductoID] || {};
      return {
        ProductoID: d.ProductoID,
        Nombre: prod.Nombre || d.ProductoID,
        Cantidad: Number(d.Cantidad) || 0,
        Subtotal: Number(d.Subtotal) || 0
      };
    });
    return {
      PedidoID: ped.ID,
      ReservaID: ped.ReservaID,
      Habitacion: String(ped.Habitacion),
      Fecha: fecha,
      ServicioID: servicioID,
      ServicioNombre: (servicios[servicioID] || {}).Nombre || servicioID,
      Estado: ped.Estado,
      Total: Number(ped.Total) || 0,
      Detalles: detalles,
      Resumen: detalles.map(function (d) { return d.Cantidad + 'x ' + d.Nombre; }).join(', ')
    };
  }).filter(function (r) {
    if (filtros.habitacion && String(r.Habitacion) !== String(filtros.habitacion)) return false;
    if (filtros.fechaDesde && r.Fecha < _fechaISO(filtros.fechaDesde)) return false;
    if (filtros.fechaHasta && r.Fecha > _fechaISO(filtros.fechaHasta)) return false;
    if (filtros.servicioID && r.ServicioID !== filtros.servicioID) return false;
    if (filtros.productoID) {
      var contiene = r.Detalles.some(function (d) { return d.ProductoID === filtros.productoID; });
      if (!contiene) return false;
    }
    return true;
  }).sort(function (a, b) { return a.Fecha < b.Fecha ? 1 : -1; });

  // Registra la consulta para auditoria.
  registrarConsultaHistorial({
    habitacion: filtros.habitacion || '',
    fechaReferencia: filtros.fechaDesde || '',
    servicioID: filtros.servicioID || '',
    tipoBusqueda: 'busqueda_historial',
    detalleResumen: 'Filtros: ' + JSON.stringify(filtros) + ' | Resultados: ' + resultados.length,
    pedidoID: '',
    reservaID: '',
    usuarioConsulta: filtros.usuarioConsulta || Session.getActiveUser().getEmail() || 'desconocido'
  });

  return resultados;
}

/**
 * Inserta una fila de auditoria en HistorialPedidos.
 * @param {Object} datos {habitacion, fechaReferencia, servicioID, tipoBusqueda,
 *   detalleResumen, pedidoID, reservaID, usuarioConsulta}
 * @return {Object} {success, id}
 */
function registrarConsultaHistorial(datos) {
  var id = generarID();
  _hoja(HOJAS.HISTORIAL).appendRow([
    id, new Date(), datos.habitacion || '', _fechaISO(datos.fechaReferencia || new Date()),
    datos.servicioID || '', datos.tipoBusqueda || '', datos.detalleResumen || '',
    datos.pedidoID || '', datos.reservaID || '', datos.usuarioConsulta || ''
  ]);
  return { success: true, id: id };
}

// ===========================================================================
// CONFIGURACION (edicion desde vista Staff)
// ===========================================================================

/**
 * Actualiza una clave de configuracion. Rol RECEPCION, ADMINISTRADOR, COCINA o RESTAURANT.
 * @param {string} clave
 * @param {*} valor
 * @param {string} email
 * @return {Object} {success, mensaje}
 */
function actualizarConfiguracion(clave, valor, email) {
  if (!_validarRolPermitido(email, ['RECEPCION', 'ADMINISTRADOR', 'COCINA', 'RESTAURANT'])) {
    return { success: false, mensaje: 'No tienes permisos para editar la configuracion.' };
  }
  if (String(clave).indexOf('PUSH_') === 0) _asegurarClavesPush(); // por si es la primera vez
  var hoja = _hoja(HOJAS.CONFIGURACION);
  var filas = _leerHojaComoObjetos(HOJAS.CONFIGURACION);
  for (var i = 0; i < filas.length; i++) {
    if (filas[i].Clave === clave) {
      hoja.getRange(filas[i]._fila, _indiceColumna(hoja, 'Valor') + 1).setValue(valor);
      _invalidarCaches(HOJAS.CONFIGURACION);
      // El token del bot no se escribe en el log (es una credencial).
      var enLog = (clave === 'PUSH_TELEGRAM_TOKEN') ? clave + ' = (oculto)' : clave + ' = ' + valor;
      registrarLog('Editar configuracion', enLog, '');
      return { success: true, mensaje: 'Configuracion actualizada.' };
    }
  }
  return { success: false, mensaje: 'Clave no encontrada.' };
}

/**
 * Devuelve TODOS los servicios (activos e inactivos) para el editor de config.
 * @return {Array<Object>}
 */
function obtenerServiciosGestion() {
  _asegurarColumnaVisibleServicios();
  _asegurarColumnaAnticipoServicios();
  _asegurarColumnaParticipantesServicios();
  _asegurarColumnaAvisoServicios();
  return _leerHojaComoObjetos(HOJAS.SERVICIOS).map(_normalizarServicio);
}

/**
 * Edita un servicio existente (todos los campos). Rol RECEPCION, ADMIN, COCINA o RESTAURANT.
 * @param {Object} datos {id, nombre, categoria, descripcion, icono, costoBase,
 *   capacidad, duracionMinutos, horarioInicio, horarioFin, requiereAprobacion,
 *   permitePrepedido, usoExclusivo, activo, visible, variantes:[{nombre,precio}]}
 * @param {string} email
 * @return {Object} {success, mensaje}
 */
function guardarServicioConfig(datos, email) {
  if (!_validarRolPermitido(email, ['RECEPCION', 'ADMINISTRADOR', 'COCINA', 'RESTAURANT'])) {
    return { success: false, mensaje: 'No tienes permisos para editar servicios.' };
  }
  _asegurarColumnaVisibleServicios();
  _asegurarColumnaAnticipoServicios();
  _asegurarColumnaParticipantesServicios();
  _asegurarColumnaAvisoServicios();
  var hoja = _hoja(HOJAS.SERVICIOS);
  var filas = _leerHojaComoObjetos(HOJAS.SERVICIOS);
  for (var i = 0; i < filas.length; i++) {
    if (filas[i].ID === datos.id) {
      _escribirCamposServicio(hoja, filas[i]._fila, datos);
      _invalidarCaches(HOJAS.SERVICIOS);
      registrarLog('Editar servicio', datos.id + ' ' + (datos.nombre || ''), '');
      return { success: true, mensaje: 'Servicio actualizado.' };
    }
  }
  return { success: false, mensaje: 'Servicio no encontrado.' };
}

/**
 * Crea un servicio nuevo. Genera ID automatico (S00X) y escribe la fila.
 * Rol RECEPCION, ADMINISTRADOR, COCINA o RESTAURANT.
 * @param {Object} datos Igual que guardarServicioConfig (sin id).
 * @param {string} email
 * @return {Object} {success, id, mensaje}
 */
function guardarServicioNuevo(datos, email) {
  if (!_validarRolPermitido(email, ['RECEPCION', 'ADMINISTRADOR', 'COCINA', 'RESTAURANT'])) {
    return { success: false, mensaje: 'No tienes permisos para crear servicios.' };
  }
  if (!datos.nombre) return { success: false, mensaje: 'El nombre es obligatorio.' };

  _asegurarColumnaVisibleServicios();
  _asegurarColumnaAnticipoServicios();
  _asegurarColumnaParticipantesServicios();
  _asegurarColumnaAvisoServicios();
  var hoja = _hoja(HOJAS.SERVICIOS);
  // Genera el proximo ID S00X.
  var max = 0;
  _leerHojaComoObjetos(HOJAS.SERVICIOS).forEach(function (s) {
    var m = String(s.ID).match(/^S(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  var nuevoID = 'S' + ('00' + (max + 1)).slice(-3);

  // Fila con la cantidad exacta de columnas actuales (rellena vacio).
  var nCols = hoja.getLastColumn();
  var vacia = [];
  for (var k = 0; k < nCols; k++) vacia.push('');
  vacia[_indiceColumna(hoja, 'ID')] = nuevoID;
  vacia[_indiceColumna(hoja, 'EsIncluible')] = 'FALSE';
  hoja.appendRow(vacia);

  _escribirCamposServicio(hoja, hoja.getLastRow(), datos);
  _invalidarCaches(HOJAS.SERVICIOS);
  registrarLog('Crear servicio', nuevoID + ' ' + datos.nombre, '');
  return { success: true, id: nuevoID, mensaje: 'Servicio creado.' };
}

/** Escribe los campos de un servicio en una fila dada (solo los presentes). */
function _escribirCamposServicio(hoja, fila, datos) {
  // Lee los encabezados UNA vez (antes se releian por cada campo).
  var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  function set(col, valor) {
    var idx = encabezados.indexOf(col);
    if (idx !== -1) hoja.getRange(fila, idx + 1).setValue(valor);
  }
  function setHora(col, valor) {
    var idx = encabezados.indexOf(col);
    if (idx !== -1) hoja.getRange(fila, idx + 1).setNumberFormat('@').setValue(valor);
  }
  if (datos.nombre !== undefined) set('Nombre', datos.nombre);
  if (datos.nombreEN !== undefined) set('NombreEN', datos.nombreEN || '');
  if (datos.categoria !== undefined) set('Categoria', datos.categoria);
  if (datos.descripcion !== undefined) set('Descripcion', datos.descripcion || '');
  if (datos.descripcionEN !== undefined) set('DescripcionEN', datos.descripcionEN || '');
  if (datos.icono !== undefined) set('Icono', datos.icono || 'fa-calendar-check');
  if (datos.costoBase !== undefined) set('CostoBase', Number(datos.costoBase) || 0);
  if (datos.capacidad !== undefined) set('Capacidad', Number(datos.capacidad) || 1);
  if (datos.duracionMinutos !== undefined) set('DuracionMinutos', Number(datos.duracionMinutos) || 30);
  if (datos.horarioInicio) setHora('HorarioInicio', datos.horarioInicio);
  if (datos.horarioFin) setHora('HorarioFin', datos.horarioFin);
  if (datos.requiereAprobacion !== undefined) set('RequiereAprobacion', datos.requiereAprobacion ? 'TRUE' : 'FALSE');
  if (datos.permitePrepedido !== undefined) set('PermitePrepedido', datos.permitePrepedido ? 'TRUE' : 'FALSE');
  if (datos.usoExclusivo !== undefined) set('UsoExclusivo', datos.usoExclusivo ? 'TRUE' : 'FALSE');
  if (datos.activo !== undefined) set('Activo', datos.activo ? 'TRUE' : 'FALSE');
  if (datos.visible !== undefined) set('Visible', datos.visible ? 'TRUE' : 'FALSE');
  if (datos.variantes !== undefined) setHora('Variantes', _serializarVariantes(datos.variantes));
  if (datos.color !== undefined) set('Color', datos.color || '');
  if (datos.anticipoMinimoHoras !== undefined) set('AnticipoMinimoHoras', Number(datos.anticipoMinimoHoras) || 0);
  if (datos.permiteParticipantes !== undefined) set('PermiteParticipantes', datos.permiteParticipantes ? 'TRUE' : 'FALSE');
  if (datos.avisoReserva !== undefined) set('AvisoReserva', datos.avisoReserva || '');
  if (datos.avisoReservaEN !== undefined) set('AvisoReservaEN', datos.avisoReservaEN || '');
}

// ===========================================================================
// GESTION DE PEDIDOS (vista Staff)
// ===========================================================================

/**
 * Devuelve los pedidos del dia con sus detalles (para cocina/recepcion).
 * @param {string} fecha
 * @return {Array<Object>}
 */
function obtenerPedidosDelDia(fecha, reservasPre) {
  fecha = _fechaISO(fecha || new Date());
  var reservas = {};
  (reservasPre || _leerHojaComoObjetos(HOJAS.RESERVAS)).forEach(function (r) { reservas[r.ID] = r; });
  var productos = {};
  _leerHojaComoObjetos(HOJAS.PRODUCTOS).forEach(function (p) { productos[p.ID] = _normalizarProducto(p); });

  var detallesPorPedido = {};
  _leerHojaComoObjetos(HOJAS.DETALLE_PEDIDOS).forEach(function (d) {
    if (!detallesPorPedido[d.PedidoID]) detallesPorPedido[d.PedidoID] = [];
    detallesPorPedido[d.PedidoID].push({
      Nombre: (productos[d.ProductoID] || {}).Nombre || d.ProductoID,
      Cantidad: Number(d.Cantidad) || 0,
      Subtotal: Number(d.Subtotal) || 0,
      Notas: d.Notas
    });
  });

  return _leerHojaComoObjetos(HOJAS.PEDIDOS).filter(function (ped) {
    var reserva = reservas[ped.ReservaID] || {};
    // No mostrar pedidos cuya reserva fue cancelada (ni sin reserva valida).
    if (!reserva.ID || _esEstadoCancelado(reserva.Estado)) return false;
    var fechaRef = reserva.Fecha ? _fechaISO(reserva.Fecha) : _fechaISO(ped.Timestamp);
    return fechaRef === fecha;
  }).map(function (ped) {
    var reserva = reservas[ped.ReservaID] || {};
    return {
      ID: ped.ID,
      ReservaID: ped.ReservaID,
      Habitacion: String(ped.Habitacion),
      Estado: ped.Estado,
      Total: Number(ped.Total) || 0,
      Notas: ped.Notas,
      Entrega: ped.Entrega || 'Restaurant',
      HoraReserva: _horaATexto(reserva.HoraInicio),
      Detalles: detallesPorPedido[ped.ID] || []
    };
  });
}

// ===========================================================================
// 6.10. SEGURIDAD Y UTILIDADES
// ===========================================================================

/**
 * Verifica que un email tenga un rol determinado y este activo.
 * @param {string} email
 * @param {string} rolRequerido
 * @return {boolean}
 */
function validarRol(email, rolRequerido) {
  if (!email) return false;
  var usuarios = _leerHojaComoObjetos(HOJAS.USUARIOS);
  for (var i = 0; i < usuarios.length; i++) {
    if (String(usuarios[i].Email).toLowerCase() === String(email).toLowerCase()) {
      return _aBooleano(usuarios[i].Activo) && usuarios[i].Rol === rolRequerido;
    }
  }
  return false;
}

/** Devuelve el usuario (o null) por email, validando que este activo. */
function obtenerUsuarioPorEmail(email) {
  if (!email) return null;
  var usuarios = _leerHojaComoObjetos(HOJAS.USUARIOS);
  for (var i = 0; i < usuarios.length; i++) {
    if (String(usuarios[i].Email).toLowerCase() === String(email).toLowerCase() &&
        _aBooleano(usuarios[i].Activo)) {
      return {
        ID: usuarios[i].ID,
        Email: usuarios[i].Email,
        Nombre: usuarios[i].Nombre,
        Rol: usuarios[i].Rol,
        HabitacionAsociada: usuarios[i].HabitacionAsociada
      };
    }
  }
  return null;
}

/**
 * Login de staff: valida email contra la hoja Usuarios. Si el usuario tiene
 * una contrasena definida en la columna Password, tambien se exige y valida;
 * si esa columna esta vacia para ese usuario, no se pide contrasena (igual
 * que el comportamiento anterior).
 * @param {string} email
 * @param {string} password
 * @return {Object} {success, usuario, mensaje}
 */
function autenticarStaff(email, password) {
  _asegurarColumnaPasswordUsuarios();
  var usuario = obtenerUsuarioPorEmail(email);
  if (!usuario) return { success: false, mensaje: 'Usuario no encontrado o inactivo.' };

  var usuarios = _leerHojaComoObjetos(HOJAS.USUARIOS);
  var fila = usuarios.filter(function (u) { return String(u.Email).toLowerCase() === String(email).toLowerCase(); })[0];
  var passwordGuardada = fila ? String(fila.Password || '') : '';
  if (passwordGuardada && passwordGuardada !== String(password || '')) {
    return { success: false, mensaje: 'Contrasena incorrecta.' };
  }

  registrarLog('Login staff', usuario.Rol, '');
  return { success: true, usuario: usuario, mensaje: 'Bienvenido ' + usuario.Nombre };
}

/**
 * Lista publica (sin login previo) de los usuarios de staff activos, para
 * pintar los botones de acceso rapido. NUNCA devuelve la contrasena, solo si
 * el usuario tiene una definida (para saber si hay que pedirla o no).
 * @return {Array} [{email, nombre, rol, requierePassword}]
 */
function obtenerUsuariosStaffLogin() {
  _asegurarColumnaPasswordUsuarios();
  var usuarios = _leerHojaComoObjetos(HOJAS.USUARIOS);
  return usuarios
    .filter(function (u) { return _aBooleano(u.Activo); })
    .map(function (u) {
      return {
        email: u.Email, nombre: u.Nombre, rol: u.Rol,
        requierePassword: !!String(u.Password || '').trim()
      };
    });
}

/** Valida que el email pertenezca a alguno de los roles indicados. */
function _validarRolPermitido(email, roles) {
  var usuario = obtenerUsuarioPorEmail(email);
  if (!usuario) return false;
  return roles.indexOf(usuario.Rol) !== -1;
}

/**
 * Genera un UUID simple.
 * @return {string}
 */
function generarID() {
  return Utilities.getUuid();
}

/**
 * Formatea un valor numerico como moneda usando MONEDA_SIMBOLO.
 * @param {number} valor
 * @return {string}
 */
function formatearMoneda(valor) {
  var simbolo = _obtenerConfigValor('MONEDA_SIMBOLO') || '$';
  var n = Number(valor) || 0;
  // Separador de miles con punto (formato CLP).
  var entero = Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return simbolo + entero;
}

/**
 * Inserta una entrada en LogActividad.
 * @param {string} accion
 * @param {string} detalle
 * @param {string} habitacion
 */
function registrarLog(accion, detalle, habitacion) {
  try {
    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (e) { email = ''; }
    _hoja(HOJAS.LOG).appendRow([
      generarID(), new Date(), email, accion, detalle || '', habitacion || ''
    ]);
  } catch (err) {
    // El log nunca debe romper el flujo principal.
  }
}
