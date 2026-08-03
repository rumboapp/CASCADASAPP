/**
 * ============================================================================
 * CASCADAS CONCIERGE — ACTO 1: SETUP DE BASE DE DATOS
 * ============================================================================
 * Archivo: Setup.gs
 *
 * Ejecuta la función `crearBaseDeDatos()` una sola vez desde el editor de
 * Google Apps Script. Creará un nuevo Spreadsheet con las 14 hojas exactas,
 * sus columnas y todos los datos iniciales poblados (incluida la carta
 * completa de 132 productos).
 *
 * Al finalizar, el ID del Spreadsheet se guarda en PropertiesService para que
 * Code.gs lo reutilice automáticamente. También se imprime en el Log.
 *
 * NOTA: Los productos NO llevan columna ImagenURL (sin fotos de platos).
 * ============================================================================
 */

/**
 * Punto de entrada del setup. Crea el Spreadsheet completo.
 * Idempotente en la práctica: si ya existe un ID guardado, avisa y no duplica.
 */
function crearBaseDeDatos() {
  // Evita crear dos veces la base de datos por accidente.
  var idExistente = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (idExistente) {
    Logger.log('Ya existe una base de datos con ID: ' + idExistente);
    Logger.log('Si deseas recrearla, borra la propiedad SPREADSHEET_ID primero.');
    return idExistente;
  }

  // Crea el Spreadsheet contenedor.
  var ss = SpreadsheetApp.create('Cascadas Concierge - Base de Datos');
  var id = ss.getId();

  // Construye cada hoja con sus encabezados y datos iniciales.
  _crearHojaConfiguracion(ss);
  _crearHojaHabitaciones(ss);
  _crearHojaServicios(ss);
  _crearHojaCategoriasProducto(ss);
  _crearHojaProductos(ss);
  _crearHojaReservas(ss);
  _crearHojaPedidos(ss);
  _crearHojaDetallePedidos(ss);
  _crearHojaUsuarios(ss);
  _crearHojaEventosBloqueos(ss);
  _crearHojaNotificaciones(ss);
  _crearHojaDisponibilidadPersonal(ss);
  _crearHojaLogActividad(ss);
  _crearHojaHistorialPedidos(ss);

  // Elimina la hoja "Hoja 1" / "Sheet1" creada por defecto.
  _eliminarHojaPorDefecto(ss);

  // Fuerza los horarios como TEXTO para que Sheets no los convierta a fechas.
  // (_normalizarHorariosComoTexto esta definida en Code.gs, mismo proyecto.)
  _normalizarHorariosComoTexto(ss);

  // Bilingue: agrega columnas EN y aplica las traducciones de fabrica.
  // (aplicarTraduccionesEN esta definida en Code.gs, mismo proyecto.)
  aplicarTraduccionesEN(ss);

  // Persiste el ID para que Code.gs lo use sin configuración manual.
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);

  Logger.log('Base de datos creada con éxito.');
  Logger.log('SPREADSHEET_ID: ' + id);
  Logger.log('URL: ' + ss.getUrl());
  return id;
}

/**
 * Helper interno: crea una hoja con encabezados y filas de datos.
 * Aplica formato de encabezado y congela la primera fila.
 * @param {Spreadsheet} ss Spreadsheet destino.
 * @param {string} nombre Nombre exacto de la hoja.
 * @param {Array<string>} encabezados Columnas exactas.
 * @param {Array<Array>} filas Datos iniciales (puede ir vacío).
 * @return {Sheet} La hoja creada.
 */
function _crearHoja(ss, nombre, encabezados, filas) {
  var hoja = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  hoja.clear();

  // Escribe encabezados.
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);

  // Estiliza la fila de encabezado.
  var rangoEncabezado = hoja.getRange(1, 1, 1, encabezados.length);
  rangoEncabezado.setFontWeight('bold');
  rangoEncabezado.setBackground('#1B3A2F');
  rangoEncabezado.setFontColor('#FFFFFF');
  hoja.setFrozenRows(1);

  // Escribe datos si existen.
  if (filas && filas.length > 0) {
    hoja.getRange(2, 1, filas.length, encabezados.length).setValues(filas);
  }

  hoja.autoResizeColumns(1, encabezados.length);
  return hoja;
}

/** 5.1. Hoja Configuracion */
function _crearHojaConfiguracion(ss) {
  var encabezados = ['Clave', 'Valor', 'Descripcion'];
  var filas = [
    ['DESAYUNO_HORARIO_INICIO', '08:30', 'Horario apertura desayuno buffet'],
    ['DESAYUNO_HORARIO_FIN', '10:30', 'Horario cierre desayuno buffet'],
    ['DESAYUNO_CAPACIDAD', 28, 'Comensales simultaneos maximo'],
    ['DESAYUNO_BLOQUE_MINUTOS', 30, 'Duracion de cada bloque de desayuno'],
    ['DESAYUNO_HABITACION_ACTIVO', 'TRUE', 'Habilita/deshabilita desayuno a habitacion'],
    ['DESAYUNO_HABITACION_COSTO', 15000, 'Costo adicional por desayuno continental en habitacion'],
    ['DESAYUNO_HABITACION_DESCRIPCION', 'Desayuno continental: te o cafe, pan, huevo, mantequilla, mermelada, algo dulce. No es a pedido, se sirve segun disponibilidad del dia.', 'Descripcion visible para huesped'],
    ['ALMUERZO_HORARIO_INICIO', '13:00', 'Horario apertura almuerzo'],
    ['ALMUERZO_HORARIO_FIN', '22:00', 'Horario cierre almuerzo'],
    ['ALMUERZO_CAPACIDAD', 28, 'Comensales simultaneos maximo'],
    ['CENA_HORARIO_INICIO', '19:00', 'Horario apertura cena'],
    ['CENA_HORARIO_FIN', '22:00', 'Horario cierre cena'],
    ['CENA_CAPACIDAD', 28, 'Comensales simultaneos maximo'],
    ['TINAJA_HORARIO_INICIO', '10:00', 'Horario apertura tinaja'],
    ['TINAJA_HORARIO_FIN', '22:00', 'Horario cierre tinaja'],
    ['TINAJA_DURACION_MINUTOS', 120, 'Duracion de la sesion de tinaja'],
    ['TINAJA_CAPACIDAD_MAX', 2, 'Personas maximas por sesion'],
    ['TINAJA_COSTO', 50000, 'Costo base tinaja'],
    ['TINAJA_IVA_PORCENTAJE', 19, 'IVA aplicado a tinaja'],
    ['BICICLETAS_CANTIDAD', 4, 'Cantidad de bicicletas disponibles'],
    ['BICICLETAS_DURACION_MINUTOS', 120, 'Duracion del arriendo'],
    ['BICICLETAS_COSTO', 0, 'Costo bicicletas'],
    ['MASAJE_DURACION_MINUTOS', 50, 'Duracion de la sesion de masaje'],
    ['MASAJE_RELAX_COSTO', 60000, 'Costo masaje de relajacion'],
    ['MASAJE_DESCOST_COSTO', 65000, 'Costo masaje descontracturante'],
    ['MASAJE_IVA_PORCENTAJE', 19, 'IVA aplicado a masajes'],
    ['PREPEDIDO_MINUTOS_LIMITE', 30, 'Minutos antes de la reserva para modificar pedido'],
    ['RESERVA_PASO_MINUTOS', 30, 'Cada cuantos minutos se ofrecen horarios de reserva'],
    ['MAX_DIAS_RESERVA_HUESPED', 3, 'Dias maximos hacia adelante que un huesped puede reservar (recepcion sin limite)'],
    ['PEDIDO_HABITACION_ACTIVO', 'TRUE', 'Habilita/deshabilita la opcion de pedir a la habitacion'],
    ['PEDIDO_HABITACION_RECARGO', 0, 'Valor que se suma al total del pedido cuando la entrega es a la habitacion'],
    ['PREPEDIDO_ACTIVO', 'TRUE', 'Habilita/deshabilita los pedidos anticipados del restaurant'],
    ['HOTEL_NOMBRE', 'Cascadas Hotel', 'Nombre del hotel'],
    ['MODO_OSCURO_DEFAULT', 'FALSE', 'Modo oscuro por defecto'],
    ['MONEDA_SIMBOLO', '$', 'Simbolo de moneda'],
    ['IDIOMA_DEFAULT', 'es', 'Idioma por defecto'],
    // Avisos al telefono con la app cerrada (se configuran desde la app)
    ['PUSH_ACTIVO', 'FALSE', 'Envia los avisos al telefono aunque la app este cerrada (Telegram)'],
    ['PUSH_TG_TOKEN', '', 'Token del bot de Telegram (lo entrega @BotFather)'],
    ['PUSH_TG_GENERAL', '', 'ID del grupo de Telegram que recibe TODOS los avisos'],
    ['PUSH_TG_RESTAURANT', '', 'ID del grupo de Telegram que recibe solo comida (gastronomia y pedidos)'],
    ['PUSH_TG_LINK_GENERAL', '', 'Enlace de invitacion al grupo general (para el QR que se imprime)'],
    ['PUSH_TG_LINK_RESTAURANT', '', 'Enlace de invitacion al grupo del restaurant (para el QR que se imprime)']
  ];
  _crearHoja(ss, 'Configuracion', encabezados, filas);
}

/** 5.2. Hoja Habitaciones (101-110, 201-210) */
function _crearHojaHabitaciones(ss) {
  var encabezados = ['Numero', 'Estado', 'Tipo', 'Capacidad', 'Notas', 'WhatsAppLink'];
  var filas = [];
  var pisos = [100, 200];
  for (var p = 0; p < pisos.length; p++) {
    for (var n = 1; n <= 10; n++) {
      filas.push([String(pisos[p] + n), 'Activa', 'Standard', 2, '', '']);
    }
  }
  _crearHoja(ss, 'Habitaciones', encabezados, filas);
}

/** 5.3. Hoja Servicios */
function _crearHojaServicios(ss) {
  var encabezados = ['ID', 'Nombre', 'Categoria', 'DuracionMinutos', 'Capacidad', 'CostoBase',
    'RequiereAprobacion', 'Activo', 'HorarioInicio', 'HorarioFin', 'Descripcion',
    'PermitePrepedido', 'EsIncluible', 'Icono', 'Variantes', 'UsoExclusivo', 'Color', 'Visible'];
  var filas = [
    ['S001', 'Desayuno Buffet', 'Gastronomia', 30, 28, 0, 'FALSE', 'TRUE', '08:30', '10:30', 'Buffet de la casa', 'FALSE', 'FALSE', 'fa-bowl-food', '', 'FALSE', '#6E7F3E', 'TRUE'],
    ['S002', 'Desayuno Habitacion', 'Gastronomia', 30, 20, 15000, 'TRUE', 'TRUE', '08:30', '10:30', 'Desayuno continental servido en habitacion. No es a pedido.', 'FALSE', 'TRUE', 'fa-mug-saucer', '', 'FALSE', '#A6524B', 'TRUE'],
    ['S003', 'Almuerzo / Cena', 'Gastronomia', 90, 28, 0, 'FALSE', 'TRUE', '13:00', '22:00', 'Restaurant: almuerzo y cena', 'TRUE', 'FALSE', 'fa-utensils', '', 'FALSE', '#9C4A63', 'TRUE'],
    ['S004', 'Cena', 'Gastronomia', 90, 28, 0, 'FALSE', 'FALSE', '19:00', '22:00', 'Fusionado en Almuerzo / Cena (S003)', 'TRUE', 'FALSE', 'fa-wine-glass', '', 'FALSE', '#556070', 'TRUE'],
    ['S005', 'Tinaja', 'Bienestar', 120, 2, 50000, 'TRUE', 'TRUE', '10:00', '22:00', 'Tinaja exterior', 'FALSE', 'TRUE', 'fa-hot-tub-person', '', 'TRUE', '#2E8A86', 'TRUE'],
    ['S006', 'Bicicletas', 'Actividad', 120, 4, 0, 'TRUE', 'TRUE', '10:00', '18:00', 'Aro 27.5', 'FALSE', 'FALSE', 'fa-bicycle', '', 'FALSE', '#3E6FA3', 'TRUE'],
    ['S007', 'Masajes', 'Bienestar', 50, 1, 60000, 'TRUE', 'TRUE', '10:00', '20:00', 'Sesion de masaje', 'FALSE', 'FALSE', 'fa-spa', 'Relajacion:60000|Descontracturante:65000', 'TRUE', '#7E4EA3', 'TRUE']
  ];
  _crearHoja(ss, 'Servicios', encabezados, filas);
}

/** 5.4. Hoja CategoriasProducto */
function _crearHojaCategoriasProducto(ss) {
  var encabezados = ['ID', 'Nombre', 'Orden', 'Visible', 'IconoFontAwesome', 'Color'];
  var filas = [
    ['CAT01', 'Entradas', 1, 'TRUE', 'fa-utensils', '#8B7355'],
    ['CAT02', 'Fondos', 2, 'TRUE', 'fa-drumstick-bite', '#8B7355'],
    ['CAT03', 'Pescados y Mariscos', 3, 'TRUE', 'fa-fish', '#2E5C8A'],
    ['CAT04', 'Carnes', 4, 'TRUE', 'fa-fire', '#8B4513'],
    ['CAT05', 'Ensaladas', 5, 'TRUE', 'fa-leaf', '#6B8E23'],
    ['CAT06', 'Pastas y Salsas', 6, 'TRUE', 'fa-bowl-food', '#D4AF37'],
    ['CAT07', 'Para los Ninos', 7, 'TRUE', 'fa-child', '#FF8C00'],
    ['CAT08', 'Postres', 8, 'TRUE', 'fa-ice-cream', '#FF69B4'],
    ['CAT09', 'Cafeteria', 9, 'TRUE', 'fa-coffee', '#6F4E37'],
    ['CAT10', 'Sandwiches', 10, 'TRUE', 'fa-bread-slice', '#D4AF37'],
    ['CAT11', 'Tablas', 11, 'TRUE', 'fa-cheese', '#D4AF37'],
    ['CAT12', 'Cervezas', 12, 'TRUE', 'fa-beer', '#D4AF37'],
    ['CAT13', 'Gin', 13, 'TRUE', 'fa-glass-martini', '#D4AF37'],
    ['CAT14', 'Vodka', 14, 'TRUE', 'fa-wine-bottle', '#D4AF37'],
    ['CAT15', 'Whisky', 15, 'TRUE', 'fa-glass-whiskey', '#D4AF37'],
    ['CAT16', 'Tequila', 16, 'TRUE', 'fa-lemon', '#D4AF37'],
    ['CAT17', 'Pisco', 17, 'TRUE', 'fa-wine-glass', '#D4AF37'],
    ['CAT18', 'Cocteles', 18, 'TRUE', 'fa-cocktail', '#D4AF37'],
    ['CAT19', 'Vinos', 19, 'TRUE', 'fa-wine-glass-alt', '#722F37'],
    ['CAT20', 'Sin Alcohol', 20, 'TRUE', 'fa-glass-water', '#87CEEB'],
    ['CAT21', 'Bar (General)', 21, 'TRUE', 'fa-glass-martini-alt', '#D4AF37']
  ];
  _crearHoja(ss, 'CategoriasProducto', encabezados, filas);
}

/** 5.5. Hoja Productos (132 productos, SIN ImagenURL) */
function _crearHojaProductos(ss) {
  var encabezados = ['ID', 'CategoriaID', 'Nombre', 'Descripcion', 'Precio', 'Disponible',
    'Visible', 'Orden', 'Etiquetas', 'TiempoPreparacionMin', 'EsMenuDelDia',
    'FechaModificacion', 'ModificadoPor'];
  var filas = [
    // ENTRADAS
    ['P001', 'CAT01', 'Crudo de filete a la antigua', 'Filete picado y mezclado con mayonesa y mostaza antigua, jugo de limon, alcaparras, pepinillos y cebolla. Todo servido con pan de la casa.', 15900, 'TRUE', 'TRUE', 1, '', 15, 'FALSE', '', ''],
    ['P002', 'CAT01', 'Carpaccio de pulpo al olivo', 'Laminas de pulpo acompanadas de salsa al olivo con crostinis y mousse de palta.', 14900, 'TRUE', 'TRUE', 2, '', 15, 'FALSE', '', ''],
    ['P003', 'CAT01', 'Tiradito de salmon con salsa acevichada', 'Cortes de salmon banados en salsa acevichada de maracuya. Todo finalizado con pico de gallo y mousse de camote.', 15900, 'TRUE', 'TRUE', 3, 'Sin Gluten', 15, 'FALSE', '', ''],
    ['P004', 'CAT01', 'Ceviche mixto', 'Atun, camaron y pulpo, marinados en leche de tigre y limon, con cebolla y cilantro.', 19900, 'TRUE', 'TRUE', 4, 'Sin Gluten', 15, 'FALSE', '', ''],
    ['P005', 'CAT01', 'Crema o sopa del dia', 'Setas | Zapallo | Esparragos | Sopa de cebolla | Ajiaco | Otros del dia.', 12900, 'TRUE', 'TRUE', 5, '', 15, 'TRUE', '', ''],
    ['P006', 'CAT01', 'Ceviche de setas y palmitos', 'Mix de champinones y palmitos marinados en leche de tigre, acompanados de camote frito y mousse de palta.', 14900, 'TRUE', 'TRUE', 6, 'Vegano', 15, 'FALSE', '', ''],
    // ENSALADAS
    ['P007', 'CAT05', 'Oriental de salmon ahumado', 'Mix verde, salmon ahumado, esparragos, palta, tomates confitados, bolitas crocantes de queso crema y jalapeno.', 14900, 'TRUE', 'TRUE', 7, 'Sin Gluten', 10, 'FALSE', '', ''],
    ['P008', 'CAT05', 'Cesar de pollo', 'Mix de hojas verdes, pollo grillado, crocante de tocino, crutones, escamas de queso parmesano, dressing Cesar.', 14900, 'TRUE', 'TRUE', 8, '', 10, 'FALSE', '', ''],
    ['P009', 'CAT05', 'Cesar de camaron', 'Mix de hojas verdes, salmon grillado, crocante de tocino, crutones, escamas de queso parmesano, dressing Cesar.', 15900, 'TRUE', 'TRUE', 9, '', 10, 'FALSE', '', ''],
    ['P010', 'CAT05', 'Huerto Las Cascadas', 'Mix de hojas verdes, palmitos, tomates confitados, palta, alcachofas, crostinis, aceitunas y vinagreta de cilantro.', 14900, 'TRUE', 'TRUE', 10, 'Vegano', 10, 'FALSE', '', ''],
    // PESCADOS Y MARISCOS
    ['P011', 'CAT03', 'Salmon grillado con miel y risotto de mote al pesto con mousse de palta', 'Salmon marinado en limon y miel a la plancha, acompanado de risotto de mote al pesto y mousse de palta.', 18900, 'TRUE', 'TRUE', 11, 'Sin Gluten', 25, 'FALSE', '', ''],
    ['P012', 'CAT03', 'Merluza Austral frita con salsa pico de gallo y pure al merquen', 'Crocante merluza frita acompanada de pure casero de papas y un toque de merquen ahumado, con salsa pico de gallo.', 17900, 'TRUE', 'TRUE', 12, '', 25, 'FALSE', '', ''],
    ['P013', 'CAT03', 'Congrio al horno con salsa de camarones y vegetales al Wok', 'Congrio horneado banado en salsa de camarones, servido con vegetales de temporada salteados al wok.', 18900, 'TRUE', 'TRUE', 13, '', 30, 'FALSE', '', ''],
    ['P014', 'CAT03', 'Pastel de jaiba chilota', 'Cremoso pastel de jaiba gratinado con queso parmesano: un clasico del sur.', 18900, 'TRUE', 'TRUE', 14, '', 30, 'FALSE', '', ''],
    // CARNES
    ['P015', 'CAT04', 'Filete acompanado de esparragos salteados y portobellos con peras', 'Medallon de filete envuelto en tocino y banado en salsa de queso gorgonzola con esparragos, portobellos y peras salteadas en mantequilla a las finas hierbas.', 19900, 'TRUE', 'TRUE', 15, 'Sin Gluten', 30, 'FALSE', '', ''],
    ['P016', 'CAT04', 'Entrana con chimichurri acompanada de papitas nativas salteadas al romero', 'Corte de entrana a la plancha y marinada con chimichurri de la casa, acompanado de papas salteadas al ajo y romero.', 21900, 'TRUE', 'TRUE', 16, 'Sin Gluten', 30, 'FALSE', '', ''],
    ['P017', 'CAT04', 'Asado de tira en salsa Cabernet Sauvignon acompanado de pastelera', 'Asado de tira en coccion lenta de vino tinto, banado en su salsa y montado con pastelera de choclo.', 22900, 'TRUE', 'TRUE', 17, '', 35, 'FALSE', '', ''],
    ['P018', 'CAT04', 'Lomo y quinoa cremosa', 'Lomo liso grillado sobre quinoa cremosa con vegetales de la estacion.', 18900, 'TRUE', 'TRUE', 18, '', 30, 'FALSE', '', ''],
    // PASTAS Y SALSAS
    ['P019', 'CAT06', 'Spaghettis, fetuccinnis o gnocchis con salsa a eleccion', 'Bolonesa | Alfredo | Pollo y pesto | Cuatro quesos', 14900, 'TRUE', 'TRUE', 19, '', 20, 'FALSE', '', ''],
    ['P020', 'CAT06', 'Cannelloni de salmon ahumado', 'Pasta casera rellena de salmon ahumado, champinones, cebolla y queso de la zona. Todo gratinado en salsa rosa y queso parmesano.', 16900, 'TRUE', 'TRUE', 20, '', 25, 'FALSE', '', ''],
    ['P021', 'CAT06', 'Rissotto de calabaza', 'Tradicional arroz cocinado lentamente en fondo de verduras de la estacion y vino blanco, finalizado con pure de calabaza.', 14900, 'TRUE', 'TRUE', 21, 'Vegano', 25, 'FALSE', '', ''],
    // POSTRES
    ['P022', 'CAT08', 'Suspiro de lucuma', 'Clasico postre peruano con crema de manjar y el sabor cremoso de las lucumas.', 6900, 'TRUE', 'TRUE', 22, '', 10, 'FALSE', '', ''],
    ['P023', 'CAT08', 'Panacotta de murta', 'Tradicional preparacion casera con el sabor de este fruto sureno de estacion.', 6900, 'TRUE', 'TRUE', 23, '', 10, 'FALSE', '', ''],
    ['P024', 'CAT08', 'Cheesecake de la casa (a eleccion segun disponibilidad)', 'Tarta de queso horneada con base de galletas y salsa.', 6900, 'TRUE', 'TRUE', 24, '', 10, 'FALSE', '', ''],
    ['P025', 'CAT08', 'Brownie con helado', 'Brownie casero con dos bolitas de helado.', 6900, 'TRUE', 'TRUE', 25, '', 10, 'FALSE', '', ''],
    // PARA LOS NINOS
    ['P026', 'CAT07', 'Pollo a la plancha 150gr.', 'Acompanado de papas fritas, arroz o pure.', 12900, 'TRUE', 'TRUE', 26, '', 15, 'FALSE', '', ''],
    ['P027', 'CAT07', 'Salmon a la plancha 150 gr.', 'Acompanado de papas fritas, arroz o pure.', 12900, 'TRUE', 'TRUE', 27, '', 15, 'FALSE', '', ''],
    ['P028', 'CAT07', 'Mini hamburguesa 170 gr.', 'Acompanado de papas fritas, arroz o pure.', 12900, 'TRUE', 'TRUE', 28, '', 15, 'FALSE', '', ''],
    ['P029', 'CAT07', 'Nuggets caseros 5 unidades', 'Acompanado de papas fritas, arroz o pure.', 12900, 'TRUE', 'TRUE', 29, '', 15, 'FALSE', '', ''],
    // CAFETERIA
    ['P030', 'CAT09', 'Espresso', '', 3500, 'TRUE', 'TRUE', 30, '', 5, 'FALSE', '', ''],
    ['P031', 'CAT09', 'Espresso doble', '', 4700, 'TRUE', 'TRUE', 31, '', 5, 'FALSE', '', ''],
    ['P032', 'CAT09', 'Cortado', '', 3900, 'TRUE', 'TRUE', 32, '', 5, 'FALSE', '', ''],
    ['P033', 'CAT09', 'Americano', '', 3600, 'TRUE', 'TRUE', 33, '', 5, 'FALSE', '', ''],
    ['P034', 'CAT09', 'Te e infusiones', '', 3400, 'TRUE', 'TRUE', 34, '', 5, 'FALSE', '', ''],
    // SANDWICHES
    ['P035', 'CAT10', 'Sandwich del lago', 'Pan ciabatta artesanal, queso crema, rucula, palta, aceituna y salmon ahumado.', 12900, 'TRUE', 'TRUE', 35, '', 15, 'FALSE', '', ''],
    ['P036', 'CAT10', 'Sandwich barros luco', 'Churrasco de filete o ave con queso.', 12500, 'TRUE', 'TRUE', 36, '', 15, 'FALSE', '', ''],
    ['P037', 'CAT10', 'Sandwich italiano', 'Churrasco de filete o ave, tomate, lechuga, mayonesa, palta.', 14500, 'TRUE', 'TRUE', 37, '', 15, 'FALSE', '', ''],
    ['P038', 'CAT10', 'Hamburguesa Cascadas', 'Pan de papa, hamburguesa, tocino, queso, cebolla morada, lechuga tomate, pepinillo, salsa golf.', 14500, 'TRUE', 'TRUE', 38, '', 20, 'FALSE', '', ''],
    ['P039', 'CAT10', 'Sandwich Calbuco', 'Pan ciabatta, pasta de ajo negro, jamon serrano, rucula, pimenton asado.', 13500, 'TRUE', 'TRUE', 39, '', 15, 'FALSE', '', ''],
    // TABLAS
    ['P040', 'CAT11', 'Tabla Osorno', 'Papas fritas, pollo salteado, cebolla caramelizada, champinones, salsa brava.', 19900, 'TRUE', 'TRUE', 40, '', 15, 'FALSE', '', ''],
    ['P041', 'CAT11', 'Tabla Cascadas', 'Charcuteria, quesos, frutas y frutos secos.', 16900, 'TRUE', 'TRUE', 41, '', 15, 'FALSE', '', ''],
    // COCTELES CLASICOS
    ['P042', 'CAT18', 'Pisco sour tradicional', '', 6200, 'TRUE', 'TRUE', 42, '', 5, 'FALSE', '', ''],
    ['P043', 'CAT18', 'Amaretto sour', '', 6200, 'TRUE', 'TRUE', 43, '', 5, 'FALSE', '', ''],
    ['P044', 'CAT18', 'Chardonnay sour', '', 6000, 'TRUE', 'TRUE', 44, '', 5, 'FALSE', '', ''],
    ['P045', 'CAT18', 'Carmenere sour', '', 6000, 'TRUE', 'TRUE', 45, '', 5, 'FALSE', '', ''],
    ['P046', 'CAT18', 'Pisco sour catedral', '', 8500, 'TRUE', 'TRUE', 46, '', 5, 'FALSE', '', ''],
    ['P047', 'CAT18', 'Alma sour', '', 12500, 'TRUE', 'TRUE', 47, '', 5, 'FALSE', '', ''],
    ['P048', 'CAT18', 'Whisky sour', '', 7200, 'TRUE', 'TRUE', 48, '', 5, 'FALSE', '', ''],
    ['P049', 'CAT18', 'Ramazzotti Spritz', '', 6900, 'TRUE', 'TRUE', 49, '', 5, 'FALSE', '', ''],
    ['P050', 'CAT18', 'Campari Tonica', '', 6900, 'TRUE', 'TRUE', 50, '', 5, 'FALSE', '', ''],
    ['P051', 'CAT18', 'Aperol Spritz', '', 6900, 'TRUE', 'TRUE', 51, '', 5, 'FALSE', '', ''],
    ['P052', 'CAT18', 'St. Germain Spritz', '', 7500, 'TRUE', 'TRUE', 52, '', 5, 'FALSE', '', ''],
    ['P053', 'CAT18', 'Pichuncho', '', 6600, 'TRUE', 'TRUE', 53, '', 5, 'FALSE', '', ''],
    ['P054', 'CAT18', 'Tropical gin', '', 9800, 'TRUE', 'TRUE', 54, '', 5, 'FALSE', '', ''],
    ['P055', 'CAT18', 'Jagermeister', '', 6600, 'TRUE', 'TRUE', 55, '', 5, 'FALSE', '', ''],
    ['P056', 'CAT18', 'Araucano', '', 6000, 'TRUE', 'TRUE', 56, '', 5, 'FALSE', '', ''],
    ['P057', 'CAT18', 'Negroni', '', 7200, 'TRUE', 'TRUE', 57, '', 5, 'FALSE', '', ''],
    ['P058', 'CAT18', 'Negroni chilensis', '', 7200, 'TRUE', 'TRUE', 58, '', 5, 'FALSE', '', ''],
    ['P059', 'CAT18', 'Cascadas', '', 7500, 'TRUE', 'TRUE', 59, '', 5, 'FALSE', '', ''],
    ['P060', 'CAT18', 'White lady', '', 7000, 'TRUE', 'TRUE', 60, '', 5, 'FALSE', '', ''],
    ['P061', 'CAT18', 'Sangria', '', 7000, 'TRUE', 'TRUE', 61, '', 5, 'FALSE', '', ''],
    ['P062', 'CAT18', 'Paloma Cascadas', '', 7500, 'TRUE', 'TRUE', 62, '', 5, 'FALSE', '', ''],
    ['P063', 'CAT18', 'Limoncello', '', 7500, 'TRUE', 'TRUE', 63, '', 5, 'FALSE', '', ''],
    ['P064', 'CAT18', 'Vermuth Rosso', '', 5500, 'TRUE', 'TRUE', 64, '', 5, 'FALSE', '', ''],
    ['P065', 'CAT18', 'Bloody Mary', '', 8000, 'TRUE', 'TRUE', 65, '', 5, 'FALSE', '', ''],
    // SIN ALCOHOL
    ['P066', 'CAT20', 'Agua mineral', '', 2900, 'TRUE', 'TRUE', 66, '', 2, 'FALSE', '', ''],
    ['P067', 'CAT20', 'Gaseosas', '', 3000, 'TRUE', 'TRUE', 67, '', 2, 'FALSE', '', ''],
    ['P068', 'CAT20', 'Jugos', '', 4500, 'TRUE', 'TRUE', 68, '', 2, 'FALSE', '', ''],
    ['P069', 'CAT20', 'Limonada clasica', '', 4000, 'TRUE', 'TRUE', 69, '', 2, 'FALSE', '', ''],
    ['P070', 'CAT20', 'Limonada especial', '', 4200, 'TRUE', 'TRUE', 70, '', 2, 'FALSE', '', ''],
    // CERVEZAS
    ['P071', 'CAT12', 'Budweisser', '', 3800, 'TRUE', 'TRUE', 71, '', 2, 'FALSE', '', ''],
    ['P072', 'CAT12', 'Heineken', '', 3800, 'TRUE', 'TRUE', 72, '', 2, 'FALSE', '', ''],
    ['P073', 'CAT12', 'Stella Artois', '', 3800, 'TRUE', 'TRUE', 73, '', 2, 'FALSE', '', ''],
    ['P074', 'CAT12', 'Tropera - artesanal local', '', 4900, 'TRUE', 'TRUE', 74, '', 2, 'FALSE', '', ''],
    ['P075', 'CAT12', 'Bosques - artesanal local', '', 6000, 'TRUE', 'TRUE', 75, '', 2, 'FALSE', '', ''],
    // GIN
    ['P076', 'CAT13', 'Azucena Citrico', '', 7500, 'TRUE', 'TRUE', 76, '', 2, 'FALSE', '', ''],
    ['P077', 'CAT13', 'Azucena Floral', '', 7500, 'TRUE', 'TRUE', 77, '', 2, 'FALSE', '', ''],
    ['P078', 'CAT13', 'Bombay Sapphire', '', 8000, 'TRUE', 'TRUE', 78, '', 2, 'FALSE', '', ''],
    ['P079', 'CAT13', 'Hendrix', '', 12000, 'TRUE', 'TRUE', 79, '', 2, 'FALSE', '', ''],
    ['P080', 'CAT13', 'Good Wind', '', 9600, 'TRUE', 'TRUE', 80, '', 2, 'FALSE', '', ''],
    // PISCO
    ['P081', 'CAT17', 'Alto del Carmen 35', '', 6500, 'TRUE', 'TRUE', 81, '', 2, 'FALSE', '', ''],
    ['P082', 'CAT17', 'Alto del Carmen 40', '', 7800, 'TRUE', 'TRUE', 82, '', 2, 'FALSE', '', ''],
    ['P083', 'CAT17', 'Bauza Especial 35', '', 5500, 'TRUE', 'TRUE', 83, '', 2, 'FALSE', '', ''],
    ['P084', 'CAT17', 'Bauza Aniversario', '', 5500, 'TRUE', 'TRUE', 84, '', 2, 'FALSE', '', ''],
    ['P085', 'CAT17', 'Alma Ultra Premium', '', 9500, 'TRUE', 'TRUE', 85, '', 2, 'FALSE', '', ''],
    ['P086', 'CAT17', 'Don Lorenzo', '', 11500, 'TRUE', 'TRUE', 86, '', 2, 'FALSE', '', ''],
    // WHISKY
    ['P087', 'CAT15', "Buchannan's", '', 9000, 'TRUE', 'TRUE', 87, '', 2, 'FALSE', '', ''],
    ['P088', 'CAT15', 'Johnnie Walker Red Label', '', 7500, 'TRUE', 'TRUE', 88, '', 2, 'FALSE', '', ''],
    ['P089', 'CAT15', 'Johnnie Walker Black Label', '', 13000, 'TRUE', 'TRUE', 89, '', 2, 'FALSE', '', ''],
    ['P090', 'CAT15', "Jack Daniel's", '', 9000, 'TRUE', 'TRUE', 90, '', 2, 'FALSE', '', ''],
    ['P091', 'CAT15', "Jack Daniel's Honey", '', 9000, 'TRUE', 'TRUE', 91, '', 2, 'FALSE', '', ''],
    ['P092', 'CAT15', "Jack Daniel's Apple", '', 9000, 'TRUE', 'TRUE', 92, '', 2, 'FALSE', '', ''],
    ['P093', 'CAT15', 'Chivas Reagal 12 anos', '', 13000, 'TRUE', 'TRUE', 93, '', 2, 'FALSE', '', ''],
    ['P094', 'CAT15', 'Chivas Reagal 18 anos', '', 25000, 'TRUE', 'TRUE', 94, '', 2, 'FALSE', '', ''],
    // VODKA
    ['P095', 'CAT14', 'Stolichnaya', '', 6000, 'TRUE', 'TRUE', 95, '', 2, 'FALSE', '', ''],
    ['P096', 'CAT14', 'Absolut', '', 7500, 'TRUE', 'TRUE', 96, '', 2, 'FALSE', '', ''],
    ['P097', 'CAT14', 'Grey Goose', '', 15000, 'TRUE', 'TRUE', 97, '', 2, 'FALSE', '', ''],
    // TEQUILA
    ['P098', 'CAT16', 'Olmeca Silver', '', 6000, 'TRUE', 'TRUE', 98, '', 2, 'FALSE', '', ''],
    ['P099', 'CAT16', 'Olmeca Reposado', '', 7000, 'TRUE', 'TRUE', 99, '', 2, 'FALSE', '', ''],
    ['P100', 'CAT16', 'Herradura Reposado', '', 9000, 'TRUE', 'TRUE', 100, '', 2, 'FALSE', '', ''],
    ['P101', 'CAT16', 'Don Julio', '', 14000, 'TRUE', 'TRUE', 101, '', 2, 'FALSE', '', ''],
    // VINOS
    ['P102', 'CAT19', 'Gala (Espumante)', '', 18900, 'TRUE', 'TRUE', 102, '', 2, 'FALSE', '', ''],
    ['P103', 'CAT19', 'Casa Donoso (Espumante)', '', 21900, 'TRUE', 'TRUE', 103, '', 2, 'FALSE', '', ''],
    ['P104', 'CAT19', 'Emiliana Adobe (Sauvignon Blanc)', '', 13800, 'TRUE', 'TRUE', 104, '', 2, 'FALSE', '', ''],
    ['P105', 'CAT19', 'Ranquilhue (Sauvignon Blanc)', '', 18600, 'TRUE', 'TRUE', 105, '', 2, 'FALSE', '', ''],
    ['P106', 'CAT19', 'Trapi Del Bueno (Sauvignon Blanc)', '', 22000, 'TRUE', 'TRUE', 106, '', 2, 'FALSE', '', ''],
    ['P107', 'CAT19', 'Garces Silva Amayna (Sauvignon Blanc)', '', 26000, 'TRUE', 'TRUE', 107, '', 2, 'FALSE', '', ''],
    ['P108', 'CAT19', 'Casa Donoso Reserva (Chardonnay)', '', 21900, 'TRUE', 'TRUE', 108, '', 2, 'FALSE', '', ''],
    ['P109', 'CAT19', 'Villard Expression (Chardonnay)', '', 20000, 'TRUE', 'TRUE', 109, '', 2, 'FALSE', '', ''],
    ['P110', 'CAT19', 'Casa Silva Reserva Cuvee (Chardonnay)', '', 17000, 'TRUE', 'TRUE', 110, '', 2, 'FALSE', '', ''],
    ['P111', 'CAT19', 'Villard Expression Pinot Noir (Tinto)', '', 20000, 'TRUE', 'TRUE', 111, '', 2, 'FALSE', '', ''],
    ['P112', 'CAT19', 'Perez Cruz Limited Edition Syrah (Tinto)', '', 31000, 'TRUE', 'TRUE', 112, '', 2, 'FALSE', '', ''],
    ['P113', 'CAT19', 'Casa Bauza Veraz Cabernet Franc (Tinto)', '', 34000, 'TRUE', 'TRUE', 113, '', 2, 'FALSE', '', ''],
    ['P114', 'CAT19', 'Rupestre Valle Itata Cinsault (Tinto)', '', 41000, 'TRUE', 'TRUE', 114, '', 2, 'FALSE', '', ''],
    ['P115', 'CAT19', 'Santa Ema Gran Reserva (Merlot)', '', 24000, 'TRUE', 'TRUE', 115, '', 2, 'FALSE', '', ''],
    ['P116', 'CAT19', 'Casa Bauza Presumido (Carmenere)', '', 25000, 'TRUE', 'TRUE', 116, '', 2, 'FALSE', '', ''],
    ['P117', 'CAT19', 'Montes Alpha Limited Edition (Carmenere)', '', 23800, 'TRUE', 'TRUE', 117, '', 2, 'FALSE', '', ''],
    ['P118', 'CAT19', 'Ventisquero Grey (Carmenere)', '', 24600, 'TRUE', 'TRUE', 118, '', 2, 'FALSE', '', ''],
    ['P119', 'CAT19', 'Casas Patronales Gran Reserva (Carmenere)', '', 20000, 'TRUE', 'TRUE', 119, '', 2, 'FALSE', '', ''],
    ['P120', 'CAT19', 'Casa Bauza Valiente (Cabernet Sauvignon)', '', 25000, 'TRUE', 'TRUE', 120, '', 2, 'FALSE', '', ''],
    ['P121', 'CAT19', 'Perez Cruz Limited Edition (Cabernet Sauvignon)', '', 30900, 'TRUE', 'TRUE', 121, '', 2, 'FALSE', '', ''],
    ['P122', 'CAT19', 'Casa Donoso Reserva (Cabernet Sauvignon)', '', 22900, 'TRUE', 'TRUE', 122, '', 2, 'FALSE', '', ''],
    ['P123', 'CAT19', 'Requingua Toro de Piedra (Cabernet Sauvignon)', '', 28900, 'TRUE', 'TRUE', 123, '', 2, 'FALSE', '', ''],
    ['P124', 'CAT19', 'Casa Bauza Atrevido (Ensamblaje)', '', 29600, 'TRUE', 'TRUE', 124, '', 2, 'FALSE', '', ''],
    ['P125', 'CAT19', 'Emiliana Coyam (Ensamblaje)', '', 35000, 'TRUE', 'TRUE', 125, '', 2, 'FALSE', '', ''],
    ['P126', 'CAT19', 'VIK Milla Cala (Ensamblaje)', '', 49900, 'TRUE', 'TRUE', 126, '', 2, 'FALSE', '', ''],
    ['P127', 'CAT19', 'Casa Donoso 1810 (Ensamblaje)', '', 28000, 'TRUE', 'TRUE', 127, '', 2, 'FALSE', '', ''],
    ['P128', 'CAT19', 'Caballo Loco apalta (Ensamblaje)', '', 65000, 'TRUE', 'TRUE', 128, '', 2, 'FALSE', '', ''],
    ['P129', 'CAT19', 'Casa Bauza Isabel (Icono)', '', 89900, 'TRUE', 'TRUE', 129, '', 2, 'FALSE', '', ''],
    ['P130', 'CAT19', 'VIK La Piu Belle (Icono)', '', 139000, 'TRUE', 'TRUE', 130, '', 2, 'FALSE', '', ''],
    ['P131', 'CAT19', 'Villard Tanagra (Icono)', '', 80000, 'TRUE', 'TRUE', 131, '', 2, 'FALSE', '', ''],
    ['P132', 'CAT19', 'Concha y Toro Don Melchor (Icono)', '', 340000, 'TRUE', 'TRUE', 132, '', 2, 'FALSE', '', '']
  ];
  _crearHoja(ss, 'Productos', encabezados, filas);
}

/** 5.6. Hoja Reservas (solo encabezados) */
function _crearHojaReservas(ss) {
  var encabezados = ['ID', 'Timestamp', 'Habitacion', 'ServicioID', 'Fecha', 'HoraInicio',
    'HoraFin', 'Personas', 'Estado', 'SolicitadoPor', 'Notas', 'PrepedidoID',
    'MotivoCancelacion', 'EsEvento'];
  _crearHoja(ss, 'Reservas', encabezados, []);
}

/** 5.7. Hoja Pedidos (solo encabezados) */
function _crearHojaPedidos(ss) {
  var encabezados = ['ID', 'ReservaID', 'Habitacion', 'Estado', 'Total', 'Timestamp', 'Notas', 'Entrega'];
  _crearHoja(ss, 'Pedidos', encabezados, []);
}

/** 5.8. Hoja DetallePedidos (solo encabezados) */
function _crearHojaDetallePedidos(ss) {
  var encabezados = ['ID', 'PedidoID', 'ProductoID', 'Cantidad', 'PrecioUnitario', 'Subtotal', 'Notas'];
  _crearHoja(ss, 'DetallePedidos', encabezados, []);
}

/** 5.9. Hoja Usuarios */
function _crearHojaUsuarios(ss) {
  var encabezados = ['ID', 'Email', 'Nombre', 'Rol', 'HabitacionAsociada', 'Activo', 'PasswordHash'];
  var filas = [
    ['U001', 'admin@cascadas.cl', 'Administrador', 'ADMINISTRADOR', '', 'TRUE', ''],
    ['U002', 'recepcion@cascadas.cl', 'Recepcion', 'RECEPCION', '', 'TRUE', ''],
    ['U003', 'cocina@cascadas.cl', 'Cocina', 'COCINA', '', 'TRUE', '']
  ];
  _crearHoja(ss, 'Usuarios', encabezados, filas);
}

/** 5.10. Hoja EventosBloqueos (solo encabezados) */
function _crearHojaEventosBloqueos(ss) {
  var encabezados = ['ID', 'Tipo', 'ServicioID', 'FechaInicio', 'FechaFin', 'HoraInicio',
    'HoraFin', 'Motivo', 'CreadoPor', 'Timestamp'];
  _crearHoja(ss, 'EventosBloqueos', encabezados, []);
}

/** 5.11. Hoja Notificaciones (solo encabezados) */
function _crearHojaNotificaciones(ss) {
  var encabezados = ['ID', 'Timestamp', 'Tipo', 'Mensaje', 'DestinatarioRol', 'Habitacion',
    'ServicioID', 'Leida', 'FechaReferencia'];
  _crearHoja(ss, 'Notificaciones', encabezados, []);
}

/** 5.12. Hoja DisponibilidadPersonal */
function _crearHojaDisponibilidadPersonal(ss) {
  var encabezados = ['ID', 'TipoPersonal', 'Fecha', 'HoraInicio', 'HoraFin', 'Disponible', 'Notas'];
  _crearHoja(ss, 'DisponibilidadPersonal', encabezados, []);
}

/** 5.13. Hoja LogActividad (solo encabezados) */
function _crearHojaLogActividad(ss) {
  var encabezados = ['ID', 'Timestamp', 'UsuarioEmail', 'Accion', 'Detalle', 'Habitacion'];
  _crearHoja(ss, 'LogActividad', encabezados, []);
}

/** 5.14. Hoja HistorialPedidos (auditoria de consultas) */
function _crearHojaHistorialPedidos(ss) {
  var encabezados = ['ID', 'Timestamp', 'Habitacion', 'FechaReferencia', 'ServicioID',
    'TipoBusqueda', 'DetalleResumen', 'PedidoID', 'ReservaID', 'UsuarioConsulta'];
  _crearHoja(ss, 'HistorialPedidos', encabezados, []);
}

/** Elimina la hoja por defecto ("Hoja 1" / "Sheet1") si aun existe. */
function _eliminarHojaPorDefecto(ss) {
  var nombresPorDefecto = ['Hoja 1', 'Hoja1', 'Sheet1', 'Sheet 1'];
  var hojas = ss.getSheets();
  if (hojas.length <= 1) return;
  for (var i = 0; i < nombresPorDefecto.length; i++) {
    var hoja = ss.getSheetByName(nombresPorDefecto[i]);
    if (hoja && ss.getSheets().length > 1) {
      ss.deleteSheet(hoja);
    }
  }
}

