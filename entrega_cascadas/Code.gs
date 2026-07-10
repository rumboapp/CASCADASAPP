// =================================================================
// CONFIGURACIÓN DE IDS GLOBALES
// =================================================================
const ID_PLANTILLA = "1CrI7WhSfXKDw1S86opgTJYKpAi1mDjXFrsKr2lEOeyY";
const ID_PLANILLA_SHEETS = "164qlshfA21LK2hIAcVlrv8rdNIamupZfF5_gSTW6zWo";
const NOMBRE_CARPETA_COTIZACIONES = "Cotizaciones Temporales";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Cotizador - Cascadas Hotel')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerConfiguracion() {
  try {
    var ss = SpreadsheetApp.openById(ID_PLANILLA_SHEETS);

    var sheetTarifas = ss.getSheetByName("Tarifas") || ss.insertSheet("Tarifas");
    var dataTarifas = sheetTarifas.getDataRange().getValues();
    var tarifas = {};
    if (dataTarifas.length <= 1) {
      var base = [
        ["Habitación Single", 150000],
        ["Habitación Matrimonial", 230000],
        ["Habitación Twin", 180000],
        ["Habitación Suite", 290000],
        ["Cama Adicional", 45000]
      ];
      base.forEach(f => { sheetTarifas.appendRow(f); tarifas[f[0]] = f[1]; });
    } else {
      for (var i = 1; i < dataTarifas.length; i++) {
        if (dataTarifas[i][0] && dataTarifas[i][0].toString().trim() !== "") {
          tarifas[dataTarifas[i][0].toString().trim()] = Number(dataTarifas[i][1]) || 0;
        }
      }
    }

    // ── Reemplazo de la categoría Single por dos variantes de cama ──
    // Cambio funcional solicitado: donde antes había "Habitación Single (1 persona)"
    // (o "Habitación Single") ahora se ofrecen dos opciones, conservando el precio.
    // No se detectan las nuevas variantes (contienen "cama") para no duplicarlas.
    var tarifasFinal = {};
    Object.keys(tarifas).forEach(function(kt) {
      var kl = kt.toLowerCase();
      if (kl.indexOf("single") >= 0 && kl.indexOf("cama") < 0) {
        var precioSingle = tarifas[kt];
        tarifasFinal["Habitación Single (1 persona - 1 Cama Matrimonial)"] = precioSingle;
        tarifasFinal["Habitación Single (1 persona - 2 Camas de 1 1/2 plaza)"] = precioSingle;
      } else {
        tarifasFinal[kt] = tarifas[kt];
      }
    });
    tarifas = tarifasFinal;

    var sheetProgramas = ss.getSheetByName("Programas") || ss.insertSheet("Programas");
    var dataProgramas = sheetProgramas.getDataRange().getValues();
    var programas = {};

    if (dataProgramas.length > 1) {
      for (var j = 1; j < dataProgramas.length; j++) {
        if (dataProgramas[j][0] && dataProgramas[j][0].toString().trim() !== "") {
          var nombreNormalizado = dataProgramas[j][0].toString().trim();
          programas[nombreNormalizado] = {
            valor: Number(dataProgramas[j][1]) || 0,
            servicios: [],
            almuerzos: 0,
            cenas: 0,
            masajes: 0,
            ocasionesEspeciales: []
          };
        }
      }
    }

    var sheetJSON = ss.getSheetByName("TicksJSON") || ss.insertSheet("TicksJSON");
    var dataJSON = sheetJSON.getDataRange().getValues();

    if (dataJSON.length > 1) {
      for (var k = 1; k < dataJSON.length; k++) {
        try {
          var nomProg = dataJSON[k][0] ? dataJSON[k][0].toString().trim() : null;
          var stringJSON = dataJSON[k][1] ? dataJSON[k][1].toString().trim() : "";
          if (nomProg && programas[nomProg]) {
            if (stringJSON !== "") {
              try {
                var configGuardada = JSON.parse(stringJSON);
                programas[nomProg].servicios = configGuardada.servicios || [];
                programas[nomProg].almuerzos = Number(configGuardada.almuerzos) || 0;
                programas[nomProg].cenas = Number(configGuardada.cenas) || 0;
                programas[nomProg].masajes = Number(configGuardada.masajes) || 0;
                programas[nomProg].ocasionesEspeciales = configGuardada.ocasionesEspeciales || [];
              } catch(errJson) {
                programas[nomProg].servicios = stringJSON.split(',');
                programas[nomProg].ocasionesEspeciales = [];
              }
            }
          }
        } catch (errorFila) {
          console.log("Error procesando fila " + k + " de TicksJSON: " + errorFila.toString());
        }
      }
    }

    // ── AMENIDADES PREDEFINIDAS ──
    var sheetAmenidades = ss.getSheetByName("Amenidades") || ss.insertSheet("Amenidades");
    var dataAmenidades = sheetAmenidades.getDataRange().getValues();
    var amenidades = {
      early:   { nombre: "Early Check In", valor: 25000, descripcion: "Ingreso a las 12:00 hrs (sujeto a disponibilidad)" },
      late:    { nombre: "Late Check Out", valor: 25000, descripcion: "Salida hasta las 15:00 hrs (sujeto a disponibilidad)" },
      tinaja:  { nombre: "Sesión de Tinaja", valor: 45000, descripcion: "Sesión de tinaja de 2 horas para dos personas previa agenda." },
      almuerzo:{ nombre: "Almuerzo Sugerencia del Chef", valor: 35000, descripcion: "Menú sugerencia de 3 tiempos de nuestro Chef para 2 personas." },
      cena:    { nombre: "Cena Sugerencia del Chef", valor: 45000, descripcion: "Menú sugerencia de 3 tiempos de nuestro Chef para 2 personas." },
      drink:   { nombre: "Welcome Drink", valor: 15000, descripcion: "Welcome drink para dos personas (Jugo, vino o espumante)." },
      masaje:  { nombre: "Masaje de Relajación", valor: 35000, descripcion: "Sesión de masaje de relajación de 45 minutos por persona." }
    };
    if (dataAmenidades.length <= 1) {
      sheetAmenidades.appendRow(["Amenidad", "PrecioNeto"]);
      Object.keys(amenidades).forEach(function(k) {
        sheetAmenidades.appendRow([amenidades[k].nombre, amenidades[k].valor]);
      });
    } else {
      // Leer valores existentes
      var nombresExistentes = [];
      for (var a = 1; a < dataAmenidades.length; a++) {
        var nomA = dataAmenidades[a][0] ? dataAmenidades[a][0].toString().trim() : "";
        var valA = Number(dataAmenidades[a][1]) || 0;
        nombresExistentes.push(nomA);
        for (var k in amenidades) {
          if (amenidades[k].nombre === nomA) { amenidades[k].valor = valA; break; }
        }
      }
      // Agregar amenidades nuevas que no existan en la hoja
      Object.keys(amenidades).forEach(function(k) {
        if (nombresExistentes.indexOf(amenidades[k].nombre) < 0) {
          sheetAmenidades.appendRow([amenidades[k].nombre, amenidades[k].valor]);
        }
      });
    }

    var logoUrl = "";
    try {
      var sheetLogo = ss.getSheetByName("LOGO");
      if (sheetLogo) {
        var dataLogo = sheetLogo.getDataRange().getValues();
        for (var l = 1; l < dataLogo.length; l++) {
          if (dataLogo[l][0] && dataLogo[l][0].toString().trim().toLowerCase() === "logo_url") {
            logoUrl = dataLogo[l][1] ? dataLogo[l][1].toString().trim() : "";
            break;
          }
        }
      }
    } catch(eLogo) {
      console.log("Error leyendo logo: " + eLogo.toString());
    }

    return { tarifas: tarifas, programas: programas, amenidades: amenidades, logoUrl: logoUrl, exito: true };
  } catch (e) {
    return { exito: false, error: e.toString(), tarifas: {}, programas: {}, amenidades: {}, logoUrl: "" };
  }
}

function guardarTarifas(t) {
  var sheet = SpreadsheetApp.openById(ID_PLANILLA_SHEETS).getSheetByName("Tarifas");
  sheet.clear();
  sheet.appendRow(["Habitacion", "PrecioNeto"]);
  for (var k in t) { sheet.appendRow([k, t[k]]); }
  return true;
}

function guardarNuevoPrograma(nombre, valor, serviciosArreglo, alms, cens, masjs, ocasionesEspecialesArreglo) {
  var ss = SpreadsheetApp.openById(ID_PLANILLA_SHEETS);
  var nombreLimpio = nombre.toString().trim();

  var sheetP = ss.getSheetByName("Programas") || ss.insertSheet("Programas");
  var dataP = sheetP.getDataRange().getValues();
  var filaPrograma = -1;
  for (var i = 1; i < dataP.length; i++) {
    if (dataP[i][0] && dataP[i][0].toString().trim().toLowerCase() === nombreLimpio.toLowerCase()) {
      filaPrograma = i + 1; break;
    }
  }
  if (filaPrograma !== -1) { sheetP.getRange(filaPrograma, 2).setValue(Number(valor)); }
  else { sheetP.appendRow([nombreLimpio, Number(valor)]); }

  var sheetJ = ss.getSheetByName("TicksJSON") || ss.insertSheet("TicksJSON");
  var dataJ = sheetJ.getDataRange().getValues();
  var filaJSON = -1;
  if (dataJ.length === 0 || (dataJ.length === 1 && dataJ[0][0] === "")) {
    sheetJ.clear();
    sheetJ.appendRow(["Nombre Programa", "ConfigJSON"]);
    dataJ = sheetJ.getDataRange().getValues();
  }
  for (var j = 1; j < dataJ.length; j++) {
    if (dataJ[j][0] && dataJ[j][0].toString().trim().toLowerCase() === nombreLimpio.toLowerCase()) {
      filaJSON = j + 1; break;
    }
  }
  var objetoConfig = {
    servicios: serviciosArreglo, almuerzos: Number(alms),
    cenas: Number(cens), masajes: Number(masjs),
    ocasionesEspeciales: ocasionesEspecialesArreglo || []
  };
  var stringParaGuardar = JSON.stringify(objetoConfig);
  if (filaJSON !== -1) { sheetJ.getRange(filaJSON, 2).setValue(stringParaGuardar); }
  else { sheetJ.appendRow([nombreLimpio, stringParaGuardar]); }
  SpreadsheetApp.flush();
  return true;
}

function eliminarProgramaDeBaseDatos(nombre) {
  try {
    var ss = SpreadsheetApp.openById(ID_PLANILLA_SHEETS);
    var nombreLimpio = nombre.toString().trim().toLowerCase();
    var sheetP = ss.getSheetByName("Programas");
    if (sheetP) {
      var dataP = sheetP.getDataRange().getValues();
      for (var i = dataP.length - 1; i >= 1; i--) {
        if (dataP[i][0] && dataP[i][0].toString().trim().toLowerCase() === nombreLimpio) sheetP.deleteRow(i + 1);
      }
    }
    var sheetJ = ss.getSheetByName("TicksJSON");
    if (sheetJ) {
      var dataJ = sheetJ.getDataRange().getValues();
      for (var j = dataJ.length - 1; j >= 1; j--) {
        if (dataJ[j][0] && dataJ[j][0].toString().trim().toLowerCase() === nombreLimpio) sheetJ.deleteRow(j + 1);
      }
    }
    SpreadsheetApp.flush();
    return true;
  } catch(e) {
    throw new Error("Error en la base de datos al intentar borrar el programa: " + e.toString());
  }
}

// =================================================================
// REGISTRO EN HISTORIAL
// =================================================================
function registrarEnHistorial(ss, datos, linkPdf) {
  try {
    var sheet = ss.getSheetByName("Historial");
    var CABECERA = ["Fecha", "Cliente", "Tipo", "Detalle", "CheckIn", "CheckOut", "Noches", "Neto", "Total", "LinkPDF"];
    var link = linkPdf || "";

    if (!sheet) {
      sheet = ss.insertSheet("Historial");
      sheet.appendRow(CABECERA);
      sheet.setFrozenRows(1);
    }

    var cabeceraActual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colNames = cabeceraActual.map(function(c){ return c.toString().toLowerCase().trim(); });
    if (colNames.indexOf("checkin") < 0) {
      sheet.insertColumnsAfter(4, 2);
      sheet.getRange(1, 1, 1, CABECERA.length).setValues([CABECERA]);
      colNames = CABECERA.map(function(c){ return c.toLowerCase(); });
    }
    if (colNames.indexOf("linkpdf") < 0) {
      sheet.getRange(1, CABECERA.length).setValue("LinkPDF");
    }

    var tz = Session.getScriptTimeZone();
    var ahora = new Date();
    var checkin  = datos.checkin  || "";
    var checkout = datos.checkout || "";
    var noches   = Number(datos.noches || 0);

    if (datos.tipo_cotizacion === "programa") {
      var neto  = Number(datos.total_programa || 0);
      var total = Math.round(neto * 1.19);
      var detalleProg = datos.nombre_programa || "Programa";
      if (datos.cantidad_programa && datos.cantidad_programa > 1) {
        detalleProg += " x" + datos.cantidad_programa;
      }
      var fila = [ahora, datos.nombre_cliente || "", "Programa",
                  detalleProg,
                  checkin, checkout, noches, neto, total, link];
      sheet.appendRow(fila);
      var uf = sheet.getLastRow();
      sheet.getRange(uf, 8).setNumberFormat("0");
      sheet.getRange(uf, 9).setNumberFormat("0");

    } else {
      if (datos.habitaciones && datos.habitaciones.length > 0) {
        var consolidadas = [];
        datos.habitaciones.forEach(function(h) {
          var tipoBase = h.tipo.replace(/\s*\(x\d+\)$/, '').trim();
          var ex = consolidadas.find(function(e){ return e.tipoBase === tipoBase && e.precio === h.precio; });
          if (ex) { ex.cantidad += (h.cantidad || 1); ex.total += h.total; }
          else { consolidadas.push({ tipoBase: tipoBase, precio: h.precio, cantidad: h.cantidad || 1, total: h.total }); }
        });

        consolidadas.forEach(function(h) {
          var etiqueta = h.tipoBase + (h.cantidad > 1 ? " x" + h.cantidad : "");
          var netoFila  = h.total;
          var totalFila = Math.round(netoFila * 1.19);
          var fila = [ahora, datos.nombre_cliente || "", "Estándar",
                      etiqueta, checkin, checkout, noches, netoFila, totalFila, link];
          sheet.appendRow(fila);
          var uf = sheet.getLastRow();
          sheet.getRange(uf, 8).setNumberFormat("0");
          sheet.getRange(uf, 9).setNumberFormat("0");
        });

        if (datos.adicionales_costo && datos.adicionales_costo.length > 0) {
          datos.adicionales_costo.forEach(function(a) {
            if (!a.detalle) return;
            var netoAd  = Number(a.total || 0);
            var totalAd = Math.round(netoAd * 1.19);
            var fila = [ahora, datos.nombre_cliente || "", "Estándar",
                        "Adicional: " + a.detalle, checkin, checkout, 0, netoAd, totalAd, link];
            sheet.appendRow(fila);
            var uf = sheet.getLastRow();
            sheet.getRange(uf, 8).setNumberFormat("0");
            sheet.getRange(uf, 9).setNumberFormat("0");
          });
        }

        // Amenidades predefinidas
        if (datos.amenidades_predefinidas && datos.amenidades_predefinidas.length > 0) {
          datos.amenidades_predefinidas.forEach(function(a) {
            var netoAd = Number(a.total || 0);
            var totalAd = Math.round(netoAd * 1.19);
            var fila = [ahora, datos.nombre_cliente || "", "Estándar",
                        "Amenidad: " + a.nombre + (a.cantidad > 1 ? " x" + a.cantidad : ""), checkin, checkout, 0, netoAd, totalAd, link];
            sheet.appendRow(fila);
            var uf = sheet.getLastRow();
            sheet.getRange(uf, 8).setNumberFormat("0");
            sheet.getRange(uf, 9).setNumberFormat("0");
          });
        }

      } else {
        var neto  = Number(datos.subtotal || 0);
        var total = Number(datos.total    || 0);
        var fila = [ahora, datos.nombre_cliente || "", "Estándar",
                    "Estándar", checkin, checkout, noches, neto, total, link];
        sheet.appendRow(fila);
        var uf = sheet.getLastRow();
        sheet.getRange(uf, 8).setNumberFormat("0");
        sheet.getRange(uf, 9).setNumberFormat("0");
      }
    }

    SpreadsheetApp.flush();
  } catch(eHist) {
    console.log("Error al registrar en historial: " + eHist.toString());
  }
}

// =================================================================
// GUARDADO DEL PDF GENERADO EN EL NAVEGADOR (formato Word manual)
// El PDF llega en base64, se archiva en una subcarpeta mensual
// (2026-07, 2026-08, ...) SIN borrar respaldos anteriores, y la
// cotización se registra en el Historial con el link a su PDF.
// Si pdfBase64 viene vacío (p. ej. al generar solo Word), únicamente
// se registra en el Historial.
// =================================================================
function guardarPdfClienteYRegistrar(datos, pdfBase64, nombreArchivo) {
  var ss = SpreadsheetApp.openById(ID_PLANILLA_SHEETS);
  var url = "";

  if (pdfBase64) {
    var carpetaRaiz;
    var carpetas = DriveApp.getFoldersByName(NOMBRE_CARPETA_COTIZACIONES);
    if (carpetas.hasNext()) { carpetaRaiz = carpetas.next(); }
    else { carpetaRaiz = DriveApp.createFolder(NOMBRE_CARPETA_COTIZACIONES); }

    var tz = Session.getScriptTimeZone();
    var ahora = new Date();
    var nombreMes = Utilities.formatDate(ahora, tz, "yyyy-MM");
    var carpetaMes;
    var subcarpetas = carpetaRaiz.getFoldersByName(nombreMes);
    if (subcarpetas.hasNext()) { carpetaMes = subcarpetas.next(); }
    else { carpetaMes = carpetaRaiz.createFolder(nombreMes); }

    // nombre con fecha y hora para distinguir versiones del mismo cliente
    var prefijo = Utilities.formatDate(ahora, tz, "yyyy-MM-dd HH.mm");
    var nombreEnDrive = prefijo + " - " + (nombreArchivo || "Cotizacion Cascadas Hotel.pdf");

    var blob = Utilities.newBlob(
      Utilities.base64Decode(pdfBase64),
      "application/pdf",
      nombreEnDrive
    );
    var pdfFile = carpetaMes.createFile(blob);
    url = "https://drive.google.com/file/d/" + pdfFile.getId() + "/view";
  }

  registrarEnHistorial(ss, datos, url);

  return { pdfUrl: url, aviso: obtenerAvisoEspacioDrive() };
}

// Aviso discreto si el Drive se está quedando sin espacio (no bloquea nada)
function obtenerAvisoEspacioDrive() {
  try {
    var about = null;
    try { about = Drive.About.get(); } catch(e1) {
      try { about = Drive.About.get({ fields: "storageQuota" }); } catch(e2) {}
    }
    if (!about) return "";
    var total = 0, usado = 0;
    if (about.quotaBytesTotal) {
      total = Number(about.quotaBytesTotal);
      usado = Number(about.quotaBytesUsed || 0) + Number(about.quotaBytesUsedInTrash || 0);
    } else if (about.storageQuota) {
      total = Number(about.storageQuota.limit || 0);
      usado = Number(about.storageQuota.usage || 0);
    }
    if (!total) return ""; // cuota ilimitada o no informada
    var libreMB = Math.round((total - usado) / (1024 * 1024));
    if (libreMB < 500) {
      return "Quedan " + (libreMB < 1024 ? libreMB + " MB" : Math.round(libreMB / 102.4) / 10 + " GB") +
             " libres en el Drive. Considere liberar espacio para que los respaldos no fallen.";
    }
    return "";
  } catch(e) {
    return "";
  }
}

// =================================================================
// BUSCADOR DE COTIZACIONES POR CLIENTE
// Busca en el Historial (el índice) y agrupa las filas que nacieron
// de una misma cotización (misma marca de tiempo + cliente).
// =================================================================
function normalizarBusqueda(s) {
  return String(s == null ? "" : s).toLowerCase()
    .replace(/á/g, "a").replace(/é/g, "e").replace(/í/g, "i")
    .replace(/ó/g, "o").replace(/ú/g, "u").replace(/ñ/g, "n").trim();
}

function buscarCotizaciones(consulta) {
  try {
    var q = normalizarBusqueda(consulta);
    if (!q || q.length < 2) return { exito: true, grupos: [] };

    var ss = SpreadsheetApp.openById(ID_PLANILLA_SHEETS);
    var sheet = ss.getSheetByName("Historial");
    if (!sheet) return { exito: true, grupos: [] };

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { exito: true, grupos: [] };

    var tz = Session.getScriptTimeZone();
    var cab = data[0].map(function(c) { return c.toString().toLowerCase().trim(); });
    var iF = cab.indexOf("fecha");    if (iF < 0) iF = 0;
    var iC = cab.indexOf("cliente");  if (iC < 0) iC = 1;
    var iT = cab.indexOf("tipo");     if (iT < 0) iT = 2;
    var iD = cab.indexOf("detalle");  if (iD < 0) iD = 3;
    var iCi = cab.indexOf("checkin");
    var iCo = cab.indexOf("checkout");
    var iTot = cab.indexOf("total");  if (iTot < 0) iTot = 8;
    var iL = cab.indexOf("linkpdf");

    var grupos = {};
    for (var i = 1; i < data.length; i++) {
      var fila = data[i];
      var cliente = fila[iC] ? fila[iC].toString().trim() : "";
      if (!cliente || normalizarBusqueda(cliente).indexOf(q) < 0) continue;

      var rawFecha = fila[iF];
      var fecha = rawFecha instanceof Date ? rawFecha : new Date(String(rawFecha));
      if (isNaN(fecha.getTime())) continue;

      var clave = fecha.getTime() + "|" + cliente;
      if (!grupos[clave]) {
        grupos[clave] = {
          ts: fecha.getTime(),
          fecha: Utilities.formatDate(fecha, tz, "dd/MM/yyyy HH:mm"),
          cliente: cliente,
          tipo: fila[iT] ? fila[iT].toString() : "",
          checkin: iCi >= 0 && fila[iCi] ? (fila[iCi] instanceof Date ? Utilities.formatDate(fila[iCi], tz, "dd/MM/yyyy") : String(fila[iCi])) : "",
          checkout: iCo >= 0 && fila[iCo] ? (fila[iCo] instanceof Date ? Utilities.formatDate(fila[iCo], tz, "dd/MM/yyyy") : String(fila[iCo])) : "",
          detalles: [],
          total: 0,
          link: ""
        };
      }
      grupos[clave].detalles.push(fila[iD] ? fila[iD].toString() : "");
      grupos[clave].total += Number(fila[iTot]) || 0;
      if (iL >= 0 && fila[iL]) grupos[clave].link = fila[iL].toString();
    }

    var lista = Object.keys(grupos).map(function(k) { return grupos[k]; });
    lista.sort(function(a, b) { return b.ts - a.ts; });
    if (lista.length > 50) lista = lista.slice(0, 50);

    return { exito: true, grupos: lista };
  } catch(e) {
    return { exito: false, error: e.toString(), grupos: [] };
  }
}

// =================================================================
// MOTOR DE GENERACIÓN
// =================================================================
function generarDocumento(datos) {
  var ss = SpreadsheetApp.openById(ID_PLANILLA_SHEETS);

  var carpetaDestino;
  var carpetas = DriveApp.getFoldersByName(NOMBRE_CARPETA_COTIZACIONES);
  if (carpetas.hasNext()) { carpetaDestino = carpetas.next(); }
  else { carpetaDestino = DriveApp.createFolder(NOMBRE_CARPETA_COTIZACIONES); }

  vaciarCarpetaPorCompleto(carpetaDestino);

  var nombreArchivo = "Cotizacion Cascadas Hotel " + datos.nombre_cliente + " " + datos.checkin.replace(/\//g, "-");

  var plantillaFile = DriveApp.getFileById(ID_PLANTILLA);
  var copiaFile = plantillaFile.makeCopy(nombreArchivo, carpetaDestino);
  var docId = copiaFile.getId();

  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();

  var fechaEmision = datos.fecha_emision_larga || " ";
  var saludo = datos.saludo || "Estimado(a)";
  var nombreCliente = datos.nombre_cliente || " ";
  var nombreProg = (datos.tipo_cotizacion === "programa" && datos.nombre_programa) ? datos.nombre_programa : " ";

  body.replaceText("{{fecha_emision}}", fechaEmision);
  body.replaceText("{{saludo}}", saludo);
  body.replaceText("{{nombre_cliente}}", nombreCliente);
  body.replaceText("{{nombre_programa}}", datos.tipo_cotizacion === "programa" ? "Programa: " + nombreProg : "");

  var AZUL_HEADER = '#424143';
  var TEXTO_BLANCO = '#C6B39B';
  var TEXTO_OSCURO = '#2e2f30';
  var FONDO_BLANCO = '#FFFFFF';

  var estiloEncabezado = {
    [DocumentApp.Attribute.BACKGROUND_COLOR]: AZUL_HEADER,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_SIZE]: 9,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    [DocumentApp.Attribute.FOREGROUND_COLOR]: TEXTO_BLANCO
  };

  var estiloCelda = {
    [DocumentApp.Attribute.BACKGROUND_COLOR]: FONDO_BLANCO,
    [DocumentApp.Attribute.FONT_SIZE]: 9,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    [DocumentApp.Attribute.FOREGROUND_COLOR]: TEXTO_OSCURO
  };

  var estiloAdicional = {
    [DocumentApp.Attribute.BACKGROUND_COLOR]: FONDO_BLANCO,
    [DocumentApp.Attribute.FONT_SIZE]: 9,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    [DocumentApp.Attribute.ITALIC]: true,
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#475569'
  };

  var estiloTotalLabel = {
    [DocumentApp.Attribute.BACKGROUND_COLOR]: AZUL_HEADER,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_SIZE]: 9,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    [DocumentApp.Attribute.FOREGROUND_COLOR]: TEXTO_BLANCO
  };

  var estiloTotalValor = {
    [DocumentApp.Attribute.BACKGROUND_COLOR]: FONDO_BLANCO,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_SIZE]: 9,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    [DocumentApp.Attribute.FOREGROUND_COLOR]: TEXTO_OSCURO
  };

  var posValores = body.findText("{{tabla_valores}}");
  if (posValores) {
    var index = body.getChildIndex(posValores.getElement().getParent());
    body.removeChild(posValores.getElement().getParent());
    var tabla = body.insertTable(index);

    if (datos.tipo_cotizacion === "estandar") {

      var hr = tabla.appendTableRow();
      ["Servicio / Habitación", "Cant.", "Check-In", "Check-Out", "Noches", "Valor por Noche (por unidad)", "Total"].forEach(function(t, idx) {
        var celda = hr.appendTableCell(t);
        celda.setAttributes(estiloEncabezado);
        celda.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        if (idx !== 0) { aplicarNoWrapCelda(celda); }
        if (idx === 1 || idx === 4) {
          try { celda.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e) {}
        }
      });

      var habitacionesConsolidadas = [];
      datos.habitaciones.forEach(function(h) {
        var tipoBase = h.tipo.replace(/\s*\(x\d+\)$/, '').trim();
        var existente = habitacionesConsolidadas.find(function(e) { return e.tipoBase === tipoBase && e.precio === h.precio; });
        if (existente) { existente.cantidad += (h.cantidad || 1); existente.total += h.total; }
        else { habitacionesConsolidadas.push({ tipoBase: tipoBase, precio: h.precio, cantidad: h.cantidad || 1, total: h.total }); }
      });

      habitacionesConsolidadas.forEach(function(h) {
        var r = tabla.appendTableRow();
        [h.tipoBase, h.cantidad.toString(), datos.checkin, datos.checkout, datos.noches.toString(), formatearMoneda(h.precio), formatearMoneda(h.total)].forEach(function(txt, idx) {
          var celda = r.appendTableCell(txt).setAttributes(estiloCelda);
          celda.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
          if (idx !== 0) { aplicarNoWrapCelda(celda); }
          if (idx === 1 || idx === 4) {
            try { celda.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e) {}
          }
        });
      });

      if (datos.adicionales_costo && datos.adicionales_costo.length > 0) {
        datos.adicionales_costo.forEach(function(a) {
          var rA = tabla.appendTableRow();
          for (var i = 0; i < 7; i++) {
            var txt = "-";
            if (i === 0) txt = "➕ " + a.detalle;
            if (i === 6) txt = formatearMoneda(a.total);
            var celda = rA.appendTableCell(txt).setAttributes(estiloAdicional);
            celda.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
          }
          aplicarNoWrapCelda(rA.getCell(6));
        });
      }

      // Amenidades predefinidas
      if (datos.amenidades_predefinidas && datos.amenidades_predefinidas.length > 0) {
        datos.amenidades_predefinidas.forEach(function(a) {
          var rA = tabla.appendTableRow();
          for (var i = 0; i < 7; i++) {
            var txt = "-";
            if (i === 0) txt = a.nombre + (a.cantidad > 1 ? " (x" + a.cantidad + ")" : "") + (a.descripcion ? " — " + a.descripcion : "");
            if (i === 1) txt = a.cantidad.toString();
            if (i === 6) txt = formatearMoneda(a.total);
            var celda = rA.appendTableCell(txt).setAttributes(estiloCelda);
            celda.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
            if (i !== 0) { aplicarNoWrapCelda(celda); }
            if (i === 1) {
              try { celda.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e) {}
            }
          }
        });
      }

      try {
        var anchosPrincipal = [115, 33, 58, 58, 44, 59, 57];
        for (var ri = 0; ri < tabla.getNumRows(); ri++) {
          var row = tabla.getRow(ri);
          for (var ci = 0; ci < row.getNumCells() && ci < anchosPrincipal.length; ci++) {
            try { row.getCell(ci).setWidth(anchosPrincipal[ci]); } catch(ew) {}
          }
        }
      } catch(eWidth) {
        console.log("Error anchos tabla principal: " + eWidth.toString());
      }

      var atributosTabla = {};
      atributosTabla[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
      tabla.setAttributes(atributosTabla);

      var ANCHO_LABEL  = 366;
      var ANCHO_VALOR  = 57;
      var ANCHO_TOTALES = ANCHO_LABEL + ANCHO_VALOR;
      var PAGINA_UTIL  = 467;
      var INDENT_DERECHA = PAGINA_UTIL - ANCHO_TOTALES;

      var indexTabla = body.getChildIndex(tabla);
      var tablaTotales = body.insertTable(indexTabla + 1);

      var totalesData = [
        ["Subtotal Neto", formatearMoneda(datos.subtotal)],
        ["IVA (19%)",     formatearMoneda(datos.iva)],
        ["Total Final",   formatearMoneda(datos.total)]
      ];

      totalesData.forEach(function(fila) {
        var row = tablaTotales.appendTableRow();

        var celdaLabel = row.appendTableCell(fila[0]);
        celdaLabel.setAttributes(estiloTotalLabel);
        celdaLabel.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        try { celdaLabel.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT); } catch(e) {}

        var celdaValor = row.appendTableCell(fila[1]);
        celdaValor.setAttributes(estiloTotalValor);
        celdaValor.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        aplicarNoWrapCelda(celdaValor);
        try { celdaValor.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT); } catch(e) {}
      });

      try {
        for (var ri = 0; ri < tablaTotales.getNumRows(); ri++) {
          var row = tablaTotales.getRow(ri);
          try { row.getCell(0).setWidth(ANCHO_LABEL); }  catch(ew) {}
          try { row.getCell(1).setWidth(ANCHO_VALOR); }  catch(ew) {}
        }
      } catch(eWT) {}

      var atributosTotales = {};
      atributosTotales[DocumentApp.Attribute.INDENT_START] = INDENT_DERECHA;
      tablaTotales.setAttributes(atributosTotales);

    } else {
      var hr = tabla.appendTableRow();
      ["Programa", "Check-In", "Check-Out", "Noches", "Valor Neto", "IVA (19%)", "Total"].forEach(function(t) {
        var celda = hr.appendTableCell(t);
        celda.setAttributes(estiloEncabezado);
        celda.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        aplicarNoWrapCelda(celda);
      });
      var r = tabla.appendTableRow();
      var neto = datos.total_programa || 0;
      var iva = Math.round(neto * 0.19);
      var total = neto + iva;
      var nombreProgCant = nombreProg;
      if (datos.cantidad_programa && datos.cantidad_programa > 1) {
        nombreProgCant += " (x" + datos.cantidad_programa + ")";
      }
      [nombreProgCant, datos.checkin, datos.checkout, datos.noches.toString(), formatearMoneda(neto), formatearMoneda(iva), formatearMoneda(total)].forEach(function(txt) {
        var celda = r.appendTableCell(txt).setAttributes(estiloCelda);
        celda.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        aplicarNoWrapCelda(celda);
      });
      var atributosProg = {};
      atributosProg[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
      tabla.setAttributes(atributosProg);
    }
  }

  var posAdj = body.findText("{{tabla_adicionales}}");
  if (posAdj) {
    var elementAdj = posAdj.getElement().getParent();
    var indexAdj = body.getChildIndex(elementAdj);
    body.removeChild(elementAdj);
    if (datos.tipo_cotizacion === "programa" && datos.ticks && datos.ticks.length > 0) {
      var tAdj = body.insertTable(indexAdj);
      var hrAdj = tAdj.appendTableRow();
      hrAdj.appendTableCell("Servicios Coberturas Especiales Incluidas").setAttributes(estiloEncabezado).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      datos.ticks.forEach(function(tk) {
        var r = tAdj.appendTableRow();
        r.appendTableCell("✓ " + tk).setAttributes(estiloCelda).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      });
      var atributosAdj = {};
      atributosAdj[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
      tAdj.setAttributes(atributosAdj);
    }
  }

  var alturaNecesaria = estimarAlturaContenido(doc);

  doc.saveAndClose();
  SpreadsheetApp.flush();
  Utilities.sleep(500);

  ajustarAlturaPagina(docId, alturaNecesaria);
  Utilities.sleep(600);

  var pdfBlob = copiaFile.getAs(MimeType.PDF).setName(nombreArchivo + ".pdf");
  var pdfFile = carpetaDestino.createFile(pdfBlob);
  var wordUrl = "https://docs.google.com/document/d/" + docId + "/export?format=docx";

  try { Drive.Files.remove(docId); } catch(errF) { try { copiaFile.setTrashed(true); } catch(e) {} }

  registrarEnHistorial(ss, datos);

  return { pdfUrl: pdfFile.getDownloadUrl().replace("?e=download&gd=true", ""), wordUrl: wordUrl };
}

function vaciarCarpetaPorCompleto(carpeta) {
  try {
    var archivos = carpeta.getFiles();
    while (archivos.hasNext()) {
      var archivo = archivos.next();
      try { Drive.Files.remove(archivo.getId()); } catch(err) { archivo.setTrashed(true); }
    }
  } catch(e) { console.log("Error al vaciar la carpeta: " + e.toString()); }
}

function aplicarCeldaInvisible(celda) {
  try { celda.setBorderColor('#FFFFFF'); } catch(e) {}
  return celda;
}

function aplicarNoWrapCelda(celda) {
  try {
    var parrafo = celda.getChild(0);
    if (parrafo && parrafo.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var textoElemento = parrafo.asParagraph().getChild(0);
      if (textoElemento && textoElemento.getType() === DocumentApp.ElementType.TEXT) {
        textoElemento.asText().setNoWrap(true);
      }
    }
  } catch(e) {}
}

function formatearMoneda(v) { return "$" + v.toLocaleString('es-CL'); }

// =================================================================
// EDICIÓN DE PLANTILLA
// =================================================================
// Extrae los tramos de formato (negrita/cursiva/subrayado) de un
// párrafo o ítem de lista, para que el PDF generado en el navegador
// respete las negritas parciales de la plantilla.
function extraerRunsDeTexto(textoElemento) {
  var runs = [];
  try {
    var texto = textoElemento.getText();
    if (!texto) return runs;
    var indices = textoElemento.getTextAttributeIndices();
    if (!indices || indices.length === 0) indices = [0];
    for (var i = 0; i < indices.length; i++) {
      var ini = indices[i];
      var fin = (i + 1 < indices.length) ? indices[i + 1] : texto.length;
      if (fin <= ini) continue;
      runs.push({
        t: texto.substring(ini, fin),
        b: textoElemento.isBold(ini) === true,
        i: textoElemento.isItalic(ini) === true,
        u: textoElemento.isUnderline(ini) === true
      });
    }
  } catch(e) {}
  return runs;
}

// Detecta negrita de un elemento mirando TODAS las señales que da Google Docs:
// el atributo del párrafo, la negrita del texto completo, y la del primer
// carácter (Docs a veces informa por una vía y a veces por otra).
function detectarBoldElemento(elemento, textoEd) {
  try {
    if (elemento.getAttributes()[DocumentApp.Attribute.BOLD] === true) return true;
  } catch(e) {}
  try {
    if (textoEd && textoEd.getText() && textoEd.isBold() === true) return true;
  } catch(e) {}
  return false;
}

function obtenerElementosPlantilla() {
  try {
    var doc = DocumentApp.openById(ID_PLANTILLA);
    var body = doc.getBody();
    var elementos = [];
    var contador = 0;
    for (var i = 0; i < body.getNumChildren(); i++) {
      var hijo = body.getChild(i);
      var tipo = hijo.getType();
      if (tipo === DocumentApp.ElementType.PARAGRAPH) {
        var parrafo = hijo.asParagraph();
        var texto = parrafo.getText();
        var heading = parrafo.getHeading();
        var tipoNombre = 'parrafo';
        if (heading === DocumentApp.ParagraphHeading.HEADING1) tipoNombre = 'heading1';
        else if (heading === DocumentApp.ParagraphHeading.HEADING2) tipoNombre = 'heading2';
        else if (heading === DocumentApp.ParagraphHeading.HEADING3) tipoNombre = 'heading3';
        var textoEdP = parrafo.editAsText();
        elementos.push({ id: contador, ruta: 'body.' + i, indiceBody: i, texto: texto, tipo: tipoNombre, editable: true, isBold: detectarBoldElemento(parrafo, textoEdP), runs: extraerRunsDeTexto(textoEdP) });
        contador++;
      } else if (tipo === DocumentApp.ElementType.LIST_ITEM) {
        var listItem = hijo.asListItem();
        var texto = listItem.getText();
        var textoEdL = listItem.editAsText();
        elementos.push({ id: contador, ruta: 'body.' + i, indiceBody: i, texto: texto, tipo: 'lista', editable: true, isBold: detectarBoldElemento(listItem, textoEdL), nestingLevel: listItem.getNestingLevel(), runs: extraerRunsDeTexto(textoEdL) });
        contador++;
      } else if (tipo === DocumentApp.ElementType.TABLE) {
        var tabla = hijo.asTable();
        var filas = tabla.getNumRows();
        var cols = tabla.getRow(0) ? tabla.getRow(0).getNumCells() : 0;
        elementos.push({ id: contador, ruta: 'body.' + i, indiceBody: i, texto: '[TABLA: ' + filas + ' filas × ' + cols + ' columnas]', tipo: 'tabla', editable: false });
        contador++;
      }
    }
    return { exito: true, elementos: elementos };
  } catch(e) {
    return { exito: false, error: e.toString(), elementos: [] };
  }
}

function guardarElementoPlantilla(indiceBody, textoNuevo, tipoElemento) {
  try {
    var doc = DocumentApp.openById(ID_PLANTILLA);
    var body = doc.getBody();
    var elemento = body.getChild(indiceBody);
    if (!elemento) return { exito: false, error: 'No se encontró el elemento en el índice ' + indiceBody };
    var tipoReal = elemento.getType();
    if (tipoReal === DocumentApp.ElementType.PARAGRAPH) {
      var parrafo = elemento.asParagraph();
      var atributos = parrafo.getAttributes();
      var atributosTexto = {};
      try { if (parrafo.getNumChildren() > 0) { var pc = parrafo.getChild(0); if (pc.getType() === DocumentApp.ElementType.TEXT) atributosTexto = pc.asText().getAttributes(); } } catch(e2) {}
      parrafo.clear();
      var nuevoTexto = parrafo.appendText(textoNuevo);
      if (Object.keys(atributosTexto).length > 0) nuevoTexto.setAttributes(atributosTexto);
      parrafo.setAttributes(atributos);
    } else if (tipoReal === DocumentApp.ElementType.LIST_ITEM) {
      var listItem = elemento.asListItem();
      var atributos = listItem.getAttributes();
      var atributosTexto = {};
      try { if (listItem.getNumChildren() > 0) { var pc = listItem.getChild(0); if (pc.getType() === DocumentApp.ElementType.TEXT) atributosTexto = pc.asText().getAttributes(); } } catch(e2) {}
      listItem.clear();
      var nuevoTexto = listItem.appendText(textoNuevo);
      if (Object.keys(atributosTexto).length > 0) nuevoTexto.setAttributes(atributosTexto);
      listItem.setAttributes(atributos);
    } else {
      return { exito: false, error: 'El elemento en índice ' + indiceBody + ' no es editable (tipo: ' + tipoReal + ')' };
    }
    doc.saveAndClose();
    return { exito: true };
  } catch(e) {
    return { exito: false, error: e.toString() };
  }
}

// =================================================================
// REPORTES
// =================================================================
function obtenerDatosReporte(filtros) {
  try {
    var ss = SpreadsheetApp.openById(ID_PLANILLA_SHEETS);
    var sheet = ss.getSheetByName("Historial");
    var tz = Session.getScriptTimeZone();

    var metricsVacias = { totalNeto: 0, totalConIva: 0, quantityTotal: 0, quantityEstandar: 0, quantityProgramas: 0, promedioNoches: 0, totalEstandar: 0, totalProgramas: 0 };

    if (!sheet) {
      sheet = ss.insertSheet("Historial");
      sheet.appendRow(["Fecha", "Cliente", "Tipo", "Detalle", "CheckIn", "CheckOut", "Noches", "Neto", "Total"]);
      sheet.setFrozenRows(1);
      return { exito: true, registros: [], metricas: metricsVacias, porDetalle: {}, porHora: {} };
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { exito: true, registros: [], metricas: metricsVacias, porDetalle: {}, porHora: {} };
    }

    var cabecera = data[0].map(function(c) { return c.toString().toLowerCase().trim(); });
    var idxFecha    = cabecera.indexOf("fecha");
    var idxCliente  = cabecera.indexOf("cliente");
    var idxTipo     = cabecera.indexOf("tipo");
    var idxDetalle  = cabecera.indexOf("detalle");
    var idxCheckIn  = cabecera.indexOf("checkin");
    var idxCheckOut = cabecera.indexOf("checkout");
    var idxNoches   = cabecera.indexOf("noches");
    var idxNeto     = cabecera.indexOf("neto");
    var idxTotal    = cabecera.indexOf("total");

    if (idxFecha    < 0) idxFecha    = 0;
    if (idxCliente  < 0) idxCliente  = 1;
    if (idxTipo     < 0) idxTipo     = 2;
    if (idxDetalle  < 0) idxDetalle  = 3;
    if (idxCheckIn  < 0) idxCheckIn  = -1;
    if (idxCheckOut < 0) idxCheckOut = -1;
    if (idxNoches   < 0) idxNoches   = 4;
    if (idxNeto     < 0) idxNeto     = 5;
    if (idxTotal    < 0) idxTotal    = 6;

    var modoFecha = (filtros && filtros.modo === "checkin") ? "checkin" : "emision";
    var fechaDesde = null, fechaHasta = null;
    if (filtros && filtros.fechaDesde) fechaDesde = new Date(filtros.fechaDesde + "T00:00:00");
    if (filtros && filtros.fechaHasta) fechaHasta = new Date(filtros.fechaHasta + "T23:59:59");

    var registros = [];
    var porHora = {};
    for (var h = 0; h < 24; h++) { porHora[h] = 0; }

    for (var i = 1; i < data.length; i++) {
      var fila = data[i];
      var rawFecha = fila[idxFecha];
      if (!rawFecha) continue;

      var fechaEmision;
      try {
        fechaEmision = rawFecha instanceof Date ? rawFecha : new Date(rawFecha.toString());
        if (isNaN(fechaEmision.getTime())) continue;
      } catch(ef) { continue; }

      var checkin = "", checkout = "";
      var fechaCheckinObj = null;

      if (idxCheckIn >= 0 && fila[idxCheckIn]) {
        var rawCI = fila[idxCheckIn];
        if (rawCI instanceof Date && !isNaN(rawCI.getTime())) {
          checkin = Utilities.formatDate(rawCI, tz, "dd/MM/yyyy");
          fechaCheckinObj = rawCI;
        } else {
          var ciStr = rawCI.toString().trim();
          checkin = ciStr;
          if (ciStr && ciStr.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
            var partes = ciStr.split('/');
            fechaCheckinObj = new Date(partes[2] + "-" + partes[1] + "-" + partes[0] + "T12:00:00");
            if (isNaN(fechaCheckinObj.getTime())) fechaCheckinObj = null;
          }
        }
      }
      if (idxCheckOut >= 0 && fila[idxCheckOut]) {
        var rawCO = fila[idxCheckOut];
        if (rawCO instanceof Date && !isNaN(rawCO.getTime())) {
          checkout = Utilities.formatDate(rawCO, tz, "dd/MM/yyyy");
        } else {
          checkout = rawCO.toString().trim();
        }
      }

      if (fechaDesde || fechaHasta) {
        if (modoFecha === "emision") {
          if (fechaDesde && fechaEmision < fechaDesde) continue;
          if (fechaHasta && fechaEmision > fechaHasta) continue;
        } else {
          if (!fechaCheckinObj) continue;
          if (fechaDesde && fechaCheckinObj < fechaDesde) continue;
          if (fechaHasta && fechaCheckinObj > fechaHasta) continue;
        }
      }

      var tipo    = fila[idxTipo]    ? fila[idxTipo].toString().trim()    : "Estándar";
      var cliente = fila[idxCliente] ? fila[idxCliente].toString().trim() : "";
      var detalle = fila[idxDetalle] ? fila[idxDetalle].toString().trim() : "";
      var noches  = Number(fila[idxNoches]) || 0;
      var neto    = Number(fila[idxNeto])   || 0;
      var total   = Number(fila[idxTotal])  || 0;

      var tipoNorm = tipo.toLowerCase().replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i').replace(/ó/g,'o').replace(/ú/g,'u');
      var esPrograma = tipoNorm === "programa" || tipoNorm.indexOf("prog") === 0;

      var horaEmision = fechaEmision.getHours();
      if (horaEmision >= 0 && horaEmision < 24) porHora[horaEmision]++;

      registros.push({
        fecha:    Utilities.formatDate(fechaEmision, tz, "dd/MM/yyyy"),
        fechaTs:  fechaEmision.getTime(),
        tipo:     esPrograma ? "Programa" : "Estándar",
        cliente:  cliente,
        detalle:  detalle,
        checkin:  checkin,
        checkout: checkout,
        noches:   noches,
        neto:     neto,
        total:    total
      });
    }

    registros.sort(function(a, b) { return b.fechaTs - a.fechaTs; });

    var cantidadTotal    = registros.length;
    var cantidadEstandar = registros.filter(function(r){ return r.tipo === "Estándar"; }).length;
    var cantidadProgramas= registros.filter(function(r){ return r.tipo === "Programa"; }).length;
    var totalNeto        = registros.reduce(function(s,r){ return s + r.neto; }, 0);
    var totalConIva      = registros.reduce(function(s,r){ return s + r.total; }, 0);
    var totalEstandar    = registros.filter(function(r){ return r.tipo === "Estándar"; }).reduce(function(s,r){ return s + r.total; }, 0);
    var totalProgramas   = registros.filter(function(r){ return r.tipo === "Programa"; }).reduce(function(s,r){ return s + r.total; }, 0);
    var promedioNoches   = cantidadTotal > 0 ? Math.round(registros.reduce(function(s,r){ return s + r.noches; }, 0) / cantidadTotal * 10) / 10 : 0;

    var porDetalle = {};
    registros.forEach(function(r) {
      var key = r.detalle || "Sin detalle";
      if (!porDetalle[key]) porDetalle[key] = 0;
      porDetalle[key]++;
    });

    return {
      exito: true,
      registros: registros,
      metricas: {
        totalNeto:         totalNeto,
        totalConIva:       totalConIva,
        cantidadTotal:     cantidadTotal,
        cantidadEstandar:  cantidadEstandar,
        cantidadProgramas: cantidadProgramas,
        promedioNoches:    promedioNoches,
        totalEstandar:     totalEstandar,
        totalProgramas:    totalProgramas
      },
      porDetalle: porDetalle,
      porHora:    porHora
    };

  } catch(e) {
    return { exito: false, error: e.toString() };
  }
}

// =================================================================
// ALTURA ADAPTATIVA
// =================================================================

function estimarAlturaContenido(doc) {
  var body = doc.getBody();
  var altura = 0;
  for (var i = 0; i < body.getNumChildren(); i++) {
    var child = body.getChild(i);
    var tipo = child.getType();
    if (tipo === DocumentApp.ElementType.PARAGRAPH) {
      var p = child.asParagraph();
      var heading = p.getHeading();
      if (heading === DocumentApp.ParagraphHeading.HEADING1) altura += 24;
      else if (heading === DocumentApp.ParagraphHeading.HEADING2) altura += 21;
      else if (heading === DocumentApp.ParagraphHeading.HEADING3) altura += 18;
      else altura += 15;
    } else if (tipo === DocumentApp.ElementType.TABLE) {
      altura += child.asTable().getNumRows() * 21;
    } else if (tipo === DocumentApp.ElementType.LIST_ITEM) {
      altura += 15;
    }
    altura += 3;
  }
  return altura + 100;
}

function ajustarAlturaPagina(docId, alturaEstimada) {
  try {
    var nuevaAltura = Math.max(Math.round(alturaEstimada) + 1600, 900);
    if (nuevaAltura > 3400) nuevaAltura = 3400;

    Docs.Documents.batchUpdate({
      requests: [{
        updateDocumentStyle: {
          documentStyle: {
            pageSize: {
              height: { magnitude: nuevaAltura, unit: "PT" }
            }
          },
          fields: "pageSize.height"
        }
      }]
    }, docId);
  } catch(e) {
    console.log("Error ajustando altura de página: " + e.toString());
  }
}
