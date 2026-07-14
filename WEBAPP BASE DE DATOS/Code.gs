const SHEET_ID = '1e9lqFYqh_ebtfrTZexm-iQys_Ii1LdobpQdxaP9O9Cw';
const SHEET_HUESPEDES = 'Huespedes';
const SHEET_PROGRAMAS = 'Programas';

/* Programas por defecto (se siembran la primera vez en la hoja "Programas").
   A partir de ahí se administran desde la app: crear, editar y borrar burbujas. */
const DEFAULT_PROGRAMAS = [
  { nombre: 'Estándar',     emoji: '🏷️' },
  { nombre: 'Residentes',   emoji: '🏠' },
  { nombre: 'Booking',      emoji: '🌐' },
  { nombre: 'Directo',      emoji: '➡️' },
  { nombre: 'Especial',     emoji: '⭐' },
  { nombre: 'Airbnb',       emoji: '🏡' },
  { nombre: 'Despegar',     emoji: '✈️' },
  { nombre: 'Expedia',      emoji: '🧳' },
  { nombre: 'Operador',     emoji: '🏢' },
  { nombre: 'Corporativo',  emoji: '🏭' },
  { nombre: 'Convenio',     emoji: '🤝' }
];

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Cascadas Hotel')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerUrlSheet() {
  return 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit';
}

/* ============================================================
   HELPERS: búsqueda de columnas por nombre (robusto)
   ============================================================ */
function getColumnIndex(headers, name) {
  const clean = name.toString().toLowerCase().replace(/[_\s]/g, '');
  for (let i = 0; i < headers.length; i++) {
    const h = (headers[i] || '').toString().toLowerCase().replace(/[_\s]/g, '');
    if (h === clean) return i;
  }
  return -1;
}

/* ============================================================
   NORMALIZACIÓN DE NOMBRES DE PROGRAMA
   Para comparar sin distinguir mayúsculas/minúsculas ni acentos,
   así "Residentes", "RESIDENTES" y "residentes" cuentan como el mismo
   programa (evita clones en el reporte y en la base de datos).
   ============================================================ */
function _normPrograma(s) {
  let t = String(s == null ? '' : s).trim().toLowerCase();
  try { t = t.normalize('NFD').replace(/[̀-ͯ]/g, ''); } catch (e) {}
  return t;
}

/* Devuelve el nombre "canónico" de un programa: si ya existe en la hoja
   Programas (comparando sin mayúsculas/acentos), devuelve ese nombre tal cual
   está guardado; si no, devuelve el texto recibido (recortado). */
function _programaCanonico(ss, nombre) {
  const objetivo = _normPrograma(nombre);
  if (!objetivo) return '';
  try {
    const sh = getProgramasSheet_(ss);
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const cand = (data[i][0] || '').toString().trim();
      if (cand && _normPrograma(cand) === objetivo) return cand;
    }
  } catch (e) {}
  return String(nombre == null ? '' : nombre).trim();
}

/* Asegura que existan las columnas nuevas al final del header.
   Se agregó "Moneda" para poder registrar tarifas en dólares. */
function ensureColumns(ss) {
  const sheet = ss.getSheetByName(SHEET_HUESPEDES);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const required = ['Check_out', 'Noches', 'Edad', 'Total_estadia', 'Moneda'];
  const missing = [];

  required.forEach(name => {
    if (getColumnIndex(headers, name) < 0) missing.push(name);
  });

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((name, i) => {
      sheet.getRange(1, startCol + i).setValue(name);
    });
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  return headers;
}

/* ============================================================
   GUARDAR
   ============================================================ */
function guardarHuesped(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_HUESPEDES);
  const headers = ensureColumns(ss);

  // Normaliza el nombre del programa al canónico de la hoja Programas, sin
  // distinguir mayúsculas/acentos (evita clones tipo "Residentes"/"RESIDENTES").
  data.programa = _programaCanonico(ss, data.programa || 'Estándar');

  const fechaCheckin = new Date(data.checkin);
  const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const mesTexto = meses[fechaCheckin.getMonth()] + ' ' + fechaCheckin.getFullYear();

  const colMap = {
    'Nombre': data.nombre,
    'Nacionalidad': data.nacionalidad,
    'Fecha_Nacimiento': data.fechaNac,
    'Check_in': data.checkin,
    'Mes': mesTexto,
    'Tarifa_por_Noche': data.tarifa,
    'Numero_contacto_o_Correo': data.contacto,
    'Motivo': data.motivo,
    'Medio': data.medio || 'No especificado',
    'Programa': data.programa || 'ESTANDAR',
    'Check_out': data.checkout || '',
    'Noches': data.noches || '',
    'Edad': data.edad || '',
    'Total_estadia': data.totalEstadia || '',
    'Moneda': data.moneda || 'CLP',
    'Pre_Reserva': '',
    'Booking': '',
    'Walk_in': ''
  };

  const rowData = headers.map(h => {
    const key = h.toString().replace(/[_\s]/g, '').toLowerCase();
    for (let k in colMap) {
      if (k.toString().replace(/[_\s]/g, '').toLowerCase() === key) return colMap[k];
    }
    return '';
  });

  sheet.appendRow(rowData);
  return { ok: true, mensaje: 'Huésped registrado' };
}

/* ============================================================
   ACTUALIZAR
   ============================================================ */
function actualizarHuesped(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_HUESPEDES);
  const headers = ensureColumns(ss);

  // Normaliza el nombre del programa al canónico de la hoja Programas, sin
  // distinguir mayúsculas/acentos (evita clones tipo "Residentes"/"RESIDENTES").
  data.programa = _programaCanonico(ss, data.programa || 'Estándar');

  const fechaCheckin = new Date(data.checkin);
  const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const mesTexto = meses[fechaCheckin.getMonth()] + ' ' + fechaCheckin.getFullYear();

  const colMap = {
    'Nombre': data.nombre,
    'Nacionalidad': data.nacionalidad,
    'Fecha_Nacimiento': data.fechaNac,
    'Check_in': data.checkin,
    'Mes': mesTexto,
    'Tarifa_por_Noche': data.tarifa,
    'Numero_contacto_o_Correo': data.contacto,
    'Motivo': data.motivo,
    'Medio': data.medio || 'No especificado',
    'Programa': data.programa || 'ESTANDAR',
    'Check_out': data.checkout || '',
    'Noches': data.noches || '',
    'Edad': data.edad || '',
    'Total_estadia': data.totalEstadia || '',
    'Moneda': data.moneda || 'CLP',
    'Pre_Reserva': '',
    'Booking': '',
    'Walk_in': ''
  };

  const rowData = headers.map(h => {
    const key = h.toString().replace(/[_\s]/g, '').toLowerCase();
    for (let k in colMap) {
      if (k.toString().replace(/[_\s]/g, '').toLowerCase() === key) return colMap[k];
    }
    return '';
  });

  const range = sheet.getRange(data.rowIndex, 1, 1, rowData.length);
  range.setValues([rowData]);
  return { ok: true, mensaje: 'Huésped actualizado' };
}

/* ============================================================
   ELIMINAR
   ============================================================ */
function eliminarHuesped(rowIndex) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_HUESPEDES);
  sheet.deleteRow(rowIndex);
  return { ok: true, mensaje: 'Huésped eliminado' };
}

/* ============================================================
   BUSCAR
   ============================================================ */
function buscarHuespedes(query, fechaFiltro) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_HUESPEDES);
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const rows = data.slice(1);

  const idx = name => getColumnIndex(headers, name);

  let resultados = rows.map((r, i) => ({
    rowIndex: i + 2,
    nombre: idx('Nombre') >= 0 ? r[idx('Nombre')] : '',
    nacionalidad: idx('Nacionalidad') >= 0 ? r[idx('Nacionalidad')] : '',
    fechaNac: idx('Fecha_Nacimiento') >= 0 ? r[idx('Fecha_Nacimiento')] : '',
    checkin: idx('Check_in') >= 0 ? r[idx('Check_in')] : '',
    checkout: idx('Check_out') >= 0 ? r[idx('Check_out')] : '',
    noches: idx('Noches') >= 0 ? r[idx('Noches')] : '',
    tarifa: idx('Tarifa_por_Noche') >= 0 ? r[idx('Tarifa_por_Noche')] : '',
    total: idx('Total_estadia') >= 0 ? r[idx('Total_estadia')] : '',
    moneda: idx('Moneda') >= 0 ? (r[idx('Moneda')] || 'CLP') : 'CLP',
    programa: idx('Programa') >= 0 ? r[idx('Programa')] : '',
    contacto: idx('Numero_contacto_o_Correo') >= 0 ? r[idx('Numero_contacto_o_Correo')] : '',
    motivo: idx('Motivo') >= 0 ? r[idx('Motivo')] : '',
    medio: idx('Medio') >= 0 ? r[idx('Medio')] : '',
    edad: idx('Edad') >= 0 ? r[idx('Edad')] : ''
  }));

  if (query && query.trim()) {
    const q = query.toLowerCase();
    resultados = resultados.filter(r => r.nombre.toLowerCase().includes(q));
  }

  if (fechaFiltro && fechaFiltro.trim()) {
    const f = String(fechaFiltro).trim();
    resultados = resultados.filter(r => {
      const cin = String(r.checkin || '').trim();
      return cin === f;
    });
  }

  return resultados;
}

/* ============================================================
   PROGRAMAS  (persistidos y editables desde la app)
   ============================================================ */
function getProgramasSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_PROGRAMAS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PROGRAMAS);
    sh.appendRow(['Nombre', 'Emoji']);
    DEFAULT_PROGRAMAS.forEach(p => sh.appendRow([p.nombre, p.emoji]));
    return sh;
  }
  // Si existe pero está vacía, sembrar por defecto
  if (sh.getLastRow() <= 1) {
    if (sh.getLastRow() === 0) sh.appendRow(['Nombre', 'Emoji']);
    DEFAULT_PROGRAMAS.forEach(p => sh.appendRow([p.nombre, p.emoji]));
  }
  return sh;
}

/* Devuelve [{nombre, emoji}] en el orden guardado. */
function obtenerProgramas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = getProgramasSheet_(ss);
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const nombre = (data[i][0] || '').toString().trim();
    if (!nombre) continue;
    out.push({ nombre: nombre, emoji: (data[i][1] || '').toString().trim() });
  }
  return out;
}

/* Crea un programa nuevo o renombra uno existente (si se envía oldNombre).
   Evita duplicados por nombre (sin distinguir mayúsculas). */
function guardarPrograma(nombre, emoji, oldNombre) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = getProgramasSheet_(ss);
  const data = sh.getDataRange().getValues();
  nombre = (nombre || '').toString().trim();
  emoji = (emoji || '').toString().trim();
  oldNombre = (oldNombre || '').toString().trim();

  if (!nombre) return { ok: false, error: 'El nombre no puede estar vacío', programas: obtenerProgramas() };

  const nombreKey = nombre.toLowerCase();

  // Buscar fila del programa a editar (si aplica)
  let filaEditar = -1;
  for (let i = 1; i < data.length; i++) {
    const n = (data[i][0] || '').toString().trim().toLowerCase();
    if (oldNombre && n === oldNombre.toLowerCase()) { filaEditar = i + 1; }
  }

  // Verificar duplicado (que no sea la misma fila que editamos)
  for (let i = 1; i < data.length; i++) {
    const n = (data[i][0] || '').toString().trim().toLowerCase();
    if (n === nombreKey && (i + 1) !== filaEditar) {
      return { ok: false, error: 'Ya existe un programa con ese nombre', programas: obtenerProgramas() };
    }
  }

  if (filaEditar > 0) {
    sh.getRange(filaEditar, 1).setValue(nombre);
    sh.getRange(filaEditar, 2).setValue(emoji);
  } else {
    sh.appendRow([nombre, emoji]);
  }
  SpreadsheetApp.flush();
  return { ok: true, programas: obtenerProgramas() };
}

/* Borra un programa por nombre. */
function eliminarPrograma(nombre) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = getProgramasSheet_(ss);
  const data = sh.getDataRange().getValues();
  const key = (nombre || '').toString().trim().toLowerCase();
  for (let i = data.length - 1; i >= 1; i--) {
    const n = (data[i][0] || '').toString().trim().toLowerCase();
    if (n === key) sh.deleteRow(i + 1);
  }
  SpreadsheetApp.flush();
  return { ok: true, programas: obtenerProgramas() };
}

/* ============================================================
   EDAD
   ============================================================ */
function calcularEdad(fechaNacStr) {
  if (!fechaNacStr) return null;
  const hoy = new Date();
  const partes = String(fechaNacStr).trim().split('-');
  if (partes.length !== 3) return null;
  const fn = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));
  if (isNaN(fn.getTime())) return null;
  let edad = hoy.getFullYear() - fn.getFullYear();
  const m = hoy.getMonth() - fn.getMonth();
  if (m < 0 || (m === 0 && hoy.getDate() < fn.getDate())) edad--;
  return edad >= 0 ? edad : null;
}

/* ============================================================
   REPORTE (versión ampliada)
   ============================================================ */
function numero_(v) {
  const n = parseFloat(String(v == null ? '' : v).replace(/[^0-9.-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function obtenerDatosReporte(filtros) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_HUESPEDES);
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const rows = data.slice(1);
  const idx = name => getColumnIndex(headers, name);

  let filtrados = rows;

  if (filtros.fechaDesde && filtros.fechaHasta) {
    const desde = String(filtros.fechaDesde).trim();
    const hasta = String(filtros.fechaHasta).trim();
    filtrados = filtrados.filter(r => {
      const fechaStr = idx('Check_in') >= 0 ? String(r[idx('Check_in')] || '').trim() : '';
      return fechaStr >= desde && fechaStr <= hasta;
    });
  }

  const totalHabitaciones = filtrados.length;

  // Moneda por fila
  const monedaDe = r => {
    const m = idx('Moneda') >= 0 ? String(r[idx('Moneda')] || 'CLP').trim().toUpperCase() : 'CLP';
    return m === 'USD' || m === 'US$' || m === 'DOLAR' || m === 'DÓLAR' ? 'USD' : 'CLP';
  };
  // Total de la estadía por fila (usa Total_estadia; si no, tarifa*noches)
  const totalDe = r => {
    let t = idx('Total_estadia') >= 0 ? numero_(r[idx('Total_estadia')]) : 0;
    if (!t) {
      const tar = idx('Tarifa_por_Noche') >= 0 ? numero_(r[idx('Tarifa_por_Noche')]) : 0;
      const noc = idx('Noches') >= 0 ? numero_(r[idx('Noches')]) : 0;
      t = tar * noc;
    }
    return t;
  };

  // Tarifa promedio (solo CLP, para no mezclar monedas)
  const tarifasCLP = filtrados.filter(r => monedaDe(r) === 'CLP')
    .map(r => idx('Tarifa_por_Noche') >= 0 ? numero_(r[idx('Tarifa_por_Noche')]) : 0)
    .filter(v => v > 0);
  const tarifaPromedio = tarifasCLP.length > 0 ? Math.round(tarifasCLP.reduce((a, b) => a + b, 0) / tarifasCLP.length) : 0;

  // Ingresos por moneda
  let ingresosCLP = 0, ingresosUSD = 0;
  filtrados.forEach(r => { (monedaDe(r) === 'USD' ? (ingresosUSD += totalDe(r)) : (ingresosCLP += totalDe(r))); });
  ingresosCLP = Math.round(ingresosCLP);
  ingresosUSD = Math.round(ingresosUSD);
  const totalTarifa = ingresosCLP; // compat: KPI "Total" en CLP

  // Noches
  const nochesArr = filtrados.map(r => idx('Noches') >= 0 ? numero_(r[idx('Noches')]) : 0);
  const nochesTotales = nochesArr.reduce((a, b) => a + b, 0);
  const nochesPromedio = totalHabitaciones > 0 ? Math.round((nochesTotales / totalHabitaciones) * 10) / 10 : 0;

  // Edad
  const edades = filtrados.map(r => calcularEdad(idx('Fecha_Nacimiento') >= 0 ? r[idx('Fecha_Nacimiento')] : '')).filter(e => e !== null);
  const edadPromedio = edades.length > 0 ? Math.round(edades.reduce((a, b) => a + b, 0) / edades.length) : 0;

  // Agrupaciones
  const porMedio = {};
  filtrados.forEach(r => {
    const m = idx('Medio') >= 0 ? (r[idx('Medio')] || 'No especificado') : 'No especificado';
    porMedio[m] = (porMedio[m] || 0) + 1;
  });

  const porMes = {};
  const ingresosMes = {};
  filtrados.forEach(r => {
    const m = idx('Mes') >= 0 ? (r[idx('Mes')] || 'Desconocido') : 'Desconocido';
    porMes[m] = (porMes[m] || 0) + 1;
    if (monedaDe(r) === 'CLP') ingresosMes[m] = (ingresosMes[m] || 0) + totalDe(r);
  });

  const porMotivo = { 'Vacaciones': 0, 'Trabajo': 0, 'Celebración': 0, 'Canje': 0, 'Premio': 0, 'Agencia': 0, 'Otro': 0 };
  filtrados.forEach(r => {
    const m = idx('Motivo') >= 0 ? String(r[idx('Motivo')] || '').trim() : '';
    const mUpper = m.toUpperCase();
    if (mUpper === 'VACACIONES') porMotivo['Vacaciones']++;
    else if (mUpper === 'TRABAJO') porMotivo['Trabajo']++;
    else if (['CUMPLEAÑOS', 'ANIVERSARIO', 'CELEBRACIÓN', 'CELEBRACION'].includes(mUpper)) porMotivo['Celebración']++;
    else if (mUpper === 'CANJE') porMotivo['Canje']++;
    else if (mUpper === 'PREMIO') porMotivo['Premio']++;
    else if (mUpper === 'AGENCIA') porMotivo['Agencia']++;
    else porMotivo['Otro']++;
  });

  // Mapa clave-normalizada -> nombre canónico (según la hoja Programas), para
  // mostrar un nombre bonito y agrupar sin distinguir mayúsculas/acentos.
  const progCanon = {};
  try {
    const shProg = getProgramasSheet_(ss);
    const dProg = shProg.getDataRange().getValues();
    for (let i = 1; i < dProg.length; i++) {
      const nm = (dProg[i][0] || '').toString().trim();
      if (nm) progCanon[_normPrograma(nm)] = nm;
    }
  } catch (e) {}

  // Agrupación por programa SIN diferenciar mayúsculas/minúsculas ni acentos.
  // Así "Residentes" y "RESIDENTES" cuentan como un solo programa (sin clones).
  const porPrograma = {};
  filtrados.forEach(r => {
    const raw = idx('Programa') >= 0 ? String(r[idx('Programa')] || '').trim() : '';
    const norm = _normPrograma(raw);
    let label;
    if (!norm || norm === 'estandar') label = progCanon['estandar'] || 'Estándar';
    else label = progCanon[norm] || raw || 'Estándar';
    porPrograma[label] = (porPrograma[label] || 0) + 1;
  });

  // Nacionalidades (top)
  const porNacionalidad = {};
  filtrados.forEach(r => {
    let n = idx('Nacionalidad') >= 0 ? String(r[idx('Nacionalidad')] || '').trim() : '';
    if (!n) n = 'Sin dato';
    porNacionalidad[n] = (porNacionalidad[n] || 0) + 1;
  });
  const nacionalidadesDistintas = Object.keys(porNacionalidad).filter(k => k !== 'Sin dato').length;
  const topNacionalidades = Object.entries(porNacionalidad).sort((a, b) => b[1] - a[1]).slice(0, 7);

  // Día de la semana del check-in
  const dias = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  const porDiaSemana = { 'Lunes':0,'Martes':0,'Miércoles':0,'Jueves':0,'Viernes':0,'Sábado':0,'Domingo':0 };
  filtrados.forEach(r => {
    const cin = idx('Check_in') >= 0 ? String(r[idx('Check_in')] || '').trim() : '';
    const p = cin.split('-');
    if (p.length === 3) {
      const d = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
      if (!isNaN(d.getTime())) porDiaSemana[dias[d.getDay()]]++;
    }
  });

  const tabla = filtrados.map(r => ({
    nombre: idx('Nombre') >= 0 ? r[idx('Nombre')] : '',
    nacionalidad: idx('Nacionalidad') >= 0 ? r[idx('Nacionalidad')] : '',
    checkin: idx('Check_in') >= 0 ? r[idx('Check_in')] : '',
    checkout: idx('Check_out') >= 0 ? r[idx('Check_out')] : '',
    noches: idx('Noches') >= 0 ? r[idx('Noches')] : '',
    tarifa: idx('Tarifa_por_Noche') >= 0 ? r[idx('Tarifa_por_Noche')] : '',
    total: idx('Total_estadia') >= 0 ? r[idx('Total_estadia')] : '',
    moneda: monedaDe(r),
    motivo: idx('Motivo') >= 0 ? r[idx('Motivo')] : '',
    medio: idx('Medio') >= 0 ? r[idx('Medio')] : '',
    programa: idx('Programa') >= 0 ? String(r[idx('Programa')] || '').trim().toUpperCase() : 'ESTANDAR'
  }));

  // Orden de meses cronológico (Mes = "Enero 2026")
  const ordenMes = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const claveMes = m => {
    const p = String(m).split(' ');
    const anio = parseInt(p[1]) || 0;
    const mi = ordenMes.indexOf(p[0]);
    return anio * 100 + (mi < 0 ? 99 : mi);
  };
  const porMesOrd = Object.entries(porMes).sort((a, b) => claveMes(a[0]) - claveMes(b[0]));
  const ingresosMesOrd = Object.entries(ingresosMes).map(([k, v]) => [k, Math.round(v)]).sort((a, b) => claveMes(a[0]) - claveMes(b[0]));

  return {
    totalHabitaciones: totalHabitaciones,
    totalTarifa: totalTarifa,
    tarifaPromedio: tarifaPromedio,
    edadPromedio: edadPromedio,
    ingresosCLP: ingresosCLP,
    ingresosUSD: ingresosUSD,
    nochesTotales: nochesTotales,
    nochesPromedio: nochesPromedio,
    nacionalidadesDistintas: nacionalidadesDistintas,
    porMedio: Object.entries(porMedio),
    porMes: porMesOrd,
    ingresosMes: ingresosMesOrd,
    porMotivo: Object.entries(porMotivo).filter(([k, v]) => v > 0),
    porPrograma: Object.entries(porPrograma),
    porNacionalidad: topNacionalidades,
    porDiaSemana: Object.entries(porDiaSemana),
    tabla: tabla
  };
}
