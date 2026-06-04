// ID de la carpeta principal de Tesorería provisto por el usuario
const CARPETA_RAIZ_ID = '1pqMjUjZ-K4Bo3lC-kSYDaGjAUxQSaRrN';
// ID del archivo de Google Sheets de destino
const SPREADSHEET_DESTINO_ID = '1OQ-BGFqYxEqy1UR6YkREXnVqwaNiFmLVEWW_Q_SB50k';
// NUEVO: ID de la carpeta específica de Matrículas provisto por el usuario
const CARPETA_MATRICULAS_ID = '1R-s385voSUlgo33xa8pqpuc_KEUEKdZS';

/**
 * 1. Crea el menú "Club online" al abrir la hoja de cálculo.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Club online')
    .addItem('Actualizar Todo 🔄', 'actualizarTodo')
    .addSeparator()
    .addItem('Enviar Mails a Deudores ✉️', 'enviarMailsDeudores')
    .addToUi();
}

/**
 * Función unificada que actualiza Socios, Deudas y Matrículas consecutivamente
 */
function actualizarTodo() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert('Proceso iniciado', 'Actualizando base de datos completa... Por favor, espera.', ui.ButtonSet.OK);
    
    // 1. Ejecuta actualización de socios
    actualizarSocios();
    
    // 2. Ejecuta antigüedad de deuda
    actualizarAntiguedadDeuda();
    
    // 3. Ejecuta la importación y conversión de Matrículas
    actualizarMatriculas();

   // 4. Ejecuta la importación de Liga Sur sin alterar tus notas manuales en las filas viejas
    actualizarSociosLigaSur();
    
    ui.alert('Proceso completado', 'Se han actualizado los Socios, Deuda, Matrículas y los nuevos de Liga Sur correctamente. ✅', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Hubo un problema durante la actualización: ' + error.toString(), ui.ButtonSet.OK);
  }
}

function actualizarSocios() {
  const nombreCarpeta = "Socios";
  const nombreHojaDestino = "Socios";
  const columnasAMantener = [1, 2, 3, 4, 9, 11, 16, 22, 28, 29];
  const filaInicioOriginal = 6;
  procesarUltimoArchivo(CARPETA_RAIZ_ID, nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener, false);
}

// 1. NUEVA FUNCIÓN DISPARADORA (Se llama desde actualizarTodo)
// 1. FUNCIÓN DISPARADORA (Asegúrate de llamarla en actualizarTodo)
function actualizarSociosLigaSur() {
  const nombreCarpeta = "Socios";
  const nombreHojaDestino = "Liga Sur Fichados";
  
  // Incluye el número 5 que corresponde a la columna E (Nº Documento / DNI)
  const columnasAMantener = [1, 2, 3, 4, 5, 9, 11, 16, 22, 28, 29];
  const filaInicioOriginal = 6;
  const carpetaSociosId = "1aHuI-BqErWxN_ShsStmwYOMtOGYbFFrl"; 
  
  procesarUltimoArchivoSur(carpetaSociosId, nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener);
}

// 2. FUNCIÓN DE PROCESAMIENTO ADAPTADA A DRIVE API V3
function procesarUltimoArchivoSur(carpetaId, nombreCarpeta, nombreHojaDestino, filaInicio, columnasAMantener) {
  const carpeta = DriveApp.getFolderById(carpetaId);
  const archivos = carpeta.getFiles();
  
  let ultimoArchivo = null;
  let ultimaFecha = new Date(0);
  
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    if (archivo.getDateCreated() > ultimaFecha) {
      ultimaFecha = archivo.getDateCreated();
      ultimoArchivo = archivo;
    }
  }
  
  if (!ultimoArchivo) {
    Logger.log("No se encontraron archivos en la carpeta Socios.");
    return;
  }
  
  let ssOrigen;
  let archivoTemporalId = null;
  const mimeType = ultimoArchivo.getMimeType();

  if (mimeType === MimeType.MICROSOFT_EXCEL || mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || ultimoArchivo.getName().endsWith('.xls')) {
    const recursoArchivo = {
      name: "Temporal_Sur_" + ultimoArchivo.getName().replace(/\.[^/.]+$/, ""),
      mimeType: "application/vnd.google-apps.spreadsheet",
      parents: [carpetaId]
    };
    
    const archivoConvertido = Drive.Files.create(recursoArchivo, ultimoArchivo.getBlob());
    archivoTemporalId = archivoConvertido.id;
    ssOrigen = SpreadsheetApp.openById(archivoTemporalId);
  } else {
    ssOrigen = SpreadsheetApp.openById(ultimoArchivo.getId());
  }
  
  const hojaOrigen = ssOrigen.getSheets()[0];
  const datosOrigen = hojaOrigen.getDataRange().getValues();
  
  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  let hojaDestino = ssDestino.getSheetByName(nombreHojaDestino);
  if (!hojaDestino) {
    hojaDestino = ssDestino.insertSheet(nombreHojaDestino);
  }
  
  // --- DETECTAR SI LA HOJA DESTINO ESTÁ VACÍA PARA CREAR ENCABEZADOS ---
  let ultimaFilaDestino = hojaDestino.getLastRow();
  const esHojaVacia = (ultimaFilaDestino === 0 || (ultimaFilaDestino === 1 && hojaDestino.getRange(1,1).getValue() === ""));

  if (esHojaVacia) {
    // La fila 5 del origen contiene los títulos de columnas (índice 4 en la matriz)
    const filaEncabezadosOrigen = datosOrigen[4]; 
    const nuevosEncabezados = [];
    
    columnasAMantener.forEach(col => {
      nuevosEncabezados.push(filaEncabezadosOrigen[col - 1]);
    });
    
    // Escribe la fila de títulos en la primera línea
    hojaDestino.getRange(1, 1, 1, nuevosEncabezados.length).setValues([nuevosEncabezados]);
    ultimaFilaDestino = 1; // Actualizamos para que los jugadores empiecen en la fila 2
  }

  // --- CONTROL DE DUPLICADOS POR DNI ---
  const datosDestinoExistentes = hojaDestino.getDataRange().getValues();
  const dnisExistentes = [];
  const indiceDniEnDestino = columnasAMantener.indexOf(5); 

  if (datosDestinoExistentes.length > 0 && datosDestinoExistentes[0][0] !== "") {
    datosDestinoExistentes.forEach(filaDestino => {
      if (filaDestino[indiceDniEnDestino]) {
        dnisExistentes.push(filaDestino[indiceDniEnDestino].toString().trim());
      }
    });
  }

  const nuevosJugadores = [];
  
  for (let i = filaInicio - 1; i < datosOrigen.length; i++) {
    const fila = datosOrigen[i];
    
    const valorColumnaV = fila[21] ? fila[21].toString() : ""; 
    const dniOrigen = fila[4] ? fila[4].toString().trim() : "";  
    
    if (valorColumnaV.includes("Sur") && dniOrigen !== "" && !dnisExistentes.includes(dniOrigen)) {
      const nuevaFila = [];
      columnasAMantener.forEach(col => {
        nuevaFila.push(fila[col - 1]);
      });
      nuevosJugadores.push(nuevaFila);
    }
  }
  
  // --- PEGAR LOS NUEVOS JUGADORES JUSTO DEBAJO ---
  if (nuevosJugadores.length > 0) {
    let filaInsercion = ultimaFilaDestino + 1;
    hojaDestino.getRange(filaInsercion, 1, nuevosJugadores.length, nuevosJugadores[0].length).setValues(nuevosJugadores);
    Logger.log("Se insertaron " + nuevosJugadores.length + " nuevos jugadores en Liga Sur Fichados.");
  } else {
    Logger.log("No se encontraron jugadores nuevos de la Liga Sur para anexar.");
  }

  if (archivoTemporalId) {
    try {
      Drive.Files.deleteResource(archivoTemporalId);
    } catch (e) {
      Logger.log("Aviso: No se pudo limpiar el archivo temporal automáticamente: " + e.toString());
    }
  }
}


function actualizarAntiguedadDeuda() {
  const nombreCarpeta = "AntiguedaddeDeuda";
  const nombreHojaDestino = "AntiguedaddeDeuda";
  const columnasAMantener = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const filaInicioOriginal = 5;
  procesarUltimoArchivo(CARPETA_RAIZ_ID, nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener, false);
}


function actualizarMatriculas() {
  const nombreHojaDestino = "Matriculas";
  // Mantiene la misma estructura de columnas y fila de inicio que AntiguedaddeDeuda
  const columnasAMantener = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const filaInicioOriginal = 5;
  
  // Llamamos al procesador indicándole la carpeta específica y activando el multiplicador por -1
  procesarUltimoArchivo(CARPETA_MATRICULAS_ID, null, nombreHojaDestino, filaInicioOriginal, columnasAMantener, true);
}

/**
 * Función auxiliar modificada encargada de buscar, filtrar, transformar y pegar en destino.
 */
function procesarUltimoArchivo(idCarpetaOrigen, nombreSubcarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener, multiplicarPorMenosUno) {
  let carpetaDestino;
  
  if (nombreSubcarpeta) {
    const carpetaRaiz = DriveApp.getFolderById(idCarpetaOrigen);
    const subcarpetas = carpetaRaiz.getFoldersByName(nombreSubcarpeta);
    if (!subcarpetas.hasNext()) {
      SpreadsheetApp.getUi().alert('Error: No se encontró la carpeta llamada "' + nombreSubcarpeta + '"');
      return;
    }
    carpetaDestino = subcarpetas.next();
  } else {
    // Si no se define subcarpeta, lee directamente el ID provisto (caso Matrículas)
    carpetaDestino = DriveApp.getFolderById(idCarpetaOrigen);
  }
  
  const archivos = carpetaDestino.getFiles();
  let ultimoArchivo = null;
  let fechaUltimoArchivo = 0;
  
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    if (archivo.getName().toLowerCase().endsWith('.xls') || archivo.getName().toLowerCase().endsWith('.xlsx')) {
      const fechaModificacion = archivo.getLastUpdated().getTime();
      if (fechaModificacion > fechaUltimoArchivo) {
        fechaUltimoArchivo = fechaModificacion;
        ultimoArchivo = archivo;
      }
    }
  }
  
  if (!ultimoArchivo) {
    SpreadsheetApp.getUi().alert('No se encontraron archivos de Excel en la carpeta configurada.');
    return;
  }
  
  let tempSpreadsheet = null;
  try {
    const blob = ultimoArchivo.getBlob();
    const config = { title: ultimoArchivo.getName() + '_temp', mimeType: MimeType.GOOGLE_SHEETS };
    const tempFile = Drive.Files.create(config, blob);
    tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
    const hojaOrigen = tempSpreadsheet.getSheets()[0];
    const datosOrigen = hojaOrigen.getDataRange().getValues();
    
    if (datosOrigen.length < filaInicioOriginal) {
      SpreadsheetApp.getUi().alert('El archivo no contiene suficientes filas.');
      eliminarArchivoTemporal(tempFile.id);
      return;
    }
    
    const matrizFiltrada = [];
    for (let i = filaInicioOriginal - 1; i < datosOrigen.length; i++) {
      const filaOriginal = datosOrigen[i];
      const nuevaFila = [];
      
      columnasAMantener.forEach(idx => {
        let valor = filaOriginal[idx - 1];
        
        // REQUERIMIENTO ESPECIAL: Si es Matrícula, validamos y multiplicamos por -1 los números != 0
        if (multiplicarPorMenosUno && typeof valor === 'number' && valor !== 0) {
          valor = valor * -1;
        }
        
        nuevaFila.push(valor !== undefined ? valor : "");
      });
      matrizFiltrada.push(nuevaFila);
    }
    
    const ssDestino = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
    let hojaDestino = ssDestino.getSheetByName(nombreHojaDestino);
    if (!hojaDestino) hojaDestino = ssDestino.insertSheet(nombreHojaDestino);
    
    hojaDestino.clearContents();
    hojaDestino.clearFormats();
    
    if (matrizFiltrada.length > 0) {
      hojaDestino.getRange(1, 1, matrizFiltrada.length, matrizFiltrada[0].length).setValues(matrizFiltrada);
    }
    eliminarArchivoTemporal(tempFile.id);
    ssDestino.toast('Hoja de "' + nombreHojaDestino + '" actualizada.', '¡Éxito!');
  } catch (error) {
    if (tempSpreadsheet) eliminarArchivoTemporal(tempSpreadsheet.getId());
    SpreadsheetApp.getUi().alert('Ocurrió un error al procesar el archivo: ' + error.message);
  }
}

function eliminarArchivoTemporal(id) {
  try { Drive.Files.remove(id); } catch(e) {
    try { DriveApp.getFileById(id).setTrashed(true); } catch(err) {}
  }
}

// ==========================================
// CONTROLADOR DE LA WEB APP MÓVIL
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Consulta de Deudas - Club SFP GONNET')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // <-- AQUÍ QUEDÓ GRABADO FIJO EL PERMISO
}

function obtenerDivisiones() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hoja = ss.getSheetByName('Divisiones');
  if (!hoja) return [];
  
  const datos = hoja.getDataRange().getValues();
  let divisiones = [];
  for (let i = 1; i < datos.length; i++) {
    let div = datos[i][0];
    if (div && divisiones.indexOf(div.trim()) === -1) {
      divisiones.push(div.trim());
    }
  }
  return divisiones;
}

function obtenerDeudasPorDivision(divisionBuscada) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hojaReporte = ss.getSheetByName('Reporte');
  if (!hojaReporte) return { nombresMeses: [], listaJugadores: [] };
  
  const datosReporte = hojaReporte.getDataRange().getValues();
  let nombresMeses = [];
  for (let col = 12; col <= 18; col++) { 
    let nombreHeader = datosReporte[0][col] ? datosReporte[0][col].toString().trim() : "Mes";
    nombresMeses.push(nombreHeader);
  }

  let listaJugadores = [];
  if (!divisionBuscada) return { nombresMeses: nombresMeses, listaJugadores: [] };
  let buscarLimpio = divisionBuscada.toLowerCase().replace(/[^a-z0-9]/g, "");

  for (let k = 1; k < datosReporte.length; k++) {
    let filaR = datosReporte[k];
    if (!filaR || filaR[7] === undefined) continue; 
    
    let divisionFilaLimpia = filaR[7].toString().toLowerCase().replace(/[^a-z0-9]/g, "");
    if (divisionFilaLimpia !== "" && divisionFilaLimpia.includes(buscarLimpio)) {
      
      // Capturamos el Total de la Deuda de la columna U (Índice 20)
      let totalDeuda = parseFloat(filaR[20]) || 0; 
      
      // NUEVO: Capturamos el valor absoluto de la Matrícula de la columna X (Índice 23)
      let matriculaValor = Math.abs(parseFloat(filaR[23])) || 0;
      
      // Filtro de visualización: si no debe nada ni tiene matrícula, pasamos de largo
      if (totalDeuda <= 0 && matriculaValor <= 0) continue; 
      
      let nombre = filaR[19] ? filaR[19].toString().trim() : "Sin Nombre";
      
      let formaPagoRaw = filaR[6] ? filaR[6].toString().trim() : "-";
      let formaPago = formaPagoRaw;
      if (formaPagoRaw.includes("Entidad de Recaudación | Pagos Virtuales del Sur")) {
        formaPago = "Debito Automatico";
      } else if (formaPagoRaw.includes("Administración/Secretaría")) {
        formaPago = "Tesoreria";
      }
      
      let descuento = filaR[8] ? filaR[8].toString().trim() : "";
      let periodoDesc = filaR[9] ? filaR[9].toString().trim() : "";
      
      let valoresMeses = [];
      let mesesConDeudaActiva = 0;
      
      for (let col = 12; col <= 18; col++) {
        let valorCelda = parseFloat(filaR[col]) || 0;
        valoresMeses.push(valorCelda);
        if (valorCelda > 0) {
          mesesConDeudaActiva++;
        }
      }
      
      let alertaCritica = (mesesConDeudaActiva > 2);
      
      listaJugadores.push({
        nombre: nombre,
        formaPago: formaPago,
        descuento: descuento,
        periodoDesc: periodoDesc,
        total: totalDeuda,
        valoresMeses: valoresMeses, 
        alerta: alertaCritica,
        matricula: matriculaValor // <-- ENVIAMOS EL VALOR A LA INTERFAZ
      });
    }
  }
  
  listaJugadores.sort((a, b) => b.total - a.total);
  return {
    nombresMeses: nombresMeses,
    listaJugadores: listaJugadores
  };
}
function enviarMailsDeudores() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hojaDeuda = ss.getSheetByName('AntiguedaddeDeuda');
  const hojaSocios = ss.getSheetByName('Socios');
  if (!hojaDeuda || !hojaSocios) return "Error: No se encuentran las hojas necesarias.";
  
  const datosDeuda = hojaDeuda.getDataRange().getValues();
  const datosSocios = hojaSocios.getDataRange().getValues();
  
  let directorioSocios = {};
  for (let j = 1; j < datosSocios.length; j++) {
    let idSocio = datosSocios[j][0] ? datosSocios[j][0].toString().trim() : "";
    let email = datosSocios[j][5];
    if (idSocio && email) {
      directorioSocios[idSocio] = email.trim();
    }
  }
  
  let correosEnviados = 0;
  for (let i = 1; i < datosDeuda.length; i++) {
    let fila = datosDeuda[i];
    let idSocioDeuda = fila[0] ? fila[0].toString().trim() : "";
    let nombreJugador = fila[1];
    if (!idSocioDeuda || idSocioDeuda.toLowerCase().includes("total")) continue;
    
    let deudaMesActual = parseFloat(fila[3]) || 0; 
    let deudaAntiguaAcumulada = 0;
    for (let col = 4; col <= 9; col++) {
      deudaAntiguaAcumulada += parseFloat(fila[col]) || 0;
    }
    
    if (deudaAntiguaAcumulada > 0) {
      let emailDestino = directorioSocios[idSocioDeuda];
      if (emailDestino && emailDestino.includes("@")) {
        let totalACobrar = deudaMesActual + deudaAntiguaAcumulada;
        let asunto = "Aviso Importante: Regularización de Cuota - Club SFP Gonnet";
        let cuerpo = `Estimado/a ${nombreJugador},\n\nLe comunicamos que registra saldo pendiente en su cuota social por más de 2 meses, acumulando un Total a Cobrar de $${totalACobrar}.\n\nSolicitamos tenga a bien acercarse a la tesoreria en 495bis y 15 bis de lunes a viernes de 18 a 20hs, escribir a tesoreriagonnet@hotmail.com o por whatsapp al 2216819698 para regularizar su situación.\n\nAtentamente,\nTesorería - Club SFP Gonnet.`;
        GmailApp.sendEmail(emailDestino, asunto, cuerpo, {
          from: "tesoreriagonnet@gmail.com"
        });
        correosEnviados++;
      }
    }
  }
  return "Proceso completado. Se enviaron exitosamente " + correosEnviados + " correos electrónicos.";
}

function ejecutarEnvioMailsUIDetalles() {
  try {
    return enviarMailsDeudores();
  } catch(e) {
    return "Error al enviar los correos: " + e.toString();
  }
}