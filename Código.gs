// ID de la carpeta principal de Tesorería provisto por el usuario
const CARPETA_RAIZ_ID = '1pqMjUjZ-K4Bo3lC-kSYDaGjAUxQSaRrN';
// ID del archivo de Google Sheets de destino
const SPREADSHEET_DESTINO_ID = '1OQ-BGFqYxEqy1UR6YkREXnVqwaNiFmLVEWW_Q_SB50k';

/**
 * 1. Crea el menú "Club online" al abrir la hoja de cálculo.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Club online')
    .addItem('Actualizar Socios y Deudas 🔄', 'actualizarTodo')
    .addSeparator()
    .addItem('Enviar Mails a Deudores ✉️', 'enviarMailsDeudores')
    .addToUi();
}

/**
 * Nueva función unificada que actualiza todo junto consecutivamente
 */
function actualizarTodo() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert('Proceso iniciado', 'Actualizando datos de socios... Por favor, espera.', ui.ButtonSet.OK);
    actualizarSocios();
    actualizarAntiguedadDeuda();
    ui.alert('Proceso completado', 'Se han actualizado los socios y la antigüedad de la deuda correctamente. ✅', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Hubo un problema durante la actualización: ' + error.toString(), ui.ButtonSet.OK);
  }
}

function actualizarSocios() {
  const nombreCarpeta = "Socios";
  const nombreHojaDestino = "Socios";
  const columnasAMantener = [1, 2, 3, 4, 9, 11, 16, 22, 28, 29];
  const filaInicioOriginal = 6;
  procesarUltimoArchivo(nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener);
}

function actualizarAntiguedadDeuda() {
  const nombreCarpeta = "AntiguedaddeDeuda";
  const nombreHojaDestino = "AntiguedaddeDeuda";
  const columnasAMantener = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const filaInicioOriginal = 5;
  procesarUltimoArchivo(nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener);
}

/**
 * Función auxiliar encargada de buscar el último archivo Excel, filtrar y pegar en destino.
 */
function procesarUltimoArchivo(nombreSubcarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener) {
  const carpetaRaiz = DriveApp.getFolderById(CARPETA_RAIZ_ID);
  const subcarpetas = carpetaRaiz.getFoldersByName(nombreSubcarpeta);
  
  if (!subcarpetas.hasNext()) {
    SpreadsheetApp.getUi().alert('Error: No se encontró la carpeta llamada "' + nombreSubcarpeta + '"');
    return;
  }
  
  const carpetaDestino = subcarpetas.next();
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
    SpreadsheetApp.getUi().alert('No se encontraron archivos de Excel en la carpeta "' + nombreSubcarpeta + '"');
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
        nuevaFila.push(filaOriginal[idx - 1] !== undefined ? filaOriginal[idx - 1] : "");
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
      
      let totalDeuda = parseFloat(filaR[20]) || 0; 
      if (totalDeuda <= 0) continue; 
      
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
        alerta: alertaCritica
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