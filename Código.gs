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
    .addItem('Actualizar Socios 👥', 'actualizarSocios')
    .addItem('Antigüedad de Deuda 📊', 'actualizarAntiguedadDeuda')
    .addSeparator()
    .addItem('Enviar Mails a Deudores ✉️', 'ejecutarEnvioMailsUI')
    .addToUi();
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
    SpreadsheetApp.getUi().alert('¡Éxito!: Hoja de "' + nombreHojaDestino + '" actualizada correctamente.');
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
      .setTitle('Consulta de Deudas - Club Online')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Obtiene las categorías de selección desde la hoja oficial "Divisiones"
 */
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

/**
 * REESTRUCTURADO Y CORREGIDO: Extrae los saldos reales desde la hoja 'Reporte'
 * mapeándolos mediante el N° de socio único obtenido entre paréntesis.
 */
/**
 * NUEVA VERSIÓN DIRECTA: Extrae y ordena la información basándose 
 * exclusivamente en las columnas de la hoja 'Reporte'.
 */
function obtenerDeudasPorDivision(divisionBuscada) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hojaReporte = ss.getSheetByName('Reporte');
  if (!hojaReporte) return [];
  
  const datosReporte = hojaReporte.getDataRange().getValues();
  let resultados = [];
  
  if (!divisionBuscada) return [];
  let buscarLimpio = divisionBuscada.toLowerCase().replace(/[^a-z0-9]/g, "");

  // Recorremos la hoja Reporte desde la fila 2 (índice 1) para omitir encabezados
  for (let k = 1; k < datosReporte.length; k++) {
    let filaR = datosReporte[k];
    if (!filaR || filaR[7] === undefined) continue; // Salta si la fila está vacía
    
    // Columna H: Divisiones
    let divisionFilaLimpia = filaR[7].toString().toLowerCase().replace(/[^a-z0-9]/g, "");
    
    // Si la fila pertenece a la división seleccionada
    if (divisionFilaLimpia !== "" && divisionFilaLimpia.includes(buscarLimpio)) {
      
      let nombre = filaR[19] ? filaR[19].toString().trim() : "Sin Nombre"; // Columna T
      let formaPago = filaR[6] ? filaR[6].toString().trim() : "-";          // Columna G
      let division = filaR[7] ? filaR[7].toString().trim() : "-";           // Columna H
      let descuento = filaR[8] ? filaR[8].toString().trim() : "";           // Columna I
      let periodoDesc = filaR[9] ? filaR[9].toString().trim() : "";         // Columna J
      let totalDeuda = parseFloat(filaR[20]) || 0;                          // Columna U
      
      // Control de Alerta: Verifica si tiene números distintos a 0 en M (Abr-26) y N (Mar-26)
      let mesM = parseFloat(filaR[12]) || 0; 
      let mesN = parseFloat(filaR[13]) || 0;
      let alertaCritica = (mesM > 0 && mesN > 0); 
      
      resultados.push({
        nombre: nombre,
        formaPago: formaPago,
        division: division,
        descuento: descuento,
        periodoDesc: periodoDesc,
        total: totalDeuda,
        alerta: alertaCritica
      });
    }
  }
  
  // ORDENAR: En función de quién debe más (Monto Total de mayor a menor)
  return resultados.sort((a, b) => b.total - a.total);
}
function ejecutarEnvioMailsUI() {
  const respuesta = enviarMailsDeudores();
  SpreadsheetApp.getUi().alert(respuesta);
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
    let deudaMesActual = parseFloat(fila[3]) || 0;
    
    if (idSocioDeuda && deudaMesActual > 0 && !idSocioDeuda.toLowerCase().includes("total")) {
      let emailDestino = directorioSocios[idSocioDeuda];
      if (emailDestino && emailDestino.includes("@")) {
        let asunto = "Aviso Importante: Regularización de Cuota - Club Online";
        let cuerpo = `Estimado/a ${nombreJugador},\n\nLe comunicamos que registra un saldo pendiente en su cuota social por un importe de $${deudaMesActual}.\n\nSolicitamos tenga a bien acercarse a la tesorería del club para regularizar su situación.\n\nAtentamente,\nDepto. de Tesorería - Club Online.`;
        MailApp.sendEmail(emailDestino, asunto, cuerpo);
        correosEnviados++;
      }
    }
  }
  return "Proceso completado. Se enviaron exitosamente " + correosEnviados + " correos electrónicos.";
}