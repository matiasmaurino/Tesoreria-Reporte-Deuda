function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Consulta de Deudas - Club SFP GONNET')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // <-- ESTA LÍNEA ES LA QUE CORRIGE EL ERROR
}
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
    .addItem('Enviar Mails a Deudores ✉️', 'enviarMailsDeudores') // <-- Modificado aquí
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
      .setTitle('Consulta de Deudas - Club SFP GONNET')
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
/**
 * EXTRAE Y MAPEA: Trae la información financiera basándose en la hoja 'Reporte'
 * e incluye dinámicamente los encabezados de los meses para evitar el error NaN.
 */
/**
 * MODIFICADO: Extrae solo jugadores con deuda > 0, mapea los 7 meses de M a S,
 * aplica el cálculo de morosidad flexible y simplifica las formas de pago.
 */
function obtenerDeudasPorDivision(divisionBuscada) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hojaReporte = ss.getSheetByName('Reporte');
  if (!hojaReporte) return { nombresMeses: [], listaJugadores: [] };
  
  const datosReporte = hojaReporte.getDataRange().getValues();
  
  // Extraer las etiquetas dinámicas de los 7 campos (Columnas M a S)
  let nombresMeses = [];
  for (let col = 12; col <= 18; col++) { // Índices 12 (M) al 18 (S)
    let nombreHeader = datosReporte[0][col] ? datosReporte[0][col].toString().trim() : "Mes";
    nombresMeses.push(nombreHeader);
  }

  let listaJugadores = [];
  if (!divisionBuscada) return { nombresMeses: nombresMeses, listaJugadores: [] };
  
  let buscarLimpio = divisionBuscada.toLowerCase().replace(/[^a-z0-9]/g, "");

  // Recorremos la hoja desde la fila 2 para extraer los registros
  for (let k = 1; k < datosReporte.length; k++) {
    let filaR = datosReporte[k];
    if (!filaR || filaR[7] === undefined) continue; // Columna H (Divisiones)
    
    let divisionFilaLimpia = filaR[7].toString().toLowerCase().replace(/[^a-z0-9]/g, "");
    
    if (divisionFilaLimpia !== "" && divisionFilaLimpia.includes(buscarLimpio)) {
      
      // REQUERIMIENTO 1: Solo procesar y mostrar jugadores con deuda total mayor a 0 (Columna U)
      let totalDeuda = parseFloat(filaR[20]) || 0; 
      if (totalDeuda <= 0) continue; 
      
      let nombre = filaR[19] ? filaR[19].toString().trim() : "Sin Nombre"; // Columna T
      
      // REQUERIMIENTO 2: Reemplazo y simplificación de las Formas de Pago (Columna G)
      let formaPagoRaw = filaR[6] ? filaR[6].toString().trim() : "-";
      let formaPago = formaPagoRaw;
      if (formaPagoRaw.includes("Entidad de Recaudación | Pagos Virtuales del Sur")) {
        formaPago = "Debito Automatico";
      } else if (formaPagoRaw.includes("Administración/Secretaría")) {
        formaPago = "Tesoreria";
      }
      
      let descuento = filaR[8] ? filaR[8].toString().trim() : "";     // Columna I
      let periodoDesc = filaR[9] ? filaR[9].toString().trim() : "";   // Columna J
      
      // REQUERIMIENTO 5: Capturar los valores de los 7 campos (Columnas M a S)
      let valoresMeses = [];
      let mesesConDeudaActiva = 0;
      
      for (let col = 12; col <= 18; col++) {
        let valorCelda = parseFloat(filaR[col]) || 0;
        valoresMeses.push(valorCelda);
        
        // Si el jugador registra deuda activa en este mes, lo contabilizamos
        if (valorCelda > 0) {
          mesesConDeudaActiva++;
        }
      }
      
      // REQUERIMIENTO 4: Alerta si el jugador debe más de 2 meses en total (sin importar cuáles)
      let alertaCritica = (mesesConDeudaActiva > 2);
      
      listaJugadores.push({
        nombre: nombre,
        formaPago: formaPago,
        descuento: descuento,
        periodoDesc: periodoDesc,
        total: totalDeuda,
        valoresMeses: valoresMeses, // Enviamos el array con los 7 montos
        alerta: alertaCritica
      });
    }
  }
  
  // Ordenar de mayor a menor monto de deuda acumulada
  listaJugadores.sort((a, b) => b.total - a.total);
  
  return {
    nombresMeses: nombresMeses,
    listaJugadores: listaJugadores
  };
}
function enviarMailsDeudores() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hojaSocios = ss.getSheetByName('Socios');
  const hojaReporte = ss.getSheetByName('Reporte'); // <-- Ahora procesamos todo desde acá
  const hojaHistorial = ss.getSheetByName('Historial de email');
  
  if (!hojaSocios || !hojaReporte) {
    return "Error: No se encuentran las hojas necesarias (Socios o Reporte).";
  }
  
  // Encabezados del historial si la hoja es nueva
  if (hojaHistorial && hojaHistorial.getLastRow() === 0) {
    hojaHistorial.appendRow(["Fecha y Hora", "ID Socio", "Nombre Jugador", "Email Destino", "Monto Reclamado", "Estado del Envío"]);
  }
  
  const datosSocios = hojaSocios.getDataRange().getValues();
  const datosReporte = hojaReporte.getDataRange().getValues();
  
  // 1. Diccionario de e-mails de la hoja 'Socios' (ID en col A, Email en col F)
  let directorioSocios = {};
  for (let j = 1; j < datosSocios.length; j++) {
    let idSocio = datosSocios[j][0] ? datosSocios[j][0].toString().trim() : "";
    let email = datosSocios[j][5];
    if (idSocio && email) {
      directorioSocios[idSocio] = email.trim();
    }
  }
  
  let correosEnviados = 0;
  let erroresAliasContador = 0;
  const fechaActual = new Date();
  
  // 2. Recorremos la hoja 'Reporte' basándonos en tu columna W
  // Columna A = Código (índice 0)
  // Columna T = Nombre completo (índice 19)
  // Columna U = INFORMAR TOTAL / Monto (índice 20)
  // Columna W = ENVIAR EMAIL (índice 22)
  for (let r = 1; r < datosReporte.length; r++) {
    let fila = datosReporte[r];
    let idSocioReporte = fila[0] ? fila[0].toString().trim() : "";
    let nombreJugador = fila[19] ? fila[19].toString().trim() : "";
    let montoTotal = parseFloat(fila[20]) || 0;
    let condicionEnviar = fila[22] ? fila[22].toString().trim().toUpperCase() : ""; // Columna W
    
    if (!idSocioReporte || idSocioReporte.toLowerCase().includes("total")) continue;
    
    // CRITERIO ULTRA SIMPLE: Si tu columna W dice "ENVIAR EMAIL", se procesa.
    if (condicionEnviar === "ENVIAR EMAIL") {
      let emailDestino = directorioSocios[idSocioReporte];
      
      if (emailDestino && emailDestino.includes("@")) {
        let asunto = "Aviso Importante: Regularización de Cuota - Club SFP Gonnet";
        let cuerpo = `Estimado/a ${nombreJugador},\n\nLe comunicamos que registra un saldo pendiente en su cuota social por más de 2 meses, acumulando un Total a Cobrar de $${montoTotal}.\n\nSolicitamos tenga a bien acercarse a la tesoreria en 495bis y 15 bis de lunes a viernes de 18 a 20hs, responder este correo o enviar uno nuevo a tesoreriagonnet@gmail.com o por ultimo escribir por whatsapp al 2216819698 para regularizar su situación.\n\nAtentamente,\nTesorería - Club SFP Gonnet.`;
        
        // SISTEMA DE ENVÍO SEGURO CON HISTORIAL
        try {
          // Intento 1: Casilla de tesorería
          GmailApp.sendEmail(emailDestino, asunto, cuerpo, {
            from: "tesoreriagonnet@gmail.com"
          });
          correosEnviados++;
          
          if (hojaHistorial) {
            hojaHistorial.appendRow([fechaActual, idSocioReporte, nombreJugador, emailDestino, montoTotal, "Enviado desde tesoreriagonnet@gmail.com"]);
          }
          
        } catch (errorAlias) {
          try {
            // Intento 2: Cuenta principal de respaldo
            GmailApp.sendEmail(emailDestino, asunto, cuerpo);
            correosEnviados++;
            erroresAliasContador++;
            
            if (hojaHistorial) {
              hojaHistorial.appendRow([fechaActual, idSocioReporte, nombreJugador, emailDestino, montoTotal, "Enviado desde cuenta principal (Fallo Alias)"]);
            }
          } catch (errorGrave) {
            console.error("Fallo de envío para ID " + idSocioReporte + ": " + errorGrave.toString());
            if (hojaHistorial) {
              hojaHistorial.appendRow([fechaActual, idSocioReporte, nombreJugador, emailDestino, montoTotal, "ERROR: " + errorGrave.toString()]);
            }
          }
        }
      }
    }
  }
  
  let mensajeFinal = "Proceso completado. Se enviaron exitosamente " + correosEnviados + " correos electrónicos según el filtro de la columna W.";
  if (erroresAliasContador > 0) {
    mensajeFinal += "\n\n⚠️ Nota: " + erroresAliasContador + " correos se enviaron desde tu cuenta principal. El alias requirió revisión.";
  }
  return mensajeFinal;
}
function ejecutarEnvioMailsUI() {
  try {
    // Llama al proceso de envío de correos
    const resultado = enviarMailsDeudores();
    
    // Muestra una ventana flotante con el resultado en la planilla
    SpreadsheetApp.getUi().alert(resultado);
  } catch(e) {
    SpreadsheetApp.getUi().alert("Error en el envío: " + e.toString());
  }
}