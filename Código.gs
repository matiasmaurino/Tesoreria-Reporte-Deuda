// ID de la carpeta principal de Tesorería provisto por el usuario
const CARPETA_RAIZ_ID = '1pqMjUjZ-K4Bo3lC-kSYDaGjAUxQSaRrN';
// ID del archivo de Google Sheets de destino
const SPREADSHEET_DESTINO_ID = '1OQ-BGFqYxEqy1UR6YkREXnVqwaNiFmLVEWW_Q_SB50k';
// NUEVO: ID de la carpeta específica de Matrículas provisto por el usuario
const CARPETA_MATRICULAS_ID = '1R-s385voSUlgo33xa8pqpuc_KEUEKdZS';
// ======= AGREGA ESTA LÍNEA AQUÍ =======
const CARPETA_FICHAJES_ID = '1YrzDCmhVjcE3_qv_uUjeNq_UqFaob_DW';

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
/**
 * Función unificada que actualiza Socios, Deudas y Matrículas consecutivamente
 * Modificada bajo la Opción B: Sin alertas de interfaz de usuario.
 */
function actualizarTodo() {
  try {
    // 1. Ejecuta actualización de socios
    actualizarSocios();
    
    // 2. Ejecuta antigüedad de deuda
    actualizarAntiguedadDeuda();

    // 3. Ejecuta la importación y conversión de Matrículas
    actualizarMatriculas();

    // 4. Ejecuta la importación de Fichajes
    actualizarFichajes();
    
    // Dejamos un registro interno en la consola en lugar de un cartel emergente
    Logger.log('Proceso completado correctamente de manera silenciosa.');
    
  } catch (error) {
    // En caso de error de ejecución, se guardará en los registros de Google Apps Script
    Logger.log('Hubo un problema durante la actualización: ' + error.toString());
  }
}
function actualizarSocios() {
  const nombreCarpeta = "Socios";
  const nombreHojaDestino = "Socios";
  const columnasAMantener = [1, 2, 3, 5, 9, 11, 16, 22, 28, 29];
  const filaInicioOriginal = 6;
  procesarUltimoArchivo(CARPETA_RAIZ_ID, nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener, false);
}

function actualizarAntiguedadDeuda() {
  const nombreCarpeta = "AntiguedaddeDeuda";
  const nombreHojaDestino = "AntiguedaddeDeuda";
  const columnasAMantener = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const filaInicioOriginal = 5;
  procesarUltimoArchivo(CARPETA_RAIZ_ID, nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener, false);
}

/**
 * NUEVA FUNCIÓN: Procesa la carpeta de Matrículas invirtiendo el signo de los números
 */
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
    nombresMeses.push(datosReporte[0][col] ? datosReporte[0][col].toString().trim() : "Mes");
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
      let matriculaValor = Math.abs(parseFloat(filaR[23])) || 0; // Columna X
      
      if (totalDeuda <= 0 && matriculaValor <= 0) continue; 
      
      let nombre = filaR[19] ? filaR[19].toString().trim() : "Sin Nombre";
      let formaPagoRaw = filaR[6] ? filaR[6].toString().trim() : "-";
      let formaPago = formaPagoRaw.includes("Entidad de Recaudación") ? "Debito Automatico" : "Tesoreria";
      
      let descuento = filaR[8] ? filaR[8].toString().trim() : "";
      let periodoDesc = filaR[9] ? filaR[9].toString().trim() : "";
      
      let valoresMeses = [];
      for (let col = 12; col <= 18; col++) {
        valoresMeses.push(parseFloat(filaR[col]) || 0);
      }
      
      // NUEVA LÓGICA DIRECTA DESDE LA PLANILLA:
      // Leemos la columna U ("CUANTOS MESES DEBE"), que es el índice 21 en la fila.
      let mesesDebidosPlanilla = parseInt(filaR[21]) || 0;
      
      // Si el número en la planilla es 2 o más, se activa la alerta crítica automáticamente
      let alertaCritica = (mesesDebidosPlanilla >= 2);
      
      listaJugadores.push({
        nombre: nombre,
        formaPago: formaPago,
        descuento: descuento,
        periodoDesc: periodoDesc,
        total: totalDeuda,
        valoresMeses: valoresMeses, 
        alerta: alertaCritica,
        matricula: matriculaValor
      });
    }
  }
  
  listaJugadores.sort((a, b) => b.total - a.total);
  return { nombresMeses: nombresMeses, listaJugadores: listaJugadores };
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

/**
 * NUEVA FUNCIÓN: Procesa los archivos .xls/.xlsx de la carpeta Fichajes
 * Copia únicamente las columnas A, B, C, F y G de forma contigua (Columnas A a E)
 * en la hoja 'Fichajes' desde la fila 7 hacia abajo, evitando duplicados por DNI.
 */
/**
 * FUNCIÓN DEFINITIVA: Procesa los archivos de la carpeta Fichajes.
 * - Escribe los datos de corrido desde la celda A1 (Fila 1: Encabezados).
 * - Extrae el Socio dentro de los paréntesis dejando ÚNICAMENTE los números enteros (elimina letras y guiones).
 * - Agrega de forma contigua las columnas hasta la G del Excel de origen.
 */
function actualizarFichajes() {
  const nombreHojaDestino = "Fichajes";
  const ssDestino = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  let hojaDestino = ssDestino.getSheetByName(nombreHojaDestino);
  if (!hojaDestino) {
    hojaDestino = ssDestino.insertSheet(nombreHojaDestino);
  }

  // 1. Evitar duplicados usando el Número de Comprobante / Nº (Columna B de destino)
  const comprobantesExistentes = new Set();
  const ultimaFilaDestino = hojaDestino.getLastRow();
  
  if (ultimaFilaDestino >= 2) {
    // Leemos la columna B (Nº de Comprobante) desde la fila 2 hasta el final actual
    const valoresDestinoB = hojaDestino.getRange(2, 2, ultimaFilaDestino - 1, 1).getValues();
    valoresDestinoB.forEach(fila => {
      let compStr = fila[0].toString().trim();
      if (compStr !== "") comprobantesExistentes.add(compStr);
    });
  }

  // 2. Buscar todos los archivos de Excel en la carpeta de Fichajes
  const carpetaDestino = DriveApp.getFolderById(CARPETA_FICHAJES_ID);
  const archivos = carpetaDestino.getFiles();
  
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName().toLowerCase();
    
    if (nombreArchivo.endsWith('.xls') || nombreArchivo.endsWith('.xlsx')) {
      let tempFileId = null;
      
      try {
        // Conversión temporal a Google Sheets
        const blob = archivo.getBlob();
        const config = { title: archivo.getName() + '_temp_fichajes', mimeType: MimeType.GOOGLE_SHEETS };
        const tempFile = Drive.Files.create(config, blob);
        tempFileId = tempFile.id;
        
        const tempSpreadsheet = SpreadsheetApp.openById(tempFileId);
        const hojaOrigen = tempSpreadsheet.getSheets()[0];
        const datosOrigen = hojaOrigen.getDataRange().getValues();
        
        // Según la estructura real: la fila 7 (índice 6) tiene las cabeceras reales
        const filaCabeceraOriginal = 6;
        const filaInicioDatosOriginal = 7; 
        
        if (datosOrigen.length <= filaInicioDatosOriginal) {
          eliminarArchivoTemporal(tempFileId);
          continue;
        }
        
        // Si la hoja destino está completamente vacía, creamos la fila 1 con los encabezados limpios
        if (hojaDestino.getLastRow() === 0) {
          const encabezadosDestino = [["Fecha", "Nº", "Cliente", "Socio", "Valores", "Importe", "Concepto"]];
          hojaDestino.getRange(1, 1, 1, 7).setValues(encabezadosDestino);
        }
        
        const registrosAAgregar = [];
        
        // Recorremos las filas de datos del archivo origen (desde índice 7 en adelante)
        for (let i = filaInicioDatosOriginal; i < datosOrigen.length; i++) {
          const filaOriginal = datosOrigen[i];
          if (!filaOriginal || filaOriginal[1] === undefined || filaOriginal[1] === "") continue; 
          
          let numComprobanteStr = filaOriginal[1].toString().trim(); // Columna B (Nº)
          
          // Validamos duplicados usando el número de comprobante único
          if (numComprobanteStr !== "" && !comprobantesExistentes.has(numComprobanteStr)) {
            
            // Tratamiento de la Fecha (Columna A - índice 0)
            let fechaValor = filaOriginal[0];
            if (fechaValor instanceof Date) {
              fechaValor = Utilities.formatDate(fechaValor, "America/Argentina/Buenos_Aires", "dd/MM/yyyy");
            } else if (typeof fechaValor === 'string' && fechaValor.includes("GMT")) {
              fechaValor = Utilities.formatDate(new Date(fechaValor), "America/Argentina/Buenos_Aires", "dd/MM/yyyy");
            }
            
            // --- SEPARACIÓN DEL CLIENTE Y LIMPIEZA DE TEXTO (SOLO NÚMEROS EN SOCIO) ---
            let textoCliente = filaOriginal[2] !== undefined ? filaOriginal[2].toString() : ""; // Columna C (Cliente)
            let parteNombre = textoCliente;
            let parteSocio = "";
            
            if (textoCliente.includes("(")) {
              parteNombre = textoCliente.split("(")[0].trim();
              
              let matchSocio = textoCliente.match(/\(([^)]+)\)/);
              if (matchSocio && matchSocio[1]) {
                // Removemos CUALQUIER carácter que no sea un número del 0 al 9 (quita "NS", "-" y espacios)
                parteSocio = matchSocio[1].replace(/[^0-9]/g, '').trim();
              }
            }
            
            // Armamos la estructura contigua de 7 columnas (A hasta G de destino)
            const nuevaFilaCompacta = [
              fechaValor,                                           // Columna A: Fecha limpia
              numComprobanteStr,                                    // Columna B: Nº Comprobante
              parteNombre,                                          // Columna C: Nombre del Cliente
              parteSocio,                                           // Columna D: Número de Socio (SOLO NÚMEROS)
              filaOriginal[3] !== undefined ? filaOriginal[3] : "", // Columna E: Valores (Efectivo)
              filaOriginal[4] !== undefined ? filaOriginal[4] : "", // Columna F: Importe (60000)
              filaOriginal[6] !== undefined ? filaOriginal[6] : ""  // Columna G: Conceptos (Columna G original, índice 6)
            ];
            
            registrosAAgregar.push(nuevaFilaCompacta);
            comprobantesExistentes.add(numComprobanteStr); 
          }
        }
        
        // 3. Pegar los registros acumulados abajo de todo de forma contigua
        if (registrosAAgregar.length > 0) {
          let filaInsercion = hojaDestino.getLastRow() + 1;
          if (filaInsercion < 2) {
            filaInsercion = 2; // Los datos empiezan en la fila 2 si los encabezados ocupan la 1
          }
          
          // Escribimos el bloque limpio de 7 columnas de ancho (A hasta G)
          hojaDestino.getRange(filaInsercion, 1, registrosAAgregar.length, 7).setValues(registrosAAgregar);
        }
        
        eliminarArchivoTemporal(tempFileId);
      } catch (error) {
        if (tempFileId) eliminarArchivoTemporal(tempFileId);
        Logger.log('Error procesando archivo de fichaje ' + archivo.getName() + ': ' + error.message);
      }
    }
  }
  ssDestino.toast('Hoja de Fichajes reestructurada con éxito desde A1.', '¡Éxito!');
}