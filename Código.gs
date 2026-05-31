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
    .addItem('Actualizar Socios', 'actualizarSocios')
    .addItem('Antigüedad de Deuda', 'actualizarAntiguedadDeuda')
    .addToUi();
}

/**
 * 2. Función para procesar y actualizar la hoja de Socios.
 */
function actualizarSocios() {
  const nombreCarpeta = "Socios";
  const nombreHojaDestino = "Socios";
  
  // Fila inicial de datos (después de eliminar 1 a 5, la fila 6 original pasa a ser la 1 de los encabezados)
  // Columnas mapeadas de Excel a mantener: A, B, C, D, I, K, P, V, AB, AC
  // En base 1 (A=1, B=2, C=3, D=4, I=9, K=11, P=16, V=22, AB=28, AC=29)
  const columnasAMantener = [1, 2, 3, 4, 9, 11, 16, 22, 28, 29];
  const filaInicioOriginal = 6; 

  procesarUltimoArchivo(nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener);
}

/**
 * 3. Función para procesar y actualizar la hoja de Antigüedad de Deuda.
 */
function actualizarAntiguedadDeuda() {
  const nombreCarpeta = "AntiguedaddeDeuda";
  const nombreHojaDestino = "AntiguedaddeDeuda";
  
  // Columnas mapeadas de Excel a mantener: de la A a la K (columnas 1 a 11 consecutivas)
  const columnasAMantener = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const filaInicioOriginal = 5; // Elimina de la 1 a la 4, la 5 son los encabezados

  procesarUltimoArchivo(nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener);
}

/**
 * Función auxiliar encargada de buscar el último archivo Excel (.xls o .xlsx), 
 * extraer sus datos aplicando los filtros de filas/columnas y pegarlos en el destino.
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
  
  // Buscar el archivo más reciente basándonos en la fecha de modificación de Drive
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
  
  // Al ser un archivo generado por un sistema (.xls), Drive lo maneja temporalmente convirtiéndolo a Google Sheets en memoria
  let tempSpreadsheet = null;
  try {
    // Intentamos la conversión directa mediante la API avanzada de Drive o blobs de Sheets
    const fileId = ultimoArchivo.getId();
    const blob = ultimoArchivo.getBlob();
    const config = {
      title: ultimoArchivo.getName() + '_temp',
      mimeType: MimeType.GOOGLE_SHEETS
    };
    
    // Crear archivo temporal convertido a Google Sheets
    const tempFile = Drive.Files.create(config, blob);
    tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
    const hojaOrigen = tempSpreadsheet.getSheets()[0];
    const datosOrigen = hojaOrigen.getDataRange().getValues();
    
    if (datosOrigen.length < filaInicioOriginal) {
      SpreadsheetApp.getUi().alert('El archivo no contiene suficientes filas según las especificaciones.');
      eliminarArchivoTemporal(tempFile.id);
      return;
    }
    
    // Filtrar filas y seleccionar solo las columnas deseadas
    const matrizFiltrada = [];
    for (let i = filaInicioOriginal - 1; i < datosOrigen.length; i++) {
      const filaOriginal = datosOrigen[i];
      const nuevaFila = [];
      
      columnasAMantener.forEach(idx => {
        // En JS los arrays inician en 0, restamos 1 al número de columna original
        nuevaFila.push(filaOriginal[idx - 1] !== undefined ? filaOriginal[idx - 1] : "");
      });
      
      matrizFiltrada.push(nuevaFila);
    }
    
    // Escribir en la hoja de destino del Spreadsheet correspondiente
    const ssDestino = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
    let hojaDestino = ssDestino.getSheetByName(nombreHojaDestino);
    
    if (!hojaDestino) {
      hojaDestino = ssDestino.insertSheet(nombreHojaDestino);
    }
    
    // Limpiamos contenido previo de la hoja de destino para evitar solapamientos viejos
    hojaDestino.clearContents();
    hojaDestino.clearFormats();
    
    // Pegar la nueva estructura de datos (encabezados incluidos en la primera posición de la matriz)
    if (matrizFiltrada.length > 0) {
      hojaDestino.getRange(1, 1, matrizFiltrada.length, matrizFiltrada[0].length).setValues(matrizFiltrada);
    }
    
    // Eliminar el archivo temporal de conversión de Drive
    eliminarArchivoTemporal(tempFile.id);
    
    SpreadsheetApp.getUi().alert('¡Éxito!: Hoja de "' + nombreHojaDestino + '" actualizada correctamente con el archivo: ' + ultimoArchivo.getName());
    
  } catch (error) {
    if (tempSpreadsheet) {
      eliminarArchivoTemporal(tempSpreadsheet.getId());
    }
    SpreadsheetApp.getUi().alert('Ocurrió un error al procesar el archivo: ' + error.message);
  }
}

function eliminarArchivoTemporal(id) {
  try {
    Drive.Files.remove(id);
  } catch(e) {
    // Si no se tiene activada la API avanzada, intenta mandarlo a la papelera como alternativa básica
    try { DriveApp.getFileById(id).setTrashed(true); } catch(err) {}
  }
}

/**
 * 1. Crea el menú "Club online" al abrir la hoja de cálculo unificando todas las opciones.
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

/**
 * 2. Función para procesar y actualizar la hoja de Socios.
 */
function actualizarSocios() {
  const nombreCarpeta = "Socios";
  const nombreHojaDestino = "Socios";
  const columnasAMantener = [1, 2, 3, 4, 9, 11, 16, 22, 28, 29];
  const filaInicioOriginal = 6; 
  procesarUltimoArchivo(nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener);
}

/**
 * 3. Función para procesar y actualizar la hoja de Antigüedad de Deuda.
 */
function actualizarAntiguedadDeuda() {
  const nombreCarpeta = "AntiguedaddeDeuda";
  const nombreHojaDestino = "AntiguedaddeDeuda";
  const columnasAMantener = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const filaInicioOriginal = 5;
  procesarUltimoArchivo(nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener);
}

/**
 * Función auxiliar para procesar el archivo Excel más reciente de Drive
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
      SpreadsheetApp.getUi().alert('El archivo no contiene suficientes filas según las especificaciones.');
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
// NUEVAS FUNCIONES PARA LA WEB APP MÓVIL Y MAILS
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Consulta de Deudas - Club Online')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function obtenerDivisiones() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hoja = ss.getSheetByName('AntiguedaddeDeuda');
  if(!hoja) return [];
  const datos = hoja.getDataRange().getValues();
  let divisiones = [];
  for (let i = 1; i < datos.length; i++) {
    let div = datos[i][0]; // Columna A
    if (div && divisiones.indexOf(div) === -1 && !div.toLowerCase().includes("total")) {
      divisiones.push(div);
    }
  }
  return divisiones.sort();
}

function obtenerDeudasPorDivision(divisionBuscada) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hojaDeuda = ss.getSheetByName('AntiguedaddeDeuda');
  if(!hojaDeuda) return [];
  const datosDeuda = hojaDeuda.getDataRange().getValues();
  let resultados = [];
  
  for (let i = 1; i < datosDeuda.length; i++) {
    let fila = datosDeuda[i];
    if (fila[0] === divisionBuscada && !fila[0].toLowerCase().includes("total")) {
      let nombre = fila[1]; // Columna B
      let totalDeuda = fila[2]; // Columna C
      let mes1 = parseFloat(fila[3]) || 0; // Columna D (Abr-26 / Mes de Alerta Mail)
      let mes2 = parseFloat(fila[4]) || 0; // Columna E (Mar-26)
      let mes3 = parseFloat(fila[5]) || 0; // Columna F (Feb-26)
      
      let mesesConDeuda = 0;
      if (mes1 > 0) mesesConDeuda++;
      if (mes2 > 0) mesesConDeuda++;
      if (mes3 > 0) mesesConDeuda++;
      
      // Indicador móvil: Debe más de 2 meses basándose en D, E y F
      let alertaCritica = (mesesConDeuda >= 2);
      
      resultados.push({
        nombre: nombre, total: totalDeuda,
        mes1: mes1, mes2: mes2, mes3: mes3,
        alerta: alertaCritica
      });
    }
  }
  return resultados;
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
  
  // Guardar correos indexándolos por el nombre completo en minúsculas
  let directorioSocios = {};
  for (let j = 1; j < datosSocios.length; j++) {
    let apellido = datosSocios[j][1]; // Columna B (Apellido)
    let nombre = datosSocios[j][2];   // Columna C (Nombres)
    let email = datosSocios[j][5];    // Columna F (Correo Electrónico del reporte de Socios reducido)
    if (apellido && nombre && email) {
      let llave = (nombre.trim() + " " + apellido.trim()).toLowerCase();
      directorioSocios[llave] = email.trim();
    }
  }
  
  let correosEnviados = 0;
  for (let i = 1; i < datosDeuda.length; i++) {
    let fila = datosDeuda[i];
    let nombreCompletoOriginal = fila[1]; // Columna B
    let deudaMesActual = parseFloat(fila[3]) || 0; // Columna D (Si debe más de 1 mes / activo)
    
    if (nombreCompletoOriginal && deudaMesActual > 0) {
      let llaveBusqueda = nombreCompletoOriginal.trim().toLowerCase();
      let emailDestino = directorioSocios[llaveBusqueda];
      
      // Intento alternativo de coincidencia cruzada invirtiendo orden
      if (!emailDestino) {
        let partes = llaveBusqueda.split(" ");
        if (partes.length >= 2) {
          let llaveInvertida = partes.slice(1).join(" ") + " " + partes[0];
          emailDestino = directorioSocios[llaveInvertida];
        }
      }
      
      if (emailDestino && emailDestino.includes("@")) {
        let asunto = "Aviso Importante: Regularización de Cuota - Club Online";
        let cuerpo = `Estimado/a ${nombreCompletoOriginal},\n\nLe comunicamos que registra un saldo pendiente en su cuota social por un importe de $${deudaMesActual}.\n\nSolicitamos tenga a bien acercarse a la tesorería del club para regularizar su situación.\n\nAtentamente,\nDepto. de Tesorería - Club Online.`;
        MailApp.sendEmail(emailDestino, asunto, cuerpo);
        correosEnviados++;
      }
    }
  }
  return "Proceso completado. Se enviaron exitosamente " + correosEnviados + " correos electrónicos.";
}