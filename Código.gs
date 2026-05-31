// ID de la carpeta principal de Tesorería provisto por el usuario
const CARPETA_RAIZ_ID = '1pqMjUjZ-K4Bo3lC-kSYDaGjAUxQSaRrN';
// ID del archivo de Google Sheets de destino
const SPREADSHEET_DESTINO_ID = '1OQ-BGFqYxEqy1UR6YkREXnVqwaNiFmLVEWW_Q_SB50k';

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
    const fileId = ultimoArchivo.getId();
    const blob = ultimoArchivo.getBlob();
    const config = {
      title: ultimoArchivo.getName() + '_temp',
      mimeType: MimeType.GOOGLE_SHEETS
    };
    
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
    
    if (!hojaDestino) {
      hojaDestino = ssDestino.insertSheet(nombreHojaDestino);
    }
    
    hojaDestino.clearContents();
    hojaDestino.clearFormats();
    
    if (matrizFiltrada.length > 0) {
      hojaDestino.getRange(1, 1, matrizFiltrada.length, matrizFiltrada[0].length).setValues(matrizFiltrada);
    }
    
    eliminarArchivoTemporal(tempFile.id);
    SpreadsheetApp.getUi().alert('¡Éxito!: Hoja de "' + nombreHojaDestino + '" actualizada correctamente.');
    
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
    try { DriveApp.getFileById(id).setTrashed(true); } catch(err) {}
  }
}

// =================================================
// FUNCIONES DE COMUNICACIÓN PARA LA WEB APP MÓVIL
// =================================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Consulta de Deudas - Club Online')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Obtiene las divisiones directamente desde la hoja limpia "Divisiones"
 */
function obtenerDivisiones() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hoja = ss.getSheetByName('Divisiones');
  if (!hoja) return [];
  
  const datos = hoja.getDataRange().getValues();
  let divisiones = [];
  
  for (let i = 1; i < datos.length; i++) {
    let div = datos[i][0]; // Columna A
    if (div && divisiones.indexOf(div.trim()) === -1) {
      divisiones.push(div.trim());
    }
  }
  return divisiones;
}

/**
 * DETECTOR DETALLADO: Escanea toda la fila para encontrar la división del jugador, 
 * solucionando desajustes en tablas dinámicas o corrimiento de celdas.
 */
function obtenerDeudasPorDivision(divisionBuscada) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hojaDeuda = ss.getSheetByName('AntiguedaddeDeuda');
  if (!hojaDeuda) return [];
  
  const datosDeuda = hojaDeuda.getDataRange().getValues();
  let resultados = [];
  
  if (!divisionBuscada) return [];
  
  // Extraemos componentes numéricos clave (ej: "2010" o "17") para mapeo tolerante
  let buscarLimpio = divisionBuscada.toLowerCase().replace(/[^a-z0-9]/g, "");
  let numeroBuscado = divisionBuscada.match(/\d+/);
  let strNumero = numeroBuscado ? numeroBuscado[0] : null;

  // Analizar los encabezados (Fila 1) para ubicar las columnas clave por texto
  let encabezados = datosDeuda[0].map(e => e.toString().toLowerCase());
  let idxNombre = encabezados.findIndex(e => e.includes("nombre") || e.includes("completo"));
  let idxTotal = encabezados.findIndex(e => e.includes("total") || e.includes("informar"));
  
  // Si no se encuentran por nombre, volvemos a los índices por defecto estables (Col B y Col C)
  if (idxNombre === -1) idxNombre = 1;
  if (idxTotal === -1) idxTotal = 2;

  // Recorremos las filas buscando coincidencias
  for (let i = 1; i < datosDeuda.length; i++) {
    let fila = datosDeuda[i];
    if (!fila || fila.length === 0) continue;
    
    // Convertimos toda la fila a texto plano continuo para analizar si la categoría está en celdas grises o movidas
    let filaTextoCompleto = fila.join(" ").toLowerCase();
    
    // Filtro de exclusión para totales generales
    if (filaTextoCompleto.includes("total general")) continue;
    
    let filaLimpia = filaTextoCompleto.replace(/[^a-z0-9]/g, "");
    
    // VALIDACIÓN DINÁMICA: Verifica si la categoría seleccionada por el técnico existe dentro de la fila actual
    let coincideTexto = filaLimpia.includes(buscarLimpio);
    let coincideNumero = strNumero && filaTextoCompleto.includes(strNumero);
    
    // Si la fila pertenece a la división buscada, extraemos al jugador
    if (coincideTexto || (divisionBuscada.toLowerCase().includes("sub") && coincideNumero)) {
      
      let nombre = fila[idxNombre] ? fila[idxNombre].toString().trim() : "";
      
      // Si la celda del nombre es igual a la división (caso de fila de separación), saltamos la línea
      if (!nombre || nombre.toLowerCase().includes("fútbol") || nombre.toLowerCase().includes("divisiones") || nombre.toLowerCase().includes("total")) continue;
      
      let totalDeuda = parseFloat(fila[idxTotal]) || 0;
      let mes1 = parseFloat(fila[3]) || 0; // Columna D
      let mes2 = parseFloat(fila[4]) || 0; // Columna E
      let mes3 = parseFloat(fila[5]) || 0; // Columna F
      
      let mesesConDeuda = 0;
      if (mes1 > 0) mesesConDeuda++;
      if (mes2 > 0) mesesConDeuda++;
      if (mes3 > 0) mesesConDeuda++;
      
      let alertaCritica = (mesesConDeuda >= 2);
      
      resultados.push({
        nombre: nombre, 
        total: totalDeuda,
        mes1: mes1, 
        mes2: mes2, 
        mes3: mes3,
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
  
  let directorioSocios = {};
  for (let j = 1; j < datosSocios.length; j++) {
    let apellido = datosSocios[j][1]; // Columna B
    let nombre = datosSocios[j][2];   // Columna C
    let email = datosSocios[j][5];    // Columna F
    if (apellido && nombre && email) {
      let llave = (nombre.trim() + " " + apellido.trim()).toLowerCase();
      directorioSocios[llave] = email.trim();
    }
  }
  
  let correosEnviados = 0;
  for (let i = 1; i < datosDeuda.length; i++) {
    let fila = datosDeuda[i];
    let nombreCompletoOriginal = fila[1]; 
    let deudaMesActual = parseFloat(fila[3]) || 0; 
    
    if (nombreCompletoOriginal && deudaMesActual > 0) {
      let llaveBusqueda = nombreCompletoOriginal.trim().toLowerCase();
      let emailDestino = directorioSocios[llaveBusqueda];
      
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