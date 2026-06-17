const CARPETA_RAIZ_ID = '1pqMjUjZ-K4Bo3lC-kSYDaGjAUxQSaRrN';
const SPREADSHEET_DESTINO_ID = '1OQ-BGFqYxEqy1UR6YkREXnVqwaNiFmLVEWW_Q_SB50k';
const CARPETA_MATRICULAS_ID = '1R-s385voSUlgo33xa8pqpuc_KEUEKdZS';
const CARPETA_FICHAJES_ID = '1YrzDCmhVjcE3_qv_uUjeNq_UqFaob_DW';
const CARPETA_INGRESOS_ID = '1wVdejUgqYbgLDN1J9m4Nbjp9nvikpaMG';

/**
 * 1. Crea el menú "Club online" al abrir la hoja de cálculo.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Club online')
    .addItem('Actualizar Todo 🔄', 'actualizarTodo')
    .addItem('Actualizar Ingresos Mensuales 💰', 'actualizarIngresos') 
    .addSeparator()
    .addItem('Enviar Mails a Deudores ✉️', 'enviarMailsDeudores')
    .addItem('Generar Plantillas WhatsApp 💬', 'generarPlantillasWhatsApp')
    .addItem('Generar PDFs de Planteles 📄', 'generarPdfPlantelesMasivo')
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
    actualizarSocios();
    
    actualizarAntiguedadDeuda();

    actualizarMatriculas();

    actualizarFichajes();

    
    Logger.log('Proceso completado correctamente de manera silenciosa.');
    
  } catch (error) {
    Logger.log('Hubo un problema durante la actualización: ' + error.toString());
  }
}
function actualizarSocios() {
  const nombreCarpeta = "Socios";
  const nombreHojaDestino = "Socios";
  const columnasAMantener = [1, 2, 3, 5, 9, 11, 16, 22, 28, 29];
  const filaInicioOriginal = 6;
  // Enviamos true en un nuevo parámetro final para activar la limpieza de cadenas de pago
  procesarUltimoArchivo(CARPETA_RAIZ_ID, nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener, false, true);
}

function actualizarAntiguedadDeuda() {
  const nombreCarpeta = "AntiguedaddeDeuda";
  const nombreHojaDestino = "AntiguedaddeDeuda";
  const columnasAMantener = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const filaInicioOriginal = 5;
  procesarUltimoArchivo(CARPETA_RAIZ_ID, nombreCarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener, false, false);
}

/**
 * NUEVA FUNCIÓN: Procesa la carpeta de Matrículas invirtiendo el signo de los números
 */
function actualizarMatriculas() {
  const nombreHojaDestino = "Matriculas";
  // Mantiene la misma estructura de columnas y fila de inicio que AntiguedaddeDeuda
  const columnasAMantener = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const filaInicioOriginal = 5;
  
  // Llamamos al procesador indicándole la carpeta específica, activando el multiplicador por -1 y desactivando la limpieza de texto
  procesarUltimoArchivo(CARPETA_MATRICULAS_ID, null, nombreHojaDestino, filaInicioOriginal, columnasAMantener, true, false);
}

/**
 * Función auxiliar modificada encargada de buscar, filtrar, transformar y pegar en destino.
 */
function procesarUltimoArchivo(idCarpetaOrigen, nombreSubcarpeta, nombreHojaDestino, filaInicioOriginal, columnasAMantener, multiplicarPorMenosUno, limpiarFormasPago) {
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
        
        // REQUERIMIENTO ESPECIAL 1: Si es Matrícula, validamos y multiplicamos por -1 los números != 0
        if (multiplicarPorMenosUno && typeof valor === 'number' && valor !== 0) {
          valor = valor * -1;
        }
        
        // REQUERIMIENTO ESPECIAL 2: Limpieza estricta de formas de pago complejas al importar Socios
        if (limpiarFormasPago && typeof valor === 'string') {
          let textoLimpio = valor.trim();
          if (textoLimpio === "Entidad de Recaudación | Pagos Virtuales del Sur (Argentina)") {
            valor = "Debito automatico";
          } else if (textoLimpio === "Administración/Secretaría") {
            valor = "Tesoreria";
          } else if (textoLimpio === "Administración/Secretaría | Pagos Virtuales del Sur (Argentina)") {
            valor = "Tesoreria";
          } else {
            valor = textoLimpio;
          }
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
    
    // =================================================================
    // NUEVO REQUERIMIENTO: REGISTRO DE ARCHIVO PROCESADO EN HOJA "Aux"
    // =================================================================
    let hojaAux = ssDestino.getSheetByName('Aux');
    if (!hojaAux) hojaAux = ssDestino.insertSheet('Aux'); // Si no existe la crea
    
    // Obtenemos todos los valores de la columna J (columna 10) para ver cuál es la última fila real ocupada
    const valoresJ = hojaAux.getRange("J:J").getValues();
    let ultimaFilaJ = 0;
    for (let f = valoresJ.length - 1; f >= 0; f--) {
      if (valoresJ[f][0] !== "") {
        ultimaFilaJ = f + 1;
        break;
      }
    }
    
    // Si la columna está completamente vacía, empezamos en la fila 1, sino en la siguiente disponible
    let filaGuardado = ultimaFilaJ === 0 ? 1 : ultimaFilaJ + 1;
    
    // Generamos un texto descriptivo detallado (Ej: "Socios: archivo_socios_2026.xlsx")
    let registroTexto = nombreHojaDestino + ": " + ultimoArchivo.getName();
    hojaAux.getRange(filaGuardado, 10).setValue(registroTexto); // Escribe en la columna J (10)
    // =================================================================
    
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

function doGet(e) {
  let JSONString = "";
  let action = e && e.parameter ? e.parameter.action : "";

  // ACCIÓN 1: Alimentar el selector dinámico del login con las categorías reales
  if (action === "obtenerCategorias") {
    let listado = [];
    const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
    const hojaDivisiones = ss.getSheetByName('Divisiones');
    if (hojaDivisiones) {
      const datos = hojaDivisiones.getDataRange().getValues();
      // Fila 0 son cabeceras. Columna A (índice 0) tiene el nombre de la categoría
      for (let i = 1; i < datos.length; i++) {
        let cat = datos[i][0] ? datos[i][0].toString().trim() : "";
        if (cat) listado.push(cat);
      }
    }
    JSONString = JSON.stringify(listado.sort());

  // ACCIÓN 2: Validar contraseña introducida contra la pestaña "Divisiones"
  } else if (action === "login") {
    let categoriaBuscada = e.parameter.categoria || "";
    let claveBuscada = e.parameter.clave || "";
    let respuesta = { exito: false, mensaje: "Categoría no encontrada" };

    const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
    const hojaDivisiones = ss.getSheetByName('Divisiones');
    if (hojaDivisiones) {
      const datos = hojaDivisiones.getDataRange().getValues();
      for (let i = 1; i < datos.length; i++) {
        let catExistente = datos[i][0] ? datos[i][0].toString().trim() : "";
        let passCorrecta = datos[i][1] ? datos[i][1].toString().trim() : ""; // Columna B (índice 1)
        
        if (catExistente.toLowerCase() === categoriaBuscada.toLowerCase()) {
          if (passCorrecta === claveBuscada.toString().trim()) {
            respuesta = { exito: true, categoria: catExistente };
          } else {
            respuesta = { exito: false, mensaje: "Contraseña incorrecta para esta división." };
          }
          break;
        }
      }
    }
    JSONString = JSON.stringify(respuesta);

  // ACCIÓN 3: Ejecutar tu motor original de deudas pasando la división autenticada
  } else if (action === "obtenerDeudas") {
    let division = e.parameter.division || "";
    
    // LLAMADA CLAVE: Se comunica directo con tu función nativa de procesamiento
    let datosDeuda = obtenerDeudasPorDivision(division); 
    JSONString = JSON.stringify(datosDeuda);
  }

  // Retorno estructurado con cabeceras JSON para evitar bloqueos de CORS en Firebase
  return ContentService.createTextOutput(JSONString)
    .setMimeType(ContentService.MimeType.JSON);
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
  
  // Forzamos explícitamente la lectura de las 29 columnas (hasta AC)
  // Usamos getDisplayValues() para traer el texto exacto visible en la pantalla (evita problemas de formato)
  const ultimaFila = hojaReporte.getLastRow();
  const datosReporte = hojaReporte.getRange(1, 1, ultimaFila, 29).getDisplayValues();
  
  let nombresMeses = [];
  for (let col = 12; col <= 18; col++) { 
    let mesTexto = datosReporte[0][col] ? datosReporte[0][col].toString().trim() : "Mes";
    mesTexto = mesTexto.replace(/[\(\)]/g, ''); 
    nombresMeses.push(mesTexto);
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
      // Si usás getDisplayValues, los valores financieros pueden venir con el signo $ o puntos, 
      // por lo que limpiamos cualquier carácter extraño antes de convertirlos a número
      let matriculaTexto = filaR[23].toString().replace(/[^0-9.,-]/g, "").replace(",", ".");
      let matriculaValor = Math.abs(parseFloat(matriculaTexto)) || 0; 
      
      if (totalDeuda < 0 && matriculaValor <= 0) continue;
      let nombre = filaR[19] ? filaR[19].toString().trim() : "Sin Nombre";
      let formaPago = filaR[6] ? filaR[6].toString().trim() : "-";
      
      // Procesamiento directo y seguro del texto de la columna AC (Índice 28)
      let ultimoPagoValor = filaR[28];
      let ultimoPagoFormateado = "-";
      if (ultimoPagoValor && ultimoPagoValor.toString().trim() !== "") {
        ultimoPagoFormateado = ultimoPagoValor.toString().trim();
      }
      
      let descuento = filaR[8] ? filaR[8].toString().trim() : "";
      let periodoDesc = filaR[9] ? filaR[9].toString().trim() : "";
      
      let valoresMeses = [];
      for (let col = 12; col <= 18; col++) {
        let montoTexto = filaR[col].toString().replace(/[^0-9.,-]/g, "").replace(",", ".");
        valoresMeses.push(parseFloat(montoTexto) || 0);
      }
      
      let mesesDebidosPlanilla = parseInt(filaR[21]) || 0;
      let alertaCritica = (mesesDebidosPlanilla >= 2);
      
      listaJugadores.push({
        nombre: nombre,
        formaPago: formaPago,
        ultimoPago: ultimoPagoFormateado,
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
function generarPlantillasWhatsApp() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hojaDeuda = ss.getSheetByName('AntiguedaddeDeuda');
  const hojaSocios = ss.getSheetByName('Socios');
  
  // 1. Validar u obtener la hoja de destino para WhatsApp
  let hojaWA = ss.getSheetByName('Plantilla WA');
  if (!hojaWA) {
    hojaWA = ss.insertSheet('Plantilla WA');
  }
  
  // Limpiamos la hoja para poner los datos nuevos y actualizados
  hojaWA.clearContents();
  hojaWA.clearFormats();
  
  // Encabezados
  hojaWA.getRange(1, 1, 1, 2).setValues([["Teléfono", "Mensaje de WhatsApp"]]).setFontWeight("bold");
  
  if (!hojaDeuda || !hojaSocios) {
    ss.toast("Error: No se encuentran las hojas necesarias.", "Error");
    return;
  }
  
  const datosDeuda = hojaDeuda.getDataRange().getValues();
  const datosSocios = hojaSocios.getDataRange().getValues();
  
  // 2. Directorio de socios: mapeamos ID de Socio con el Teléfono (Columna 5 -> Índice 4)
  let directorioSocios = {};
  for (let j = 1; j < datosSocios.length; j++) {
    let idSocio = datosSocios[j][0] ? datosSocios[j][0].toString().trim() : "";
    let telefono = datosSocios[j][4] ? datosSocios[j][4].toString().trim() : "Sin Teléfono"; // Columna 5 es índice 4
    
    if (idSocio) {
      directorioSocios[idSocio] = telefono;
    }
  }
  
  let registrosWA = [];
  
  // Función auxiliar interna para dar formato de dinero: $X.XXX (sin decimales)
  const formatearDinero = (valor) => {
    return "$" + Math.round(valor).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  };
  
  // 3. Procesamos los deudores
  for (let i = 1; i < datosDeuda.length; i++) {
    let fila = datosDeuda[i];
    let idSocioDeuda = fila[0] ? fila[0].toString().trim() : "";
    let nombreJugador = fila[1];
    
    if (!idSocioDeuda || idSocioDeuda.toLowerCase().includes("total")) continue;
    
    let deudaMesActual = parseFloat(fila[3]) || 0; 
    let deudaAntiguaAcumulada = 0;
    let mesesDebidos = 0;
    
    // Contamos meses y sumamos deuda antigua (columnas de la 4 a la 9 en adelante)
    for (let col = 4; col <= 9; col++) {
      let valorCol = parseFloat(fila[col]) || 0;
      if (valorCol > 0) {
        mesesDebidos++;
        deudaAntiguaAcumulada += valorCol;
      }
    }
    
    // Si tiene deudas en los meses anteriores (Filtro idéntico al mail corporativo)
    if (deudaAntiguaAcumulada > 0) {
      // Sumamos 1 mes más al contador si también debe el mes actual
      if (deudaMesActual > 0) mesesDebidos++; 
      
      let telefonoDestino = directorioSocios[idSocioDeuda] || "Revisar Teléfono";
      let totalACobrar = deudaMesActual + deudaAntiguaAcumulada;
      
      // Aplicamos el formato solicitado sin decimales: $XX.XXX
      let totalFormateado = formatearDinero(totalACobrar);
      
      // Armamos la plantilla con el texto enriquecido para WhatsApp
      let cuerpoWhatsApp = `Estimado/a *${nombreJugador}*,\n\nLe comunicamos que registra saldo pendiente en su cuota social por más de 2 meses, acumulando un *Total a Cobrar de ${totalFormateado}*.\n\nSolicitamos tenga a bien acercarse a la tesorería en 495bis y 15 bis de lunes a viernes de 18 a 20hs, escribir a tesoreriagonnet@hotmail.com o por whatsapp al 2216819698 para regularizar su situación.\n\nAtentamente,\n*Tesorería - Club SFP Gonnet.*`;
      
      registrosWA.push({
        telefono: telefonoDestino,
        mensaje: cuerpoWhatsApp,
        meses: mesesDebidos // Guardamos el número de meses adeudados para poder ordenar
      });
    }
  }
  
  // 4. ORDENAR de mayor a menor según la cantidad de meses que deben
  registrosWA.sort((a, b) => b.meses - a.meses);
  
  // Convertimos el array de objetos a la estructura pura de filas de la planilla [Telefono, Mensaje]
  let matrizFinal = registrosWA.map(item => [item.telefono, item.mensaje]);
  
  // 5. Imprimir y formatear la hoja final
  if (matrizFinal.length > 0) {
    hojaWA.getRange(2, 1, matrizFinal.length, 2).setValues(matrizFinal);
    
    // Mejoras de visualización para que los humanos trabajen cómodos
    hojaWA.getRange(2, 2, matrizFinal.length, 1).setWrap(true); // Ajuste automático de línea para leer el mensaje completo
    hojaWA.setColumnWidth(1, 130); // Ancho para ver el teléfono
    hojaWA.setColumnWidth(2, 550); // Ancho amplio para el texto
    
    ss.toast("Se generaron " + matrizFinal.length + " mensajes ordenados por mora en 'Plantilla WA'.", "¡Éxito!");
  } else {
    ss.toast("No se encontraron deudores que cumplan las condiciones.", "Aviso");
  }
}

function generarPdfPlantelesMasivo() {
  const CARPETA_PDF_DESTINO_ID = '19heazUq7HMUunRm2Z3Mdy-2D0g2A4iQn';
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hojaReporte = ss.getSheetByName('Reporte');
  if (!hojaReporte) {
    SpreadsheetApp.getUi().alert('Error: No se encontró la hoja llamada "Reporte".');
    return;
  }
  
  const datos = hojaReporte.getDataRange().getValues();
  if (datos.length <= 1) {
    SpreadsheetApp.getUi().alert('La hoja "Reporte" no contiene datos suficientes.');
    return;
  }
  
  // 1. Obtener o validar carpeta de destino en Google Drive
  let carpetaDestino;
  try {
    carpetaDestino = DriveApp.getFolderById(CARPETA_PDF_DESTINO_ID);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error: No se pudo acceder a la carpeta de Drive provista. Verifica los permisos o el ID.');
    return;
  }

  // Borrar todo el contenido de la carpeta destino antes de empezar
  const archivosExistentes = carpetaDestino.getFiles();
  while (archivosExistentes.hasNext()) {
    archivosExistentes.next().setTrashed(true);
  }
  const subCarpetasExistentes = carpetaDestino.getFolders();
  while (subCarpetasExistentes.hasNext()) {
    subCarpetasExistentes.next().setTrashed(true);
  }
  
  const filaCabecera = datos[0];
  const indicesColumnas = [1, 2, 4, 5, 6, 20, 21];
  const cabecerasFiltradas = indicesColumnas.map(idx => filaCabecera[idx] ? filaCabecera[idx].toString().replace(/[\(\)]/g, '').trim() : "");
  // 2. Agrupar filas de jugadores por Plantel (La columna H sigue siendo el índice 7 para agrupar internamente)
  let plantelesMap = {};
  for (let i = 1; i < datos.length; i++) {
    let fila = datos[i];
    if (!fila || fila[7] === undefined || fila[7].toString().trim() === "") continue;
    
    let nombrePlantel = fila[7].toString().trim();
    if (!plantelesMap[nombrePlantel]) {
      plantelesMap[nombrePlantel] = [];
    }
    
    // Extraemos los valores de las columnas seleccionadas
    let filaFiltrada = indicesColumnas.map(idx => {
      let val = fila[idx];
      
      if (val instanceof Date) {
        return Utilities.formatDate(val, "America/Argentina/Buenos_Aires", "dd/MM/yyyy");
      }
      
      if (typeof val === 'string') {
        return val.trim();
      }
      
      // Formatear montos de dinero para las columnas financieras excepto la última de meses enteros
      if (typeof val === 'number') {
        if (idx === 21) { // Última columna del array indicesColumnas (Mapea la columna V de la hoja)
          return Math.round(val).toString();
        } else if (idx >= 12) {
          return "$" + Math.round(val).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
        }
      }
      
      return val !== undefined && val !== null ? val.toString().trim() : "";
    });
    
    plantelesMap[nombrePlantel].push(filaFiltrada);
  }
  
  let totalCreados = 0;
  const listaPlanteles = Object.keys(plantelesMap);
  const fechaHoyStr = Utilities.formatDate(new Date(), "America/Argentina/Buenos_Aires", "dd/MM/yyyy");

  // 4. Iterar sobre cada plantel para armar y exportar su PDF sin la columna H
  listaPlanteles.forEach(plantel => {
    let filasJugadores = plantelesMap[plantel];
    
    let htmlContent = `
    <html>
    <head>
      <style>
        @page {
          size: A4 landscape;
          margin: 8mm 6mm;
        }
 
        body {
          font-family: Arial, sans-serif;
          color: #222222;
          margin: 0;
          padding: 0;
          font-size: 10pt; /* Tamaño base estándar para todas las letras del reporte */
        }
        .header {
        
          background-color: #ffffff; /* Fondo blanco limpio para contraste */
          color: #222222;
          padding: 0 0 10px 0;
          border-bottom: 3px solid #111111;
          margin-bottom: 12px;
        }
        .header table {
          width: 100%;
          border-collapse: collapse;
 
        }
        .header td {
          border: none;
          padding: 0;
          vertical-align: bottom;
        }
        .titulo {
          font-size: 14pt;
          /* El único con tamaño destacado según instrucción 1 */
          font-weight: bold;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          margin-bottom: 4px;
        }
        .meta-texto {
          font-size: 10pt;
          /* Igual al tamaño base */
          color: #222222;
          /* Texto en negro para máxima visibilidad */
          font-weight: normal;
        }
        .meta-texto strong {
          font-weight: bold;
        }
        .logo-container {
          text-align: right;
          width: 65px;
        }
        .logo-club {
          height: 48px;
          width: auto;
        }
        table.datos-table {
          width: 100%;
          border-collapse: collapse;
          margin-top: 5px;
        }
        th {
          background-color: #f2f2f2;
          color: #111111;
          font-weight: bold;
          border: 1px solid #cccccc;
          padding: 4px 3px;
          font-size: 10pt;
          /* Igual al tamaño base */
          text-transform: uppercase;
          text-align: left;
        }
        td {
          border: 1px solid #e0e0e0;
          padding: 4px 3px;
          text-align: left;
          font-size: 10pt; /* Igual al tamaño base */
          white-space: nowrap;
        }
        tr:nth-child(even) {
          background-color: #fafafa;
        }
        .monto {
          text-align: right;
          font-weight: bold;
        }
        /* Alineación e identidad estricta de número entero para la columna de meses debidos */
        .meses-der {
          text-align: right !important;
          font-weight: bold;
        }
      </style>
    </head>
    <body>
      <div class="header">
        <table>
          <tr>
            <td>
              <div class="titulo">Reporte de Control de Deuda</div>
              <div class="meta-texto">
                 Plantel: <strong>${plantel}</strong>  &nbsp;|&nbsp; Fecha de Emisión: <strong>${fechaHoyStr}</strong>
              </div>
            </td>
                      </tr>
        </table>
      </div>
      
      <table class="datos-table">
        <thead>
           <tr>
    `;
    
    // Inyectar cabeceras sin la columna H
    cabecerasFiltradas.forEach(cab => {
      htmlContent += `<th>${cab}</th>`;
    });
    htmlContent += `
          </tr>
        </thead>
        <tbody>
    `;
    // Inyectar filas procesadas
    filasJugadores.forEach(fila => {
      htmlContent += `<tr>`;
      fila.forEach((celda, idxCelda) => {
        let clases = [];
        
        // Si es la última columna (Índice de celda final), aplicar alineación numérica derecha sin formatos raros
        if (idxCelda === fila.length - 1) {
          clases.push('meses-der');
         } else if (celda.toString().indexOf('$') !== -1) {
          clases.push('monto');
        }
        
        let claseAtributo = clases.length > 0 ? ` class="${clases.join(' ')}"` : '';
        htmlContent += `<td${claseAtributo}>${celda}</td>`;
      });
      htmlContent += `</tr>`;
    });
    htmlContent += `
        </tbody>
      </table>
    </body>
    </html>
    `;
    // 5. Crear archivo HTML temporal y convertir a PDF nativo
    let blobHtml = Utilities.newBlob(htmlContent, "text/html", "temp.html");
    let archivoTemp = carpetaDestino.createFile(blobHtml);
    
    let pdfBlob = archivoTemp.getAs(MimeType.PDF).setName(`Reporte_Deudas_${plantel.replace(/[^a-zA-Z0-9]/g, "_")}.pdf`);
    carpetaDestino.createFile(pdfBlob);
    
    archivoTemp.setTrashed(true);
    totalCreados++;
  });
  ss.toast(`Se generaron ${totalCreados} PDFs optimizados con los nuevos lineamientos gráficos.`, "¡Éxito!");
}
// Carga dinámicamente las categorías reales en el select del Login
function obtenerCategoriasLogin() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hojaDivisiones = ss.getSheetByName('Divisiones');
  
  if (!hojaDivisiones) return [];
  
  const datos = hojaDivisiones.getDataRange().getValues();
  let categorias = [];
  
  // Asumiendo que Fila 0 son cabeceras y Columna A (índice 0) tiene el nombre de la categoría
  for (let i = 1; i < datos.length; i++) {
    let cat = datos[i][0] ? datos[i][0].toString().trim() : "";
    if (cat) {
      categorias.push(cat);
    }
  }
  return categorias.sort();
}

// Valida las credenciales introducidas por el técnico contra la pestaña Divisiones
function verificarCredencialesTecnico(categoria, contraseniaIntroducida) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hojaDivisiones = ss.getSheetByName('Divisiones');
  
  if (!hojaDivisiones) {
    return { exito: false, mensaje: "Error interno: No se encontró la hoja Divisiones." };
  }
  
  const datos = hojaDivisiones.getDataRange().getValues();
  
  for (let i = 1; i < datos.length; i++) {
    let catExistente = datos[i][0] ? datos[i][0].toString().trim() : "";
    let passCorrecta = datos[i][1] ? datos[i][1].toString().trim() : ""; // Columna B (índice 1)
    
    if (catExistente.toLowerCase() === categoria.toLowerCase()) {
      if (passCorrecta === contraseniaIntroducida.toString().trim()) {
        return { exito: true, categoria: catExistente };
      } else {
        return { exito: false, mensaje: "Contraseña incorrecta. Inténtalo de nuevo." };
      }
    }
  }
  
  return { exito: false, mensaje: "La categoría seleccionada no fue encontrada." };
}
/**
 * FUNCIÓN ACTUALIZADA: Procesa todos los archivos de la carpeta Ingresos.
 * - Copia absolutamente TODAS las filas desde la fila 7 hasta la última celda de cada archivo.
 * - Separa la Columna E (Cuenta) en dos columnas en base al paréntesis "( )".
 * - Limpia y recorta la Columna H (columna G en el índice) eliminando todo lo posterior al símbolo ">".
 * - Aplica reemplazo estricto para "Ingreso en Cuenta de Entidad Financiera > Pagos Virtuales del Sur..." por "Debito Automatico".
 * - Pega el bloque de corrido uno abajo del otro en la hoja 'Ingresos', desde la fila 2 en adelante.
 */
function actualizarIngresos() {
  const nombreHojaDestino = "Ingresos";
  const ssDestino = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  let hojaDestino = ssDestino.getSheetByName(nombreHojaDestino);
  
  if (!hojaDestino) {
    hojaDestino = ssDestino.insertSheet(nombreHojaDestino);
  }

  // Localizar y recorrer la carpeta de Ingresos en Google Drive
  const carpetaOrigen = DriveApp.getFolderById(CARPETA_INGRESOS_ID);
  const archivos = carpetaOrigen.getFiles();
  let archivosProcesadosContador = 0;
  
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName().toLowerCase();
    
    // Validar que sea un archivo Excel
    if (nombreArchivo.endsWith('.xls') || nombreArchivo.endsWith('.xlsx')) {
      let tempFileId = null;
      try {
        // Conversión temporal interna a formato Google Sheets para poder estructurar la lectura de rangos
        const blob = archivo.getBlob();
        const config = { title: archivo.getName() + '_temp_ingresos', mimeType: MimeType.GOOGLE_SHEETS };
        const tempFile = Drive.Files.create(config, blob);
        tempFileId = tempFile.id;
        
        const tempSpreadsheet = SpreadsheetApp.openById(tempFileId);
        const hojaOrigen = tempSpreadsheet.getSheets()[0];
        const datosOrigen = hojaOrigen.getDataRange().getValues();
        
        const filaInicioDatosOriginal = 7; // Fila 7 real (Índice 6 en la matriz)
        
        if (datosOrigen.length >= filaInicioDatosOriginal) {
          
          // Si la hoja destino está vacía, armamos la cabecera en la Fila 1 adaptada al desglose de cuenta
          if (hojaDestino.getLastRow() === 0) {
            const encabezadosPorDefecto = [["Fecha", "Nº Comprobante", "Detalle / Cliente", "Categoría", "Cuenta Nombre", "Cuenta Detalle ( )", "Importe", "Concepto", "Forma de Pago"]];
            hojaDestino.getRange(1, 1, 1, encabezadosPorDefecto[0].length).setValues(encabezadosPorDefecto);
          }
          
          const registrosAAgregar = [];
          
          // Recorremos y extraemos TODAS las filas desde la 7
          for (let i = filaInicioDatosOriginal - 1; i < datosOrigen.length; i++) {
            const filaOriginal = datosOrigen[i];
            
            // Formateo rápido para celdas tipo fecha (Columna A)
            let fechaValor = filaOriginal[0];
            if (fechaValor instanceof Date) {
              fechaValor = Utilities.formatDate(fechaValor, "America/Argentina/Buenos_Aires", "dd/MM/yyyy");
            }
            
            // ---- REQUERIMIENTO ANTERIOR: SEPARACIÓN DE LA COLUMNA E (Índice 4) ----
            let cuentaTextoOriginal = filaOriginal[4] ? filaOriginal[4].toString().trim() : "";
            let parteTextoInicial = cuentaTextoOriginal;
            let parteDentroParentesis = "";
            
            const coincidenciaParentesis = cuentaTextoOriginal.match(/(.*?)\((.*?)\)/);
            if (coincidenciaParentesis) {
              parteTextoInicial = coincidenciaParentesis[1].trim();     
              parteDentroParentesis = coincidenciaParentesis[2].trim(); 
            }
            // ------------------------------------------------------------------------
            
            // ---- NUEVO REQUERIMIENTO: LIMPIEZA DE LA COLUMNA H (Índice 6 en la fila original) ----
            let textoColumnaH = filaOriginal[6] ? filaOriginal[6].toString().trim() : "";

// 1. REGLA DEL > (Se ejecuta primero: corta el texto si tiene el símbolo)
if (textoColumnaH.includes(">")) {
  textoColumnaH = textoColumnaH.split(">")[0].trim();
}

// 2. REEMPLAZO DE TEXTO (Se ejecuta segundo: evalúa la frase ya recortada)
if (textoColumnaH === "Ingreso en Cuenta de Entidad Financiera") {
  textoColumnaH = "Debito Automatico";
}
            // --------------------------------------------------------------------------------------
            
            // Reconstruimos la fila para inyectar en la sábana contigua de destino
            const nuevaFilaMatriz = [];
            
            for (let j = 0; j < filaOriginal.length; j++) {
              if (j === 0) {
                nuevaFilaMatriz.push(fechaValor ? fechaValor : ""); // Fecha formateada
              } else if (j === 4) {
                // En la posición de la columna E original, metemos los dos campos del desglose
                nuevaFilaMatriz.push(parteTextoInicial);
                nuevaFilaMatriz.push(parteDentroParentesis);
              } else if (j === 6) {
                // En la posición de la columna H original (índice 6), metemos el nuevo texto limpio
                nuevaFilaMatriz.push(textoColumnaH);
              } else {
                nuevaFilaMatriz.push(filaOriginal[j]);
              }
            }
            
            registrosAAgregar.push(nuevaFilaMatriz);
          }
          
          // Inyección contigua en la parte inferior de la hoja 'Ingresos'
          if (registrosAAgregar.length > 0) {
            let filaInsercion = hojaDestino.getLastRow() + 1;
            if (filaInsercion < 2) {
              filaInsercion = 2; // Respetamos encabezados
            }
            
            // Pega de golpe todo el bloque de filas transformadas
            hojaDestino.getRange(filaInsercion, 1, registrosAAgregar.length, registrosAAgregar[0].length).setValues(registrosAAgregar);
            archivosProcesadosContador++;
            
            // =================================================================
            // REGISTRO AUDITORÍA: Guarda el historial en la hoja "Aux" (Columna J)
            // =================================================================
            let hojaAux = ssDestino.getSheetByName('Aux');
            if (!hojaAux) hojaAux = ssDestino.insertSheet('Aux');
            
            const valoresJ = hojaAux.getRange("J:J").getValues();
            let ultimaFilaJ = 0;
            for (let f = valoresJ.length - 1; f >= 0; f--) {
              if (valoresJ[f][0] !== "") {
                ultimaFilaJ = f + 1;
                break;
              }
            }
            let filaGuardadoAux = ultimaFilaJ === 0 ? 1 : ultimaFilaJ + 1;
            hojaAux.getRange(filaGuardadoAux, 10).setValue(nombreHojaDestino + ": " + archivo.getName());
            // =================================================================
          }
        }
        
        eliminarArchivoTemporal(tempFileId);
      } catch (error) {
        if (tempFileId) eliminarArchivoTemporal(tempFileId);
        Logger.log('Error procesando archivo de ingresos ' + archivo.getName() + ': ' + error.message);
      }
    }
  }
  
  if (archivosProcesadosContador > 0) {
    ssDestino.toast('Se importaron los ingresos limpiando y recortando la columna H.', '¡Éxito!');
  } else {
    ssDestino.toast('No se encontraron archivos de Excel para procesar.', 'Aviso');
  }
}