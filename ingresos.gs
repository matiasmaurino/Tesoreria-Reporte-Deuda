/**
 * Importa los archivos de la carpeta de Ingresos acumulándolos en la hoja "Ingresos".
 * Control de duplicados:
 *   - Por archivo: si el nombre ya figura en ImportacionArchivosClubOnline, se saltea.
 *   - Por fila: NO se agrega una fila si su columna C (Nº de recibo) ya existe en la hoja
 *     (o ya apareció antes dentro del mismo archivo).
 */
function actualizarIngresos() {
  const TIEMPO_INICIO = new Date().getTime();
  const LIMITE_MS = 5 * 60 * 1000;

  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const CARPETA_INGRESOS_ID = '1hKRtI2E_kTJY_Q51eKckTQ_dTU0MBE_q';
  const NOMBRE_CARPETA_PROCESADOS = 'Procesados';
  const MAX_ARCHIVOS_POR_TANDA = 1;
  const FILA_INICIO_DATOS = 7;
  const COL_RECIBO = 2; // columna C (A = 0)

  let hojaIngresos = ss.getSheetByName('Ingresos');
  if (!hojaIngresos) {
    hojaIngresos = ss.insertSheet('Ingresos');
  }

  // Nº de recibo ya cargados en la hoja Ingresos (columna C)
  const recibosExistentes = new Set();
  const ultimaFilaIng = hojaIngresos.getLastRow();
  if (ultimaFilaIng >= 1) {
    hojaIngresos.getRange(1, COL_RECIBO + 1, ultimaFilaIng, 1).getValues().forEach(f => {
      const clave = claveRecibo_(f[0]);
      if (clave) recibosExistentes.add(clave);
    });
  }

  // Archivos ya importados (control por nombre)
  const hojaControl = ss.getSheetByName(HOJA_CONTROL_IMPORTACIONES);
  let archivosProcesados = [];
  if (hojaControl && hojaControl.getLastRow() > 1) {
    archivosProcesados = hojaControl.getRange(2, 1, hojaControl.getLastRow() - 1, 1)
      .getValues()
      .map(row => String(row[0]).trim());
  }

  const carpeta = DriveApp.getFolderById(CARPETA_INGRESOS_ID);

  // Buscar o crear la subcarpeta "Procesados"
  let carpetaProcesados;
  const subcarpetas = carpeta.getFoldersByName(NOMBRE_CARPETA_PROCESADOS);
  carpetaProcesados = subcarpetas.hasNext() ? subcarpetas.next() : carpeta.createFolder(NOMBRE_CARPETA_PROCESADOS);

  const archivos = carpeta.getFiles();

  let archivosIgnoradosPorDuplicado = 0;
  let archivosProcesadosEnEstaTanda = 0;
  let filasAgregadasTotal = 0;
  let filasDuplicadasTotal = 0;
  let filasSinReciboTotal = 0;

  while (archivos.hasNext()) {
    if (new Date().getTime() - TIEMPO_INICIO > LIMITE_MS) {
      Logger.log('Se acerca el límite de tiempo. Cortando ejecución de forma segura.');
      break;
    }
    if (archivosProcesadosEnEstaTanda >= MAX_ARCHIVOS_POR_TANDA) break;

    const archivo = archivos.next();
    const nombreArchivo = archivo.getName().trim();
    const nombreMin = nombreArchivo.toLowerCase();
    if (!nombreMin.endsWith('.xls') && !nombreMin.endsWith('.xlsx')) continue;

    if (archivosProcesados.includes(nombreArchivo)) {
      archivosIgnoradosPorDuplicado++;
      continue;
    }

    let tempId = null;
    try {
      // Conversión a Google Sheets (sintaxis Drive v3, igual que el resto del proyecto)
      const tempFile = Drive.Files.create(
        { name: 'TEMP_' + nombreArchivo, mimeType: MimeType.GOOGLE_SHEETS },
        archivo.getBlob()
      );
      tempId = tempFile.id;

      const datosOrigen = leerHojaEstable_(tempId);
      const nuevasFilas = [];
      let duplicadas = 0;
      let sinRecibo = 0;

      for (let i = FILA_INICIO_DATOS - 1; i < datosOrigen.length; i++) {
        const fila = datosOrigen[i];
        const clave = claveRecibo_(fila[COL_RECIBO]);
        if (!clave) { sinRecibo++; continue; }                       // filas vacías / totales
        if (recibosExistentes.has(clave)) { duplicadas++; continue; } // ya cargado
        recibosExistentes.add(clave);
        nuevasFilas.push(fila);
      }

      if (nuevasFilas.length > 0) {
        const filaDestino = hojaIngresos.getLastRow() + 1;
        hojaIngresos.getRange(filaDestino, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
      }

      Logger.log(nombreArchivo + ': ' + nuevasFilas.length + ' filas agregadas, ' +
        duplicadas + ' duplicadas omitidas, ' + sinRecibo + ' sin Nº de recibo omitidas.');
      filasAgregadasTotal += nuevasFilas.length;
      filasDuplicadasTotal += duplicadas;
      filasSinReciboTotal += sinRecibo;

      registrarImportacionClubOnline(nombreArchivo);
      archivo.moveTo(carpetaProcesados);
      archivosProcesadosEnEstaTanda++;

    } catch (e) {
      Logger.log('Error procesando el archivo ' + nombreArchivo + ': ' + e.message);
    } finally {
      if (tempId) eliminarArchivoTemporal(tempId);
    }
  }

  if (archivosProcesadosEnEstaTanda > 0) {
    ss.toast('Ingresos: ' + filasAgregadasTotal + ' filas nuevas, ' + filasDuplicadasTotal +
      ' recibos repetidos omitidos.', '✅ Tanda completada');
  } else if (archivosIgnoradosPorDuplicado > 0) {
    ss.toast('No hay archivos nuevos. Se omitieron ' + archivosIgnoradosPorDuplicado + ' archivo(s) ya procesados.', 'ℹ️ Ingresos');
  }
}

/**
 * Limpieza única: borra de la hoja Ingresos las filas cuyo Nº de recibo (columna C)
 * ya apareció más arriba. Conserva siempre la primera aparición.
 * Correrla una vez desde el editor para sacar los duplicados ya cargados.
 */
function limpiarDuplicadosIngresos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const hoja = ss.getSheetByName('Ingresos');
  if (!hoja) return;
  const ultima = hoja.getLastRow();
  if (ultima < 2) return;

  const valoresC = hoja.getRange(1, 3, ultima, 1).getValues();
  const vistos = new Set();
  const aBorrar = []; // números de fila (1 = primera)
  valoresC.forEach((f, i) => {
    const clave = claveRecibo_(f[0]);
    if (!clave) return;
    if (vistos.has(clave)) aBorrar.push(i + 1); else vistos.add(clave);
  });

  // Borra de abajo hacia arriba, agrupando filas contiguas
  for (let k = aBorrar.length - 1; k >= 0; ) {
    let fin = aBorrar[k], ini = fin;
    while (k - 1 >= 0 && aBorrar[k - 1] === ini - 1) { ini--; k--; }
    hoja.deleteRows(ini, fin - ini + 1);
    k--;
  }
  Logger.log('Duplicados eliminados en Ingresos: ' + aBorrar.length);
  ss.toast('Se eliminaron ' + aBorrar.length + ' filas duplicadas por Nº de recibo.', 'Ingresos');
}

/** Normaliza el Nº de recibo para comparar (texto sin espacios de más). */
function claveRecibo_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim().replace(/\s+/g, ' ');
}