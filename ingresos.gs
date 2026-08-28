function actualizarIngresos() {
  const TIEMPO_INICIO = new Date().getTime();
  const LIMITE_MS = 5 * 60 * 1000;

  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const CARPETA_INGRESOS_ID = '1hKRtI2E_kTJY_Q51eKckTQ_dTU0MBE_q';
  const NOMBRE_CARPETA_PROCESADOS = 'Procesados';
  const MAX_ARCHIVOS_POR_TANDA = 1;
  const FILA_INICIO_DATOS = 7;

  let hojaIngresos = ss.getSheetByName('Ingresos');
  if (!hojaIngresos) {
    hojaIngresos = ss.insertSheet('Ingresos');
  }

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
  if (subcarpetas.hasNext()) {
    carpetaProcesados = subcarpetas.next();
  } else {
    carpetaProcesados = carpeta.createFolder(NOMBRE_CARPETA_PROCESADOS);
  }

  const archivos = carpeta.getFiles();

  let archivosIgnoradosPorDuplicado = 0;
  let archivosProcesadosEnEstaTanda = 0;

  while (archivos.hasNext()) {
    if (new Date().getTime() - TIEMPO_INICIO > LIMITE_MS) {
      Logger.log('Se acerca el límite de tiempo. Cortando ejecución de forma segura.');
      break;
    }

    if (archivosProcesadosEnEstaTanda >= MAX_ARCHIVOS_POR_TANDA) break;

    const archivo = archivos.next();
    const nombreArchivo = archivo.getName().trim();

    // Si por algún motivo la subcarpeta "Procesados" aparece en la iteración, la salteamos
    if (archivo.getId() === carpetaProcesados.getId()) continue;

    if (archivosProcesados.includes(nombreArchivo)) {
      archivosIgnoradosPorDuplicado++;
      continue;
    }

    let archivoConvertidoId = null;

    try {
      const recursoConvertido = Drive.Files.copy(
        { title: 'TEMP_' + nombreArchivo, mimeType: MimeType.GOOGLE_SHEETS },
        archivo.getId(),
        { convert: true }
      );
      archivoConvertidoId = recursoConvertido.id;

      const ssTemp = SpreadsheetApp.openById(archivoConvertidoId);
      const hojaTemp = ssTemp.getSheets()[0];
      const ultimaFila = hojaTemp.getLastRow();
      const ultimaColumna = hojaTemp.getLastColumn();

      if (ultimaFila >= FILA_INICIO_DATOS) {
        const cantidadFilas = ultimaFila - FILA_INICIO_DATOS + 1;
        const datos = hojaTemp.getRange(FILA_INICIO_DATOS, 1, cantidadFilas, ultimaColumna).getValues();
        const filaDestino = hojaIngresos.getLastRow() + 1;
        hojaIngresos.getRange(filaDestino, 1, datos.length, datos[0].length).setValues(datos);
      } else {
        Logger.log(`El archivo ${nombreArchivo} no tiene datos a partir de la fila ${FILA_INICIO_DATOS}.`);
      }

      if (hojaControl) {
        hojaControl.appendRow([nombreArchivo, new Date()]);
      }

      // Mover el archivo original a "Procesados"
      archivo.moveTo(carpetaProcesados);

      archivosProcesadosEnEstaTanda++;

    } catch (e) {
      Logger.log(`Error procesando el archivo ${nombreArchivo}: ${e.message}`);
    } finally {
      if (archivoConvertidoId) {
        Drive.Files.remove(archivoConvertidoId);
      }
    }
  }

  if (archivosProcesadosEnEstaTanda > 0) {
    SpreadsheetApp.getUi().alert(`✅ Tanda completada: Se procesaron ${archivosProcesadosEnEstaTanda} archivo(s) correctamente.`);
  } else if (archivosIgnoradosPorDuplicado > 0) {
    SpreadsheetApp.getUi().alert(`ℹ️ No hay archivos nuevos. Se omitieron ${archivosIgnoradosPorDuplicado} archivo(s) ya procesados.`);
  }
}