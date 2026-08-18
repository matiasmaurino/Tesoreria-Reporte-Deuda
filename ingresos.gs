/**
 * Importa los archivos de la carpeta de Ingresos acumulándolos en la hoja "Ingresos"
 * a partir de la fila 7, leyendo de la columna A a la M (13 columnas).
 */
function actualizarIngresos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
  const CARPETA_INGRESOS_ID = '1wVdejUgqYbgLDN1J9m4Nbjp9nvikpaMG';
  
  let hojaIngresos = ss.getSheetByName('Ingresos');
  if (!hojaIngresos) {
    hojaIngresos = ss.insertSheet('Ingresos');
  }

  // Lista de control de importaciones
  const hojaControl = ss.getSheetByName(HOJA_CONTROL_IMPORTACIONES);
  let archivosProcesados = [];
  if (hojaControl && hojaControl.getLastRow() > 1) {
    archivosProcesados = hojaControl.getRange(2, 1, hojaControl.getLastRow() - 1, 1)
      .getValues()
      .map(row => String(row[0]).trim());
  }

  const carpeta = DriveApp.getFolderById(CARPETA_INGRESOS_ID);
  const archivos = carpeta.getFiles();

  let totalFilasImportadas = 0;
  let archivosIgnoradosPorDuplicado = 0;
  let archivosProcesadosEnEstaTanda = 0;

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName().trim();

    // Validar duplicado
    if (archivosProcesados.includes(nombreArchivo)) {
      archivosIgnoradosPorDuplicado++;
      continue;
    }

    try {
      let datosArchivo = [];
      const mimeType = archivo.getMimeType();

      if (mimeType === MimeType.GOOGLE_SHEETS) {
        const tempSS = SpreadsheetApp.open(archivo);
        const tempSheet = tempSS.getSheets()[0];
        if (tempSheet.getLastRow() >= 7) {
          datosArchivo = tempSheet.getRange(7, 1, tempSheet.getLastRow() - 6, 13).getValues();
        }
      } else if (mimeType === MimeType.MICROSOFT_EXCEL || mimeType === MimeType.CSV) {
        const blob = archivo.getBlob();
        const config = { title: '[TEMP] ' + nombreArchivo, parents: [{ id: CARPETA_INGRESOS_ID }] };
        const tempFile = Drive.Files.insert(config, blob, { convert: true });
        const tempSS = SpreadsheetApp.openById(tempFile.id);
        const tempSheet = tempSS.getSheets()[0];

        if (tempSheet.getLastRow() >= 7) {
          datosArchivo = tempSheet.getRange(7, 1, tempSheet.getLastRow() - 6, 13).getValues();
        }

        // Eliminar archivo temporal
        Drive.Files.remove(tempFile.id);
      }

      if (datosArchivo.length > 0) {
        const valoresColumnaA = hojaIngresos.getRange("A:A").getValues();
        let ultimaFilaA = 6;
        for (let i = valoresColumnaA.length - 1; i >= 0; i--) {
          if (valoresColumnaA[i][0] !== "" && valoresColumnaA[i][0] !== null) {
            ultimaFilaA = Math.max(i + 1, 6);
            break;
          }
        }

        hojaIngresos.getRange(ultimaFilaA + 1, 1, datosArchivo.length, 13)
          .setValues(datosArchivo);

        registrarImportacionClubOnline(nombreArchivo);
        totalFilasImportadas += datosArchivo.length;
        archivosProcesadosEnEstaTanda++;
        Logger.log(`✅ Archivo "${nombreArchivo}" importado exitosamente (${datosArchivo.length} filas).`);
      }
    } catch (e) {
      Logger.log('Error procesando el archivo ' + nombreArchivo + ': ' + e.toString());
    }
  }

  // Alertas al finalizar
  if (totalFilasImportadas > 0) {
    SpreadsheetApp.getUi().alert(`✅ Importación completada.\nSe agregaron ${totalFilasImportadas} filas de ${archivosProcesadosEnEstaTanda} archivo(s).`);
  } else if (archivosIgnoradosPorDuplicado > 0 && archivosProcesadosEnEstaTanda === 0) {
    SpreadsheetApp.getUi().alert(`ℹ️ Los archivos de la carpeta ya figuran registrados en "ImportacionArchivosClubOnline".\nSi deseas reimportarlos, borra sus filas en esa pestaña.`);
  } else {
    SpreadsheetApp.getUi().alert(`⚠️ No se encontraron archivos o filas válidas desde la fila 7.`);
  }
}