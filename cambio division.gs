function guardarCambioDivision() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var interfaz = sheet.getSheetByName("Formulario");
  var destino = sheet.getSheetByName("Cambio Division");
  
  // Obtener los datos del formulario
  var jugador = interfaz.getRange("C3").getValue();
  var divActual = interfaz.getRange("C5").getValue();
  var divNueva = interfaz.getRange("C7").getValue();
  
  // Validar que los campos no estén vacíos
  if (jugador == "" || divNueva == "") {
    SpreadsheetApp.getUi().alert("Por favor, selecciona un jugador y su nueva división.");
    return;
  }
  
  // Registrar en la hoja Cambio Division (Fecha, Jugador, Div Anterior, Div Nueva)
  var fecha = new Date();
  destino.appendRow([fecha, jugador, divActual, divNueva]);
  
  // OPCIONAL: Si también quieres actualizar automáticamente la división en la hoja 'Reporte':
  var hojaReporte = sheet.getSheetByName("Reporte");
  var datosReporte = hojaReporte.getDataRange().getValues();
  for (var i = 1; i < datosReporte.length; i++) {
    if (datosReporte[i][0] == jugador) { // Asumiendo que el nombre está en la columna A
      hojaReporte.getRange(i + 1, 2).setValue(divNueva); // Cambia la columna B (2)
      break;
    }
  }
  
  // Limpiar el formulario
  interfaz.getRange("C3").clearContent();
  interfaz.getRange("C7").clearContent();
  
  SpreadsheetApp.getUi().alert("¡Cambio registrado y actualizado con éxito!");
}