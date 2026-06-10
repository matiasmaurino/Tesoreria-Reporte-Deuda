function clasificarAdjuntos() {
  // 1. CONDICIONES: Define qué palabra buscar en el archivo y a qué carpeta de Drive va
  var configuracion = [
    { palabraClave: "AntiguedaddeDeuda", nombreCarpeta: "AntiguedaddeDeuda" },
    { palabraClave: "Fichajes", nombreCarpeta: "Fichajes" },
    { palabraClave: "Socios",    nombreCarpeta: "Socios" }
  ];
  
  // 2. Busca correos recibidos en la última hora con archivos adjuntos
  // Al ejecutar manualmente, si hace más de 1 hora que no te llega un mail, la lista estará vacía
  var busqueda = 'has:attachment newer_than:1h';
  var hilos = GmailApp.search(busqueda);
  
  Logger.log("=== INICIANDO VERIFICACIÓN MANUAL DE CORREOS ===");
  Logger.log("Hilos de correo detectados en la última hora: " + hilos.length);
  
  if (hilos.length === 0) {
    Logger.log("Aviso: No hay correos nuevos con adjuntos en la última hora. La función terminó de forma segura sin procesar nada.");
    return; // Evita que el script continúe y falle por falta de datos
  }
  
  var archivosProcesadosContador = 0;
  
  for (var i = 0; i < hilos.length; i++) {
    var mensajes = hilos[i].getMessages();
    for (var j = 0; j < mensajes.length; j++) {
      var adjuntos = mensajes[j].getAttachments();
      for (var k = 0; k < adjuntos.length; k++) {
        var adjunto = adjuntos[k];
        var nombreArchivo = adjunto.getName().toLowerCase();
        
        // 3. Revisa si el nombre del archivo coincide con alguna de tus palabras clave
        for (var c = 0; c < configuracion.length; c++) {
          var regla = configuracion[c];
          if (nombreArchivo.indexOf(regla.palabraClave.toLowerCase()) !== -1) {
            
            // Intenta buscar o crear la carpeta de forma segura
            var destino;
            try {
              var carpetas = DriveApp.getFoldersByName(regla.nombreCarpeta);
              destino = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(regla.nombreCarpeta);
            } catch (errCarpeta) {
              Logger.log("Error crítico al acceder o crear la carpeta '" + regla.nombreCarpeta + "': " + errCarpeta.message);
              continue;
            }
            
            // Evita duplicados
            var archivosExistentes = destino.getFilesByName(adjunto.getName());
            if (!archivosExistentes.hasNext()) {
              destino.createFile(adjunto);
              Logger.log("¡ÉXITO! Guardado con seguridad: " + adjunto.getName() + " -> En carpeta: " + regla.nombreCarpeta);
              archivosProcesadosContador++;
            } else {
              Logger.log("Omitido: El archivo '" + adjunto.getName() + "' ya existe en la carpeta destino.");
            }
            break;
          }
        }
      }
    }
  }
  Logger.log("=== FIN DEL PROCESO === Total de archivos nuevos guardados en Drive: " + archivosProcesadosContador);
}