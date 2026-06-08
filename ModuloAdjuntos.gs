function clasificarAdjuntos() {
  // 1. CONDICIONES: Define qué palabra buscar en el archivo y a qué carpeta de Drive va
  // Puedes añadir todas las carpetas y palabras clave que quieras aquí abajo:
  var configuracion = [
    { palabraClave: "AntiguedaddeDeuda", nombreCarpeta: "AntiguedaddeDeuda" },
    { palabraClave: "Fichajes", nombreCarpeta: "Fichajes" },
    { palabraClave: "Socios",    nombreCarpeta: "Socios" }
  ];
  
  // 2. Busca correos recibidos en la última hora con archivos adjuntos
  // Ajusta "newer_than:1h" según la frecuencia con la que ejecutes el script
  var busqueda = 'has:attachment newer_than:1h'; 
  var hilos = GmailApp.search(busqueda);
  
  for (var i = 0; i < hilos.length; i++) {
    var mensajes = hilos[i].getMessages();
    for (var j = 0; j < mensajes.length; j++) {
      var adjuntos = mensajes[j].getAttachments();
      
      for (var k = 0; k < adjuntos.length; k++) {
        var adjunto = adjuntos[k];
        var nombreArchivo = adjunto.getName().toLowerCase(); // Convierte a minúsculas para evitar errores
        
        // 3. Revisa si el nombre del archivo coincide con alguna de tus palabras clave
        for (var c = 0; c < configuracion.length; c++) {
          var regla = configuracion[c];
          
          if (nombreArchivo.indexOf(regla.palabraClave.toLowerCase()) !== -1) {
            // Encuentra o crea la carpeta de destino
            var carpetas = DriveApp.getFoldersByName(regla.nombreCarpeta);
            var destino = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(regla.nombreCarpeta);
            
            // Evita duplicados
            var archivosExistentes = destino.getFilesByName(adjunto.getName());
            if (!archivosExistentes.hasNext()) {
              destino.createFile(adjunto);
              Logger.log("Guardado " + adjunto.getName() + " en " + regla.nombreCarpeta);
            }
            break; // Ya encontró su carpeta, pasa al siguiente archivo
          }
        }
      }
    }
  }
}