function enviarCorreos() {
  // Obtener la hoja activa
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtener la hoja "Lista_Estudiantes"
  var hojaEstudiantes = hojaActiva.getSheetByName("Lista_Estudiantes");
  if (!hojaEstudiantes) {
    Logger.log("Error: No se encontró la hoja 'Lista_Estudiantes'. Por favor, revisa el nombre de la hoja.");
    return;  // Detener el script si la hoja no se encuentra
  }
  
  // Obtener la hoja "Visualizaciones"
  var hojaVisualizaciones = hojaActiva.getSheetByName("Visualizaciones");
  if (!hojaVisualizaciones) {
    Logger.log("Error: No se encontró la hoja 'Visualizaciones'. Por favor, revisa el nombre de la hoja.");
    return;  // Detener el script si la hoja no se encuentra
  }

  // Obtener los datos de los estudiantes (se asume que las primeras dos columnas tienen los datos)
  var datosEstudiantes = hojaEstudiantes.getRange(2, 1, hojaEstudiantes.getLastRow() - 1, 2).getValues();
  
  // Verifica si los datos de los estudiantes se obtuvieron correctamente
  if (datosEstudiantes.length === 0) {
    Logger.log("Error: No se encontraron datos en la hoja 'Lista_Estudiantes'.");
    return;
  }
  
  // Obtener todas las visualizaciones disponibles de la columna B
  var visualizaciones = hojaVisualizaciones.getRange(2, 2, hojaVisualizaciones.getLastRow() - 1, 1).getValues().flat(); // Cambiado a columna B (índice 2)
  
  // Verificar si las visualizaciones se obtuvieron correctamente
  if (visualizaciones.length < 2) {
    Logger.log("Error: No hay suficientes visualizaciones en la hoja 'Visualizaciones'. Se requieren al menos dos.");
    return;
  }

  // Iterar por cada estudiante de la prueba
  datosEstudiantes.forEach(function(fila) {
    var correoEstudiante = fila[1]; // Correo del estudiante
    var nombreEstudiante = fila[0]; // Nombre del estudiante

    // Seleccionar dos visualizaciones de manera aleatoria sin repetir
    var opcionIndices = [];
    while (opcionIndices.length < 2) {
      var index = Math.floor(Math.random() * visualizaciones.length);
      if (!opcionIndices.includes(index)) {
        opcionIndices.push(index);
      }
    }

    var opcionA = visualizaciones[opcionIndices[0]];
    var opcionB = visualizaciones[opcionIndices[1]];

    // Crear el contenido del correo
    var asunto = "Visualización de datos :: PEC1 :: opciones buena práctica";
    var mensaje = "Hola, " + nombreEstudiante + ",\n\n" +
      "OPCIONES PAC1: Te damos dos visualizaciones y debes elegir una para hacer el análisis tal y como está explicado en el enunciado de la PAC1.\n" +
      "Para la visualización mala puedes elegirla tú mismo.  Tus dos opciones son:\n\n" +
      "OPCIÓN A: " + opcionA + "\n\n" +
      "OPCIÓN B: " + opcionB + "\n\n" +
      "Recuerda que lo que valoramos es comentar todos los puntos que se piden, dar referencias de los elementos de visualización de las lecturas teóricas y adaptarse al tiempo solicitado.\n" +
      "Si tienes dudas, a partir del lunes, puedes ponerlas en el foro para que el resto de compañeros de curso puedan beneficiarse.\n\n" +
      "¡Buen trabajo!\nJosé Manuel";

    // Enviar el correo electrónico usando Gmail
    try {
      MailApp.sendEmail(correoEstudiante, asunto, mensaje);
      Logger.log("Correo enviado a: " + correoEstudiante);
    } catch (error) {
      Logger.log("Error al enviar el correo a: " + correoEstudiante + " - " + error);
    }
  });

  Logger.log("Proceso completado.");
}