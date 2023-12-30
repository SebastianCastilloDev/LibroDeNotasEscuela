function libroDeNotas() {
  generarAsignatura()
}

//alumnos, asignaturas, calificaciones, profesora, curso
function generarAsignatura() {
  const datos = data()

  for (let i = 0; i < datos.asignaturas.length; i++) {
    generarNuevaHoja(datos.asignaturas[i])
    template(datos.asignaturas[i], datos.colores[i], datos)
  }
}

function validacion(datos) {
  let validado = false
  if (datos.nombres == []) {
    Browser.msgBox('Advertencia', 'Nombres de alumnos vacío.', Browser.Buttons.OK)
  }
  if (datos.apPaterno == []) {
    Browser.msgBox('Advertencia', 'Apellido paterno de alumnos vacío.', Browser.Buttons.OK)
  }
  if (datos.apMaterno == []) {
    Browser.msgBox('Advertencia', 'Apellido materno de alumnos vacío.', Browser.Buttons.OK)
  }
  if (datos.rut == []) {
    Browser.msgBox('Advertencia', 'R.U.T. alumnos vacío.', Browser.Buttons.OK)
  }
  if (datos.asignaturas == []) {
    Browser.msgBox('Advertencia', 'Asignaturas vacío.', Browser.Buttons.OK)
  }
  if (datos.profesora === []) {
    Browser.msgBox('Advertencia', 'Nombre profesora vacío.', Browser.Buttons.OK)
  }
  if (datos.curso == []) {
    Browser.msgBox('Advertencia', 'Nombre curso vacío.', Browser.Buttons.OK)
  }
  if (datos.calificaciones == []) {
    Browser.msgBox('Advertencia', 'Número de calificaciones vacío.', Browser.Buttons.OK)
  }

  if (datos.nombres != [] &&
    datos.apPaterno != [] &&
    datos.apMaterno != [] &&
    datos.rut != [] &&
    datos.asignaturas != [] &&
    datos.profesora != [] &&
    datos.curso != [] &&
    !isNaN(datos.calificaciones[0])
  ) {

    validado = true
  }

  return validado
}

function template(asignatura, color, datos) {

  const parametros = configuracion()

  const profesora = datos.profesora[0]
  const curso = datos.curso[0]
  const calificaciones = datos.calificaciones[0]
  const alumnos = nombreCompleto(datos)
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const hoja = ss.getSheetByName(asignatura)

  cabeceraTemplate(hoja, asignatura, profesora, curso)
  colorCeldaAsignatura(hoja, color)
  nombresAlumnosTemplate(hoja, parametros, alumnos)
  actividadTemplate(hoja, alumnos, parametros)
  notasTemplate(hoja, calificaciones, alumnos, parametros, asignatura)
  formatoCondicional(calificaciones, alumnos)
}

// Poner color a una celda
function ponerColorCelda(celda, color) {
  celda.setBackground(color)
}

function colorPestana(pestana, color) {
  spreadsheet.getSheetByName(pestana).setTabColor(color);
}

function rangosA1(fila, columna) {
  ss = SS()
  hoja = ss.getActiveSheet()
}

function obtenerNombresHojas() {
  const ss = SS()
  const hojas = ss.getSheets()
  const nombresHojas = []
  for (let i = 0; i < hojas.length; i++) {
    nombresHojas.push(hojas[i].getName())
  }
  return (nombresHojas)
}

function obtenerColorCelda() {
  const ss = SpreadsheetApp.openById('19mfnLjE0CiATe5-3qjnRJKK1Xlv5xgdLLU30QyewmLs')
  const nombresHojas = ["LENGUAJE", "MATEMÁTICAS", "HISTORIA", "CIENCIAS NATURALES", "TECNOLOGÍA", "MÚSICA", "ARTES VISUALES", "EDUCACIÓN FÍSICA", "RELIGIÓN", "ORIENTACIÓN"]

  colores = []

  for (let i = 0; i < nombresHojas.length; i++) {
    colores.push(ss.getSheetByName(nombresHojas[i]).getRange('A2').getBackground())
  }

  return (colores)
}

function nombreYcolor() {
  const nombres = ["LENGUAJE", "MATEMÁTICAS", "HISTORIA", "CIENCIAS NATURALES", "TECNOLOGÍA", "MÚSICA", "ARTES VISUALES", "EDUCACIÓN FÍSICA", "RELIGIÓN", "ORIENTACIÓN"]
  const colores = obtenerColorCelda()

  const nombre_color = []

  for (let i = 0; i < nombres.length; i++) {
    nombre_color.push({
      nombreHoja: nombres[i],
      color: colores[i]
    })
  }
}

function formatoCondicional(calificaciones, alumnos) {
  const hoja = SpreadsheetApp.getActiveSheet()
  let conditionalFormatRules = hoja.getConditionalFormatRules();
  Logger.log(conditionalFormatRules)
  Logger.log(hoja)

  rango1 = hoja.getRange(9, calificaciones + 2, alumnos.length, 1)
  rango2 = hoja.getRange(9, 2 * calificaciones + 3, alumnos.length, 1)
  rango3 = hoja.getRange(9, 2 * calificaciones + 4, alumnos.length, 1)

  conditionalFormatRules = hoja.getConditionalFormatRules();

  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([rango1, rango2, rango3])
    .whenNumberGreaterThanOrEqualTo(4)
    .setBold(true)
    .setFontColor('#0000FF')
    .build());
  hoja.setConditionalFormatRules(conditionalFormatRules);

  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([rango1, rango2, rango3])
    .whenNumberLessThan(4)
    .setBold(true)
    .setFontColor('#FF0000')
    .build());
  hoja.setConditionalFormatRules(conditionalFormatRules);

}

function SS() {
  //return SpreadsheetApp.openById(sheetId())
  return SpreadsheetApp.getActiveSpreadsheet()
}

function generarNuevaHoja(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.insertSheet(nombre, ss.getNumSheets());
}

function nombreCompleto(data) {
  const nombresCompletos = []
  for (let i = 0; i < data.nombres.length; i++) {
    nombresCompletos.push(data.apPaterno[i] + " " + data.apMaterno[i] + " " + data.nombres[i])
  }
  return nombresCompletos.sort()
}

function limpiarLibro() {
  const ss = SS()
  const confirmar = Browser.inputBox('Ingrese la palabra "Borrar" para eliminar todas las hojas del libro')
  if (confirmar === "Borrar") {
    const hojas = ss.getSheets()
    for (let i = 0; i < hojas.length; i++) {
      if (hojas[i].getName() != "Datos") {
        const hoja = ss.getSheetByName(hojas[i].getName())
        ss.deleteSheet(hoja)
      }
    }
  }
}

function cabeceraTemplate(hoja, asignatura, profesora, curso, colores) {
  hoja.getRange("A2").setValue(asignatura)
  hoja.getRange("A4").setValue("PROFESORA")
  hoja.getRange("B4").setValue(profesora)
  hoja.getRange("A5").setValue("CURSO")
  hoja.getRange("B5").setValue(curso)
}

function nombresAlumnosTemplate(hoja, parametros, alumnos) {
  const FILA_INICIO_ALUMNOS = 9
  hoja.getRange("A8").setValue("APELLIDOS NOMBRE")
  hoja.getRange("A8").setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)
  hoja.setColumnWidth(1, parametros.anchoNombre)
  for (let i = 0; i < alumnos.length; i++) {
    hoja.getRange(FILA_INICIO_ALUMNOS + i, 1).setValue(i + 1 + ". " + alumnos[i])
    hoja.getRange(FILA_INICIO_ALUMNOS + i, 1).setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)
  }
}

function actividadTemplate(hoja, alumnos, parametros, calificaciones) {
  const FILA_INICIO_ALUMNOS = 9
  hoja.getRange(FILA_INICIO_ALUMNOS + alumnos.length, 1).setValue("ACTIVIDAD - FECHA")
  hoja.getRange(FILA_INICIO_ALUMNOS + alumnos.length, 1).setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)
  hoja.setRowHeight(FILA_INICIO_ALUMNOS + alumnos.length, parametros.alturaActividadFecha)

  for (j = 0; j < 2 * calificaciones + 3; j++) {
    hoja.getRange(FILA_INICIO_ALUMNOS + alumnos.length, j + 2).setTextRotation(90)
  }
}

function notasTemplate(hoja, calificaciones, alumnos, parametros) {
  const FILA_INICIO_ALUMNOS = 9

  //Generar lista fila de notas primer semestre
  for (let j = 0; j < calificaciones; j++) {
    hoja.getRange(8, 2 + j).setValue("N" + (j + 1))
    hoja.setColumnWidth(2 + j, parametros.anchoNota)
    hoja.getRange(8, 2 + j).setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)
  }

  // Formato promedio primer semestre
  hoja.getRange(8, 2 + calificaciones).setValue('X')
  hoja.getRange(8, 2 + calificaciones).setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)
  hoja.getRange(8, 2 + calificaciones).setFontWeight('bold')
  hoja.setColumnWidth(2 + calificaciones, parametros.anchoNota)

  // Incorporar formula de promedios primer semestre
  for (let i = 0; i < alumnos.length; i++) {
    const celdaInicio = hoja.getRange(FILA_INICIO_ALUMNOS + i, 2).getA1Notation()
    const celdaFin = hoja.getRange(FILA_INICIO_ALUMNOS + i, calificaciones + 1).getA1Notation()
    const formula = '=IFERROR(IF(AVERAGE(' + celdaInicio + ':' + celdaFin + ')>' + parametros.notaMaxima + ',"NOTA MAYOR A 7",IF(AVERAGE(' + celdaInicio + ':' + celdaFin + ')<' + parametros.notaMinima + ',"NOTA MENOR A 2",AVERAGE(' + celdaInicio + ':' + celdaFin + '))),"SIN NOTAS")'
    hoja.getRange(FILA_INICIO_ALUMNOS + i, calificaciones + 2).setFormula(formula)
    hoja.getRange(FILA_INICIO_ALUMNOS + i, calificaciones + 2).setFontWeight('bold')
  }
  // Generar lista fila de notas segundo semestre
  for (let j = 0; j < calificaciones; j++) {
    hoja.getRange(8, 2 + j + 1 + calificaciones).setValue("N" + (j + 1))
    hoja.getRange(8, 2 + j + 1 + calificaciones).setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)
    hoja.setColumnWidth(2 + j + 1 + calificaciones, parametros.anchoNota)
  }
  // Formato promedio segundo semestre
  hoja.getRange(8, 2 + 2 * calificaciones + 1).setValue('X')
  hoja.getRange(8, 2 + 2 * calificaciones + 1).setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)
  hoja.getRange(8, 2 + 2 * calificaciones + 1).setFontWeight('bold')
  hoja.setColumnWidth(2 + 2 * calificaciones + 1, parametros.anchoNota)

  // Incorporar formula de promedios segundo semestre
  for (let i = 0; i < alumnos.length; i++) {
    const celdaInicio = hoja.getRange(FILA_INICIO_ALUMNOS + i, 3 + calificaciones).getA1Notation()
    const celdaFin = hoja.getRange(FILA_INICIO_ALUMNOS + i, 2 * calificaciones + 2).getA1Notation()
    const formula = '=IFERROR(IF(AVERAGE(' + celdaInicio + ':' + celdaFin + ')>' + parametros.notaMaxima + ',"NOTA MAYOR A 7",IF(AVERAGE(' + celdaInicio + ':' + celdaFin + ')<' + parametros.notaMinima + ',"NOTA MENOR A 2",AVERAGE(' + celdaInicio + ':' + celdaFin + '))),"SIN NOTAS")'
    hoja.getRange(FILA_INICIO_ALUMNOS + i, 2 * calificaciones + 3).setFormula(formula)
    hoja.getRange(FILA_INICIO_ALUMNOS + i, 2 * calificaciones + 3).setFontWeight('bold')
  }

  // Formato promedio final
  hoja.getRange(8, 2 + 2 * calificaciones + 2).setValue('X FINAL')
  hoja.getRange(8, 2 + 2 * calificaciones + 2).setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)
  hoja.getRange(8, 2 + 2 * calificaciones + 2).setFontWeight('bold')
  hoja.setColumnWidth(2 + 2 * calificaciones + 2, parametros.anchoPromedioFinal)

  // Incorporar formula promedio final
  for (let i = 0; i < alumnos.length; i++) {
    const celdaPromedio1S = hoja.getRange(FILA_INICIO_ALUMNOS + i, calificaciones + 2).getA1Notation()
    const celdaPromedio2S = hoja.getRange(FILA_INICIO_ALUMNOS + i, 2 * calificaciones + 3).getA1Notation()
    const formula = '=IFERROR(IF(AND(ISNUMBER(' + celdaPromedio1S + '),ISNUMBER(' + celdaPromedio2S + ')),AVERAGE(' + celdaPromedio1S + ',' + celdaPromedio2S + '),""),"")'
    hoja.getRange(FILA_INICIO_ALUMNOS + i, 2 * calificaciones + 4).setFormula(formula)
    hoja.getRange(FILA_INICIO_ALUMNOS + i, 2 * calificaciones + 4).setFontWeight('bold')
  }

  // Configurando la celda "PRIMER SEMESTRE"
  hoja.getRange(7, 2).setValue("PRIMER SEMESTRE")
  hoja.getRange(7, 2).setBackground('#fff2cc')
  hoja.getRange(7, 2, 1, calificaciones + 1).merge()
  hoja.getRange(7, 2, 1, calificaciones + 1).setValue('PRIMER SEMESTRE')
  hoja.getRange(7, 2, 1, calificaciones + 1).setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)

  // Configurando la celda "SEGUNDO SEMESTRE"
  hoja.getRange(7, 3 + calificaciones, 1, calificaciones + 1).setValue("SEGUNDO SEMESTRE")
  hoja.getRange(7, 3 + calificaciones, 1, calificaciones + 1).setBackground('#fff2cc')
  hoja.getRange(7, 3 + calificaciones, 1, calificaciones + 1).merge()
  hoja.getRange(7, 3 + calificaciones, 1, calificaciones + 1).setValue('SEGUNDO SEMESTRE')
  hoja.getRange(7, 3 + calificaciones, 1, calificaciones + 1).setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)

  for (let i = 0; i < alumnos.length + 1; i++) {
    for (let j = 0; j < 2 * calificaciones + 3; j++) {
      hoja.getRange(FILA_INICIO_ALUMNOS + i, j + 2).setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)
    }
  }

  //centrando los valores de las celdas horizontal y verticalmente
  hoja.getRange(7, 2, 3 + alumnos.length, 3 + 2 * calificaciones).setHorizontalAlignment("center")
  hoja.getRange(7, 2, 3 + alumnos.length, 3 + 2 * calificaciones).setVerticalAlignment("center")

  // Formato con dos decimales a las celdas notas primer semestre
  hoja.getRange(FILA_INICIO_ALUMNOS, 2, alumnos.length, calificaciones).setNumberFormat('0.00')
  // Formato con un decimal a los promedios primer semestre
  hoja.getRange(FILA_INICIO_ALUMNOS, calificaciones + 2, alumnos.length, 1).setNumberFormat('0.0')
  // Formato con dos decimales a las celdas notas segundo semestre
  hoja.getRange(FILA_INICIO_ALUMNOS, 3 + calificaciones, alumnos.length, calificaciones).setNumberFormat('0.00')
  // Formato con un decimal a los promedios segundo semestre
  hoja.getRange(FILA_INICIO_ALUMNOS, 2 * calificaciones + 3, alumnos.length, 1).setNumberFormat('0.0')
  // Formato con un decimal a los promedios finales
  hoja.getRange(FILA_INICIO_ALUMNOS, 2 * calificaciones + 4, alumnos.length, 1).setNumberFormat('0.0')
  Logger.log(hoja.getRange(FILA_INICIO_ALUMNOS, 2 * calificaciones + 3, alumnos.length, 1).getA1Notation())

}

function colorCeldaAsignatura(hoja, color) {
  hoja.getRange('A2').setBackground(color)
  hoja.setTabColor(color)
}

function data() {
  return {
    nombres: obtenerData(1),
    apPaterno: obtenerData(2),
    apMaterno: obtenerData(3),
    rut: obtenerData(4),
    genero: obtenerData(5),
    asignaturas: obtenerData(6),
    colores: obtenerColores(7),
    profesora: obtenerData(8),
    curso: obtenerData(9),
    calificaciones: obtenerData(10)
  }
}

function obtenerData(columna) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const hoja = ss.getSheetByName("Datos")
  let data = []
  let i = 4
  while (hoja.getRange(i, columna).getValue() != "") {
    data.push(hoja.getRange(i, columna).getValue())
    i++
  }
  return data
}

function obtenerColores(columna) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const hoja = ss.getSheetByName("Datos")
  let colores = []
  let i = 4
  while (hoja.getRange(i, columna - 1).getValue() != "") {
    colores.push(hoja.getRange(i, columna).getBackground())
    i++
  }
  return colores
}

//documento 
function generateDoc(docName) {
  const informeTemplate = DriveApp.getFileById(docConfig().templateId)
  const doc = informeTemplate.makeCopy()
  doc.setName(docName)
  const id = doc.getId()
  const url = doc.getUrl()
  return { id, url }
}

function openDocument(id) {
  return DocumentApp.openById(id)
}

function generateDocByStudent() {
  const carpetaInformes = DriveApp.getFolderById(docConfig().folderId)
  const alumnos = data()
  const nombres = alumnos.nombres
  const apPaterno = alumnos.apPaterno
  const apMaterno = alumnos.apMaterno
  const rut = alumnos.rut
  const genero = alumnos.genero
  const curso = alumnos.curso[0]

  const carpetaCurso = carpetaInformes.createFolder(curso + " " + new Date())
  for (let i = 0; i < nombres.length; i++) {
    const nombreCompleto = nombres[i] + " " + apPaterno[i] + " " + apMaterno[i]
    const nombreCompletoPatMatNom = apPaterno[i] + " " + apMaterno[i] + " " + nombres[i]
    const rutFormateado = formatoRut(rut[i].toString())
    const generoAlumno = genero[i]
    const archivoDoc = generateDoc(nombreCompleto)
    const idDoc = archivoDoc.id
    const urlDoc = archivoDoc.url

    DriveApp.getFileById(idDoc).moveTo(carpetaCurso)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos').getRange(4 + i, 11).setValue(urlDoc)

    buildDocument(idDoc, nombreCompleto, rutFormateado, generoAlumno, curso, nombreCompletoPatMatNom)
  }
}

function generateDocByStudent2() {
  const carpetaInformes   = DriveApp.getFolderById(docConfig().folderId)
  const alumnos           = data()
  const nombres           = alumnos.nombres
  const apPaterno         = alumnos.apPaterno
  const apMaterno         = alumnos.apMaterno
  const rut               = alumnos.rut
  const genero            = alumnos.genero
  const curso             = alumnos.curso[0]

  const carpetaCurso      = carpetaInformes.createFolder(curso + " " + new Date())

  for (let i = 0; i < nombres.length; i++) {
    const nombreCompleto          = nombres[i] + " " + apPaterno[i] + " " + apMaterno[i]
    const nombreCompletoPatMatNom = apPaterno[i] + " " + apMaterno[i] + " " + nombres[i]
    const rutFormateado           = formatoRut(rut[i].toString())
    const generoAlumno            = genero[i]
    const archivoDoc              = generateDoc(nombreCompleto)
    const idDoc                   = archivoDoc.id
    const urlDoc                  = archivoDoc.url

    DriveApp.getFileById(idDoc).moveTo(carpetaCurso)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos').getRange(4 + i, 11).setValue(urlDoc)

    buildDocument2(idDoc, nombreCompleto, rutFormateado, generoAlumno, curso, nombreCompletoPatMatNom)
  }
}


function buildDocument(id, nombreCompleto, rutFormateado, generoAlumno, curso, nombreCompletoPatMatNom) {
  const doc = openDocument(id)
  const body = doc.getBody()

  // Tamaño oficio
  const pageSize = {}
  pageSize[DocumentApp.Attribute.PAGE_HEIGHT] = 28.35 * 33
  pageSize[DocumentApp.Attribute.PAGE_WIDTH]  = 28.35 * 21.59
  body.setAttributes(pageSize)

  body.setMarginBottom(28.35)

  const titleTemplate = "INFORME DE NOTAS"
  const p1Template = "El  Establecimiento Educacional  N º 1788 “EDEN” San Ramón, reconocido oficialmente por el Ministerio de Educación, según Resolución Exenta N°4291 de 2005, R.B.D 25.417-7 otorga el presente Informe Educacional a:"
  p2Template = generoAlumno == "H" ? 
                              "DON: " + nombreCompleto.toUpperCase() + ", R.U.N.: " + rutFormateado 
                              : 
                              "DOÑA: " + nombreCompleto.toUpperCase() + " R.U.N.: " + rutFormateado

  const p3Template = "Estudiante de " + curso + " que de acuerdo al Plan y Programa de Estudio aprobado por Decreto Supremo 83/2015 y 67/2020 y su Reglamento de Evaluación y Promoción Escolar ha obtenido la siguiente evaluación semestral parcial y final que se indica:"

  resetBody(body)

  const title = body.appendParagraph(titleTemplate)
  const p1 = body.appendParagraph(p1Template)
  const p2 = body.appendParagraph(p2Template)
  const p3 = body.appendParagraph(p3Template)

  title.setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)

  p1.setHeading(DocumentApp.ParagraphHeading.NORMAL)
    .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY)
    .setSpacingAfter(5)

  p2.setHeading(DocumentApp.ParagraphHeading.NORMAL)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setSpacingAfter(10)
    .setSpacingBefore(10)

  p3.setHeading(DocumentApp.ParagraphHeading.NORMAL)
    .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY)
    .setSpacingAfter(10)

  const datosTablaNotas = datosTablaNotasDocumento(nombreCompletoPatMatNom)

  const datosTablaResumen = [
    ['Situación Final', ""],
    ['% Asistencia', ""],
    ['Observaciones: ', ""]
  ]

  const tablaNotas = body.appendTable(datosTablaNotas);
  const tablaResumen = body.appendTable(datosTablaResumen);

  tablaNotas.setColumnWidth(0, 170)

  tablaResumen.setColumnWidth(0, 170)

  const p4 = body.appendParagraph('________________________________                      ________________________________')
  const p5 = body.appendParagraph('Timbre, Nombre, Apellido y Firma.                                Timbre, Nombre, Apellido y Firma.')
  const p6 = body.appendParagraph('Profesora                                                                          Jefa U.T.P.')

  const p7 = body.appendParagraph('________________________________')
  const p8 = body.appendParagraph('Timbre, Nombre, Apellido y Firma.')
  const p9 = body.appendParagraph('Coordinadora Directiva')


  p4.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setSpacingBefore(18)
  p5.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  p6.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  p7.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setSpacingBefore(18)
  p8.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  p9.setAlignment(DocumentApp.HorizontalAlignment.CENTER)

  const fecha = new Date()
  const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
  const p10 = body.appendParagraph('San Ramón, ' + meses[fecha.getMonth()] + " de " + fecha.getFullYear())

  p10.setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .setSpacingBefore(0)

}

function buildDocument2(id, nombreCompleto, rutFormateado, generoAlumno, curso, nombreCompletoPatMatNom) {
  const doc   = openDocument(id)
  const body  = doc.getBody()

  // Tamaño oficio
  const pageSize = {}
  pageSize[DocumentApp.Attribute.PAGE_HEIGHT]   = 28.35 * 33
  pageSize[DocumentApp.Attribute.PAGE_WIDTH]    = 28.35 * 21.59
  body.setAttributes(pageSize)
  body.setMarginBottom(28.35)

  const titleTemplate = "INFORME DE NOTAS"
  const p1Template = "El  Establecimiento Educacional  N º 1788 “EDEN” San Ramón, reconocido oficialmente por el Ministerio de Educación, según Resolución Exenta N°4291 de 2005, R.B.D 25.417-7 otorga el presente Informe Educacional a:"
  p2Template = generoAlumno == "H" ? "DON: " + nombreCompleto.toUpperCase() + ", R.U.N.: " + rutFormateado : "DOÑA: " + nombreCompleto.toUpperCase() + " R.U.N.: " + rutFormateado

  const p3Template = "Estudiante de " + curso + " que de acuerdo al Plan y Programa de Estudio aprobado por Decreto Supremo 83/2015 y 67/2020 y su Reglamento de Evaluación y Promoción Escolar ha obtenido la siguiente evaluación semestral parcial y final que se indica:"

  resetBody(body)

  const title = body.appendParagraph(titleTemplate)
  const p1 = body.appendParagraph(p1Template)
  const p2 = body.appendParagraph(p2Template)
  const p3 = body.appendParagraph(p3Template)

  title.setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)

  p1.setHeading(DocumentApp.ParagraphHeading.NORMAL)
    .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY)
    .setSpacingAfter(5)

  p2.setHeading(DocumentApp.ParagraphHeading.NORMAL)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setSpacingAfter(10)
    .setSpacingBefore(10)

  p3.setHeading(DocumentApp.ParagraphHeading.NORMAL)
    .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY)
    .setSpacingAfter(10)

  const datosTablaNotas = datosTablaNotasDocumento2(nombreCompletoPatMatNom)

  const datosTablaResumen = [
    ['Situación Final', ""],
    ['% Asistencia', ""],
    ['Observaciones: ', ""]
  ]

  const tablaNotas = body.appendTable(datosTablaNotas);
  const tablaResumen = body.appendTable(datosTablaResumen);

  tablaNotas.setColumnWidth(0, 170)

  tablaResumen.setColumnWidth(0, 170)

  const p4 = body.appendParagraph('________________________________                      ________________________________')
  const p5 = body.appendParagraph('Timbre, Nombre, Apellido y Firma.                                Timbre, Nombre, Apellido y Firma.')
  const p6 = body.appendParagraph('Profesora                                                                          Jefa U.T.P.')

  const p7 = body.appendParagraph('________________________________')
  const p8 = body.appendParagraph('Timbre, Nombre, Apellido y Firma.')
  const p9 = body.appendParagraph('Coordinadora Directiva')


  p4.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setSpacingBefore(18)
  p5.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  p6.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  p7.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setSpacingBefore(18)
  p8.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  p9.setAlignment(DocumentApp.HorizontalAlignment.CENTER)

  const fecha = new Date()
  const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
  const p10 = body.appendParagraph('San Ramón, ' + meses[fecha.getMonth()] + " de " + fecha.getFullYear())

  p10.setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .setSpacingBefore(0)

}


function datosTablaNotasDocumento2(alumno) {

  const FILA_INICIO_ALUMNOS   = 9
  const ss                    = SpreadsheetApp.getActiveSpreadsheet()
  const datos                 = data()
  const asignaturas           = datos.asignaturas
  const alumnos               = datos.nombres
  const numeroNotas           = datos.calificaciones[0]

  const alumnosCurso          = ss.getSheetByName("Lenguaje").getRange(FILA_INICIO_ALUMNOS,1,alumnos.length,1).getValues().flat()
  console.log(alumnosCurso)
  const indiceAlumno = alumnosCurso.findIndex(elemento => elemento.includes(alumno))

  const datosTablaNotas = []

  datosTablaNotas.push(['Subsector', 'X 1S', 'N1', 'N2', 'N3', 'N4', 'N5', 'N6', 'X 2S', 'X F'])

  let notas = []
  asignaturas.forEach(
    asignatura => {
      notas.push(
        ss
          .getSheetByName(asignatura)
          .getRange(FILA_INICIO_ALUMNOS, 2, alumnos.length, numeroNotas * 2 + 4) // aca vamos a seleccionar al alumno para poner su nota
          .getValues()
      )
    }
  )


  for (let i = 0; i < asignaturas.length; i++) {
    let notasPorAsignatura = notas[i][indiceAlumno].slice(6, notas[i][indiceAlumno].length - 1)
    fila = []
    fila.push(asignaturas[i])
    for (let i = 0; i < notasPorAsignatura.length; i++) {
      fila.push(notasPorAsignatura[i])
    }
    datosTablaNotas.push(fila)
  }

  for (let i = 0; i < datosTablaNotas.length; i++) {
    for (let j = 0; j < datosTablaNotas[i].length; j++) {
      console.log(datosTablaNotas[i][j])
      if (!isNaN(datosTablaNotas[i][j])) {
        datosTablaNotas[i][j] = Math.round(datosTablaNotas[i][j] * 10) / 10
      }
    }
  }

  let promedio1 = 0
  let promedio2 = 0
  let promedio3 = 0
  
  for (let i = 0; i < asignaturas.length; i++) {
    promedio1 += datosTablaNotas[i+1][1] 
    promedio2 += datosTablaNotas[i+1][8] 
    promedio3 += datosTablaNotas[i+1][9] 
  }
  promedio1 = Math.round(promedio1 / (datosTablaNotas.length-1) * 10) / 10
  promedio2 = Math.round(promedio2 / (datosTablaNotas.length-1) * 10) / 10
  promedio3 = Math.round(promedio3 / (datosTablaNotas.length-1) * 10) / 10

  datosTablaNotas.push(['Promedio Final', promedio1, '', '', '', '', '', '', promedio2, promedio3])

  console.log(datosTablaNotas)
  return datosTablaNotas
}

function datosTablaNotasDocumento(alumno) {
  const FILA_INICIO_ALUMNOS = 9
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const datos = data()
  const asignaturas = datos.asignaturas
  const alumnos = datos.nombres
  const numeroNotas = datos.calificaciones[0]

  const datosTablaNotas = []

  datosTablaNotas.push(['Subsector', 'N1', 'N2', 'N3', 'N4', 'N5', 'N6', 'X'])

  let promedio = 0
  let contadorPromedio = 0
  for (let j = 0; j < asignaturas.length; j++) {
    const hoja = ss.getSheetByName(asignaturas[j])
    const notas = []
    for (let i = 0; i < alumnos.length; i++) {
      if (hoja.getRange(FILA_INICIO_ALUMNOS + i, 1).getValue().includes(alumno)) {
        notas.push(asignaturas[j])
        //numero de notas + 1 para capturar la fila de promedio
        for (let k = 0; k < numeroNotas; k++) {
          if (isNaN(hoja.getRange(FILA_INICIO_ALUMNOS + i, k + 2).getValue()) || hoja.getRange(FILA_INICIO_ALUMNOS + i, k + 2).getValue() == 0) {
            notas.push("")
          } else {
            notas.push((hoja.getRange(FILA_INICIO_ALUMNOS + i, k + 2).getValue()).toFixed(1))
          }
        }

        let promedioParcial = hoja.getRange(FILA_INICIO_ALUMNOS + i, 8).getValue()
        if (typeof promedioParcial === "number") {
          if (hoja.getName() === "ORIENTACIÓN" || hoja.getName() === "RELIGIÓN") {
            notas.push(notaConcepto(Math.round(promedioParcial * 10) / 10))
          } else {
            promedioParcial = Math.round(promedioParcial * 10) / 10
            notas.push(promedioParcial)
            promedio += promedioParcial
            contadorPromedio += 1
          }
        } else {
          notas.push("")
        }
      }
    }
    datosTablaNotas.push(notas)
  }
  promedio /= contadorPromedio
  datosTablaNotas.push(['Promedio Final', "", "", "", "", "", "", promedio.toFixed(1)])
  return datosTablaNotas
}





















































function resetBody(body) {
  body.clear()
}

function formatoRut(rut) {
  // rutLimpio es el rut sin puntos ni guiones
  const rutLimpio = rut.replace(/\./g, '').replace(/\-/g, '').trim().toLowerCase();
  const lastDigit = rutLimpio.substr(-1, 1);
  const rutDigit = rutLimpio.substr(0, rutLimpio.length - 1)
  let rutFormateado = '';
  for (let i = rutDigit.length; i > 0; i--) {
    const e = rutDigit.charAt(i - 1);
    rutFormateado = e.concat(rutFormateado);
    if (i % 3 === 0) {
      rutFormateado = '.'.concat(rutFormateado);
    }
  }
  return rutFormateado.concat('-').concat(lastDigit);
}


function notaConcepto(nota) {

  if (nota >= 2 && nota < 4) {
    return "I"
  } else if (nota >= 4 && nota < 5) {
    return "S"
  } else if (nota >= 5 && nota < 6) {
    return "B"
  } else if (nota >= 6 && nota <= 7) {
    return "MB"
  } else {
    return "Error: valor fuera de rango"
  }
}

function conceptoNota(concepto) {
  if (concepto.toLowerCase() === "i") {
    return 3
  } else if (concepto.toLowerCase() === "s") {
    return 4.5
  } else if (concepto.toLowerCase() === "b") {
    return 5.5
  } else if (concepto.toLowerCase() === "mb") {
    return 6.5
  }
}

function configuracion() {
  const parametros = {
    anchoNombre: 250,
    anchoNota: 40,
    anchoPromedioFinal: 100,
    alturaFila: 21,
    alturaActividadFecha: 112,
    notaMinima: 2,
    notaMaxima: 7
  }
  return parametros
}

function docConfig(){
  return {
    id:'',
    folderId:'',
    templateId:''
  }
}
