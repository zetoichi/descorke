/*
  1. Creá una hoja de cálculo y copiá su ID de la URL.
  2. En Apps Script: Configuración del proyecto > Propiedades de la secuencia
     de comandos. Agregá FORM_SHEETS_ID con el ID de la hoja.
  3. Pegá este archivo e implementalo como Aplicación web.
     Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.
  4. Copiá la URL terminada en /exec en CLUB.formEndpoint dentro de index.html.
  5. Para actualizarlo: Gestionar implementaciones > Editar > Nueva versión > Implementar.
*/

const FORM_SHEETS_ID = PropertiesService.getScriptProperties().getProperty("FORM_SHEETS_ID");
const SHEET_NAME = "Respuestas";
const NOTIFICATION_EMAILS = [
  "sampedrochristian@yahoo.com.ar",
  "celeste.credidio@gmail.com",
  "alejandro.zapiola@gmail.com",
];
const HEADERS = [
  "Fecha",
  "Nombre",
  "Teléfono",
  "Email",
  "Membresía",
  "Relación con el vino",
  "Tipos de vino",
  "No le gusta",
  "Ocasiones",
  "Vinos favoritos",
  "Estado del aviso",
];

function doPost(e) {
  const values = e && e.parameter ? e.parameter : {};
  if (values.website) return output("ok");

  const email = String(values.email || "").trim();
  if (!values.nombre || !values.telefono || !/^\S+@\S+\.\S+$/.test(email)) {
    return output("invalid");
  }

  if (!FORM_SHEETS_ID) throw new Error("Falta configurar la Script Property FORM_SHEETS_ID");

  const spreadsheet = SpreadsheetApp.openById(FORM_SHEETS_ID);
  const sheet =
    spreadsheet.getSheetByName(SHEET_NAME) ||
    spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else if (!sheet.getRange(1, HEADERS.length).getValue()) {
    sheet.getRange(1, HEADERS.length).setValue(HEADERS[HEADERS.length - 1]);
  }

  sheet.appendRow([
    new Date(),
    safeCell(values.nombre, 120),
    safeCell(values.telefono, 60),
    safeCell(email, 254),
    safeCell(values.membresia, 80),
    safeCell(values.relacion, 500),
    safeCell(values.tipos, 500),
    safeCell(values.no_gusta, 1000),
    safeCell(values.ocasiones, 500),
    safeCell(values.favoritos, 1000),
    "Pendiente",
  ]);
  const statusCell = sheet.getRange(sheet.getLastRow(), HEADERS.length);

  try {
    sendNotification(values, email);
    statusCell.setValue("Enviado");
  } catch (error) {
    const message = error && error.message ? error.message : String(error);
    statusCell.setValue(`Error: ${message}`.slice(0, 500));
    SpreadsheetApp.flush();
    console.error("La respuesta se guardó, pero no se pudo enviar el aviso", error);
    throw error;
  }

  return output("ok");
}

function sendNotification(values, email) {
  const name = cleanText(values.nombre, 120);
  const subject = `[La Orden] Nueva respuesta de ${name.replace(/[\r\n]+/g, " ")}`;
  const body = [
    "Se recibió una nueva respuesta del formulario:",
    "",
    `Nombre: ${name}`,
    `Teléfono: ${cleanText(values.telefono, 60)}`,
    `Email: ${email}`,
    `Membresía: ${cleanText(values.membresia, 80)}`,
    `Relación con el vino: ${cleanText(values.relacion, 500)}`,
    `Tipos de vino: ${cleanText(values.tipos, 500)}`,
    `No le gusta: ${cleanText(values.no_gusta, 1000)}`,
    `Ocasiones: ${cleanText(values.ocasiones, 500)}`,
    `Vinos favoritos: ${cleanText(values.favoritos, 1000)}`,
  ].join("\n");

  MailApp.sendEmail(NOTIFICATION_EMAILS.join(","), email, subject, body);
}

function cleanText(value, maxLength) {
  return String(value || "")
    .trim()
    .slice(0, maxLength);
}

function safeCell(value, maxLength) {
  const text = cleanText(value, maxLength);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function output(message) {
  return ContentService.createTextOutput(message);
}

function testSafeCell() {
  if (safeCell("=1+1", 100) !== "'=1+1")
    throw new Error("Formula sanitization failed");
  if (safeCell(" Malbec ", 100) !== "Malbec")
    throw new Error("Whitespace trimming failed");
}
