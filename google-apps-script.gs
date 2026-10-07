/*
  1. Creá una hoja de cálculo y copiá su ID de la URL.
  2. En la hoja: Extensiones > Apps Script. Pegá este archivo y reemplazá SPREADSHEET_ID.
  3. Implementar > Nueva implementación > Aplicación web.
     Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.
  4. Copiá la URL terminada en /exec en CLUB.formEndpoint dentro de index.html.
*/

const SPREADSHEET_ID = "17nkyZ9FcBoqYQsEe7tNhK88YiwXIrA50RE2s-UNfsog";
const SHEET_NAME = "Respuestas";
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
];

function doPost(e) {
  const values = e && e.parameter ? e.parameter : {};
  if (values.website) return output("ok");

  const email = String(values.email || "").trim();
  if (!values.nombre || !values.telefono || !/^\S+@\S+\.\S+$/.test(email)) {
    return output("invalid");
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet =
    spreadsheet.getSheetByName(SHEET_NAME) ||
    spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

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
  ]);

  return output("ok");
}

function safeCell(value, maxLength) {
  const text = String(value || "")
    .trim()
    .slice(0, maxLength);
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
