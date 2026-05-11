const path = require("path");
const XLSX = require("xlsx");
const { coerceText, normalizeKey } = require("./utils");

function loadSpreadsheet(spreadsheetFile, config) {
  if (!spreadsheetFile || !spreadsheetFile.buffer) {
    throw new Error("A spreadsheet file is required.");
  }

  const workbook = readWorkbook(spreadsheetFile);
  const requestedSheetName = config.spreadsheet.sheetName;
  const sheetName = workbook.SheetNames.includes(requestedSheetName) ? requestedSheetName : workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("The spreadsheet does not contain any sheets.");
  }

  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    defval: "",
    raw: false,
    blankrows: false,
  });

  const headerRows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: false,
    blankrows: false,
  });

  const headers = (headerRows[0] || []).map((header) => coerceText(header)).filter(Boolean);
  const normalizedRows = rows
    .map((row, index) => decorateRow(row, index + 2))
    .filter((row) => (config.spreadsheet.skipEmptyRows ? !row.__isEmpty : true));

  return {
    headers,
    rows: normalizedRows,
    sheetName,
    workbookSheetNames: workbook.SheetNames,
  };
}

function readWorkbook(spreadsheetFile) {
  const extension = path.extname(spreadsheetFile.originalname || "").toLowerCase();

  if (extension === ".csv") {
    const csvText = decodeCsvBuffer(spreadsheetFile.buffer);
    return XLSX.read(csvText, { type: "string" });
  }

  return XLSX.read(spreadsheetFile.buffer, { type: "buffer" });
}

function decodeCsvBuffer(buffer) {
  if (!buffer || buffer.length === 0) {
    return "";
  }

  if (hasBom(buffer, [0xff, 0xfe])) {
    return new TextDecoder("utf-16le").decode(buffer.subarray(2));
  }

  if (hasBom(buffer, [0xfe, 0xff])) {
    return new TextDecoder("utf-16be").decode(buffer.subarray(2));
  }

  const utf8Text = stripBom(buffer.toString("utf8"));
  if (!utf8Text.includes("\uFFFD")) {
    return utf8Text;
  }

  return stripBom(new TextDecoder("windows-1254").decode(buffer));
}

function hasBom(buffer, signature) {
  return signature.every((byte, index) => buffer[index] === byte);
}

function stripBom(text) {
  return String(text || "").replace(/^\uFEFF/, "");
}

function decorateRow(row, rowNumber) {
  const normalizedRow = {};
  const lookup = {};
  let isEmpty = true;

  for (const [key, value] of Object.entries(row)) {
    const text = coerceText(value);
    normalizedRow[key] = text;
    lookup[normalizeKey(key)] = text;

    if (text !== "") {
      isEmpty = false;
    }
  }

  normalizedRow.__lookup = lookup;
  normalizedRow.__rowNumber = rowNumber;
  normalizedRow.__isEmpty = isEmpty;
  return normalizedRow;
}

module.exports = {
  loadSpreadsheet,
};
