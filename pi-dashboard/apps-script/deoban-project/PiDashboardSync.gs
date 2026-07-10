const SHEET_NAME = "더본외식Lowdata";
const HEADER_ROW = 1;

const COLS = {
  hold: "D",
  brand: "F",
  name: "G",
  bizNo: "I",
  ownerContact: "U",
  partner: "Y",
  installOwner: "Z",
  installDate: "AA",
  progress: "AE"
};

const EXCLUDE_HOLD_VALUES = ["설치제외", "보류", "폐점", "9월 1일 이후"];

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return jsonOutput({ error: "시트를 찾을 수 없습니다: " + SHEET_NAME, stores: [] });
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const numRows = Math.max(lastRow - HEADER_ROW, 0);
  const dataRows = sheet.getRange(HEADER_ROW + 1, 1, numRows, lastCol).getValues();

  const idx = {};
  Object.keys(COLS).forEach(function (key) {
    idx[key] = colLetterToIndex(COLS[key]) - 1;
  });

  const stores = [];
  dataRows.forEach(function (row) {
    const hold = String(row[idx.hold] || "").trim();
    if (EXCLUDE_HOLD_VALUES.indexOf(hold) !== -1) return;

    const name = row[idx.name];
    if (!name) return;

    stores.push({
      brand: String(row[idx.brand] || "").trim() || "더본외식",
      name: String(name).trim(),
      bizNo: String(row[idx.bizNo] || "").trim(),
      ownerContact: String(row[idx.ownerContact] || ""),
      partner: String(row[idx.partner] || ""),
      installOwner: String(row[idx.installOwner] || ""),
      installDate: formatDate(row[idx.installDate]),
      progress: String(row[idx.progress] || "")
    });
  });

  return jsonOutput({ stores: stores });
}

function colLetterToIndex(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col;
}

function formatDate(v) {
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return v ? String(v) : "";
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
