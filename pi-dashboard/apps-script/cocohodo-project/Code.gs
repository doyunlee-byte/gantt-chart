const SHEET_NAME = "가맹점리스트 취합";
const HEADER_ROW = 4;

const HEADER_MAP = {
  name: "매장명",
  bizNo: "사업자번호",
  ownerContact: "대표자연락처",
  installOwner: "설치담당자",
  installDate: "설치예정일",
  progress: "운영관리등록여부"
};

function doGet(e) {
  const ss = SpreadsheetApp.openById("1GRfnkDylr6w4VLUU-Yecd71EUjm7V72KHlL7QoMKz4E");
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return jsonOutput({ error: "시트를 찾을 수 없습니다: " + SHEET_NAME, stores: [] });
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headerRowValues = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const colIndex = {};
  Object.keys(HEADER_MAP).forEach(function (key) {
    const headerText = HEADER_MAP[key];
    const idx = headerRowValues.findIndex(function (h) {
      return String(h).trim() === headerText;
    });
    colIndex[key] = idx;
  });

  const dataRows = sheet
    .getRange(HEADER_ROW + 1, 1, Math.max(lastRow - HEADER_ROW, 0), lastCol)
    .getValues();

  const stores = [];
  dataRows.forEach(function (row) {
    const name = colIndex.name >= 0 ? row[colIndex.name] : "";
    if (!name) return;

    stores.push({
      brand: "코코호도",
      name: String(name).trim(),
      bizNo: colIndex.bizNo >= 0 ? String(row[colIndex.bizNo]).trim() : "",
      ownerContact: colIndex.ownerContact >= 0 ? String(row[colIndex.ownerContact]) : "",
      installOwner: colIndex.installOwner >= 0 ? String(row[colIndex.installOwner]) : "",
      installDate: colIndex.installDate >= 0 ? formatDate(row[colIndex.installDate]) : "",
      progress: colIndex.progress >= 0 ? String(row[colIndex.progress]) : "",
      partner: ""
    });
  });

  return jsonOutput({ stores: stores });
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
