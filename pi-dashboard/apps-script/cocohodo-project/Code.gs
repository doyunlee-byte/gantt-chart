const SHEET_NAME = "가맹점리스트 취합";
const HEADER_ROW = 4;

const HEADER_MAP = {
  name: "매장명",
  bizNo: "사업자번호",
  ownerContact: "대표자연락처",
  installOwner: "설치담당자",
  installDate: "설치예정일",
  progress: "운영관리등록여부",
  orderType: "발주유형",
  storeId: "매장아이디",
  naverId: "네이버ID",
  posCount: "포스 대수",
  onlineBizReg: "전상등록여부",
  naverConnect: "네이버커넥트"
};

const PRESET_SHEET_NAME = "장비프리셋";
const PRESET_HEADER_ROW = 15;
const SPREADSHEET_ID = "1GRfnkDylr6w4VLUU-Yecd71EUjm7V72KHlL7QoMKz4E";

function doGet(e) {
  const mode = e && e.parameter ? e.parameter.mode : "";
  if (mode === "preset") {
    return jsonOutput(getEquipmentPreset());
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
      partner: "",
      orderType: colIndex.orderType >= 0 ? String(row[colIndex.orderType]).trim() : "",
      storeId: colIndex.storeId >= 0 ? String(row[colIndex.storeId]).trim() : "",
      naverId: colIndex.naverId >= 0 ? String(row[colIndex.naverId]).trim() : "",
      posCount: colIndex.posCount >= 0 ? row[colIndex.posCount] : "",
      onlineBizReg: colIndex.onlineBizReg >= 0 ? String(row[colIndex.onlineBizReg]).trim() : "",
      naverConnect: colIndex.naverConnect >= 0 ? String(row[colIndex.naverConnect]).trim() : ""
    });
  });

  return jsonOutput({ stores: stores });
}

/**
 * "장비프리셋" 탭을 읽어 depth2(C열) 기준으로 그룹핑한 라벨(H열) 목록을 반환합니다.
 * depth2는 병합셀이라 빈 값인 행은 바로 위 행의 depth2 값을 이어받습니다.
 */
function getEquipmentPreset() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(PRESET_SHEET_NAME);
  if (!sheet) {
    return { error: "시트를 찾을 수 없습니다: " + PRESET_SHEET_NAME, groups: [] };
  }

  const lastRow = sheet.getLastRow();
  const numRows = lastRow - PRESET_HEADER_ROW;
  if (numRows <= 0) {
    return { groups: [] };
  }

  // A=no, B=depth1, C=depth2, D=depth3, E=tab/filter/innerPage, F=section, G=ui, H=label, I=option
  const dataRows = sheet.getRange(PRESET_HEADER_ROW + 1, 1, numRows, 9).getValues();

  const groups = [];
  const groupIndexByDepth2 = {};
  let lastDepth2 = "";

  dataRows.forEach(function (row) {
    const depth2Cell = String(row[2]).trim();
    if (depth2Cell) lastDepth2 = depth2Cell;
    const depth2 = lastDepth2;
    if (!depth2) return; // depth2를 아직 만나지 못한 행은 스킵

    const label = String(row[7]).trim();
    if (!label) return; // 빈 label은 스킵

    if (!(depth2 in groupIndexByDepth2)) {
      groupIndexByDepth2[depth2] = groups.length;
      groups.push({ depth2: depth2, labels: [] });
    }
    groups[groupIndexByDepth2[depth2]].labels.push(label);
  });

  return { groups: groups };
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
