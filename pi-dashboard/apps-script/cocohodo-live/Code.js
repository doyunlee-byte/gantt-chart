const SHEET_NAME = "가맹점리스트 취합";
const HEADER_ROW = 4;

// 대시보드가 원하는 필드명 : 실제 시트 헤더에 들어있는 텍스트
const HEADER_MAP = {
  name: "매장명",
  bizNo: "사업자번호",
  ownerContact: "대표자연락처",
  installOwner: "설치담당자",     // 시트에 아직 없다면 직접 컬럼을 추가해주세요
  installDate: "설치예정일",      // 시트에 아직 없다면 직접 컬럼을 추가해주세요
  progress: "운영관리등록여부"    // 진행 상태를 대표하는 컬럼 (필요시 다른 컬럼으로 교체)
};

const PRESET_SHEET_NAME = "장비프리셋";
const PRESET_HEADER_ROW = 15;

function doGet(e) {
  const mode = e && e.parameter ? e.parameter.mode : "";
  if (mode === "preset") {
    return jsonOutput(getEquipmentPreset());
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return jsonOutput({ error: "시트를 찾을 수 없습니다: " + SHEET_NAME, stores: [] });
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headerRowValues = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

  // 헤더 텍스트 → 열 인덱스(0-based) 매핑
  const colIndex = {};
  Object.keys(HEADER_MAP).forEach(function (key) {
    const headerText = HEADER_MAP[key];
    const idx = headerRowValues.findIndex(function (h) {
      return String(h).trim() === headerText;
    });
    colIndex[key] = idx; // 못 찾으면 -1
  });

  const dataRows = sheet
    .getRange(HEADER_ROW + 1, 1, Math.max(lastRow - HEADER_ROW, 0), lastCol)
    .getValues();

  const stores = [];
  dataRows.forEach(function (row) {
    const name = colIndex.name >= 0 ? row[colIndex.name] : "";
    if (!name) return; // 매장명 없는 빈 행은 스킵

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

/**
 * "장비프리셋" 탭을 읽어 depth2(C열) 기준으로 그룹핑한 라벨(H열) 목록을 반환합니다.
 * depth2는 병합셀이라 빈 값인 행은 바로 위 행의 depth2 값을 이어받습니다.
 */
function getEquipmentPreset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
    const groupLabels = groups[groupIndexByDepth2[depth2]].labels;
    if (groupLabels.indexOf(label) === -1) {
      groupLabels.push(label); // 그룹 내 중복 label 제거 — 처음 등장한 것만 유지
    }
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
