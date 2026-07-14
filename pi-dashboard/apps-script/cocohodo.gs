/**
 * 코코호도 "가맹점리스트 취합" 탭 → 대시보드 연동용 Web App
 *
 * [배포 방법]
 * 1) 코코호도 스프레드시트 열기 → 확장 프로그램 > Apps Script
 * 2) 기본 Code.gs 내용을 전부 지우고 이 파일 내용을 붙여넣기
 * 3) 우측 상단 "배포" > "새 배포" > 유형: 웹 앱
 *    - 실행 계정: 나(본인)
 *    - 액세스 권한: 링크가 있는 모든 사용자 (조직 내부용이면 "조직 내" 선택 가능)
 * 4) 배포 후 나오는 웹 앱 URL을 복사
 * 5) 대시보드 index.html 상단 PROJECT_CONFIG.cocohodo.webAppUrl 에 붙여넣기
 *
 * [주의]
 * - 헤더 행이 4행에 있다고 가정합니다 (실제 시트 구조 기준). 다르면 HEADER_ROW 값을 수정하세요.
 * - 컬럼은 "헤더 텍스트"로 찾습니다. 시트의 실제 헤더 텍스트와 아래 HEADER_MAP의 값이
 *   정확히 일치해야 합니다. 다르면 HEADER_MAP만 수정하면 됩니다 (열 위치가 바뀌어도 안전).
 */

const SHEET_NAME = "가맹점리스트 취합";
const HEADER_ROW = 4;

// 대시보드가 원하는 필드명 : 실제 시트 헤더에 들어있는 텍스트
const HEADER_MAP = {
  name: "매장명",
  bizNo: "사업자번호",
  ownerContact: "대표자연락처",
  installOwner: "설치담당",       // 시트에 아직 없다면 직접 컬럼을 추가해주세요
  installDate: "설치확정일",      // 시트에 아직 없다면 직접 컬럼을 추가해주세요
  progress: "운영관리등록여부",   // 진행 상태를 대표하는 컬럼 (필요시 다른 컬럼으로 교체)
  orderType: "발주유형",
  storeId: "매장아이디",
  naverId: "네이버ID",
  posCount: "포스 수",
  onlineBizReg: "전산등록 여부",
  naverConnect: "네이버커넥트",
  centerName: "센터",               // AG열 (실제 헤더: "센터")
  installCompleteDate: "설치완료일", // AJ열
  installStatus: "설치 여부"         // AK열 (실제 헤더: "설치 여부", 공백 포함)
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
  // 셀 안에 줄바꿈(\n)이 들어간 헤더도 안전하게 매칭되도록 양쪽 다 줄바꿈 제거 후 비교
  const colIndex = {};
  Object.keys(HEADER_MAP).forEach(function (key) {
    const headerText = normalizeHeaderText(HEADER_MAP[key]);
    const idx = headerRowValues.findIndex(function (h) {
      return normalizeHeaderText(h) === headerText;
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
      partner: "",
      orderType: colIndex.orderType >= 0 ? String(row[colIndex.orderType]).trim() : "",
      storeId: colIndex.storeId >= 0 ? String(row[colIndex.storeId]).trim() : "",
      naverId: colIndex.naverId >= 0 ? String(row[colIndex.naverId]).trim() : "",
      posCount: colIndex.posCount >= 0 ? row[colIndex.posCount] : "",
      onlineBizReg: colIndex.onlineBizReg >= 0 ? String(row[colIndex.onlineBizReg]).trim() : "",
      naverConnect: colIndex.naverConnect >= 0 ? String(row[colIndex.naverConnect]).trim() : "",
      centerName: colIndex.centerName >= 0 ? String(row[colIndex.centerName]).trim() : "",
      installCompleteDate: colIndex.installCompleteDate >= 0 ? formatDate(row[colIndex.installCompleteDate]) : "",
      installStatus: colIndex.installStatus >= 0 ? String(row[colIndex.installStatus]).trim() : ""
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

function normalizeHeaderText(v) {
  return String(v).replace(/\n/g, '').trim();
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
