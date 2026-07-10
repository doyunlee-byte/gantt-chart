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
  installOwner: "설치담당자",     // 시트에 아직 없다면 직접 컬럼을 추가해주세요
  installDate: "설치예정일",      // 시트에 아직 없다면 직접 컬럼을 추가해주세요
  progress: "운영관리등록여부"    // 진행 상태를 대표하는 컬럼 (필요시 다른 컬럼으로 교체)
};

function doGet(e) {
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
