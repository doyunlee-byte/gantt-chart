/**
 * 더본외식 "더본외식Lowdata" 탭 → 대시보드 연동용 Web App
 *
 * [배포 방법]
 * 1) 더본외식 스프레드시트 열기 → 확장 프로그램 > Apps Script
 * 2) 기본 Code.gs 내용을 전부 지우고 이 파일 내용을 붙여넣기
 * 3) 우측 상단 "배포" > "새 배포" > 유형: 웹 앱
 *    - 실행 계정: 나(본인)
 *    - 액세스 권한: 링크가 있는 모든 사용자 (조직 내부용이면 "조직 내" 선택 가능)
 * 4) 배포 후 나오는 웹 앱 URL을 복사
 * 5) 대시보드 index.html 상단 PROJECT_CONFIG.deoban.webAppUrl 에 붙여넣기
 *
 * [열 매핑 — 도윤님이 알려준 실제 컬럼 위치]
 *   D열: 보류여부 (제외 필터)
 *   F열: 브랜드
 *   G열: 점포명
 *   I열: 사업자등록번호2
 *   U열: 가맹계약자(점주) 연락처
 *   Y열: 협력사
 *   Z열: 설치 담당자
 *   AA열: 설치 예정일
 *   AE열: 점포 진척도
 *
 * 열 위치가 바뀌면 아래 COLS 값(알파벳 열 문자)만 수정하면 됩니다.
 */

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

// 이 값들이 보류여부(D열)에 들어있으면 대시보드 목록에서 제외합니다.
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
    idx[key] = colLetterToIndex(COLS[key]) - 1; // 0-based
  });

  const stores = [];
  dataRows.forEach(function (row) {
    const hold = String(row[idx.hold] || "").trim();
    if (EXCLUDE_HOLD_VALUES.indexOf(hold) !== -1) return; // 제외 대상

    const name = row[idx.name];
    if (!name) return; // 빈 행 스킵

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
