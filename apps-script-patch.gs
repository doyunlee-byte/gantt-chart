/**
 * ═══════════════════════════════════════════════════════════════
 *  gantt-chart Apps Script 패치
 *  기존 Apps Script 파일에 아래 코드를 추가하세요.
 *
 *  추가 방법:
 *  1. Google Apps Script 에디터 열기
 *     (Sheets 메뉴 → 확장 프로그램 → Apps Script)
 *  2. 기존 doGet() 함수 안에 ★ 표시 부분 한 줄 추가
 *  3. 파일 하단에 나머지 함수들 전체 붙여넣기
 *  4. 저장 후 "새 배포" → 웹 앱으로 배포 (액세스: 모든 사용자)
 * ═══════════════════════════════════════════════════════════════
 */


// ── 1. 기존 doGet() 안에 아래 한 줄 추가 ─────────────────────────
//
//   function doGet(e) {
//     const action = e.parameter.action;
//     const tab    = e.parameter.tab;
//
//  ★ if (action === 'loadSettings') return _loadSettingsJSON();  // ← 이 줄 추가
//
//     if (tab) { ... }   // 기존 CSV 핸들러
//     ...
//   }
//
// ─────────────────────────────────────────────────────────────────


// ── 2. 파일 하단에 아래 함수들 전체 붙여넣기 ─────────────────────

/**
 * POST 핸들러 — saveSettings 처리
 * fetch(..., { method:'POST', body: JSON.stringify({action,cls,fix,pcs}) })
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'saveSettings') {
      _saveSettingsJSON(body);
      _saveFixSettingsTabular(body);
      return _jsonResponse({ ok: true });
    }

    return _jsonResponse({ error: 'unknown action: ' + action });
  } catch (err) {
    return _jsonResponse({ error: err.message });
  }
}

/**
 * _Settings 시트에 JSON 블롭으로 저장 (내부용)
 */
function _saveSettingsJSON(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('_Settings');
  if (!sh) {
    sh = ss.insertSheet('_Settings');
    sh.appendRow(['key', 'value', 'updated_at']);
    sh.setColumnWidth(2, 600);
  }

  const val = JSON.stringify({ cls: body.cls, fix: body.fix, pcs: body.pcs });
  const now = new Date().toISOString();
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'ganttSettings') {
      sh.getRange(i + 1, 2, 1, 2).setValues([[val, now]]);
      return;
    }
  }
  sh.appendRow(['ganttSettings', val, now]);
}

/**
 * FixSettings 시트에 가독성 있는 테이블 형식으로 저장
 * 컬럼: 센터명 | 프로젝트명 | 계획인원 | 계획건/일 | 실제인원 | 실제건/일 | 픽스여부 | 업데이트
 */
function _saveFixSettingsTabular(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('FixSettings');
  if (!sh) {
    sh = ss.insertSheet('FixSettings');
  }
  sh.clearContents();
  sh.appendRow(['센터명', '프로젝트명', '계획인원', '계획건/일', '실제인원', '실제건/일', '픽스여부', '업데이트']);

  const fix = body.fix || {};
  const pcs = body.pcs || {};
  const now = new Date().toISOString();
  const rows = [];

  for (const key of Object.keys(fix)) {
    const sep = key.indexOf('_');
    if (sep === -1) continue;
    const projId     = key.substring(0, sep);
    const centerName = key.substring(sep + 1);
    const s = pcs[key] || {};
    rows.push([
      centerName,
      projId,
      s.planPersonnel != null ? s.planPersonnel : '',
      s.planDaily     != null ? s.planDaily     : '',
      s.realPersonnel != null ? s.realPersonnel : '',
      s.realDaily     != null ? s.realDaily     : '',
      fix[key] ? 'TRUE' : 'FALSE',
      now
    ]);
  }

  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 8).setValues(rows);
  }
}

/**
 * _Settings 시트에서 JSON 블롭 읽어서 반환
 * GET ?action=loadSettings 로 호출됨
 */
function _loadSettingsJSON() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('_Settings');
  if (!sh) return _jsonResponse({});

  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'ganttSettings') {
      try {
        return _jsonResponse(JSON.parse(rows[i][1]));
      } catch (_) {}
    }
  }
  return _jsonResponse({});
}

/** JSON ContentService 헬퍼 */
function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
