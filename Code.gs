/**
 * ═══════════════════════════════════════════════════════════════
 *  gantt-chart Apps Script — Code.gs
 *  Google Apps Script 에디터에 이 파일 전체를 붙여넣고
 *  "새 배포" → 웹 앱 (액세스: 모든 사용자)으로 배포하세요.
 * ═══════════════════════════════════════════════════════════════
 */

// ── GET 핸들러 ──────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  const tab    = e.parameter.tab;

  if (action === 'loadSettings') return _loadSettingsJSON();
  if (action === 'createSlide')  return _createSlide();

  if (tab) return _getSheetCSV(tab);

  return ContentService
    .createTextOutput('OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── POST 핸들러 — saveSettings 처리 ────────────────────────────
// fetch(..., { method:'POST', body: JSON.stringify({action,cls,fix,pcs}) })
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'saveSettings') {
      _saveSettingsJSON(body);
      _saveFixSettingsTabular(body);
      return _jsonResponse({ ok: true });
    }

    if (action === 'updateBoardCell') {
      _updateBoardCell(body);
      return _jsonResponse({ ok: true });
    }

    if (action === 'appendBoardRow') {
      _appendBoardRow(body);
      return _jsonResponse({ ok: true });
    }

    if (action === 'updateBoardRow') {
      _updateBoardRow(body);
      return _jsonResponse({ ok: true });
    }

    return _jsonResponse({ error: 'unknown action: ' + action });
  } catch (err) {
    return _jsonResponse({ error: err.message });
  }
}

// ── 시트 → CSV 반환 ────────────────────────────────────────────
function _getSheetCSV(tabName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(tabName);
    if (!sh) return ContentService
      .createTextOutput('')
      .setMimeType(ContentService.MimeType.TEXT);

    const data = sh.getDataRange().getValues();
    const csv  = data.map(row =>
      row.map(cell => {
        const s = String(cell);
        return s.includes(',') || s.includes('"') || s.includes('\n')
          ? '"' + s.replace(/"/g, '""') + '"'
          : s;
      }).join(',')
    ).join('\n');

    return ContentService
      .createTextOutput(csv)
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput('')
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ── Google Slides 생성 ─────────────────────────────────────────
function _createSlide() {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const ssName   = ss.getName();
    const pres     = SlidesApp.create(ssName + ' — 요약 ' + _today());
    const slide    = pres.getSlides()[0];

    // 제목 슬라이드 기본 텍스트
    slide.getShapes()[0].getText().setText(ssName + ' 프로젝트 현황');
    slide.getShapes()[1].getText().setText(_today() + ' 기준');

    // 데이터 시트 목록으로 슬라이드 추가
    const sheets = ss.getSheets().filter(sh =>
      !['_Settings', 'FixSettings'].includes(sh.getName())
    );

    sheets.forEach(sh => {
      const newSlide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      const shapes   = newSlide.getShapes();
      shapes[0].getText().setText(sh.getName());

      const data  = sh.getDataRange().getValues();
      const lines = data.slice(0, 20).map(row => row.slice(0, 5).join('  |  '));
      shapes[1].getText().setText(lines.join('\n'));
    });

    return _jsonResponse({ url: pres.getUrl() });
  } catch (err) {
    return _jsonResponse({ error: err.message });
  }
}

// ── _Settings 시트에 JSON 블롭으로 저장 ────────────────────────
function _saveSettingsJSON(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName('_Settings');
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
  let sh   = ss.getSheetByName('FixSettings');
  if (!sh) sh = ss.insertSheet('FixSettings');
  sh.clearContents();
  sh.appendRow(['센터명', '프로젝트명', '계획인원', '계획건/일', '실제인원', '실제건/일', '픽스여부', '업데이트']);

  const fix  = body.fix || {};
  const pcs  = body.pcs || {};
  const now  = new Date().toISOString();
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

// ── 태스크보드 쓰기 ────────────────────────────────────────────

function _updateBoardCell(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('태스크보드');
  if (!sh) throw new Error('태스크보드 시트 없음');
  sh.getRange(body.row, body.col).setValue(body.value);
}

function _appendBoardRow(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('태스크보드');
  if (!sh) throw new Error('태스크보드 시트 없음');
  sh.appendRow(body.data);
}

function _updateBoardRow(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('태스크보드');
  if (!sh) throw new Error('태스크보드 시트 없음');
  sh.getRange(body.row, 1, 1, body.data.length).setValues([body.data]);
}

// ── 헬퍼 ───────────────────────────────────────────────────────

function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _today() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
