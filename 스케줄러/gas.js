// ══════════════════════════════════════════════════════
//  팀 스케줄러 - Google Apps Script 백엔드
//
//  사용법:
//  1. Google Sheets 새 파일 열기
//  2. 확장프로그램 > Apps Script
//  3. 이 파일 전체 붙여넣기 (기존 내용 덮어씌우기)
//  4. 저장 (Ctrl+S)
//  5. 배포 > 새 배포
//     - 유형: 웹 앱
//     - 다음 사용자로 실행: 나 (내 Google 계정)
//     - 액세스 권한: 모든 사용자 (익명 포함)
//  6. 배포 > 배포 URL 복사
//  7. index.html의 API_URL 변수에 붙여넣기
// ══════════════════════════════════════════════════════

const SHEET_NAME = 'schedules';

function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    if (action === 'read') {
      result = readSchedule(e.parameter.team, e.parameter.week);
    } else if (action === 'write') {
      result = writeSchedule(
        e.parameter.team,
        e.parameter.member,
        e.parameter.week,
        e.parameter.slots,
        e.parameter.notes || ''
      );
    } else {
      result = { error: '알 수 없는 action: ' + action };
    }
  } catch(err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── 시트 가져오기 (없으면 생성) ── */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['team', 'member', 'week', 'slots', 'notes', 'updated']);
    sheet.setFrozenRows(1);
    // 열 너비 조정
    sheet.setColumnWidth(4, 300);
    sheet.setColumnWidth(5, 200);
  }
  return sheet;
}

/* ── 읽기 ── */
function readSchedule(team, week) {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === team && rows[i][2] === week) {
      let slots = {};
      try { slots = JSON.parse(rows[i][3] || '{}'); } catch(_) {}
      result.push({
        member:  rows[i][1],
        slots:   slots,
        notes:   rows[i][4] || '',
        updated: rows[i][5] || ''
      });
    }
  }
  return result;
}

/* ── 쓰기 (없으면 추가, 있으면 덮어쓰기) ── */
function writeSchedule(team, member, week, slotsStr, notes) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  let rowIdx  = -1;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === team && rows[i][1] === member && rows[i][2] === week) {
      rowIdx = i + 1; // 시트는 1-indexed
      break;
    }
  }

  const now = new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });

  if (rowIdx === -1) {
    sheet.appendRow([team, member, week, slotsStr, notes, now]);
  } else {
    sheet.getRange(rowIdx, 4).setValue(slotsStr);
    sheet.getRange(rowIdx, 5).setValue(notes);
    sheet.getRange(rowIdx, 6).setValue(now);
  }

  return { success: true };
}
