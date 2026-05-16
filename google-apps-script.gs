// =============================================================
// 사랑과평안의교회 · 업무현황 취합 - Apps Script 백엔드
// =============================================================

const SHEET_ID = '1pcmN5NVxDMIGSsATuWyzxHVOVVq4cfTafKH4esR4nws';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const isUpdate = saveSubmission(data);
    return response({ success: true, message: isUpdate ? '재제출 완료' : '저장 완료', isUpdate });
  } catch (err) {
    return response({ success: false, message: err.message });
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('제출내역');
    if (!sheet || sheet.getLastRow() < 2) {
      return response({ success: true, data: [] });
    }
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const records = rows.slice(1).map(row =>
      Object.fromEntries(headers.map((h, i) => [h, row[i]]))
    );
    return response({ success: true, data: records });
  } catch (err) {
    return response({ success: false, message: err.message });
  }
}

// 제출 저장 — 같은 부서/팀이면 해당 행 덮어쓰기, 없으면 추가
// returns true if overwrite, false if new row
function saveSubmission(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, '제출내역');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['제출일시', '부서/팀', '작성자', '직분', '업무수', '개선요청', '업무상세(JSON)']);
  }

  const works = data.works || [];
  const ts = new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });
  const newRow = [ts, data.dept || '', data.name || '', data.role || '', works.length, data.feedback || '', JSON.stringify(works)];

  // 중복 체크 — 같은 부서/팀 행이 있으면 덮어쓰기
  let isUpdate = false;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const deptValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = 0; i < deptValues.length; i++) {
      if (deptValues[i][0] === data.dept) {
        sheet.getRange(i + 2, 1, 1, newRow.length).setValues([newRow]);
        isUpdate = true;
        break;
      }
    }
  }
  if (!isUpdate) sheet.appendRow(newRow);

  // 업무상세 시트에 플랫 구조로 저장
  saveWorkDetails(ss, data, works, ts);

  return isUpdate;
}

// 업무상세 시트 — 업무 항목별 1행 저장
function saveWorkDetails(ss, data, works, ts) {
  const sheet = getOrCreateSheet(ss, '업무상세');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['제출일시', '부서/팀', '작성자', '순번', '업무명', '주기', '업무범위', '관리항목', '예산']);
  }

  // 해당 부서 기존 행 역순 삭제
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const deptValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = deptValues.length - 1; i >= 0; i--) {
      if (deptValues[i][0] === data.dept) sheet.deleteRow(i + 2);
    }
  }

  // 업무별 행 추가
  works.forEach(function(w, idx) {
    sheet.appendRow([ts, data.dept, data.name || '', idx + 1, w.title, w.cycle, w.scope, w.manage, w.budget]);
  });
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function response(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
