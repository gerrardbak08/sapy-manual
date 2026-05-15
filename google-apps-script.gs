// =============================================================
// 사랑과평안의교회 · 업무현황 취합 - Apps Script 백엔드
// =============================================================
// 설정: SHEET_ID에 본인 스프레드시트 ID 입력 후 웹앱 배포
// =============================================================

const SHEET_ID = '1pcmN5NVxDMIGSsATuWyzxHVOVVq4cfTafKH4esR4nws';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    saveSubmission(data);
    return response({ success: true, message: '저장 완료' });
  } catch (err) {
    return response({ success: false, message: err.message });
  }
}

function doGet(e) {
  // 관리자 대시보드용 현황 조회
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('제출내역');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const records = rows.slice(1).map(row =>
    Object.fromEntries(headers.map((h, i) => [h, row[i]]))
  );
  return response({ success: true, data: records });
}

function saveSubmission(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, '제출내역');

  // 헤더가 없으면 생성
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['제출일시', '부서/팀', '작성자', '직분', '업무수', '개선요청', '업무상세(JSON)']);
  }

  const works = data.works || [];
  sheet.appendRow([
    new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }),
    data.dept     || '',
    data.name     || '',
    data.role     || '',
    works.length,
    data.feedback || '',
    JSON.stringify(works)
  ]);
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function response(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
