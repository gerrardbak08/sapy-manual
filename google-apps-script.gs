/**
 * ============================================================
 * 사랑과평안의교회 — 부서 업무 현황 제출 백엔드
 * Google Apps Script Web App
 * ============================================================
 *
 * [배포 방법]
 * 1. Google Sheets 메뉴 → 확장 프로그램 → Apps Script
 * 2. 이 파일의 코드를 붙여넣기
 * 3. 배포 → 새 배포 → 유형: 웹 앱
 * 4. 설정:
 *    - 설명: 사평교 업무현황 API (버전명)
 *    - 다음 사용자로 실행: 나(스크립트 소유자)
 *    - 액세스 권한: 모든 사용자 (익명 포함)
 * 5. 배포 후 생성된 URL이 Web App URL
 *
 * [필요 권한]
 * - Google Sheets (스프레드시트 읽기/쓰기)
 * - 외부 POST 요청 수신 (자동 허용)
 *
 * [Web App URL 형식]
 * https://script.google.com/macros/s/{DEPLOYMENT_ID}/exec
 *
 * [API 사용]
 * POST {URL}          — 부서 업무현황 제출
 * GET  {URL}          — 제출 현황 조회 (관리자 대시보드용)
 * GET  {URL}?dept=WOR — 특정 부서 제출 여부 조회
 *
 * [주의사항]
 * - 배포 URL은 공개되지 않도록 관리할 것
 * - 스크립트를 수정한 뒤에는 반드시 새 배포(버전 업)를 생성해야 반영됨
 * ============================================================
 */

// ============================================================
// 상수 정의
// ============================================================

/** 스프레드시트 시트명 */
var SHEET = {
  CURRENT:   '최신현황',   // 부서별 최신 현황 요약
  WORK:      '업무현황',   // 개별 업무 항목 상세
  SUBMITTED: '제출내역'    // 제출 이력 로그
};

/** 부서코드 매핑 테이블 */
var DEPT_CODE = {
  '예배부':    'WOR',
  '재정부':    'FIN',
  '새가족부':  'NEW',
  '교육부':    'EDU',
  '친교봉사부': 'FEL',
  '시설관리부': 'FAC',
  '환경미화부': 'ENV',
  '경조사부':  'CEL',
  '문화사역팀': 'CUL',
  '문서사역팀': 'DOC',
  '성가대':    'CHO',
  '찬양팀':    'PRS'
};

// ============================================================
// 시트 컬럼 인덱스 (0-based)
// ============================================================

/** 최신현황 시트 컬럼 순서 */
var COL_CURRENT = {
  DEPT_CODE:    0,  // 부서코드
  DEPT_NAME:    1,  // 부서명
  LEADER:       2,  // 부서장
  DEPUTY:       3,  // 차장
  MEMBER_COUNT: 4,  // 인원
  MEETING:      5,  // 정기 모임
  WORK_COUNT:   6   // 업무 수
};

/** 업무현황 시트 컬럼 순서 */
var COL_WORK = {
  ROW_ID:        0,  // 행ID
  DEPT_NAME:     1,  // 부서명
  SUBMITTER:     2,  // 작성자
  POSITION:      3,  // 직책
  WORK_TITLE:    4,  // 업무명
  CYCLE:         5,  // 주기
  SCOPE:         6,  // 업무범위
  MANAGEMENT:    7,  // 관리업무
  COLLABORATORS: 8,  // 협력부서
  BUDGET:        9   // 예산
};

/** 제출내역 시트 컬럼 순서 */
var COL_SUBMIT = {
  TIMESTAMP:   0,  // 제출일시
  DEPT_CODE:   1,  // 부서코드
  DEPT_NAME:   2,  // 부서명
  SUBMITTER:   3,  // 작성자
  POSITION:    4,  // 직책
  WORK_COUNT:  5,  // 업무 수
  IMPROVEMENT: 6   // 개선 요청사항
};

// ============================================================
// CORS 헬퍼
// ============================================================

/**
 * CORS 허용 헤더를 포함한 ContentService 출력 객체를 생성한다.
 * Google Apps Script는 OPTIONS preflight를 직접 처리하지 않으므로
 * 모든 응답에 CORS 헤더를 포함한다.
 *
 * @param {Object} payload - JSON으로 직렬화할 응답 객체
 * @returns {TextOutput} CORS 헤더가 포함된 JSON 응답
 */
function buildJsonResponse(payload) {
  var output = ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
// 유틸리티 함수
// ============================================================

/**
 * 현재 스프레드시트에서 지정된 이름의 시트를 반환한다.
 * 시트가 없으면 오류를 던진다.
 *
 * @param {string} sheetName - 시트 이름
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('시트를 찾을 수 없습니다: ' + sheetName);
  }
  return sheet;
}

/**
 * 타임스탬프 + 부서코드 조합으로 고유 행 ID를 생성한다.
 * 예: 20260515143022_WOR
 *
 * @param {Date} now - 기준 날짜/시간
 * @param {string} deptCode - 부서코드
 * @returns {string} 고유 행 ID
 */
function buildRowId(now, deptCode) {
  var pad = function(n) { return String(n).padStart(2, '0'); };
  var ts = String(now.getFullYear())
    + pad(now.getMonth() + 1)
    + pad(now.getDate())
    + pad(now.getHours())
    + pad(now.getMinutes())
    + pad(now.getSeconds());
  return ts + '_' + deptCode;
}

/**
 * 날짜를 한국식 문자열(yyyy-MM-dd HH:mm:ss)로 포매팅한다.
 *
 * @param {Date} date
 * @returns {string}
 */
function formatTimestamp(date) {
  var pad = function(n) { return String(n).padStart(2, '0'); };
  return date.getFullYear()
    + '-' + pad(date.getMonth() + 1)
    + '-' + pad(date.getDate())
    + ' ' + pad(date.getHours())
    + ':' + pad(date.getMinutes())
    + ':' + pad(date.getSeconds());
}

/**
 * 예산 배열을 사람이 읽기 좋은 문자열로 변환한다.
 * 예: "주보 인쇄비(월 15,000원) / 기타(연 50,000원)"
 *
 * @param {Array} budget - [{item, amount}, ...]
 * @returns {string}
 */
function formatBudget(budget) {
  if (!budget || budget.length === 0) return '';
  return budget.map(function(b) {
    return (b.item || '') + '(' + (b.amount || '') + ')';
  }).join(' / ');
}

/**
 * 협력부서 배열을 쉼표로 구분된 문자열로 변환한다.
 *
 * @param {Array} collaborators - string[]
 * @returns {string}
 */
function formatCollaborators(collaborators) {
  if (!collaborators || collaborators.length === 0) return '';
  return collaborators.join(', ');
}

// ============================================================
// 데이터 저장 함수
// ============================================================

/**
 * 제출내역 시트에 제출 요약 행을 추가한다.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 제출내역 시트
 * @param {string} timestamp - 제출일시 문자열
 * @param {string} deptCode  - 부서코드
 * @param {Object} data      - 파싱된 폼 데이터
 */
function appendSubmitRecord(sheet, timestamp, deptCode, data) {
  var row = new Array(7);
  row[COL_SUBMIT.TIMESTAMP]   = timestamp;
  row[COL_SUBMIT.DEPT_CODE]   = deptCode;
  row[COL_SUBMIT.DEPT_NAME]   = data.department   || '';
  row[COL_SUBMIT.SUBMITTER]   = data.name          || '';
  row[COL_SUBMIT.POSITION]    = data.position      || '';
  row[COL_SUBMIT.WORK_COUNT]  = (data.works && data.works.length) ? data.works.length : 0;
  row[COL_SUBMIT.IMPROVEMENT] = data.improvement   || '';
  sheet.appendRow(row);
}

/**
 * 업무현황 시트에 각 업무 항목을 개별 행으로 추가한다.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 업무현황 시트
 * @param {string} rowId     - 고유 행ID (타임스탬프+부서코드)
 * @param {string} deptCode  - 부서코드 (행ID 접두어에 이미 포함되어 있으나 별도 저장)
 * @param {Object} data      - 파싱된 폼 데이터
 * @returns {number} 추가된 업무 항목 수
 */
function appendWorkItems(sheet, rowId, deptCode, data) {
  var works = data.works;
  if (!works || works.length === 0) return 0;

  works.forEach(function(work, index) {
    var row = new Array(10);
    // 항목이 여러 개일 경우 인덱스를 suffix로 붙여 구분
    row[COL_WORK.ROW_ID]        = rowId + '_' + (index + 1);
    row[COL_WORK.DEPT_NAME]     = data.department          || '';
    row[COL_WORK.SUBMITTER]     = data.name                || '';
    row[COL_WORK.POSITION]      = data.position            || '';
    row[COL_WORK.WORK_TITLE]    = work.title               || '';
    row[COL_WORK.CYCLE]         = work.cycle               || '';
    row[COL_WORK.SCOPE]         = work.scope               || '';
    row[COL_WORK.MANAGEMENT]    = work.management          || '';
    row[COL_WORK.COLLABORATORS] = formatCollaborators(work.collaborators);
    row[COL_WORK.BUDGET]        = formatBudget(work.budget);
    sheet.appendRow(row);
  });

  return works.length;
}

/**
 * 최신현황 시트에서 해당 부서 행을 찾거나 새로 만들고,
 * 부서장/차장 이름 및 업무 수를 업데이트한다.
 *
 * 직책이 '부서장'이면 부서장 컬럼을, '차장'이면 차장 컬럼을 갱신한다.
 * 업무 수는 이번 제출의 업무 개수로 덮어쓴다(누적 아님).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 최신현황 시트
 * @param {string} deptCode - 부서코드
 * @param {Object} data     - 파싱된 폼 데이터
 */
function upsertCurrentStatus(sheet, deptCode, data) {
  var deptName  = data.department || '';
  var submitter = data.name       || '';
  var position  = data.position   || '';
  var workCount = (data.works && data.works.length) ? data.works.length : 0;

  var lastRow = sheet.getLastRow();
  var targetRowIndex = -1; // 1-based 행 번호

  // 기존 행 탐색: 부서코드 또는 부서명으로 매칭
  if (lastRow >= 2) {
    // 헤더(1행) 제외하고 읽기
    var range = sheet.getRange(2, 1, lastRow - 1, 7);
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
      var rowDeptCode = String(values[i][COL_CURRENT.DEPT_CODE]).trim();
      var rowDeptName = String(values[i][COL_CURRENT.DEPT_NAME]).trim();
      if (rowDeptCode === deptCode || rowDeptName === deptName) {
        targetRowIndex = i + 2; // +2 = 헤더 1행 + 0-based 인덱스 보정
        break;
      }
    }
  }

  if (targetRowIndex === -1) {
    // 해당 부서 행 없음 → 새 행 추가
    var newRow = new Array(7).fill('');
    newRow[COL_CURRENT.DEPT_CODE] = deptCode;
    newRow[COL_CURRENT.DEPT_NAME] = deptName;
    if (position === '부서장') {
      newRow[COL_CURRENT.LEADER] = submitter;
    } else if (position === '차장') {
      newRow[COL_CURRENT.DEPUTY] = submitter;
    }
    newRow[COL_CURRENT.WORK_COUNT] = workCount;
    sheet.appendRow(newRow);
  } else {
    // 기존 행 업데이트
    var rowRange = sheet.getRange(targetRowIndex, 1, 1, 7);
    var rowValues = rowRange.getValues()[0];

    // 부서코드/부서명이 비어있으면 채워 넣기
    if (!rowValues[COL_CURRENT.DEPT_CODE]) {
      rowValues[COL_CURRENT.DEPT_CODE] = deptCode;
    }
    if (!rowValues[COL_CURRENT.DEPT_NAME]) {
      rowValues[COL_CURRENT.DEPT_NAME] = deptName;
    }

    // 직책에 따라 부서장 또는 차장 갱신
    if (position === '부서장') {
      rowValues[COL_CURRENT.LEADER] = submitter;
    } else if (position === '차장') {
      rowValues[COL_CURRENT.DEPUTY] = submitter;
    }

    // 업무 수 갱신 (이번 제출 기준으로 덮어쓰기)
    rowValues[COL_CURRENT.WORK_COUNT] = workCount;

    rowRange.setValues([rowValues]);
  }
}

// ============================================================
// 메인 진입점
// ============================================================

/**
 * HTTP POST 요청 처리 — 부서 업무현황 제출
 *
 * 요청 본문(JSON):
 * {
 *   "department": "예배부",
 *   "position": "부서장",
 *   "name": "홍길동",
 *   "works": [{ title, cycle, scope, management, collaborators, budget }],
 *   "improvement": "..."
 * }
 *
 * 성공 응답: { "success": true, "message": "제출이 완료되었습니다", "rowCount": N }
 * 실패 응답: { "success": false, "error": "오류 메시지" }
 *
 * @param {GoogleAppsScript.Events.DoPost} e
 * @returns {TextOutput}
 */
function doPost(e) {
  try {
    // ── 1. 요청 본문 파싱 ──────────────────────────────────────
    var rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : null;
    if (!rawBody) {
      return buildJsonResponse({ success: false, error: '요청 본문이 비어 있습니다.' });
    }

    var data;
    try {
      data = JSON.parse(rawBody);
    } catch (parseErr) {
      return buildJsonResponse({ success: false, error: 'JSON 파싱 오류: ' + parseErr.message });
    }

    // ── 2. 필수 필드 유효성 검사 ───────────────────────────────
    if (!data.department) {
      return buildJsonResponse({ success: false, error: '부서명(department)은 필수입니다.' });
    }
    if (!data.name) {
      return buildJsonResponse({ success: false, error: '작성자 이름(name)은 필수입니다.' });
    }
    if (!data.position) {
      return buildJsonResponse({ success: false, error: '직책(position)은 필수입니다.' });
    }

    // ── 3. 부서코드 조회 ───────────────────────────────────────
    var deptCode = DEPT_CODE[data.department];
    if (!deptCode) {
      // 매핑에 없는 부서는 부서명 앞 3자 대문자로 임시 코드 생성
      deptCode = data.department.substring(0, 3).toUpperCase();
    }

    // ── 4. 공통 타임스탬프 생성 ────────────────────────────────
    var now       = new Date();
    var timestamp = formatTimestamp(now);
    var rowId     = buildRowId(now, deptCode);

    // ── 5. 시트 참조 획득 ──────────────────────────────────────
    var submitSheet  = getSheet(SHEET.SUBMITTED);
    var workSheet    = getSheet(SHEET.WORK);
    var currentSheet = getSheet(SHEET.CURRENT);

    // ── 6. 제출내역 기록 ───────────────────────────────────────
    appendSubmitRecord(submitSheet, timestamp, deptCode, data);

    // ── 7. 업무현황 기록 ───────────────────────────────────────
    var savedWorkCount = appendWorkItems(workSheet, rowId, deptCode, data);

    // ── 8. 최신현황 업데이트 ───────────────────────────────────
    upsertCurrentStatus(currentSheet, deptCode, data);

    // ── 9. 성공 응답 반환 ──────────────────────────────────────
    return buildJsonResponse({
      success:  true,
      message:  '제출이 완료되었습니다',
      rowCount: savedWorkCount
    });

  } catch (err) {
    // 예상치 못한 오류 처리
    return buildJsonResponse({
      success: false,
      error:   '서버 오류가 발생했습니다: ' + err.message
    });
  }
}

/**
 * HTTP GET 요청 처리 — 제출 현황 조회 (관리자 대시보드용)
 *
 * 쿼리 파라미터:
 *   dept={부서코드}  → 특정 부서의 제출 여부만 반환
 *   (없음)           → 전체 부서 현황 반환
 *
 * 전체 조회 응답:
 * {
 *   "success": true,
 *   "departments": [
 *     { "code": "WOR", "name": "예배부", "submitted": true,  "leader": "홍길동", "workCount": 5 },
 *     { "code": "FIN", "name": "재정부", "submitted": false, "leader": "",       "workCount": 0 },
 *     ...
 *   ]
 * }
 *
 * 특정 부서 조회 응답:
 * { "success": true, "code": "WOR", "name": "예배부", "submitted": true }
 *
 * @param {GoogleAppsScript.Events.DoGet} e
 * @returns {TextOutput}
 */
function doGet(e) {
  try {
    var params    = (e && e.parameter) ? e.parameter : {};
    var queryDept = params.dept || null;

    // 최신현황 시트 읽기
    var currentSheet = getSheet(SHEET.CURRENT);
    var lastRow      = currentSheet.getLastRow();

    // 시트에 저장된 부서 데이터를 맵으로 구성
    var submittedMap = {}; // { deptCode: { name, leader, deputy, workCount } }

    if (lastRow >= 2) {
      var dataRange  = currentSheet.getRange(2, 1, lastRow - 1, 7);
      var dataValues = dataRange.getValues();
      dataValues.forEach(function(row) {
        var code = String(row[COL_CURRENT.DEPT_CODE]).trim();
        if (code) {
          submittedMap[code] = {
            name:      String(row[COL_CURRENT.DEPT_NAME]).trim(),
            leader:    String(row[COL_CURRENT.LEADER]).trim(),
            deputy:    String(row[COL_CURRENT.DEPUTY]).trim(),
            workCount: Number(row[COL_CURRENT.WORK_COUNT]) || 0
          };
        }
      });
    }

    // ── 특정 부서 단건 조회 ────────────────────────────────────
    if (queryDept) {
      var found    = submittedMap[queryDept] || null;
      var deptName = '';
      // DEPT_CODE 역방향 탐색으로 부서명 확인
      Object.keys(DEPT_CODE).forEach(function(name) {
        if (DEPT_CODE[name] === queryDept) deptName = name;
      });

      return buildJsonResponse({
        success:   true,
        code:      queryDept,
        name:      found ? found.name : deptName,
        submitted: !!found,
        leader:    found ? found.leader    : '',
        deputy:    found ? found.deputy    : '',
        workCount: found ? found.workCount : 0
      });
    }

    // ── 전체 부서 현황 조회 ────────────────────────────────────
    var allDepts = Object.keys(DEPT_CODE).map(function(name) {
      var code  = DEPT_CODE[name];
      var info  = submittedMap[code] || null;
      return {
        code:      code,
        name:      name,
        submitted: !!info,
        leader:    info ? info.leader    : '',
        deputy:    info ? info.deputy    : '',
        workCount: info ? info.workCount : 0
      };
    });

    return buildJsonResponse({
      success:     true,
      departments: allDepts
    });

  } catch (err) {
    return buildJsonResponse({
      success: false,
      error:   '조회 중 오류가 발생했습니다: ' + err.message
    });
  }
}
