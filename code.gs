// ==========================================
// 1. 웹앱 화면 표시
// ==========================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('나세나반 다이어리')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 2. 보안 설정 관련 함수
//    스프레드시트 URL은 Script Properties(서버)에만 저장됨
//    → HTML 소스코드에 절대 노출되지 않음
// ==========================================

/**
 * 스프레드시트가 이미 설정되어 있는지 확인
 * (HTML 로딩 시 최초 1회 호출)
 * @returns {boolean}
 */
function isSetupComplete() {
  const url = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_URL');
  return !!url;
}

/**
 * 스프레드시트 URL을 Script Properties에 저장
 * URL은 서버(GAS)에만 보관 → 소스코드/localStorage에 URL 자체는 저장하지 않음
 * localStorage에는 "설정 완료 여부 플래그"만 저장됨
 * @param {string} spreadsheetUrl
 * @returns {{success: boolean, message: string}}
 */
function setupSpreadsheet(spreadsheetUrl) {
  try {
    // 실제로 열 수 있는지 검증 (권한 없거나 잘못된 URL이면 여기서 에러)
    SpreadsheetApp.openByUrl(spreadsheetUrl);
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_URL', spreadsheetUrl);
    return { success: true, message: '스프레드시트 연동이 완료되었습니다.' };
  } catch (e) {
    return { success: false, message: '스프레드시트를 열 수 없습니다. URL과 공유 권한을 확인해주세요. (' + e.message + ')' };
  }
}

/**
 * 설정 초기화 (스프레드시트 변경 시 사용)
 */
function resetSetup() {
  PropertiesService.getScriptProperties().deleteProperty('SPREADSHEET_URL');
  return '설정이 초기화되었습니다.';
}

// ==========================================
// 3. 스프레드시트 공통 헬퍼 함수
// ==========================================

function getSpreadsheet() {
  const url = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_URL');
  if (!url) throw new Error('스프레드시트가 설정되지 않았습니다. 초기 설정을 진행해주세요.');
  return SpreadsheetApp.openByUrl(url);
}

function getSheet(sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('"' + sheetName + '" 시트를 찾을 수 없습니다. 스프레드시트 시트명을 확인해주세요.');
  return sheet;
}

// ==========================================
// 4. 학생 정보 불러오기
// ==========================================

function getStudentList() {
  const sheet = getSheet('학생정보');
  const data = sheet.getDataRange().getValues();
  const students = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    students.push({
      num:          data[i][0],
      name:         data[i][1],
      phone:        data[i][3] || '',
      parentPhone1: data[i][4] || '',
      parentPhone2: data[i][5] || '',
      note:         data[i][6] || ''
    });
  }
  return students;
}

// ==========================================
// 5. 각 메뉴별 데이터 저장 함수 (Create)
// ==========================================

// [출결기록] 저장
function saveAttendanceRecord(data) {
  const sheet = getSheet('출결기록');
  const id = 'ATT-' + new Date().getTime();
  sheet.appendRow([
    id, data.date, data.studentNum, data.studentName,
    data.status, data.reason, data.proof
  ]);
  return '출결 기록이 저장되었습니다.';
}

// [수업기록] 저장
function saveClassRecord(data) {
  const sheet = getSheet('수업기록');
  const id = 'CLS-' + new Date().getTime();
  sheet.appendRow([
    id, data.date,
    data.periods.join(', '),
    data.subjects.join(', '),
    data.unit, data.topic, data.reflection, data.link, data.files
  ]);
  return '수업 기록이 저장되었습니다.';
}

// [일상기록] 저장
function saveDailyRecord(data) {
  const sheet = getSheet('일상기록');
  const id = 'DLY-' + new Date().getTime();
  sheet.appendRow([id, data.date, data.keyword, data.content, data.link, data.files]);
  return '일상 기록이 저장되었습니다.';
}

// [학생기록] 저장
function saveStudentRecord(data) {
  const sheet = getSheet('학생기록');
  const id = 'STU-' + new Date().getTime();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([id, timestamp, data.studentNum, data.studentName, data.category, data.content, '']);
  return '학생 기록이 저장되었습니다.';
}

// [상담기록] 저장 (다중 학생 → 각각 행으로 저장)
function saveCounselRecord(data) {
  const sheet = getSheet('상담기록');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  data.studentTags.forEach(function(studentStr) {
    const numMatch = studentStr.match(/^(\d+)번/);
    const studentNum  = numMatch ? numMatch[1] : '';
    const studentName = studentStr.replace(/^\d+번\s*/, '');
    const id = 'CNS-' + new Date().getTime() + '-' + studentNum;
    sheet.appendRow([
      id, timestamp, studentNum, studentName,
      data.targetType, data.method, data.content, ''
    ]);
  });

  return data.studentTags.length + '건의 상담 기록이 저장되었습니다.';
}

// ==========================================
// 6. 학생 대시보드: 특정 학생의 전체 기록 조회
// ==========================================

/**
 * 학생 번호로 출결기록, 학생기록, 상담기록을 한 번에 조회
 *   출결기록:  A=ID, B=날짜, C=번호, D=이름, E=출결상태, F=사유, G=증빙자료
 *   학생기록:  A=ID, B=기록일시, C=번호, D=이름, E=분류, F=내용, G=지도내용
 *   상담기록:  A=ID, B=상담일시, C=번호, D=이름, E=상담대상, F=방법, G=내용, H=추후계획
 */
function getStudentAllRecords(studentNum) {
  var ss     = getSpreadsheet();
  var numInt = parseInt(String(studentNum), 10);  // 숫자/문자열 모두 처리
  var result = { attendance: [], records: [], counsel: [] };

  function fmtDate(v) {
    return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v);
  }
  function fmtDatetime(v) {
    return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(v);
  }
  function matchNum(cellVal) {
    return parseInt(String(cellVal), 10) === numInt;
  }

  // 출결기록 (C열 = 번호 = index 2)
  var attSheet = ss.getSheetByName('출결기록');
  if (attSheet && attSheet.getLastRow() > 1) {
    var attData = attSheet.getDataRange().getValues();
    for (var i = 1; i < attData.length; i++) {
      if (!attData[i][0]) continue;
      if (matchNum(attData[i][2])) {
        result.attendance.push({
          date:   fmtDate(attData[i][1]),
          status: attData[i][4],
          reason: attData[i][5] || '',
          proof:  attData[i][6] || ''
        });
      }
    }
  }

  // 학생기록 (C열 = 번호 = index 2)
  var recSheet = ss.getSheetByName('학생기록');
  if (recSheet && recSheet.getLastRow() > 1) {
    var recData = recSheet.getDataRange().getValues();
    for (var j = 1; j < recData.length; j++) {
      if (!recData[j][0]) continue;
      if (matchNum(recData[j][2])) {
        result.records.push({
          timestamp: fmtDatetime(recData[j][1]),
          category:  recData[j][4],
          content:   recData[j][5] || ''
        });
      }
    }
  }

  // 상담기록 (C열 = 번호 = index 2)
  var cnslSheet = ss.getSheetByName('상담기록');
  if (cnslSheet && cnslSheet.getLastRow() > 1) {
    var cnslData = cnslSheet.getDataRange().getValues();
    for (var k = 1; k < cnslData.length; k++) {
      if (!cnslData[k][0]) continue;
      if (matchNum(cnslData[k][2])) {
        result.counsel.push({
          timestamp:  fmtDatetime(cnslData[k][1]),
          targetType: cnslData[k][4],
          method:     cnslData[k][5],
          content:    cnslData[k][6] || ''
        });
      }
    }
  }

  return result;
}

// ==========================================
// 7. 출결기록 전체 목록 조회 (리스트 화면용)
// ==========================================

function getAttendanceList() {
  var attSheet = getSpreadsheet().getSheetByName('출결기록');
  if (!attSheet || attSheet.getLastRow() <= 1) return [];
  var data = attSheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      id:          data[i][0],
      date:        data[i][1] instanceof Date
                     ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
                     : String(data[i][1]),
      studentNum:  data[i][2],
      studentName: data[i][3],
      status:      data[i][4],
      reason:      data[i][5] || '',
      proof:       data[i][6] || ''
    });
  }
  return result;
}
