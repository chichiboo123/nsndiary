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
// ==========================================
function isSetupComplete() {
  const url = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_URL');
  return !!url;
}

function setupSpreadsheet(spreadsheetUrl) {
  try {
    SpreadsheetApp.openByUrl(spreadsheetUrl);
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_URL', spreadsheetUrl);
    var m = spreadsheetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    var ssId = m ? m[1] : '';
    return { success: true, message: '스프레드시트 연동이 완료되었습니다.', spreadsheetId: ssId };
  } catch (e) {
    return { success: false, message: '스프레드시트를 열 수 없습니다. URL과 공유 권한을 확인해주세요. (' + e.message + ')', spreadsheetId: '' };
  }
}

function verifyAccess(spreadsheetId) {
  if (!spreadsheetId) return false;
  var storedUrl = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_URL');
  if (!storedUrl) return false;
  var m = storedUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  var storedId = m ? m[1] : null;
  return storedId === String(spreadsheetId);
}

function resetSetup() {
  PropertiesService.getScriptProperties().deleteProperty('SPREADSHEET_URL');
  return '설정이 초기화되었습니다.';
}

// ==========================================
// 3. 스프레드시트 공통 헬퍼
// ==========================================

// 실행 내 스프레드시트 캐시 (동일 google.script.run 호출 내 openByUrl 중복 방지)
var _cachedSS = null;

function getSpreadsheet() {
  if (_cachedSS) return _cachedSS;
  const url = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_URL');
  if (!url) throw new Error('스프레드시트가 설정되지 않았습니다. 초기 설정을 진행해주세요.');
  _cachedSS = SpreadsheetApp.openByUrl(url);
  return _cachedSS;
}

function getSheet(sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('"' + sheetName + '" 시트를 찾을 수 없습니다. 스프레드시트 시트명을 확인해주세요.');
  return sheet;
}

// 공통 날짜 포맷
function _fmt(v, tz) {
  return v instanceof Date ? Utilities.formatDate(v, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v || '');
}
function _fmtDt(v, tz) {
  return v instanceof Date ? Utilities.formatDate(v, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(v || '');
}

// ==========================================
// 4. 학생 정보
// ==========================================
// 학생정보 시트: A=번호, B=이름, C=(비움), D=학생연락처, E=보호자1, F=보호자2, G=특이사항
function getStudentList() {
  // CacheService로 5분 캐싱 (학생 정보는 자주 바뀌지 않음)
  var cache = CacheService.getScriptCache();
  var cached = cache.get('STUDENT_LIST');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }
  const sheet = getSheet('학생정보');
  const data = sheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    students.push({
      num: data[i][0], name: data[i][1],
      phone: data[i][3] || '', parentPhone1: data[i][4] || '',
      parentPhone2: data[i][5] || '', note: data[i][6] || ''
    });
  }
  try { cache.put('STUDENT_LIST', JSON.stringify(students), 300); } catch(e) {}
  return students;
}

// ==========================================
// 4-1. 초기 데이터 일괄 조회 (학생 + 일정 한 번에)
// ==========================================
// 프론트에서 loadAppData() 시 google.script.run 1회로 두 데이터 수신
function getInitialData() {
  return {
    students: getStudentList(),
    schedules: getScheduleList()
  };
}

// 캐시를 무시하고 스프레드시트에서 직접 최신 학생 목록 반환 (새로고침용)
function refreshStudentList() {
  CacheService.getScriptCache().remove('STUDENT_LIST');
  return getStudentList();
}

// ==========================================
// 4-2. 학생 추가 / 삭제
// ==========================================
function addStudent(data) {
  var sheet = getSheet('학생정보');
  var rows = sheet.getDataRange().getValues();
  var num = parseInt(String(data.studentNum), 10);
  if (!num || num < 1) throw new Error('올바른 번호를 입력해주세요.');
  // 중복 번호 확인
  for (var i = 1; i < rows.length; i++) {
    if (parseInt(String(rows[i][0]), 10) === num) {
      throw new Error('이미 같은 번호의 학생이 있습니다. (' + num + '번)');
    }
  }
  sheet.appendRow([data.studentNum, data.name, '', data.phone || '', data.parentPhone1 || '', data.parentPhone2 || '', data.note || '']);
  CacheService.getScriptCache().remove('STUDENT_LIST');
  return '학생이 추가되었습니다.';
}

function deleteStudent(studentNum) {
  var sheet = getSheet('학생정보');
  var rows = sheet.getDataRange().getValues();
  var num = parseInt(String(studentNum), 10);
  for (var i = 1; i < rows.length; i++) {
    if (parseInt(String(rows[i][0]), 10) === num) {
      sheet.deleteRow(i + 1);
      CacheService.getScriptCache().remove('STUDENT_LIST');
      return '학생이 삭제되었습니다.';
    }
  }
  throw new Error('학생을 찾을 수 없습니다. (번호: ' + studentNum + ')');
}

// ==========================================
// 5. 일정 조회
// ==========================================
// 일정 시트 컬럼: A=날짜, B=제목, C=분류(학교일정/학급일정/행사/기타), D=내용
function getScheduleList() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('일정');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue;
    var dateCell = data[i][0];
    var dateStr = '';
    if (dateCell instanceof Date) {
      dateStr = Utilities.formatDate(dateCell, tz, 'yyyy-MM-dd');
    } else {
      var text = String(dateCell || '').trim();
      var m1 = text.match(/^(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
      if (m1) {
        dateStr = m1[1] + '-' + ('0'+m1[2]).slice(-2) + '-' + ('0'+m1[3]).slice(-2);
      } else {
        var m2 = text.match(/^(\d{1,2})[.\-\/](\d{1,2})/);
        if (m2) {
          dateStr = new Date().getFullYear() + '-' + ('0'+m2[1]).slice(-2) + '-' + ('0'+m2[2]).slice(-2);
        } else {
          dateStr = text;
        }
      }
    }
    result.push({
      rowIndex: i + 1,
      date:     dateStr,
      title:    String(data[i][1] || ''),
      category: String(data[i][2] || '기타'),
      content:  String(data[i][3] || '')
    });
  }
  result.sort(function(a, b) { return a.date.localeCompare(b.date); });
  return result;
}

// ==========================================
// 6. 일정 추가/수정, 학생정보 수정
// ==========================================
function addScheduleItem(data) {
  var sheet = getSheet('일정');
  sheet.appendRow([data.date, data.title, data.category || '기타', data.content || '']);
  return '일정이 추가되었습니다.';
}

// 수업기록: A=ID, B=날짜, C=교시, D=과목, E=단원/차시, F=배움주제, G=수업내용, H=성찰, I=링크, J=파일URL
function updateClassRecord(data) {
  var sheet = getSheet('수업기록');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      // data.files 가 undefined면 기존 값(rows[i][9]) 유지
      var filesVal = (data.files !== undefined && data.files !== null) ? data.files : (rows[i][9] || '');
      sheet.getRange(i + 1, 2, 1, 9).setValues([[
        data.date, data.periods || '', data.subjects || '', rows[i][4],
        data.topic || '', data.content || '', data.reflection || '', data.link || '', filesVal
      ]]);
      return '수업 기록이 수정되었습니다.';
    }
  }
  throw new Error('수업 기록을 찾을 수 없습니다. (ID: ' + data.id + ')');
}

function updateScheduleItem(data) {
  var sheet = getSheet('일정');
  var row = parseInt(data.rowIndex, 10);
  if (!row || row < 2) throw new Error('올바른 행 번호가 아닙니다.');
  sheet.getRange(row, 1, 1, 4).setValues([[data.date, data.title, data.category || '기타', data.content || '']]);
  return '일정이 수정되었습니다.';
}

function deleteScheduleItem(rowIndex) {
  var sheet = getSheet('일정');
  var row = parseInt(rowIndex, 10);
  if (!row || row < 2) throw new Error('올바른 행 번호가 아닙니다.');
  sheet.deleteRow(row);
  return '일정이 삭제되었습니다.';
}

function updateStudent(data) {
  var sheet = getSheet('학생정보');
  var rows = sheet.getDataRange().getValues();
  var num = parseInt(String(data.studentNum), 10);
  for (var i = 1; i < rows.length; i++) {
    if (parseInt(String(rows[i][0]), 10) === num) {
      sheet.getRange(i + 1, 1, 1, 7).setValues([[
        data.studentNum, data.name, '', data.phone || '', data.parentPhone1 || '', data.parentPhone2 || '', data.note || ''
      ]]);
      // 학생 목록 캐시 무효화
      CacheService.getScriptCache().remove('STUDENT_LIST');
      return '학생 정보가 수정되었습니다.';
    }
  }
  throw new Error('학생을 찾을 수 없습니다. (번호: ' + data.studentNum + ')');
}

// ==========================================
// 7. 각 기록 저장 (Create)
// ==========================================
// 출결기록: A=ID, B=날짜, C=번호, D=이름, E=출결상태, F=사유, G=증빙자료
function saveAttendanceRecord(data) {
  const sheet = getSheet('출결기록');
  const id = 'ATT-' + new Date().getTime();
  sheet.appendRow([id, data.date, data.studentNum, data.studentName, data.status, data.reason, data.proof]);
  return '출결 기록이 저장되었습니다.';
}

function updateAttendanceRecord(data) {
  var sheet = getSheet('출결기록');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.getRange(i + 1, 2, 1, 6).setValues([[
        data.date, data.studentNum, data.studentName, data.status, data.reason || '', data.proof || 'X'
      ]]);
      return '출결 기록이 수정되었습니다.';
    }
  }
  throw new Error('출결 기록을 찾을 수 없습니다. (ID: ' + data.id + ')');
}

function saveBulkAttendanceRecords(records) {
  var sheet = getSheet('출결기록');
  var ts = new Date().getTime();
  for (var i = 0; i < records.length; i++) {
    var data = records[i];
    var id = 'ATT-' + (ts + i);
    sheet.appendRow([id, data.date, data.studentNum, data.studentName, data.status, data.reason, data.proof]);
  }
  return records.length + '건의 출결 기록이 저장되었습니다.';
}

// 수업기록: A=ID, B=날짜, C=교시, D=과목, E=단원/차시, F=배움주제, G=수업내용, H=성찰, I=링크, J=파일URL
function saveClassRecord(data) {
  const sheet = getSheet('수업기록');
  const id = 'CLS-' + new Date().getTime();
  sheet.appendRow([id, data.date, data.periods.join(', '), data.subjects.join(', '),
    data.unit, data.topic, data.content || '', data.reflection, data.link, data.files]);
  return '수업 기록이 저장되었습니다.';
}

// 일상기록: A=ID, B=날짜, C=키워드, D=내용, E=링크, F=파일URL
// 일상기록: A=ID, B=날짜, C=키워드, D=내용, E=링크, F=파일URL
function saveDailyRecord(data) {
  const sheet = getSheet('일상기록');
  const id = 'DLY-' + new Date().getTime();
  sheet.appendRow([id, data.date, data.keyword, data.content, data.link, data.files]);
  return '일상 기록이 저장되었습니다.';
}

function updateDailyRecord(data) {
  var sheet = getSheet('일상기록');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.getRange(i + 1, 2, 1, 4).setValues([[data.date, data.keyword, data.content, data.link]]);
      return '일상 기록이 수정되었습니다.';
    }
  }
  throw new Error('일상 기록을 찾을 수 없습니다. (ID: ' + data.id + ')');
}

// 학생기록: A=ID, B=기록일시, C=번호, D=이름, E=분류, F=내용, G=지도내용
function saveStudentRecord(data) {
  const sheet = getSheet('학생기록');
  const id = 'STU-' + new Date().getTime();
  const ts = data.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.appendRow([id, ts, data.studentNum, data.studentName, data.category, data.content, '']);
  return '학생 기록이 저장되었습니다.';
}

function updateStudentRecord(data) {
  var sheet = getSheet('학생기록');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.getRange(i + 1, 2, 1, 1).setValues([[data.date]]);
      sheet.getRange(i + 1, 5, 1, 2).setValues([[data.category, data.content]]);
      return '학생 기록이 수정되었습니다.';
    }
  }
  throw new Error('학생 기록을 찾을 수 없습니다. (ID: ' + data.id + ')');
}

// 상담기록: A=ID, B=상담일시, C=번호(쉼표구분), D=이름(쉼표구분), E=상담대상, F=방법, G=내용, H=추후계획
function saveCounselRecord(data) {
  const sheet = getSheet('상담기록');
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const id = 'CNS-' + new Date().getTime();
  // 여러 학생을 쉼표로 묶어 한 행으로 저장
  const nums  = data.studentTags.map(function(s){ var m=s.match(/^(\d+)번/); return m?m[1]:''; });
  const names = data.studentTags.map(function(s){ return s.replace(/^\d+번\s*/, ''); });
  sheet.appendRow([id, ts, nums.join(', '), names.join(', '), data.targetType, data.method, data.content, '']);
  return '상담 기록이 저장되었습니다.';
}

function updateCounselRecord(data) {
  var sheet = getSheet('상담기록');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.getRange(i + 1, 5, 1, 3).setValues([[data.targetType, data.method, data.content]]);
      return '상담 기록이 수정되었습니다.';
    }
  }
  throw new Error('상담 기록을 찾을 수 없습니다. (ID: ' + data.id + ')');
}

// ==========================================
// 8. 기록 삭제 (ID로 행 삭제)
// ==========================================
function deleteRecord(sheetName, rowId) {
  var sheet = getSheet(sheetName);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(rowId)) {
      sheet.deleteRow(i + 1);
      return '삭제되었습니다.';
    }
  }
  throw new Error('해당 기록을 찾을 수 없습니다. (ID: ' + rowId + ')');
}

// ==========================================
// 8. 기록 목록 조회
// ==========================================
function getAttendanceList() {
  var sheet = getSpreadsheet().getSheetByName('출결기록');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      id: String(data[i][0]), date: _fmt(data[i][1], tz),
      studentNum: data[i][2], studentName: String(data[i][3] || ''),
      status: String(data[i][4] || ''), reason: String(data[i][5] || ''),
      proof: String(data[i][6] || '')
    });
  }
  return result;
}

function getClassList() {
  var sheet = getSpreadsheet().getSheetByName('수업기록');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      id: String(data[i][0]), date: _fmt(data[i][1], tz),
      periods: String(data[i][2] || ''), subjects: String(data[i][3] || ''),
      topic: String(data[i][5] || ''), content: String(data[i][6] || ''),
      reflection: String(data[i][7] || ''),
      link: String(data[i][8] || ''), files: String(data[i][9] || '')
    });
  }
  return result;
}

function getDailyList() {
  var sheet = getSpreadsheet().getSheetByName('일상기록');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      id: String(data[i][0]), date: _fmt(data[i][1], tz),
      keyword: String(data[i][2] || ''), content: String(data[i][3] || ''),
      link: String(data[i][4] || ''), files: String(data[i][5] || '')
    });
  }
  return result;
}

function getStudentRecordList() {
  var sheet = getSpreadsheet().getSheetByName('학생기록');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      id: String(data[i][0]), timestamp: _fmtDt(data[i][1], tz),
      studentNum: data[i][2], studentName: String(data[i][3] || ''),
      category: String(data[i][4] || ''), content: String(data[i][5] || '')
    });
  }
  return result;
}

function getCounselList() {
  var sheet = getSpreadsheet().getSheetByName('상담기록');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    // C열(번호), D열(이름)은 단일 또는 쉼표구분 다중값 모두 지원
    var numsRaw  = String(data[i][2] || '');
    var namesRaw = String(data[i][3] || '');
    var nums  = numsRaw.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
    var names = namesRaw.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
    var students = nums.map(function(n, idx){ return { num: n, name: names[idx] || '' }; });
    result.push({
      id: String(data[i][0]), timestamp: _fmtDt(data[i][1], tz),
      studentNum: nums[0] || '', studentName: names[0] || '',
      students: students,
      targetType: String(data[i][4] || ''), method: String(data[i][5] || ''),
      content: String(data[i][6] || '')
    });
  }
  return result;
}

// ==========================================
// 9. 학생 대시보드: 특정 학생 전체 기록
// ==========================================
function getStudentAllRecords(studentNum) {
  var ss     = getSpreadsheet();
  var numInt = parseInt(String(studentNum), 10);
  var tz     = Session.getScriptTimeZone();
  var result = { attendance: [], records: [], counsel: [] };
  function matchNum(v) { return parseInt(String(v), 10) === numInt; }

  var attSheet = ss.getSheetByName('출결기록');
  if (attSheet && attSheet.getLastRow() > 1) {
    var d = attSheet.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (!d[i][0]) continue;
      if (matchNum(d[i][2])) {
        result.attendance.push({ date: _fmt(d[i][1], tz), status: d[i][4], reason: d[i][5]||'', proof: d[i][6]||'' });
      }
    }
  }
  var recSheet = ss.getSheetByName('학생기록');
  if (recSheet && recSheet.getLastRow() > 1) {
    var d2 = recSheet.getDataRange().getValues();
    for (var j = 1; j < d2.length; j++) {
      if (!d2[j][0]) continue;
      if (matchNum(d2[j][2])) {
        result.records.push({ timestamp: _fmtDt(d2[j][1], tz), category: d2[j][4], content: d2[j][5]||'' });
      }
    }
  }
  var cnslSheet = ss.getSheetByName('상담기록');
  if (cnslSheet && cnslSheet.getLastRow() > 1) {
    var d3 = cnslSheet.getDataRange().getValues();
    for (var k = 1; k < d3.length; k++) {
      if (!d3[k][0]) continue;
      // 쉼표구분 다중 학생번호 처리
      var cnslNums = String(d3[k][2]).split(',').map(function(n){ return parseInt(n.trim(), 10); });
      if (cnslNums.indexOf(numInt) !== -1) {
        result.counsel.push({ timestamp: _fmtDt(d3[k][1], tz), targetType: d3[k][4], method: d3[k][5], content: d3[k][6]||'' });
      }
    }
  }
  return result;
}

// ==========================================
// 10. 파일 업로드 (Google Drive)
// ==========================================
function uploadFileToDrive(base64Data, fileName, mimeType) {
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName);
  var folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
  var folder = null;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch(e) { folder = null; }
  }
  if (!folder) {
    folder = DriveApp.createFolder('나세나반 다이어리 첨부파일');
    PropertiesService.getScriptProperties().setProperty('DRIVE_FOLDER_ID', folder.getId());
  }
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var fileId = file.getId();
  // 파일 ID 추출이 쉽도록 일관된 URL 형식으로 반환
  return { url: 'https://drive.google.com/file/d/' + fileId + '/view', name: fileName };
}

// ==========================================
// 11. 전체 검색
// ==========================================
function globalSearch(keyword) {
  if (!keyword || keyword.trim() === '') return [];
  var kw = keyword.toLowerCase().trim();
  var ss = getSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var results = [];

  function search(sheetName, pageId, typeName, titleCol, contentCol, dateCol) {
    try {
      var s = ss.getSheetByName(sheetName);
      if (!s || s.getLastRow() <= 1) return;
      var d = s.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        if (!d[i][0]) continue;
        var t = String(d[i][titleCol] || '');
        var c = String(d[i][contentCol] || '');
        if (t.toLowerCase().indexOf(kw) >= 0 || c.toLowerCase().indexOf(kw) >= 0) {
          results.push({ type: typeName, pageId: pageId,
            date: _fmt(d[i][dateCol], tz), label: t || '(제목없음)',
            content: c.length > 80 ? c.substring(0,80) + '…' : c });
        }
      }
    } catch(e) {}
  }

  // 일정: A=날짜(0), B=제목(1), D=내용(3)
  search('일정',    'page-schedule',       '일정',   1, 3, 0);
  // 수업기록: B=날짜(1), F=배움주제(5), G=수업내용(6), H=성찰(7)
  search('수업기록', 'page-class',         '수업기록', 5, 6, 1);
  // 일상기록: B=날짜(1), C=키워드(2), D=내용(3)
  search('일상기록', 'page-daily',         '일상기록', 2, 3, 1);
  // 학생기록: B=기록일시(1), D=이름(3), F=내용(5)
  search('학생기록', 'page-student-record','학생기록', 3, 5, 1);
  // 상담기록: B=상담일시(1), D=이름(3), G=내용(6)
  search('상담기록', 'page-counsel',       '상담기록', 3, 6, 1);

  return results;
}
// ==========================================
// 12. 디지털 도구 조회
// ==========================================
// 디지털 도구 시트 컬럼: A=분류, B=도구명, C=URL, D=설명
function getToolList() {
  var ss = getSpreadsheet();
  // 주의: 스프레드시트의 시트 이름이 정확히 '디지털 도구'여야 합니다. 띄어쓰기 확인!
  var sheet = ss.getSheetByName('디지털 도구'); 
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  var data = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    if (!data[i][1]) continue; // 도구명이 비어있으면 건너뜀
    result.push({
      category: String(data[i][0] || '기타'),
      name: String(data[i][1] || ''),
      url: String(data[i][2] || '#'),
      description: String(data[i][3] || '')
    });
  }
  return result;
}
// ==========================================
// 13. 자리배치 데이터 저장 및 불러오기
// ==========================================
function saveSeatingData(dataString) {
  // 스크립트 내부 속성(Properties)에 데이터를 안전하게 저장합니다.
  PropertiesService.getScriptProperties().setProperty('SEATING_DATA', dataString);
  return '자리 배치가 안전하게 저장되었습니다.';
}

function getSeatingData() {
  // 저장된 데이터를 불러옵니다. 없으면 빈 객체를 반환합니다.
  return PropertiesService.getScriptProperties().getProperty('SEATING_DATA') || '{}';
}

// ==========================================
// 14. 제출확인 데이터 저장 및 불러오기
// ==========================================
function saveSubmissionData(dataString) {
  PropertiesService.getScriptProperties().setProperty('SUBMISSION_DATA', dataString);
  return '제출 데이터가 저장되었습니다.';
}

function getSubmissionData() {
  return PropertiesService.getScriptProperties().getProperty('SUBMISSION_DATA') || '[]';
}

// ==========================================
// 15. 주간학습안내 이미지 저장 및 불러오기
// ==========================================
function saveWeeklyNotice(dataObj) {
  // dataObj: { url: string, name: string }
  PropertiesService.getScriptProperties().setProperty('WEEKLY_NOTICE', JSON.stringify(dataObj));
  return '주간학습안내가 저장되었습니다.';
}

function getWeeklyNotice() {
  return PropertiesService.getScriptProperties().getProperty('WEEKLY_NOTICE') || '';
}

function clearWeeklyNotice() {
  PropertiesService.getScriptProperties().deleteProperty('WEEKLY_NOTICE');
  return '주간학습안내가 삭제되었습니다.';
}
