/**
 * 꼬꼬챌린지 데이터 관리 (Google Apps Script)
 * 학교별 4개 시트 (총 12개 시트 구조):
 *  - A초, B초, C초 (일일기록: 학생 위, 교사 아래 t-전화번호4자리)
 *  - A초_월별성장, B초_월별성장, C초_월별성장 (내몸탐험)
 *  - A초_학생설문응답, B초_학생설문응답, C초_학생설문응답
 *  - A초_교사설문응답, B초_교사설문응답, C초_교사설문응답
 */

const TZ = 'Asia/Seoul';
const SCHOOLS = ['A초', 'B초', 'C초'];

const HEADERS = {
  '일별기록': ['중복키','일시','학교','학생키','개인번호','이름','구분','날짜','포인트','스티커'],
  '월별성장': ['중복키','측정일시','학교','학생키','개인번호','이름','구분','측정월','키(cm)','몸무게(kg)','BMI'],
  '설문응답': ['응답일시','학교','학생키','개인번호','이름','구분','설문구분',
               '문항1','문항2','문항3','문항4','문항5','문항6',
               '문항7','문항8','문항9','문항10','문항11','문항12']
};

function getSS() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1iAt-AhzYwCIW-si0-BtmVKI1vUujTHzX-e2w_t_3Uvc');
  } catch(e) {
    return SpreadsheetApp.openById('1iAt-AhzYwCIW-si0-BtmVKI1vUujTHzX-e2w_t_3Uvc');
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📁 꼬꼬챌린지 관리')
    .addItem('① 학교별 12개 시트 준비', 'setup')
    .addSeparator()
    .addItem('💾 드라이브에 백업', 'backupToDrive')
    .addItem('🧹 일별기록 정렬 (학생 위 / 교사 아래)', 'sortDailyAll')
    .addToUi();
}

function setup() {
  var SS = getSS();
  
  // 글로벌 통합 시트 자동 정리 (생성 방지 및 삭제)
  var deleteGlobalTabs = [
    '학생기록', '교사기록', '학생일별스티커', '교사일별스티커',
    '학생설문응답', '교사설문응답', '명단', '일별기록', '월별성장', '설문응답'
  ];
  deleteGlobalTabs.forEach(function(name) {
    var target = SS.getSheetByName(name);
    if (target) {
      try { SS.deleteSheet(target); } catch(e) {}
    }
  });

  // 학교별 4개 시트 생성 (총 3학교 * 4 = 12개 시트)
  SCHOOLS.forEach(function(school) {
    // 1. 일일기록 ({학교})
    var sh1 = SS.getSheetByName(school) || SS.insertSheet(school);
    sh1.getRange(1, 1, 1, 10).setValues([HEADERS['일별기록']])
      .setFontWeight('bold').setBackground('#e8f0fe');
    sh1.setFrozenRows(1);

    // 2. 내몸탐험 ({학교}_월별성장)
    var sh2 = SS.getSheetByName(school + '_월별성장') || SS.getSheetByName(school + '_내몸탐험') || SS.insertSheet(school + '_월별성장');
    sh2.getRange(1, 1, 1, 11).setValues([HEADERS['월별성장']])
      .setFontWeight('bold').setBackground('#e2f0d9');
    sh2.setFrozenRows(1);

    // 3. 학생설문응답 ({학교}_학생설문응답)
    var sh3 = SS.getSheetByName(school + '_학생설문응답') || SS.insertSheet(school + '_학생설문응답');
    sh3.getRange(1, 1, 1, 19).setValues([HEADERS['설문응답']])
      .setFontWeight('bold').setBackground('#fff2cc');
    sh3.setFrozenRows(1);

    // 4. 교사설문응답 ({학교}_교사설문응답)
    var sh4 = SS.getSheetByName(school + '_교사설문응답') || SS.insertSheet(school + '_교사설문응답');
    sh4.getRange(1, 1, 1, 19).setValues([HEADERS['설문응답']])
      .setFontWeight('bold').setBackground('#f8cbad');
    sh4.setFrozenRows(1);
  });
  Logger.log('학교별 4개 시트 (총 12개 시트) 준비 완료');
}

function backupToDrive() {
  var SS = getSS();
  var name = '[꼬꼬챌린지 백업] ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HHmmss');
  DriveApp.getFileById(SS.getId()).makeCopy(name);
  Logger.log('백업 완료: ' + name);
}

/* ── 공통 ────────────────────────────── */

function parseId(rawId, fallbackSchool) {
  var raw = String(rawId || '').trim();
  var school = '';
  var m = raw.match(/^([A-Da-d])초?[_-]/);
  if (m) school = m[1].toUpperCase() + '초';

  var id = raw.replace(/^[A-Da-d]초[_-]/i, '').replace(/^[A-Da-d][_-]/i, '').trim();

  if (!school) {
    school = String(fallbackSchool || '').trim();
    if (school && school.slice(-1) !== '초') school += '초';
  }
  if (SCHOOLS.indexOf(school) < 0) school = 'A초';

  var type;
  if (/^T-?\d+$/i.test(id) || /^t-?\d+$/i.test(id)) {
    type = '2.교사';
    var numOnly = id.replace(/^[Tt]-?/i, '');
    id = 't-' + numOnly;
  } else if (/^guest$/i.test(id)) {
    type = '3.게스트';
    id = 'guest';
  } else {
    type = '1.학생';
  }

  return { school: school, id: id, type: type, key: school + '-' + id };
}

function resolveName(key, given, id) {
  return String(given || '').trim() || id;
}

function upsert(sh, dupKey, row) {
  if (!sh) return '시트없음';
  var last = sh.getLastRow();
  var targetKey = String(dupKey).trim();
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === targetKey) {
        sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
        return '덮어씀';
      }
    }
  }
  appendOrInsertRow(sh, row);
  return '추가함';
}

function countStickers(key, schoolName) {
  var SS = getSS();
  var sh = SS.getSheetByName(schoolName || 'A초');
  if (!sh || sh.getLastRow() < 2) return 0;
  var targetKey = String(key || '').trim();
  if (!targetKey) return 0;

  var lastRow = sh.getLastRow();
  var v = sh.getRange(2, 1, lastRow - 1, Math.min(10, sh.getLastColumn())).getValues();
  var n = 0;
  for (var i = 0; i < v.length; i++) {
    var studentKey = String(v[i][3] || v[i][2] || '').trim();
    if (studentKey === targetKey || studentKey.indexOf(targetKey) > -1) {
      var sVal = Number(v[i][v[i].length - 1]);
      n += (!isNaN(sVal) && sVal > 0) ? sVal : 1;
    }
  }
  return n;
}

function levelOf(n) {
  if (n >= 45) return '4단계 꼬꼬대장';
  if (n >= 24) return '3단계 튼튼이';
  if (n >= 8)  return '2단계 삐약이';
  return '1단계 알콩이';
}

function ymd(d)   { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function stamp(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm:ss'); }

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function appendOrInsertRow(sheet, rowData) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    sheet.appendRow(rowData);
    return;
  }
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();
  var targetRow = -1;
  for (var i = 1; i < values.length; i++) {
    var val = String(values[i][0] || '').trim();
    if (!val) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow > 1) {
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

// 학생(1.학생)은 위, 교사(2.교사)는 아래로 일일기록 정렬
function sortDailySheet(sh) {
  if (!sh || sh.getLastRow() < 3) return;
  var headers = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0];
  var firstColName = String(headers[0] || '').trim();

  if (firstColName === '중복키') {
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
      .sort([{ column: 7, ascending: true }, { column: 2, ascending: true }]);
  } else {
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
      .sort([{ column: 5, ascending: true }, { column: 1, ascending: true }]);
  }
}

function sortDailyAll() {
  var SS = getSS();
  SCHOOLS.forEach(function(school) {
    var sh = SS.getSheetByName(school);
    if (sh) sortDailySheet(sh);
  });
}

/* ── 앱 요청 처리 ────────────────────── */

function doPost(e) {
  var lock = LockService.getScriptLock();
  var got = lock.tryLock(30000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.action === 'log_mission') return json(saveDaily(d));
    if (d.action === 'update_bmi')  return json(saveGrowth(d));
    if (d.action === 'save_survey' || d.action === 'submit_survey') return json(saveSurvey(d));
    return json({ success: true, message: '수신 완료' });
  } catch (err) {
    return json({ success: false, error: String(err) });
  } finally {
    if (got) lock.releaseLock();
  }
}

function saveDaily(d) {
  var SS = getSS();
  var p = parseId(d.studentId, d.school);
  var schoolName = p.school || "A초";
  var name = resolveName(p.key, d.name, p.id);
  var pts = Number(d.points || 0);
  var sticker = (pts >= 50 || Number(d.dailySticker || 0) >= 1) ? 1 : 0;
  var now = new Date(), date = ymd(now);
  var dupKey = p.key + '|' + date;

  var sh = SS.getSheetByName(schoolName) || SS.insertSheet(schoolName);
  var headers = [];
  if (sh.getLastRow() >= 1) {
    headers = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0];
  }
  var firstColName = String(headers[0] || '').trim();

  if (firstColName === '중복키') {
    var row = [dupKey, stamp(now), p.school, p.key, p.id, name, p.type, date, pts, sticker];
    upsert(sh, dupKey, row);
  } else {
    var row = [stamp(now), p.school, "'" + p.id, name, p.type, sticker > 0 ? sticker : pts];
    appendOrInsertRow(sh, row);
  }

  // 학생(1.학생) 위, 교사(2.교사) 아래 자동 정렬
  sortDailySheet(sh);

  var total = countStickers(p.key, schoolName);
  return { success: true, 이름: name, 구분: p.type, 포인트: pts, 총스티커: total, 레벨: levelOf(total) };
}

function saveGrowth(d) {
  var SS = getSS();
  var p = parseId(d.studentId, d.school);
  var schoolName = p.school || "A초";
  var name = resolveName(p.key, d.name, p.id);
  var now = new Date();
  var month = Utilities.formatDate(now, TZ, 'yyyy-MM');
  var h = Number(d.height || 0), w = Number(d.weight || 0);
  var bmi = (h > 0 && w > 0) ? Math.round(w / Math.pow(h / 100, 2) * 10) / 10 : '';

  var sh = SS.getSheetByName(schoolName + '_월별성장') || SS.getSheetByName(schoolName + '_내몸탐험') || SS.insertSheet(schoolName + '_월별성장');
  upsert(sh, p.key + '|' + month,
    [p.key + '|' + month, stamp(now), p.school, p.key, p.id, name, p.type, month, h, w, bmi]);

  var total = countStickers(p.key, schoolName);
  return { success: true, 이름: name, 구분: p.type, BMI: bmi, 총스티커: total, 레벨: levelOf(total) };
}

function saveSurvey(d) {
  var SS = getSS();
  var p = parseId(d.studentId, d.school);
  var schoolName = p.school || "A초";
  var isTeacher = p.type === '2.교사';
  
  var targetSheetName = schoolName + (isTeacher ? "_교사설문응답" : "_학생설문응답");
  var sh = SS.getSheetByName(targetSheetName) || SS.insertSheet(targetSheetName);

  var name = resolveName(p.key, d.name, p.id);
  var rawAnswers = d.answers || d.surveyAnswers || [];
  var ansList = [];
  if (Object.prototype.toString.call(rawAnswers) === '[object Array]') {
    ansList = rawAnswers;
  } else if (typeof rawAnswers === 'string') {
    try {
      ansList = JSON.parse(rawAnswers);
      if (Object.prototype.toString.call(ansList) !== '[object Array]') ansList = [rawAnswers];
    } catch (x) { ansList = [rawAnswers]; }
  }
  
  var row = [stamp(new Date()), p.school, p.key, p.id, name, p.type, d.surveyType || '사전설문'];
  for (var i = 0; i < 12; i++) {
    var val = (ansList && ansList[i] !== undefined && ansList[i] !== null) ? String(ansList[i]).trim() : '';
    row.push(val);
  }
  appendOrInsertRow(sh, row);
  
  return { success: true, message: '설문 저장 완료 (' + targetSheetName + ')' };
}

function doGet(e) {
  try {
    var raw = String((e.parameter && e.parameter.studentId) || '').trim();
    var p = parseId(raw, e.parameter && e.parameter.school);
    var schoolName = p.school || 'A초';
    var total = countStickers(p.key, schoolName);
    var SS = getSS();
    var preSurveyDone = false;
    var isTeacher = p.type === '2.교사';
    var surveySheetName = schoolName + (isTeacher ? '_교사설문응답' : '_학생설문응답');
    var sh = SS.getSheetByName(surveySheetName) || SS.getSheetByName(schoolName + '_설문응답');
    if (sh && sh.getLastRow() >= 2) {
      var sRows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
      for (var r = 0; r < sRows.length; r++) {
        var rowSch = String(sRows[r][1] || '').trim();
        var rowKey = String(sRows[r][2] || '').trim();
        var rowId = String(sRows[r][3] || '').trim();
        if (rowKey === p.key || (rowId === p.id && (rowSch === p.school || !rowSch))) {
          preSurveyDone = true;
          break;
        }
      }
    }
    return json({ success: true, student: {
      id: raw, name: resolveName(p.key, '', p.id),
      totalSticker: total, level: levelOf(total), preSurveyDone: preSurveyDone } });
  } catch (err) {
    return json({ success: true, student: { id: 'guest', name: '학생', totalSticker: 0, preSurveyDone: false } });
  }
}