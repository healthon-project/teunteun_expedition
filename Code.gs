/**
 * 꼬꼬챌린지 데이터 관리 (Google Apps Script)
 * 학교별 시트 명확 매핑 & 데이터 분리:
 *  1) {학교} (예: A초, B초, C초, D초): 일일 로그인/미션 기록 (당일 1줄만 보존 / 스티커 1개 / 학생 위, 교사 t-전화번호4자리 아래 정렬)
 *  2) {학교}_내몸탐험 (예: A초_내몸탐험): 월1회 신체기록 (키, 몸무게, BMI, 총스티커, 레벨 / 교사 t-전화번호4자리)
 *  3) {학교}_설문응답 (예: A초_설문응답): 학생 + 교사 통합 설문응답 (교사 t-전화번호4자리 / '사전설문' 제거 / 맨윗줄 헤더 보존)
 */

const TZ = 'Asia/Seoul';
const SCHOOLS = ['A초', 'B초', 'C초', 'D초'];

const HEADERS = {
  '일별기록': ['중복키','일시','학교','학생키','개인번호','이름','구분','날짜','포인트','스티커'],
  '월별성장': ['중복키','측정일시','학교','학생키','개인번호','이름','구분','측정월','키(cm)','몸무게(kg)','BMI','총스티커','레벨'],
  '설문응답': ['응답일시','학교','개인번호','이름','구분','문항1','문항2','문항3','문항4','문항5','문항6','문항7','문항8','문항9','문항10','문항11','문항12']
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
    .addItem('① 시트 연결 및 헤더 점검', 'setup')
    .addSeparator()
    .addItem('💾 드라이브에 백업', 'backupToDrive')
    .addItem('🧹 일별기록 정렬 (학생 위 / 교사 아래)', 'sortDailyAll')
    .addToUi();
}

function parseSchoolName(schoolInput) {
  var str = String(schoolInput || 'A초').trim().toUpperCase();
  var m = str.match(/^([A-Za-z0-9가-힣]+)초?/);
  if (m) {
    var name = m[1];
    if (name.indexOf('초') < 0) name += '초';
    return name;
  }
  return 'A초';
}

/**
 * 1. 일일기록 시트 전용 탐색 (예: "A초", "B초", "C초")
 * 언더바(_)가 없는 순수 학교 시트에만 연결되도록 안전 보장!
 */
function getDailySheet(SS, school) {
  var name = parseSchoolName(school);
  
  // 1) 정확한 이름 매칭 ("A초")
  var sh = SS.getSheetByName(name);
  if (sh) return sh;
  
  // 2) 대소문자 및 트림 매칭 (언더바 없는 탭 검색)
  var sheets = SS.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    if (sName.toUpperCase() === name.toUpperCase()) return sheets[i];
  }
  
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    if (sName.toUpperCase().indexOf(name.toUpperCase()) === 0 && sName.indexOf('_') < 0) {
      return sheets[i];
    }
  }

  // 기존 시트 없을 경우 생성
  var newSh = SS.insertSheet(name);
  newSh.getRange(1, 1, 1, 10).setValues([HEADERS['일별기록']]).setFontWeight('bold').setBackground('#e8f0fe');
  newSh.setFrozenRows(1);
  return newSh;
}

/**
 * 2. 내몸탐험 시트 전용 탐색 (예: "A초_내몸탐험")
 */
function getGrowthSheet(SS, school) {
  var name = parseSchoolName(school);
  var sh = SS.getSheetByName(name + '_내몸탐험');
  if (sh) return sh;

  var sheets = SS.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    if (sName.toUpperCase().indexOf(name.toUpperCase()) === 0 && (sName.indexOf('내몸탐험') >= 0 || sName.indexOf('성장') >= 0)) {
      return sheets[i];
    }
  }

  var newSh = SS.insertSheet(name + '_내몸탐험');
  newSh.getRange(1, 1, 1, 13).setValues([HEADERS['월별성장']]).setFontWeight('bold').setBackground('#e2f0d9');
  newSh.setFrozenRows(1);
  return newSh;
}

/**
 * 3. 설문응답 시트 전용 탐색 (예: "A초_설문응답")
 */
function getSurveySheet(SS, school) {
  var name = parseSchoolName(school);
  var sh = SS.getSheetByName(name + '_설문응답');
  if (sh) return sh;

  var sheets = SS.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    if (sName.toUpperCase().indexOf(name.toUpperCase()) === 0 && sName.indexOf('설문') >= 0) {
      return sheets[i];
    }
  }

  var newSh = SS.insertSheet(name + '_설문응답');
  newSh.getRange(1, 1, 1, 17).setValues([HEADERS['설문응답']]).setFontWeight('bold').setBackground('#fff2cc');
  newSh.setFrozenRows(1);
  return newSh;
}

function setup() {
  var SS = getSS();
  SCHOOLS.forEach(function(school) {
    var sh1 = getDailySheet(SS, school);
    if (sh1.getLastRow() === 0) {
      sh1.getRange(1, 1, 1, 10).setValues([HEADERS['일별기록']]).setFontWeight('bold').setBackground('#e8f0fe');
      sh1.setFrozenRows(1);
    }

    var sh2 = getGrowthSheet(SS, school);
    if (sh2.getLastRow() === 0) {
      sh2.getRange(1, 1, 1, 13).setValues([HEADERS['월별성장']]).setFontWeight('bold').setBackground('#e2f0d9');
      sh2.setFrozenRows(1);
    }

    var sh3 = getSurveySheet(SS, school);
    cleanAndFixSurveySheet(sh3);
  });
  Logger.log('학교별 시트 연결 점검 완료');
}

function cleanAndFixSurveySheet(sh) {
  if (!sh) return;
  
  var expectedHeaders = HEADERS['설문응답'];
  var lastRow = sh.getLastRow();
  
  if (lastRow === 0) {
    sh.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders])
      .setFontWeight('bold').setBackground('#fff2cc');
    sh.setFrozenRows(1);
    return;
  }

  var firstCell = String(sh.getRange(1, 1).getValue() || '').trim();

  if (firstCell !== '응답일시') {
    sh.insertRowBefore(1);
    sh.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders])
      .setFontWeight('bold').setBackground('#fff2cc');
    sh.setFrozenRows(1);
    lastRow = sh.getLastRow();
  } else {
    sh.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders])
      .setFontWeight('bold').setBackground('#fff2cc');
    sh.setFrozenRows(1);
  }

  if (lastRow >= 2) {
    var maxCols = Math.max(18, sh.getLastColumn());
    var dataRange = sh.getRange(2, 1, lastRow - 1, maxCols);
    var rows = dataRange.getValues();
    var modified = false;

    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      var foundIdx = -1;
      for (var c = 0; c < row.length; c++) {
        if (String(row[c] || '').trim() === '사전설문') {
          foundIdx = c;
          break;
        }
      }
      if (foundIdx !== -1) {
        row.splice(foundIdx, 1);
        while (row.length < 17) row.push('');
        rows[r] = row.slice(0, 17);
        modified = true;
      }
    }

    if (modified) {
      sh.getRange(2, 1, rows.length, 17).setValues(rows);
    }
  }
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
  var m = raw.match(/^([A-Za-z0-9가-힣]+)초?[_-]/);
  if (m) {
    school = m[1].toUpperCase();
    if (school.indexOf('초') < 0) school += '초';
  }

  var id = raw.replace(/^[A-Za-z0-9가-힣]+초[_-]/i, '').replace(/^[A-Za-z0-9가-힣]+[_-]/i, '').trim();

  if (!school) {
    school = String(fallbackSchool || '').trim();
    if (school && school.indexOf('초') < 0) school += '초';
  }
  if (!school) school = 'A초';

  var type;
  if (/^T-?\d+$/i.test(id) || /^t-?\d+$/i.test(id) || raw.indexOf('교사') >= 0 || raw.indexOf('teacher') >= 0) {
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

function formatDateStr(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, TZ, 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  var m = s.match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : s;
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

function upsertDailyRow(sh, p, name, date, now, pts, sticker) {
  if (!sh) return;
  var lastRow = sh.getLastRow();
  var todayDateStr = ymd(now);
  var targetKey = String(p.key || '').trim();
  var targetId = String(p.id || '').trim();
  var finalSticker = 1;

  if (lastRow >= 2) {
    var data = sh.getRange(2, 1, lastRow - 1, Math.max(10, sh.getLastColumn())).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rDate1 = formatDateStr(row[0]);
      var rDate2 = formatDateStr(row[1]);
      var rDate8 = formatDateStr(row[7]);
      var isToday = (rDate1 === todayDateStr || rDate2 === todayDateStr || rDate8 === todayDateStr);

      var rDupKey = String(row[0] || '').trim();
      var rKey = String(row[3] || '').trim();
      var rId1 = String(row[2] || '').trim().replace(/^'/, '');
      var rId2 = String(row[4] || '').trim().replace(/^'/, '');

      var isSameStudent = (rDupKey.indexOf(targetKey) === 0 || rKey === targetKey || rId1 === targetId || rId2 === targetId);

      if (isToday && isSameStudent) {
        var targetRowIndex = i + 2;
        var headers = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0];
        var firstColName = String(headers[0] || '').trim();

        if (firstColName === '중복키') {
          var updatedRow = [targetKey + '|' + todayDateStr, stamp(now), p.school, targetKey, targetId, name, p.type, todayDateStr, pts, finalSticker];
          sh.getRange(targetRowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
        } else {
          var updatedRow = [stamp(now), p.school, "'" + targetId, name, p.type, finalSticker];
          sh.getRange(targetRowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
        }
        return 'updated';
      }
    }
  }

  var headers = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0];
  var firstColName = String(headers[0] || '').trim();
  if (firstColName === '중복키') {
    var newRow = [targetKey + '|' + todayDateStr, stamp(now), p.school, targetKey, targetId, name, p.type, todayDateStr, pts, finalSticker];
    appendOrInsertRow(sh, newRow);
  } else {
    var newRow = [stamp(now), p.school, "'" + targetId, name, p.type, finalSticker];
    appendOrInsertRow(sh, newRow);
  }
  return 'inserted';
}

function countStickers(key, schoolName) {
  var SS = getSS();
  var sh = getDailySheet(SS, schoolName || 'A초');
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

// 학생(1.학생)은 위, 교사(2.교사)는 아래로 정렬
function sortDailySheet(sh) {
  if (!sh || sh.getLastRow() < 3) return;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var typeColIdx = -1;
  var dateColIdx = 1;

  for (var c = 0; c < headers.length; c++) {
    var hName = String(headers[c] || '').trim();
    if (hName === '구분') typeColIdx = c + 1;
    if (hName === '일시' || hName === '날짜' || hName === '응답일시') dateColIdx = c + 1;
  }

  if (typeColIdx > 0) {
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
      .sort([{ column: typeColIdx, ascending: true }, { column: dateColIdx, ascending: true }]);
  }
}

function sortDailyAll() {
  var SS = getSS();
  SCHOOLS.forEach(function(school) {
    var sh = getDailySheet(SS, school);
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

// 1. 일일기록 저장 ({학교}: A초, B초, C초 시트에 무조건 저장을 보장!)
function saveDaily(d) {
  var SS = getSS();
  var p = parseId(d.studentId, d.school);
  var schoolName = p.school || "A초";
  var isTeacher = p.type === '2.교사' || String(d.role || d.type || '').indexOf('교사') >= 0;
  
  var name = resolveName(p.key, d.name, p.id);
  var idToSave = p.id;

  if (isTeacher) {
    p.type = '2.교사';
    if (idToSave.indexOf('t-') !== 0) idToSave = 't-' + idToSave.replace(/^[Tt]-?/, '');
    if (name.indexOf('t-') !== 0) name = 't-' + name;
    p.id = idToSave;
    p.key = schoolName + '-' + idToSave;
  }

  var pts = Number(d.points || 0);
  var sticker = 1;
  var now = new Date(), date = ymd(now);

  // 무조건 순수 학교 시트(A초, B초, C초 등)에 기록! (절대 _내몸탐험이나 _설문응답으로 가지 않음)
  var sh = getDailySheet(SS, schoolName);
  upsertDailyRow(sh, p, name, date, now, pts, sticker);
  sortDailySheet(sh);

  var total = countStickers(p.key, schoolName);
  return { success: true, 이름: name, 구분: p.type, 포인트: pts, 총스티커: total, 레벨: levelOf(total) };
}

// 2. 내몸탐험 저장 ({학교}_내몸탐험: 월1회 키, 몸무게, BMI, 총스티커, 레벨)
function saveGrowth(d) {
  var SS = getSS();
  var p = parseId(d.studentId, d.school);
  var schoolName = p.school || "A초";
  var isTeacher = p.type === '2.교사' || String(d.role || d.type || '').indexOf('교사') >= 0;

  var name = resolveName(p.key, d.name, p.id);
  var idToSave = p.id;

  if (isTeacher) {
    p.type = '2.교사';
    if (idToSave.indexOf('t-') !== 0) idToSave = 't-' + idToSave.replace(/^[Tt]-?/, '');
    if (name.indexOf('t-') !== 0) name = 't-' + name;
    p.id = idToSave;
    p.key = schoolName + '-' + idToSave;
  }

  var now = new Date();
  var month = Utilities.formatDate(now, TZ, 'yyyy-MM');
  var h = Number(d.height || 0), w = Number(d.weight || 0);
  var bmi = (h > 0 && w > 0) ? Math.round(w / Math.pow(h / 100, 2) * 10) / 10 : '';

  var total = countStickers(p.key, schoolName);
  var lvl = levelOf(total);

  var sh = getGrowthSheet(SS, schoolName);
  
  if (sh.getLastRow() >= 1) {
    var headers = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0];
    var firstColName = String(headers[0] || '').trim();

    if (firstColName === '중복키') {
      upsert(sh, p.key + '|' + month,
        [p.key + '|' + month, stamp(now), p.school, p.key, idToSave, name, p.type, month, h, w, bmi, total, lvl]);
    } else {
      upsert(sh, p.key + '|' + month,
        [stamp(now), p.school, "'" + idToSave, name, p.type, month, h, w, bmi, total, lvl]);
    }
  } else {
    sh.getRange(1, 1, 1, 13).setValues([HEADERS['월별성장']]).setFontWeight('bold').setBackground('#e2f0d9');
    sh.setFrozenRows(1);
    upsert(sh, p.key + '|' + month,
      [p.key + '|' + month, stamp(now), p.school, p.key, idToSave, name, p.type, month, h, w, bmi, total, lvl]);
  }

  return { success: true, 이름: name, 구분: p.type, BMI: bmi, 총스티커: total, 레벨: lvl };
}

// 3. 설문응답 저장 ({학교}_설문응답)
function saveSurvey(d) {
  var SS = getSS();
  var p = parseId(d.studentId, d.school);
  var schoolName = p.school || "A초";
  var isTeacher = p.type === '2.교사' || String(d.role || d.type || '').indexOf('교사') >= 0;

  var name = resolveName(p.key, d.name, p.id);
  var idToSave = p.id;

  if (isTeacher) {
    p.type = '2.교사';
    if (idToSave.indexOf('t-') !== 0) idToSave = 't-' + idToSave.replace(/^[Tt]-?/, '');
    if (name.indexOf('t-') !== 0) name = 't-' + name;
    p.id = idToSave;
  }

  var sh = getSurveySheet(SS, schoolName);
  cleanAndFixSurveySheet(sh);

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

  var row = [stamp(new Date()), p.school, "'" + idToSave, name, p.type];
  for (var i = 0; i < 12; i++) {
    var val = (ansList && ansList[i] !== undefined && ansList[i] !== null) ? String(ansList[i]).trim() : '';
    if (val === '사전설문') val = '';
    row.push(val);
  }
  appendOrInsertRow(sh, row);
  
  return { success: true, message: '설문 저장 완료 (' + sh.getName() + ')' };
}

function doGet(e) {
  try {
    var raw = String((e.parameter && e.parameter.studentId) || '').trim();
    var p = parseId(raw, e.parameter && e.parameter.school);
    var schoolName = p.school || 'A초';
    var total = countStickers(p.key, schoolName);
    var SS = getSS();
    var preSurveyDone = false;
    var isTeacher = p.type === '2.교사' || String(e.parameter && e.parameter.role || '').indexOf('교사') >= 0;

    var sh = getSurveySheet(SS, schoolName);
    if (sh && sh.getLastRow() >= 2) {
      var sRows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
      for (var r = 0; r < sRows.length; r++) {
        var rowSch = String(sRows[r][1] || '').trim();
        var rowId = String(sRows[r][2] || '').trim().replace(/^'/, '');
        var targetId = p.id;
        if (isTeacher && targetId.indexOf('t-') !== 0) targetId = 't-' + targetId.replace(/^[Tt]-?/, '');
        if ((rowId === targetId || rowId === p.id) && (rowSch === p.school || !rowSch)) {
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