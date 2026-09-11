/**
 * 꼬꼬챌린지 데이터 관리
 * 시트: 명단 / 일별기록 / 월별성장 / 설문응답
 */

const TZ = 'Asia/Seoul';
const OLD_SCHOOLS = ['A초', 'B초', 'C초', 'D초'];

const SH = { roster: '명단', daily: '일별기록', growth: '월별성장', survey: '설문응답' };

const HEADERS = {
  '명단':     ['학생키','학교','개인번호','이름','구분','학년'],
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
    .addItem('① 시트 준비', 'setup')
    .addItem('② 기존자료 변환', 'migrate')
    .addSeparator()
    .addItem('💾 드라이브에 백업', 'backupToDrive')
    .addItem('🧹 일별기록 정렬', 'sortDaily')
    .addToUi();
}

function setup() {
  var SS = getSS();
  Object.entries(HEADERS).forEach(function (pair) {
    var name = pair[0], header = pair[1];
    var sh = SS.getSheetByName(name) || SS.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
  });
  Logger.log('시트 준비 완료');
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
  if (OLD_SCHOOLS.indexOf(school) < 0) school = 'A초';

  var type;
  if (/^T-?\d+$/i.test(id)) { type = '2.교사'; id = 'T-' + id.replace(/^T-?/i, ''); }
  else if (/^guest$/i.test(id)) { type = '3.게스트'; id = 'guest'; }
  else { type = '1.학생'; }

  return { school: school, id: id, type: type, key: school + '-' + id };
}

var _roster = null;
function rosterName(key) {
  if (!_roster) {
    _roster = {};
    var SS = getSS();
    var sh = SS.getSheetByName(SH.roster);
    if (sh && sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
        if (r[0]) _roster[String(r[0]).trim()] = String(r[3]).trim();
      });
    }
  }
  return _roster[key] || '';
}

function resolveName(key, given, id) {
  return rosterName(key) || String(given || '').trim() || id;
}

function upsert(sh, dupKey, row) {
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
  sh.appendRow(row);
  return '추가함';
}

function countStickers(key) {
  var SS = getSS();
  var sh = SS.getSheetByName(SH.daily);
  if (!sh || sh.getLastRow() < 2) return 0;
  var targetKey = String(key || '').trim();
  if (!targetKey) return 0;

  var lastRow = sh.getLastRow();
  var v = sh.getRange(2, 4, lastRow - 1, 7).getValues();
  var n = 0;
  for (var i = 0; i < v.length; i++) {
    var studentKey = String(v[i][0] || '').trim();
    if (studentKey === targetKey) {
      var sVal = Number(v[i][6]);
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
  var name = resolveName(p.key, d.name, p.id);
  var pts = Number(d.points || 0);
  if (pts < 50 && Number(d.dailySticker || 0) < 1) {
    return { success: true, message: '포인트 미달' };
  }
  var now = new Date(), date = ymd(now);
  upsert(SS.getSheetByName(SH.daily), p.key + '|' + date,
    [p.key + '|' + date, stamp(now), p.school, p.key, p.id, name, p.type, date, pts, 1]);
  var total = countStickers(p.key);
  return { success: true, 이름: name, 총스티커: total, 레벨: levelOf(total) };
}

function saveGrowth(d) {
  var SS = getSS();
  var p = parseId(d.studentId, d.school);
  var name = resolveName(p.key, d.name, p.id);
  var now = new Date();
  var month = Utilities.formatDate(now, TZ, 'yyyy-MM');
  var h = Number(d.height || 0), w = Number(d.weight || 0);
  var bmi = (h > 0 && w > 0) ? Math.round(w / Math.pow(h / 100, 2) * 10) / 10 : '';
  upsert(SS.getSheetByName(SH.growth), p.key + '|' + month,
    [p.key + '|' + month, stamp(now), p.school, p.key, p.id, name, p.type, month, h, w, bmi]);
  var total = countStickers(p.key);
  return { success: true, 이름: name, BMI: bmi, 총스티커: total, 레벨: levelOf(total) };
}

function saveSurvey(d) {
  var SS = getSS();
  var p = parseId(d.studentId, d.school);
  var schoolName = p.school || "A초";
  var targetName = schoolName + "_설문응답";
  var sh = SS.getSheetByName(targetName) || SS.getSheetByName(p.type === '2.교사' ? '교사설문응답' : '학생설문응답') || SS.getSheetByName(SH.survey);
  if (!sh) {
    sh = SS.insertSheet(targetName);
    sh.appendRow(["응답일시", "학교", "개인번호", "이름", "설문구분", "문항1", "문항2", "문항3", "문항4", "문항5", "문항6", "문항7", "문항8", "문항9", "문항10", "문항11", "문항12"]);
  }
  
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
  
  var headers = sh.getRange(1, 1, 1, Math.max(17, sh.getLastColumn())).getValues()[0];
  var hasSurveyTypeCol = String(headers[4] || '').indexOf("설문구분") > -1 || String(headers[0] || '').indexOf("응답일시") > -1;
  
  var row = [];
  if (hasSurveyTypeCol) {
    row = [stamp(new Date()), p.school, p.key, p.id, name, d.surveyType || '사전설문'];
  } else {
    row = [stamp(new Date()), p.school, p.key, p.id, name, p.type, d.surveyType || '사전설문'];
  }
  
  for (var i = 0; i < 12; i++) {
    var val = (ansList && ansList[i] !== undefined && ansList[i] !== null) ? String(ansList[i]).trim() : '';
    row.push(val);
  }
  sh.appendRow(row);
  return { success: true, message: '설문 저장 완료' };
}

function doGet(e) {
  try {
    var raw = String((e.parameter && e.parameter.studentId) || '').trim();
    var p = parseId(raw, e.parameter && e.parameter.school);
    var total = countStickers(p.key);
    var SS = getSS();
    var preSurveyDone = false;
    var surveySheets = ['학생설문응답', '교사설문응답', SH.survey];
    for (var s = 0; s < surveySheets.length; s++) {
      var sh = SS.getSheetByName(surveySheets[s]);
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
      if (preSurveyDone) break;
    }
    return json({ success: true, student: {
      id: raw, name: resolveName(p.key, '', p.id),
      totalSticker: total, level: levelOf(total), preSurveyDone: preSurveyDone } });
  } catch (err) {
    return json({ success: true, student: { id: 'guest', name: '학생', totalSticker: 0, preSurveyDone: false } });
  }
}

/* ── 기존자료 변환 ───────────────────── */

function migrate() { runMigrate(); }

function runMigrate() {
  var SS = getSS();
  var log = { daily: 0, growth: 0, dupDaily: 0, dupGrowth: 0, survey: 0 };
  SS.getSheets().forEach(function (sh) {
    Logger.log('시트 [' + sh.getName() + '] ' + sh.getLastRow() + '행');
  });
  migDaily(log);
  migGrowth(log);
  migSurvey(log);
  sortDaily();
  Logger.log('결과 ' + JSON.stringify(log));
  return log;
}

function migDaily(log) {
  var SS = getSS();
  var out = SS.getSheetByName(SH.daily);
  var seen = {}, order = [];

  OLD_SCHOOLS.forEach(function (school) {
    var sh = SS.getSheetByName(school);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (r) {
      if (!r[0] || !r[2]) return;
      if (Number(r[4] || 0) < 1) return;
      var when = new Date(r[0]);
      if (isNaN(when.getTime())) return;

      var p = parseId(r[2], r[1] || school);
      var date = ymd(when);
      var k = p.key + '|' + date;

      if (seen[k]) log.dupDaily++; else order.push(k);
      seen[k] = [k, stamp(when), p.school, p.key, p.id,
                 resolveName(p.key, r[3], p.id), p.type, date, 100, 1];
    });
  });

  var rows = order.map(function (k) { return seen[k]; })
                  .sort(function (a, b) { return a[1] < b[1] ? -1 : 1; });

  if (out.getLastRow() >= 2) {
    out.getRange(2, 1, out.getLastRow() - 1, out.getLastColumn()).clearContent();
  }

  if (rows.length) {
    out.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    log.daily = rows.length;
  }
}

function migGrowth(log) {
  var SS = getSS();
  var out = SS.getSheetByName(SH.growth);
  var seen = {}, order = [];

  OLD_SCHOOLS.forEach(function (school) {
    var sh = SS.getSheetByName(school + '_월별성장');
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues().forEach(function (r) {
      if (!r[0] || !r[2]) return;
      var when = new Date(r[0]);
      if (isNaN(when.getTime())) return;

      var p = parseId(r[2], r[1] || school);
      var ym = Utilities.formatDate(when, TZ, 'yyyy-MM');
      var mn = parseInt(String(r[4] || '').replace(/월/g, '').trim(), 10);
      if (mn >= 1 && mn <= 12) ym = when.getFullYear() + '-' + ('0' + mn).slice(-2);

      var k = p.key + '|' + ym;
      var H = Number(r[5] || 0), W = Number(r[6] || 0);
      var bmi = (H > 0 && W > 0) ? Math.round(W / Math.pow(H / 100, 2) * 10) / 10 : '';

      if (seen[k]) log.dupGrowth++; else order.push(k);
      seen[k] = [k, stamp(when), p.school, p.key, p.id,
                 resolveName(p.key, r[3], p.id), p.type, ym, H, W, bmi];
    });
  });

  var rows = order.map(function (k) { return seen[k]; })
                  .sort(function (a, b) { return a[1] < b[1] ? -1 : 1; });

  if (out.getLastRow() >= 2) {
    out.getRange(2, 1, out.getLastRow() - 1, out.getLastColumn()).clearContent();
  }

  if (rows.length) {
    out.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    log.growth = rows.length;
  }
}

function migSurvey(log) {
  var SS = getSS();
  var out = SS.getSheetByName(SH.survey);
  var rows = [];

  OLD_SCHOOLS.forEach(function (school) {
    var sh = SS.getSheetByName(school + '_설문응답');
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, 17).getValues().forEach(function (r) {
      if (!r[0] || !r[2]) return;
      var p = parseId(r[2], r[1] || school);
      var row = [r[0], p.school, p.key, p.id,
                 resolveName(p.key, r[3], p.id), p.type, r[4] || '사전설문'];
      for (var i = 5; i < 17; i++) row.push(r[i] !== undefined ? String(r[i]) : '');
      rows.push(row);
    });
  });

  if (out.getLastRow() >= 2) {
    out.getRange(2, 1, out.getLastRow() - 1, out.getLastColumn()).clearContent();
  }

  if (rows.length) {
    out.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    log.survey = rows.length;
  }
}

function sortDaily() {
  var SS = getSS();
  var sh = SS.getSheetByName(SH.daily);
  if (!sh || sh.getLastRow() < 3) return;
  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
    .sort([{ column: 7, ascending: true },
           { column: 4, ascending: true },
           { column: 8, ascending: true }]);
}