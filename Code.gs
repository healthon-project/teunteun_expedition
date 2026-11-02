/**
 * 꼬꼬챌린지 데이터 관리 (Google Apps Script)
 * 학교별 시트 명확 매핑 & 데이터 분리:
 *  1) {학교} 일일기록 (예: "A초 일일기록", "A초_일일기록", "A초", "일일기록", GID: 265693353): 일일 로그인/미션 기록 (당일 1줄만 보존 / 스티커 1개 / 학생 위, 교사 t-전화번호4자리 아래 정렬)
 *  2) {학교}_내몸탐험 (예: "A초_내몸탐험", "A초 내몸탐험"): 월1회 신체기록 (키, 몸무게, BMI, 총스티커, 레벨 / 교사 t-전화번호4자리)
 *  3) {학교}_설문응답 (예: "A초_설문응답", "A초 설문응답"): 학생 + 교사 통합 사전설문응답 (교사 t-전화번호4자리 / 맨윗줄 헤더 보존)
 *  4) {학교} 사후설문응답 (예: "A초 사후설문응답", "B초 사후설문응답", "C초 사후설문응답"): 학생 + 교사 통합 사후설문응답
 */

const TZ = 'Asia/Seoul';
const SCHOOLS = ['A초', 'B초', 'C초', 'D초'];
const POST_SURVEY_SCHOOLS = ['A초', 'B초', 'C초'];

const HEADERS = {
  '일별기록': ['중복키','일시','학교','학생키','개인번호','이름','구분','날짜','포인트','스티커'],
  '월별성장': ['측정일시','학교','개인번호','이름','측정월','키(cm)','몸무게(kg)','BMI','총스티커','레벨'],
  '설문응답': ['응답일시','학교','개인번호','이름','구분','문항1','문항2','문항3','문항4','문항5','문항6','문항7','문항8','문항9','문항10','문항11','문항12'],
  '사후설문응답': ['응답일시','학교','개인번호','이름','구분','문항1','문항2','문항3','문항4','문항5','문항6','문항7','문항8','문항9','문항10','문항11','문항12','문항13','문항14','문항15','문항16']
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
    .addItem('① 시트 연결 및 헤더 점검 (사후설문 탭 3개 생성 + 오분류 복구)', 'setup')
    .addItem('🎯 사후설문 탭 3개(A·B·C초 사후설문응답) 즉시 생성', 'createPostSurveyTabsMenu')
    .addItem('🔄 타학교 오분류 데이터 자동 복구 (A초→B·C·D초)', 'relocateMisroutedRecordsMenu')
    .addSeparator()
    .addItem('💾 드라이브에 백업', 'backupToDrive')
    .addItem('🧹 일별기록 정렬 (학생 위 / 교사 아래)', 'sortDailyAll')
    .addToUi();
}

function parseSchoolName(schoolInput) {
  var str = String(schoolInput || 'A초').trim();
  var m = str.match(/([A-Da-d])초?/);
  if (m) {
    return m[1].toUpperCase() + '초';
  }
  var m2 = str.match(/([A-Za-z0-9가-힣]+)초?/);
  if (m2) {
    var letter = m2[1].toUpperCase();
    if (letter.indexOf('초') < 0) letter += '초';
    return letter;
  }
  return 'A초';
}

// 다른 학교 시트인지 판별하여 학교 간 시트 혼선 원천 차단
function isOtherSchoolSheet(sheetName, targetSchool) {
  var sUpper = String(sheetName || '').trim().toUpperCase();
  var tUpper = parseSchoolName(targetSchool).toUpperCase();
  for (var i = 0; i < SCHOOLS.length; i++) {
    var sc = SCHOOLS[i].toUpperCase();
    if (sc !== tUpper) {
      var letter = sc.replace('초', '');
      if (sUpper.indexOf(sc) >= 0 || sUpper === letter || sUpper.indexOf(letter + '_') === 0 || sUpper.indexOf(letter + ' ') === 0) {
        return true;
      }
    }
  }
  return false;
}

/**
 * 1. 일일기록 시트 전용 탐색 (학교별 엄격 분리 보장)
 */
function getDailySheet(SS, school) {
  var name = parseSchoolName(school); // 예: "A초", "B초", "C초", "D초"
  var sheets = SS.getSheets();
  var nameUpper = name.toUpperCase();
  var letterOnly = nameUpper.replace('초', '');

  // 0) A초인 경우에만 GID 265693353 직접 일치 시트 탐색
  if (nameUpper === 'A초' || nameUpper === 'A') {
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === 265693353 && !isOtherSchoolSheet(sheets[i].getName(), name)) {
        return sheets[i];
      }
    }
  }

  // 1) 띄어쓰기/언더바 포함 대소문자/공백 무관 직접 시트 검색
  var candidates = [
    nameUpper + ' 일일기록',
    nameUpper + '_일일기록',
    nameUpper + '일일기록',
    nameUpper + ' 일별기록',
    nameUpper + '_일별기록',
    nameUpper + '일별기록',
    nameUpper,
    '일일기록 ' + nameUpper,
    '일일기록_' + nameUpper
  ];

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    var sUpper = sName.toUpperCase();
    if (candidates.indexOf(sUpper) >= 0) {
      return sheets[i];
    }
  }

  // 2) 해당 학교 이름으로 시작/포함하고 "일일기록" 또는 "기록" 포함 시트 탐색 (타학교 시트 제외)
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    if (isOtherSchoolSheet(sName, name)) continue;
    var sUpper = sName.toUpperCase();
    if ((sUpper.indexOf(nameUpper) >= 0 || sUpper.indexOf(letterOnly + '초') >= 0) &&
        (sUpper.indexOf('일일기록') >= 0 || sUpper.indexOf('일별기록') >= 0 || sUpper.indexOf('기록') >= 0) &&
        sUpper.indexOf('내몸탐험') < 0 && sUpper.indexOf('성장') < 0 && sUpper.indexOf('설문') < 0) {
      return sheets[i];
    }
  }

  // 3) 해당 학교 이름("B초" 등)으로 시작하고 "내몸탐험", "성장", "설문"이 포함되지 않은 시트 탐색
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    if (isOtherSchoolSheet(sName, name)) continue;
    var sUpper = sName.toUpperCase();
    if (sUpper.indexOf(nameUpper) === 0 && sUpper.indexOf('내몸탐험') < 0 && sUpper.indexOf('성장') < 0 && sUpper.indexOf('설문') < 0) {
      return sheets[i];
    }
  }

  // 4) 기존 시트가 없을 때 해당 학교 전용 일일기록 시트 신규 생성 (타학교 시트로 넘어가지 않도록 보장)
  var newSh = SS.insertSheet(name + '_일일기록');
  newSh.getRange(1, 1, 1, 10).setValues([HEADERS['일별기록']]).setFontWeight('bold').setBackground('#e8f0fe');
  newSh.setFrozenRows(1);
  return newSh;
}

/**
 * 2. 내몸탐험/월별성장 시트 전용 탐색 (학교별 엄격 분리 보장)
 */
function getGrowthSheet(SS, school) {
  var name = parseSchoolName(school);
  var sheets = SS.getSheets();
  var nameUpper = name.toUpperCase();
  var letterOnly = nameUpper.replace('초', '');

  var candidates = [
    nameUpper + '_내몸탐험',
    nameUpper + ' 내몸탐험',
    nameUpper + '내몸탐험',
    nameUpper + '_성장',
    nameUpper + ' 성장',
    nameUpper + '_월별성장',
    nameUpper + ' 월별성장',
    nameUpper + '월별성장',
    '내몸탐험_' + nameUpper,
    '내몸탐험 ' + nameUpper,
    '월별성장_' + nameUpper,
    '월별성장 ' + nameUpper
  ];

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    var sUpper = sName.toUpperCase();
    if (candidates.indexOf(sUpper) >= 0) {
      return sheets[i];
    }
  }

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    if (isOtherSchoolSheet(sName, name)) continue;
    var sUpper = sName.toUpperCase();
    if ((sUpper.indexOf(nameUpper) >= 0 || sUpper.indexOf(letterOnly + '초') >= 0) && 
        (sUpper.indexOf('내몸탐험') >= 0 || sUpper.indexOf('성장') >= 0 || sUpper.indexOf('월별성장') >= 0)) {
      return sheets[i];
    }
  }

  var newSh = SS.insertSheet(name + '_월별성장');
  newSh.getRange(1, 1, 1, HEADERS['월별성장'].length).setValues([HEADERS['월별성장']]).setFontWeight('bold').setBackground('#e2f0d9');
  newSh.setFrozenRows(1);
  return newSh;
}

/**
 * 3. 설문응답 시트 전용 탐색 (예: "A초_설문응답", "A초 설문응답")
 */
function getSurveySheet(SS, school) {
  var name = parseSchoolName(school);
  var sheets = SS.getSheets();
  var nameUpper = name.toUpperCase();

  var candidates = [
    nameUpper + '_설문응답',
    nameUpper + ' 설문응답',
    nameUpper + '설문응답',
    nameUpper + '_설문',
    nameUpper + ' 설문'
  ];

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    var sUpper = sName.toUpperCase();
    if (candidates.indexOf(sUpper) >= 0) {
      return sheets[i];
    }
  }

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    var sUpper = sName.toUpperCase();
    if (sUpper.indexOf(nameUpper) === 0 && sUpper.indexOf('설문') >= 0 && sUpper.indexOf('사후') < 0) {
      return sheets[i];
    }
  }

  var newSh = SS.insertSheet(name + '_설문응답');
  newSh.getRange(1, 1, 1, 17).setValues([HEADERS['설문응답']]).setFontWeight('bold').setBackground('#fff2cc');
  newSh.setFrozenRows(1);
  return newSh;
}

/**
 * 4. 사후설문응답 시트 전용 탐색 (예: "A초 사후설문응답", "B초 사후설문응답", "C초 사후설문응답", "A초_사후설문응답")
 */
function getPostSurveySheet(SS, school) {
  var name = parseSchoolName(school);
  var sheets = SS.getSheets();
  var nameUpper = name.toUpperCase();

  var candidates = [
    nameUpper + ' 사후설문응답',
    nameUpper + '_사후설문응답',
    nameUpper + '사후설문응답',
    nameUpper + ' 사후설문',
    nameUpper + '_사후설문',
    nameUpper + '사후설문'
  ];

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    var sUpper = sName.toUpperCase();
    if (candidates.indexOf(sUpper) >= 0) {
      return sheets[i];
    }
  }

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim();
    var sUpper = sName.toUpperCase();
    if (sUpper.indexOf(nameUpper) === 0 && sUpper.indexOf('사후') >= 0) {
      return sheets[i];
    }
  }

  var newSh = SS.insertSheet(name + ' 사후설문응답');
  newSh.getRange(1, 1, 1, HEADERS['사후설문응답'].length).setValues([HEADERS['사후설문응답']]).setFontWeight('bold').setBackground('#d9ead3');
  newSh.setFrozenRows(1);
  return newSh;
}

function ensurePostSurveySheets(SS) {
  SS = SS || getSS();
  POST_SURVEY_SCHOOLS.forEach(function(school) {
    var shPost = getPostSurveySheet(SS, school);
    cleanAndFixSurveySheet(shPost);
  });
}

function createPostSurveyTabsMenu() {
  var SS = getSS();
  ensurePostSurveySheets(SS);
  SpreadsheetApp.getUi().alert('완료: [A초 사후설문응답], [B초 사후설문응답], [C초 사후설문응답] 탭 3개가 생성/연결되었습니다!');
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
    cleanAndFixGrowthSheet(sh2);

    var sh3 = getSurveySheet(SS, school);
    cleanAndFixSurveySheet(sh3);
  });
  ensurePostSurveySheets(SS);
  relocateMisroutedRecords(SS);
  Logger.log('학교별 시트 연결 점검 및 오분류 데이터 복구 완료 (A초/B초/C초 사후설문응답 포함)');
}

function normalizeMonth(rawMonth, now) {
  var nowObj = now || new Date();
  if (!rawMonth) {
    return Utilities.formatDate(nowObj, TZ, 'yyyy-MM');
  }
  var str = String(rawMonth).trim();
  if (/^\d{4}-\d{2}$/.test(str)) {
    return str;
  }
  var mMatch = str.match(/^(\d{1,2})$/) || str.match(/^(\d{1,2})월$/);
  if (mMatch) {
    var yearStr = Utilities.formatDate(nowObj, TZ, 'yyyy');
    var mNum = parseInt(mMatch[1], 10);
    var mPadded = (mNum < 10 ? '0' : '') + mNum;
    return yearStr + '-' + mPadded;
  }
  var dateMatch = str.match(/(\d{4}-\d{2})/);
  if (dateMatch) {
    return dateMatch[1];
  }
  return Utilities.formatDate(nowObj, TZ, 'yyyy-MM');
}

function cleanAndFixGrowthSheet(sh) {
  if (!sh) return;
  var lastRow = sh.getLastRow();
  var defaultHeaders = HEADERS['월별성장']; // ['측정일시','학교','개인번호','이름','측정월','키(cm)','몸무게(kg)','BMI','총스티커','레벨']

  if (lastRow === 0) {
    sh.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders])
      .setFontWeight('bold').setBackground('#e2f0d9');
    sh.setFrozenRows(1);
    return;
  }

  var maxCols = Math.max(defaultHeaders.length, sh.getLastColumn());
  var headerRow = sh.getRange(1, 1, 1, maxCols).getValues()[0];

  // 불필요하게 추가되었던 '중복키', '학생키', '구분' 열 위치 탐색
  var removeIndices = [];
  for (var c = 0; c < headerRow.length; c++) {
    var hName = String(headerRow[c] || '').trim();
    if (hName === '중복키' || hName === '학생키' || hName === '구분') {
      removeIndices.push(c);
    }
  }

  if (removeIndices.length > 0) {
    // 헤더에서 해당 항목 제거
    for (var i = removeIndices.length - 1; i >= 0; i--) {
      headerRow.splice(removeIndices[i], 1);
    }

    if (headerRow.length === 0) {
      headerRow = defaultHeaders;
    }

    sh.getRange(1, 1, 1, headerRow.length).setValues([headerRow])
      .setFontWeight('bold').setBackground('#e2f0d9');
    sh.setFrozenRows(1);

    if (lastRow >= 2) {
      var dataRange = sh.getRange(2, 1, lastRow - 1, maxCols);
      var rows = dataRange.getValues();

      for (var r = 0; r < rows.length; r++) {
        var row = rows[r];
        // 데이터 행에서도 동일한 열 제거 (뒤에서부터 제거하여 인덱스 밀림 방지)
        for (var i = removeIndices.length - 1; i >= 0; i--) {
          var idx = removeIndices[i];
          if (idx < row.length) {
            row.splice(idx, 1);
          }
        }
        rows[r] = row;
      }

      sh.getRange(2, 1, rows.length, headerRow.length).setValues(rows);
      if (maxCols > headerRow.length) {
        sh.getRange(1, headerRow.length + 1, lastRow, maxCols - headerRow.length).clearContent();
      }
    }
  } else {
    sh.setFrozenRows(1);
  }
}

function cleanAndFixSurveySheet(sh) {
  if (!sh) return;
  
  var isPostSheet = String(sh.getName() || '').indexOf('사후') >= 0;
  var expectedHeaders = isPostSheet ? HEADERS['사후설문응답'] : HEADERS['설문응답'];
  var expectedColCount = expectedHeaders.length; // 사전 17열(문항1~12), 사후 21열(문항1~16)
  var headerBg = isPostSheet ? '#d9ead3' : '#fff2cc';
  var lastRow = sh.getLastRow();
  
  if (lastRow === 0) {
    sh.getRange(1, 1, 1, expectedColCount).setValues([expectedHeaders])
      .setFontWeight('bold').setBackground(headerBg);
    sh.setFrozenRows(1);
    return;
  }

  var firstCell = String(sh.getRange(1, 1).getValue() || '').trim();

  if (firstCell !== '응답일시') {
    sh.insertRowBefore(1);
    sh.getRange(1, 1, 1, expectedColCount).setValues([expectedHeaders])
      .setFontWeight('bold').setBackground(headerBg);
    sh.setFrozenRows(1);
    lastRow = sh.getLastRow();
  } else {
    sh.getRange(1, 1, 1, expectedColCount).setValues([expectedHeaders])
      .setFontWeight('bold').setBackground(headerBg);
    sh.setFrozenRows(1);
  }

  if (lastRow >= 2) {
    var maxCols = Math.max(expectedColCount + 1, sh.getLastColumn());
    var dataRange = sh.getRange(2, 1, lastRow - 1, maxCols);
    var rows = dataRange.getValues();
    var modified = false;

    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      var foundIdx = -1;
      for (var c = 0; c < Math.min(7, row.length); c++) {
        var cellTxt = String(row[c] || '').trim();
        if (cellTxt === '사전설문' || cellTxt === '사후설문') {
          foundIdx = c;
          break;
        }
      }
      if (foundIdx !== -1) {
        row.splice(foundIdx, 1);
        while (row.length < expectedColCount) row.push('');
        rows[r] = row.slice(0, expectedColCount);
        modified = true;
      }
    }

    if (modified) {
      sh.getRange(2, 1, rows.length, expectedColCount).setValues(rows);
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
  var origRaw = raw;
  var school = '';
  var explicitFallback = String(fallbackSchool || '').trim();
  if (explicitFallback) explicitFallback = parseSchoolName(explicitFallback);

  if (raw.indexOf('_') > 0) {
    var parts = raw.split('_');
    school = parseSchoolName(parts[0].trim());
    raw = parts.slice(1).join('_').trim();
  }

  var m = raw.match(/^([A-Za-z0-9가-힣]+)초?[_-]/);
  if (m && !school) {
    school = parseSchoolName(m[1]);
  }

  var id = raw.replace(/^[A-Za-z0-9가-힣]+초?[_-]/i, '').trim();

  // 만약 학번 접두어가 기본값 'A초'로 되어 있더라도, 전달된 school 파라미터가 B초/C초/D초라면 선택된 학교(B초/C초/D초)를 우선 적용!
  if (explicitFallback && explicitFallback !== 'A초' && (!school || school === 'A초')) {
    school = explicitFallback;
  }

  if (!school) {
    school = explicitFallback || 'A초';
  }

  var isTeacherMatch = /^T-?\d+$/i.test(id) || /^t-?\d+$/i.test(id) || 
                       origRaw.indexOf('교사') >= 0 || origRaw.indexOf('teacher') >= 0 || 
                       origRaw.indexOf('T-') >= 0 || origRaw.indexOf('t-') >= 0;

  var type;
  if (isTeacherMatch) {
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

/**
 * 각 학교 시트(설문응답, 월별성장, 일일기록)에서 해당 학생(학번+이름)의 기록 가중치 점수를 계산합니다.
 * - 설문응답 일치: +100점
 * - 월별성장 일치: +20점/건
 * - 일일기록 일치: +2점/건
 * 이를 통해 B초/C초/D초 학생의 기록이 오늘 A초로 1~2건 잘못 넘어왔더라도
 * 기존 기록이 훨씬 많은 원래 소속 학교(B초, C초, D초)를 100% 정확히 판별합니다.
 */
function scoreStudentInSchool(SS, schoolName, targetId, targetName) {
  var cleanId = String(targetId || '').trim().replace(/^'/, '');
  var cleanName = String(targetName || '').trim().replace(/^[Tt]-/, '');
  if (!cleanId || cleanId === 'guest') return { idAndNameScore: 0, idOnlyScore: 0 };

  var idAndNameScore = 0;
  var idOnlyScore = 0;

  var sheetWeights = [
    { sh: getSurveySheet(SS, schoolName), weight: 100 },
    { sh: getGrowthSheet(SS, schoolName), weight: 20 },
    { sh: getDailySheet(SS, schoolName), weight: 2 }
  ];

  for (var s = 0; s < sheetWeights.length; s++) {
    var sh = sheetWeights[s].sh;
    var w = sheetWeights[s].weight;
    if (!sh || sh.getLastRow() < 2) continue;
    var lastRow = sh.getLastRow();
    var maxCols = Math.min(7, sh.getLastColumn());
    var data = sh.getRange(2, 1, lastRow - 1, maxCols).getValues();

    for (var r = 0; r < data.length; r++) {
      var info = extractRowIdAndName(data[r]);
      if (info.id === cleanId) {
        idOnlyScore += w;
        if (cleanName && cleanName !== '참여자' && cleanName !== '탐험대원' && cleanName !== cleanId) {
          if (info.name === cleanName) {
            idAndNameScore += w;
          }
        }
      }
    }
  }
  return { idAndNameScore: idAndNameScore, idOnlyScore: idOnlyScore };
}

function resolveActualSchool(SS, p, givenName) {
  if (!p || !p.id || p.id === 'guest') return p;
  // 이미 B초, C초, D초로 명시된 경우 그대로 유지
  if (p.school && p.school !== 'A초') return p;

  var cleanName = String(givenName || '').trim().replace(/^[Tt]-/, '');
  var hasMeaningfulName = (cleanName && cleanName !== '참여자' && cleanName !== '탐험대원' && cleanName !== p.id);

  var aScoreObj = scoreStudentInSchool(SS, 'A초', p.id, cleanName);
  var bestSchool = 'A초';
  var bestScore = hasMeaningfulName ? aScoreObj.idAndNameScore : aScoreObj.idOnlyScore;

  var otherSchools = ['B초', 'C초', 'D초'];
  for (var i = 0; i < otherSchools.length; i++) {
    var sc = otherSchools[i];
    var scScoreObj = scoreStudentInSchool(SS, sc, p.id, cleanName);
    var scScore = hasMeaningfulName ? scScoreObj.idAndNameScore : scScoreObj.idOnlyScore;
    if (scScore > bestScore) {
      bestScore = scScore;
      bestSchool = sc;
    }
  }

  if (bestSchool !== 'A초') {
    p.school = bestSchool;
    p.key = bestSchool + '-' + p.id;
  }
  return p;
}

/**
 * A초 시트(일일기록, 월별성장, 설문응답)에 잘못 넘어간 B초/C초/D초 학생 데이터를
 * 원래 소속 학교(B초, C초, D초) 시트로 자동 이동 및 복구합니다.
 */
function relocateMisroutedRecords(SS) {
  SS = SS || getSS();
  var movedCount = 0;

  // 모든 학교의 성장시트 헤더를 먼저 정렬하여 열 구조 일치 보장
  SCHOOLS.forEach(function(sc) {
    cleanAndFixGrowthSheet(getGrowthSheet(SS, sc));
  });

  // 1) A초, B초, C초, D초 전체에서 각 학생(학번|이름)의 학교별 기록 가중치 점수 집계
  var studentSchoolScores = {}; // key: "id|name" -> { 'A초': score, 'B초': score, 'C초': score, 'D초': score }
  var weights = [
    { getter: getSurveySheet, w: 100 },
    { getter: getGrowthSheet, w: 20 },
    { getter: getDailySheet, w: 2 }
  ];

  for (var i = 0; i < SCHOOLS.length; i++) {
    var sc = SCHOOLS[i];
    for (var k = 0; k < weights.length; k++) {
      var sh = weights[k].getter(SS, sc);
      var w = weights[k].w;
      if (!sh || sh.getLastRow() < 2) continue;
      var data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(7, sh.getLastColumn())).getValues();
      for (var r = 0; r < data.length; r++) {
        var info = extractRowIdAndName(data[r]);
        if (info.id && info.name) {
          var key = info.id + '|' + info.name;
          if (!studentSchoolScores[key]) {
            studentSchoolScores[key] = { 'A초': 0, 'B초': 0, 'C초': 0, 'D초': 0 };
          }
          studentSchoolScores[key][sc] = (studentSchoolScores[key][sc] || 0) + w;
        }
      }
    }
  }

  // 2) B초, C초, D초 점수가 A초 점수보다 높은 학생들을 타학교 소속 맵(otherStudentMap)으로 확정
  var otherStudentMap = {}; // key: "id|name" -> "B초" | "C초" | "D초"
  var studentKeys = Object.keys(studentSchoolScores);
  for (var m = 0; m < studentKeys.length; m++) {
    var sKey = studentKeys[m];
    var scMap = studentSchoolScores[sKey];
    var bestSc = 'A초';
    var maxScScore = scMap['A초'] || 0;
    var otherList = ['B초', 'C초', 'D초'];
    for (var j = 0; j < otherList.length; j++) {
      var candSc = otherList[j];
      if ((scMap[candSc] || 0) > maxScScore) {
        maxScScore = scMap[candSc];
        bestSc = candSc;
      }
    }
    if (bestSc !== 'A초') {
      otherStudentMap[sKey] = bestSc;
    }
  }

  // 3) A초_일일기록, A초_월별성장, A초_설문응답, A초 사후설문응답에서 오분류 행 탐색 및 원래 학교(B초/C초/D초)로 이동
  movedCount += moveMisroutedFromSheet(SS, getDailySheet(SS, 'A초'), 'daily', otherStudentMap);
  movedCount += moveMisroutedFromSheet(SS, getGrowthSheet(SS, 'A초'), 'growth', otherStudentMap);
  movedCount += moveMisroutedFromSheet(SS, getSurveySheet(SS, 'A초'), 'survey', otherStudentMap);
  movedCount += moveMisroutedFromSheet(SS, getPostSurveySheet(SS, 'A초'), 'post_survey', otherStudentMap);
  return movedCount;
}

function relocateMisroutedRecordsMenu() {
  var cnt = relocateMisroutedRecords(getSS());
  SpreadsheetApp.getUi().alert('오분류 데이터 자동 복구 완료: 총 ' + cnt + '건의 타학교(B초/C초/D초) 기록을 원래 학교 시트로 이동했습니다.');
}

function extractRowIdAndName(row) {
  var id = '';
  var name = '';
  var school = '';
  for (var c = 0; c < Math.min(7, row.length); c++) {
    var val = String(row[c] || '').trim().replace(/^'/, '');
    if (!val) continue;
    if (val === 'A초' || val === 'B초' || val === 'C초' || val === 'D초') {
      school = val;
    } else if (/^(\d{4,5}|[Tt]-\d{4})$/.test(val)) {
      if (!id) id = val.replace(/^T-/, 't-');
    } else if (/^[A-D]초-(\d{4,5}|[Tt]-\d{4})$/i.test(val)) {
      var parts = val.split('-');
      if (!id) id = parts.slice(1).join('-').replace(/^T-/, 't-');
    } else if (/^[가-힣]{2,5}$/.test(val) || /^[Tt]-[가-힣]{2,5}$/.test(val)) {
      if (val !== '사전설문' && val !== '사후설문' && val !== '일일기록' && val !== '내몸탐험') {
        name = val.replace(/^[Tt]-/, '');
      }
    }
  }
  return { id: id, name: name, school: school };
}

function moveMisroutedFromSheet(SS, shA, sheetType, otherStudentMap) {
  if (!shA || shA.getLastRow() < 2) return 0;
  var lastRow = shA.getLastRow();
  var lastCol = shA.getLastColumn();
  var headers = shA.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = shA.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var moved = 0;

  // 아래 행부터 역순으로 확인하여 삭제 시 행 밀림 방지
  for (var r = rows.length - 1; r >= 0; r--) {
    var row = rows[r];
    var info = extractRowIdAndName(row);
    var targetSchool = '';

    if (info.school && info.school !== 'A초') {
      targetSchool = info.school;
    } else if (info.id && info.name && otherStudentMap[info.id + '|' + info.name]) {
      targetSchool = otherStudentMap[info.id + '|' + info.name];
    }

    if (targetSchool && targetSchool !== 'A초') {
      // 행 내의 'A초'를 targetSchool로 교정
      var fixedRow = row.slice();
      for (var c = 0; c < fixedRow.length; c++) {
        var cellStr = String(fixedRow[c] || '');
        if (cellStr === 'A초') {
          fixedRow[c] = targetSchool;
        } else if (cellStr.indexOf('A초-') === 0) {
          fixedRow[c] = cellStr.replace(/^A초-/, targetSchool + '-');
        }
        var hName = String(headers[c] || '').trim();
        if ((hName === '개인번호' || hName === '학번') && fixedRow[c]) {
          fixedRow[c] = "'" + String(fixedRow[c]).replace(/^'/, '');
        }
      }

      var targetSh = null;
      if (sheetType === 'daily') targetSh = getDailySheet(SS, targetSchool);
      else if (sheetType === 'growth') targetSh = getGrowthSheet(SS, targetSchool);
      else if (sheetType === 'survey') targetSh = getSurveySheet(SS, targetSchool);
      else if (sheetType === 'post_survey') targetSh = getPostSurveySheet(SS, targetSchool);

      if (targetSh) {
        appendOrInsertRow(targetSh, fixedRow);
        shA.deleteRow(r + 2);
        moved++;
      }
    }
  }
  return moved;
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

function buildDailyRow(headers, p, name, date, now, pts, sticker) {
  var targetKey = String(p.key || '').trim();
  var targetId = String(p.id || '').trim();
  var todayDateStr = date;
  var timestampStr = stamp(now);
  var finalSticker = 1;

  var row = [];
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    if (h === '중복키') {
      row.push(targetKey + '|' + todayDateStr);
    } else if (h === '일시' || h === '응답일시') {
      row.push(timestampStr);
    } else if (h === '학교') {
      row.push(p.school);
    } else if (h === '학생키') {
      row.push(targetKey);
    } else if (h === '개인번호' || h === '학번' || h === 'ID') {
      row.push("'" + targetId);
    } else if (h === '이름') {
      row.push(name);
    } else if (h === '구분') {
      row.push(p.type);
    } else if (h === '날짜') {
      row.push(todayDateStr);
    } else if (h === '포인트') {
      row.push(pts);
    } else if (h === '스티커' || h === '일별스티커') {
      row.push(finalSticker);
    } else {
      row.push('');
    }
  }
  return row;
}

function upsertDailyRow(sh, p, name, date, now, pts, sticker) {
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow === 0) {
    sh.getRange(1, 1, 1, 10).setValues([HEADERS['일별기록']]).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
    lastRow = 1;
  }

  var headers = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0];
  var todayDateStr = ymd(now);
  var targetKey = String(p.key || '').trim();
  var targetId = String(p.id || '').trim();
  var numIdOnly = targetId.replace(/^[Tt]-?/i, '');

  if (lastRow >= 2) {
    var data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      var isToday = false;
      for (var col = 0; col < row.length; col++) {
        var formatted = formatDateStr(row[col]);
        if (formatted === todayDateStr) {
          isToday = true;
          break;
        }
      }

      var isSameStudent = false;
      for (var col = 0; col < row.length; col++) {
        var cellVal = String(row[col] || '').trim().replace(/^'/, '');
        var numCellOnly = cellVal.replace(/^[Tt]-?/i, '');
        var cellIsTeacher = cellVal.indexOf('t-') >= 0 || cellVal.indexOf('T-') >= 0 || cellVal.indexOf('교사') >= 0;

        if (p.type === '2.교사') {
          if (cellIsTeacher && (cellVal === targetKey || cellVal === targetId || (numIdOnly.length >= 4 && numCellOnly === numIdOnly))) {
            isSameStudent = true;
            break;
          }
        } else {
          if (!cellIsTeacher && (cellVal === targetKey || cellVal === targetId || (numIdOnly.length >= 4 && numCellOnly === numIdOnly))) {
            isSameStudent = true;
            break;
          }
        }
      }

      if (isToday && isSameStudent) {
        var targetRowIndex = i + 2;
        var updatedRow = buildDailyRow(headers, p, name, date, now, pts, sticker);
        sh.getRange(targetRowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
        return 'updated';
      }
    }
  }

  var newRow = buildDailyRow(headers, p, name, date, now, pts, sticker);
  appendOrInsertRow(sh, newRow);
  return 'inserted';
}

function countStickers(key, schoolName) {
  var SS = getSS();
  var sh = getDailySheet(SS, schoolName || 'A초');
  if (!sh || sh.getLastRow() < 2) return 0;
  var targetKey = String(key || '').trim();
  if (!targetKey) return 0;

  var p = parseId(targetKey, schoolName);
  var fullKey = p.key;
  var idOnly = p.id;
  var isTeacher = (p.type === '2.교사');

  var lastRow = sh.getLastRow();
  var v = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  var count = 0;

  for (var i = 0; i < v.length; i++) {
    var row = v[i];
    var rowStr = row.join(' ');
    var isMatch = false;

    if (isTeacher) {
      if (rowStr.indexOf(fullKey) >= 0 || rowStr.indexOf(idOnly) >= 0 || 
          (rowStr.indexOf(idOnly.replace(/^[Tt]-?/i, '')) >= 0 && rowStr.indexOf('교사') >= 0)) {
        isMatch = true;
      }
    } else {
      if (rowStr.indexOf('2.교사') < 0 && rowStr.indexOf('t-') < 0 && rowStr.indexOf('T-') < 0) {
        if (rowStr.indexOf(fullKey) >= 0 || rowStr.indexOf(idOnly) >= 0) {
          isMatch = true;
        } else {
          var numOnly = idOnly.replace(/^[Tt]-?/i, '');
          if (numOnly.length >= 4 && rowStr.indexOf(numOnly) >= 0) {
            isMatch = true;
          }
        }
      }
    }

    if (isMatch) {
      var sVal = Number(row[row.length - 1]);
      count += (!isNaN(sVal) && sVal > 0) ? sVal : 1;
    }
  }
  return count;
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
    if (d.action === 'save_survey' || d.action === 'submit_survey' || d.action === 'save_post_survey' || d.action === 'submit_post_survey') return json(saveSurvey(d));
    return json({ success: true, message: '수신 완료' });
  } catch (err) {
    return json({ success: false, error: String(err) });
  } finally {
    if (got) lock.releaseLock();
  }
}

// 1. 일일기록 저장 ("A초 일일기록", "A초_일일기록", "A초", "일일기록" 시트 중 어느것이든 무조건 정밀 저장을 보장!)
function saveDaily(d) {
  var SS = getSS();
  var p = parseId(d.studentId, d.school);
  p = resolveActualSchool(SS, p, d.name);
  var schoolName = p.school || "A초";
  var isTeacher = (p.type === '2.교사') || 
                  (String(d.role || d.type || '').indexOf('교사') >= 0) || 
                  (String(d.studentId || '').toUpperCase().indexOf('_T-') >= 0) ||
                  (String(d.studentId || '').toLowerCase().indexOf('t-') === 0);

  var name = resolveName(p.key, d.name, p.id);
  var idToSave = p.id;

  if (isTeacher) {
    p.type = '2.교사';
    if (idToSave.indexOf('t-') !== 0) idToSave = 't-' + idToSave.replace(/^[Tt]-?/, '');
    if (name.indexOf('t-') !== 0 && name.indexOf('T-') !== 0) name = 't-' + name;
    p.id = idToSave;
    p.key = schoolName + '-' + idToSave;
  }

  var pts = Number(d.points || 0);
  var sticker = 1;
  var now = new Date(), date = ymd(now);

  // 무조건 일일기록 시트에 기록!
  var sh = getDailySheet(SS, schoolName);
  upsertDailyRow(sh, p, name, date, now, pts, sticker);
  sortDailySheet(sh);

  var total = countStickers(p.key, schoolName);
  return { success: true, 학교: schoolName, 이름: name, 구분: p.type, 포인트: pts, 총스티커: total, 레벨: levelOf(total) };
}

function buildGrowthRow(headers, p, name, month, now, hVal, wVal, bmiVal, total, lvl) {
  var targetKey = String(p.key || '').trim();
  var targetId = String(p.id || '').trim();
  var timestampStr = stamp(now);

  var row = [];
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    if (h === '중복키') {
      row.push(targetKey + '|' + month);
    } else if (h === '측정일시' || h === '일시' || h === '응답일시') {
      row.push(timestampStr);
    } else if (h === '학교') {
      row.push(p.school);
    } else if (h === '학생키') {
      row.push(targetKey);
    } else if (h === '개인번호' || h === '학번' || h === 'ID') {
      row.push("'" + targetId);
    } else if (h === '이름') {
      row.push(name);
    } else if (h === '측정월' || h === '월' || h === '날짜') {
      row.push(month);
    } else if (h === '키(cm)' || h === '키') {
      row.push(hVal);
    } else if (h === '몸무게(kg)' || h === '몸무게') {
      row.push(wVal);
    } else if (h === 'BMI' || h === 'bmi') {
      row.push(bmiVal);
    } else if (h === '총스티커' || h === '스티커') {
      row.push(total);
    } else if (h === '레벨' || h === '단계') {
      row.push(lvl);
    } else {
      row.push('');
    }
  }
  return row;
}

function upsertGrowthRow(sh, p, month, newRow) {
  if (!sh) return '시트없음';
  var lastRow = sh.getLastRow();
  var targetDupKey = (p.key + '|' + month).trim();
  var numIdOnly = String(p.id || '').replace(/^[Tt]-?/i, '').trim();

  if (lastRow >= 2) {
    var headers = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0];
    var data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var isSameMonth = false;
      var isSameStudent = false;

      for (var col = 0; col < row.length; col++) {
        var cellVal = String(row[col] || '').trim().replace(/^'/, '');
        if (cellVal === targetDupKey || cellVal === month || cellVal.indexOf(month) >= 0) {
          isSameMonth = true;
        }
        var numCellOnly = cellVal.replace(/^[Tt]-?/i, '');
        if (cellVal === p.key || cellVal === p.id || (numIdOnly.length >= 4 && numCellOnly === numIdOnly) || cellVal.indexOf(p.key) === 0) {
          isSameStudent = true;
        }
      }

      if (isSameMonth && isSameStudent) {
        var targetRowIndex = i + 2;
        sh.getRange(targetRowIndex, 1, 1, newRow.length).setValues([newRow]);
        return '덮어씀';
      }
    }
  }

  appendOrInsertRow(sh, newRow);
  return '추가함';
}

// 2. 내몸탐험 저장 ({학교}_내몸탐험: 월1회 키, 몸무게, BMI, 총스티커, 레벨)
function saveGrowth(d) {
  var SS = getSS();
  relocateMisroutedRecords(SS);
  var p = parseId(d.studentId, d.school);
  p = resolveActualSchool(SS, p, d.name);
  var schoolName = p.school || "A초";
  var isTeacher = (p.type === '2.교사') || 
                  (String(d.role || d.type || '').indexOf('교사') >= 0) || 
                  (String(d.studentId || '').toUpperCase().indexOf('_T-') >= 0) ||
                  (String(d.studentId || '').toLowerCase().indexOf('t-') === 0);

  var name = resolveName(p.key, d.name, p.id);
  var idToSave = p.id;

  if (isTeacher) {
    p.type = '2.교사';
    if (idToSave.indexOf('t-') !== 0) idToSave = 't-' + idToSave.replace(/^[Tt]-?/, '');
    if (name.indexOf('t-') !== 0 && name.indexOf('T-') !== 0) name = 't-' + name;
    p.id = idToSave;
    p.key = schoolName + '-' + idToSave;
  }

  var now = new Date();
  var month = normalizeMonth(d.month, now);
  var h = Number(d.height || 0), w = Number(d.weight || 0);
  var bmi = Number(d.bmi) || ((h > 0 && w > 0) ? Math.round(w / Math.pow(h / 100, 2) * 10) / 10 : '');

  var total = Number(d.stickerCount) || countStickers(p.key, schoolName);
  var lvl = levelOf(total);

  var sh = getGrowthSheet(SS, schoolName);
  cleanAndFixGrowthSheet(sh);

  var headers = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0];
  var rowToInsert = buildGrowthRow(headers, p, name, month, now, h, w, bmi, total, lvl);
  
  var status = upsertGrowthRow(sh, p, month, rowToInsert);

  return { success: true, status: status, 학교: schoolName, 이름: name, 구분: p.type, BMI: bmi, 총스티커: total, 레벨: lvl };
}

function upsertSurveyRow(sh, p, idToSave, row) {
  if (!sh) return '시트없음';
  var last = sh.getLastRow();
  var targetId = String(idToSave || '').trim().replace(/^'/, '');
  var targetSchool = String(p.school || '').trim();
  var targetKey = targetSchool + '-' + targetId;

  if (last >= 2) {
    var data = sh.getRange(2, 1, last - 1, Math.min(5, sh.getLastColumn())).getValues();
    for (var i = 0; i < data.length; i++) {
      var rKey = String(data[i][0] || '').trim();
      var rSchool = String(data[i][1] || '').trim();
      var rId = String(data[i][2] || '').trim().replace(/^'/, '');
      if (rKey === targetKey || (rSchool === targetSchool && rId === targetId)) {
        sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
        return '덮어씀';
      }
    }
  }
  appendOrInsertRow(sh, row);
  return '추가함';
}

// 3. 설문응답 저장 ({학교}_설문응답 및 {학교} 사후설문응답)
function saveSurvey(d) {
  var SS = getSS();
  ensurePostSurveySheets(SS);

  var p = parseId(d.studentId, d.school);
  p = resolveActualSchool(SS, p, d.name);
  var schoolName = p.school || "A초";
  var isTeacher = (p.type === '2.교사') || 
                  (String(d.role || d.type || '').indexOf('교사') >= 0) || 
                  (String(d.studentId || '').toUpperCase().indexOf('_T-') >= 0) ||
                  (String(d.studentId || '').toLowerCase().indexOf('t-') === 0);

  var name = resolveName(p.key, d.name, p.id);
  var idToSave = p.id;

  if (isTeacher) {
    p.type = '2.교사';
    if (idToSave.indexOf('t-') !== 0) idToSave = 't-' + idToSave.replace(/^[Tt]-?/, '');
    if (name.indexOf('t-') !== 0 && name.indexOf('T-') !== 0) name = 't-' + name;
    p.id = idToSave;
    p.key = schoolName + '-' + idToSave;
  }

  var rawAnswers = d.answers || d.surveyAnswers || [];
  var ansList = [];
  if (Object.prototype.toString.call(rawAnswers) === '[object Array]') {
    ansList = rawAnswers;
  } else if (typeof rawAnswers === 'object') {
    var keys = Object.keys(rawAnswers).sort(function(a, b) {
      var na = parseInt(String(a).replace(/[^0-9]/g, '')) || 0;
      var nb = parseInt(String(b).replace(/[^0-9]/g, '')) || 0;
      return na - nb;
    });
    ansList = keys.map(function(k) { return rawAnswers[k]; });
  }

  // 사후설문 여부 판별 (action, isPostSurvey, surveyType 또는 문항 수가 13개 이상인 경우 무조건 사후설문 전용 탭으로 분리!)
  var isPost = (d.isPostSurvey === true) ||
               (String(d.isPostSurvey) === 'true') ||
               (String(d.surveyType || '').indexOf('사후') >= 0) ||
               (d.action === 'save_post_survey' || d.action === 'submit_post_survey') ||
               (ansList.length > 12);

  var sh = isPost ? getPostSurveySheet(SS, schoolName) : getSurveySheet(SS, schoolName);
  cleanAndFixSurveySheet(sh);

  var row = [stamp(new Date()), p.school, "'" + idToSave, name, p.type];
  var questionCount = isPost ? 16 : 12;
  for (var i = 0; i < questionCount; i++) {
    row.push((ansList[i] !== undefined && ansList[i] !== null) ? String(ansList[i]) : '');
  }

  var status = upsertSurveyRow(sh, p, idToSave, row);
  return { success: true, status: status, 설문종류: isPost ? '사후설문' : '사전설문', 시트: sh.getName(), 이름: name, 구분: p.type };
}

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    var action = params.action || '';
    if (action === 'init_post_survey_tabs') {
      var SS0 = getSS();
      ensurePostSurveySheets(SS0);
      return json({ success: true, message: 'A초/B초/C초 사후설문응답 탭 3개 생성 완료' });
    }
    if (action === 'get_student') {
      var SS = getSS();
      ensurePostSurveySheets(SS);
      var studentId = params.studentId || '';
      var school = params.school || 'A초';
      var givenName = params.name || '';
      var p = parseId(studentId, school);
      p = resolveActualSchool(SS, p, givenName);
      var totalStickers = countStickers(p.key, p.school);
      var totalPts = totalStickers * 100;
      var lvlName = levelOf(totalStickers);
      var lvlNum = 1;
      if (totalStickers >= 45) lvlNum = 4;
      else if (totalStickers >= 24) lvlNum = 3;
      else if (totalStickers >= 8) lvlNum = 2;

      return json({
        success: true,
        student: {
          studentId: studentId,
          school: p.school,
          name: p.id,
          totalPoints: totalPts,
          stickerCount: totalStickers,
          level: lvlNum,
          levelName: lvlName
        }
      });
    }
    return json({ success: true, message: '꼬꼬챌린지 API 실행 중' });
  } catch(err) {
    return json({ success: false, error: String(err) });
  }
}