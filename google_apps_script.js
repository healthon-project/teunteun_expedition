/**
 * 튼튼탐험대 (TeunTeun Expedition) - 구글 앱스 스크립트 (Google Apps Script)
 * 학생용/교사역 분리 및 개인번호(4자리) 기반 누적 합산 스키마 (데이트 파싱 버그 수정 버젼)
 */

function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : null;
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    setupSheets(sheet);
  } catch(err) {
    return createJsonResponse({
      success: false,
      message: "초기화 에러: " + err.message
    });
  }
  
  if (action === 'get_leaderboard') {
    return handleGetLeaderboard(sheet);
  } else if (action === 'get_student') {
    var studentId = (e && e.parameter) ? e.parameter.studentId : "";
    return handleGetStudent(sheet, studentId);
  }
  
  return createJsonResponse({
    success: false, 
    message: "잘못된 action 요청입니다. (get_leaderboard, get_student 필요)"
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 10초 동시성 대기
  } catch(err) {
    return createJsonResponse({
      success: false,
      message: "동시 요청이 많아 처리가 지연되었습니다. 잠시 후 다시 시도해주세요."
    });
  }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    
    try {
      setupSheets(sheet);
    } catch(err) {
      return createJsonResponse({
        success: false,
        message: "초기화 에러: " + err.message
      });
    }
    
    var postData;
    try {
      postData = JSON.parse(e.postData.contents);
    } catch(err) {
      return createJsonResponse({
        success: false, 
        message: "JSON 파싱 에러: " + err.toString()
      });
    }
    
    var action = postData.action;
    
    if (action === 'register') {
      return handleRegister(sheet, postData);
    } else if (action === 'log_mission') {
      return handleLogMission(sheet, postData);
    } else if (action === 'submit_survey') {
      return handleSubmitSurvey(sheet, postData);
    }
    
    return createJsonResponse({ success: false, message: "알 수 없는 요청: " + action });
  } finally {
    lock.releaseLock();
  }
}

// 헬퍼: ID 구조 파싱 및 매핑 정보 반환 (한국어 시트 매핑 + 학교 정보 파싱)
function getParticipantDetails(studentId, schoolFromData) {
  var rawId = studentId ? studentId.toString().trim() : "";
  var school = schoolFromData || "";
  var cleanId = rawId;
  
  if (rawId.indexOf("_") > -1) {
    var parts = rawId.split("_");
    if (!school) school = parts[0];
    cleanId = parts[1];
  }
  if (!school) school = "A초";
  
  var isTeacher = cleanId.indexOf("T-") === 0;
  if (isTeacher) cleanId = cleanId.substring(2);
  
  return {
    rawId: rawId,
    school: school,
    isTeacher: isTeacher,
    cleanId: cleanId,
    profileSheet: isTeacher ? "교사기록" : "학생기록",
    dailySheet: isTeacher ? "교사일별스티커" : "학생일별스티커",
    surveySheet: isTeacher ? "교사설문응답" : "학생설문응답"
  };
}

// 헬퍼: 학번/개인번호 파싱 닉네임 생성
function getStudentNickname(cleanId, name, school) {
  var prefix = school ? "[" + school + "] " : "";
  if (cleanId.length === 4) {
    var grade = cleanId.charAt(0);
    var classNum = cleanId.charAt(1);
    var num = parseInt(cleanId.substring(2));
    return prefix + grade + "학년 " + classNum + "반 " + name;
  } else if (cleanId.length === 5) {
    var grade = cleanId.charAt(0);
    var classNum = parseInt(cleanId.substring(1, 3));
    var num = parseInt(cleanId.substring(3));
    return prefix + grade + "학년 " + classNum + "반 " + name;
  }
  return prefix + name;
}

// 시트 초기화 및 헤더 생성 (학교 컬럼 추가)
function setupSheets(sheet) {
  if (!sheet) {
    throw new Error("스프레드시트를 활성화할 수 없습니다.");
  }

  // 1. 기존 사용하지 않는 구버전 삭제 대상 탭 자동 삭제
  var deleteTabs = [
    "시트1", "Sheet1", "MissionLogs", 
    "Students", "Teachers", 
    "StudentDailyPoints", "TeacherDailyPoints", 
    "StudentMonthlySummary", "TeacherMonthlySummary",
    "StudentSurveyResponses", "TeacherSurveyResponses",
    "학생일별포인트", "교사일별포인트"
  ];
  deleteTabs.forEach(function(name) {
    var target = sheet.getSheetByName(name);
    if (target) {
      try {
        sheet.deleteSheet(target);
      } catch(e) {}
    }
  });

  // 2. 신규 6개 탭 구조 초기화 (학교 컬럼 B열 추가)
  var tabConfigs = [
    { name: "학생기록", headers: ["일시", "학교", "개인번호", "이름", "키(cm)", "몸무게(kg)", "BMI", "월총스티커", "누적총스티커", "레벨"], color: "#E2F0D9" },
    { name: "교사기록", headers: ["일시", "학교", "개인번호", "이름", "키(cm)", "몸무게(kg)", "BMI", "월총스티커", "누적총스티커", "레벨"], color: "#DDEBF7" },
    { name: "학생일별스티커", headers: ["일시", "학교", "개인번호", "이름", "일별스티커"], color: "#F2F2F2" },
    { name: "교사일별스티커", headers: ["일시", "학교", "개인번호", "이름", "일별스티커"], color: "#FFF2CC" },
    { name: "학생설문응답", headers: ["일시", "학교", "개인번호", "이름", "문항1", "문항2", "문항3", "문항4", "문항5", "문항6", "문항7", "문항8", "문항9", "문항10", "문항11", "문항12"], color: "#E2EFDA" },
    { name: "교사설문응답", headers: ["일시", "학교", "개인번호", "이름", "문항1", "문항2", "문항3", "문항4", "문항5", "문항6", "문항7", "문항8", "문항9", "문항10", "문항11", "문항12"], color: "#F8CBAD" }
  ];

  tabConfigs.forEach(function(config) {
    var targetSheet = sheet.getSheetByName(config.name);
    if (!targetSheet) {
      targetSheet = sheet.insertSheet(config.name);
    }
    
    // 헤더 행 설정 및 디자인 적용
    targetSheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
    targetSheet.getRange(1, 1, 1, config.headers.length).setFontWeight("bold").setBackground(config.color);
    
    // 개인번호 열(C열) 텍스트 포맷 설정으로 자릿수 보존
    targetSheet.getRange("C:C").setNumberFormat("@");
  });
}

// 헬퍼: 오늘자 일일 포인트 행 중복 정리 및 100P 상한 적용
function syncDailyRowForToday(dailySheet, p, name, pointsDelta, todayStr, todayDateStr, isSetAbsolute) {
  var dRows = dailySheet.getDataRange().getValues();
  var matchingRowIndices = [];
  
  for (var j = 1; j < dRows.length; j++) {
    var dDateStr = formatDateToYYYYMMDD(dRows[j][0]);
    var dSchool = dRows[j][1] ? dRows[j][1].toString().trim() : "";
    var dId = dRows[j][2] ? dRows[j][2].toString().trim() : "";
    if (dDateStr === todayDateStr && dId === p.cleanId && (dSchool === p.school || !dSchool)) {
      matchingRowIndices.push(j + 1);
    }
  }
  
  var targetPoints = 0;
  if (matchingRowIndices.length > 0) {
    var primaryRowIndex = matchingRowIndices[0];
    var curPoints = parseInt(dailySheet.getRange(primaryRowIndex, 5).getValue()) || 0;
    // If curPoints is stored as 1 sticker, convert to points for delta calculation
    if (curPoints < 10) curPoints = curPoints * 100;
    targetPoints = isSetAbsolute ? Math.min(100, Math.max(0, pointsDelta)) : Math.min(100, Math.max(0, curPoints + pointsDelta));
    var targetStickers = (data && data.dailySticker !== undefined && data.dailySticker !== null) ? Number(data.dailySticker) : 1;
    
    dailySheet.getRange(primaryRowIndex, 1).setValue(todayStr);
    dailySheet.getRange(primaryRowIndex, 2).setValue(p.school);
    if (name) dailySheet.getRange(primaryRowIndex, 4).setValue(name);
    dailySheet.getRange(primaryRowIndex, 5).setValue(targetStickers);
    
    for (var k = matchingRowIndices.length - 1; k > 0; k--) {
      dailySheet.deleteRow(matchingRowIndices[k]);
    }
  } else {
    targetPoints = Math.min(100, Math.max(0, pointsDelta));
    var targetStickers = (data && data.dailySticker !== undefined && data.dailySticker !== null) ? Number(data.dailySticker) : 1;
    dailySheet.appendRow([todayStr, p.school, "'" + p.cleanId, name, targetStickers]);
  }
  
  SpreadsheetApp.flush();
}

// 1. 참여자 가입/기록 저장 API
function handleRegister(sheet, data) {
  var p = getParticipantDetails(data.studentId, data.school);
  var profileSheet = sheet.getSheetByName(p.profileSheet);
  
  var name = (data.name ? data.name : "").toString().trim();
  var height = parseFloat(data.height) || 0;
  var weight = parseFloat(data.weight) || 0;
  var bmi = (height > 0) ? parseFloat((weight / ((height / 100) * (height / 100))).toFixed(1)) : 0;
  
  var todayStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  var todayDateStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  var currentMonthStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM");
  
  var rows = profileSheet.getDataRange().getValues();
  var foundRowIndex = -1;
  
  for (var i = 1; i < rows.length; i++) {
    var cellValue = rows[i][0];
    var rowMonthStr = formatDateToYYYYMM(cellValue);
    var rowSchool = rows[i][1] ? rows[i][1].toString().trim() : "";
    var rowId = rows[i][2] ? rows[i][2].toString().trim() : "";
    if (rowId === p.cleanId && (rowSchool === p.school || !rowSchool) && rowMonthStr === currentMonthStr) {
      foundRowIndex = i + 1;
      break;
    }
  }
  
  if (foundRowIndex > -1) {
    profileSheet.getRange(foundRowIndex, 1).setValue(todayStr);
    profileSheet.getRange(foundRowIndex, 2).setValue(p.school);
    if (name) profileSheet.getRange(foundRowIndex, 4).setValue(name);
    if (height > 0) profileSheet.getRange(foundRowIndex, 5).setValue(height);
    if (weight > 0) profileSheet.getRange(foundRowIndex, 6).setValue(weight);
    
    var curH = parseFloat(profileSheet.getRange(foundRowIndex, 5).getValue()) || 0;
    var curW = parseFloat(profileSheet.getRange(foundRowIndex, 6).getValue()) || 0;
    var curBmi = (curH > 0) ? parseFloat((curW / ((curH / 100) * (curH / 100))).toFixed(1)) : 0;
    profileSheet.getRange(foundRowIndex, 7).setValue(curBmi);
  } else {
    var prevH = height;
    var prevW = weight;
    
    if (prevH <= 0 || prevW <= 0) {
      for (var r = rows.length - 1; r >= 1; r--) {
        var rSch = rows[r][1] ? rows[r][1].toString().trim() : "";
        var rId = rows[r][2] ? rows[r][2].toString().trim() : "";
        if (rId === p.cleanId && (rSch === p.school || !rSch)) {
          if (prevH <= 0 && parseFloat(rows[r][4]) > 0) prevH = parseFloat(rows[r][4]);
          if (prevW <= 0 && parseFloat(rows[r][5]) > 0) prevW = parseFloat(rows[r][5]);
          if (prevH > 0 && prevW > 0) break;
        }
      }
    }
    
    var calcBmi = (prevH > 0 && prevW > 0) ? parseFloat((prevW / ((prevH / 100) * (prevH / 100))).toFixed(1)) : 0;
    profileSheet.appendRow([todayStr, p.school, "'" + p.cleanId, name, prevH, prevW, calcBmi, 0, 0, "알콩이"]);
  }
  
  var dailySheet = sheet.getSheetByName(p.dailySheet);
  syncDailyRowForToday(dailySheet, p, name, 0, todayStr, todayDateStr, false);
  updateProfilePointsRealtime(sheet, p, name, todayStr, currentMonthStr);
  
  return createJsonResponse({
    success: true,
    message: "새로운 참여자 기록이 추가되고 정상 등록되었습니다!",
    isNew: true
  });
}

// 2. 미션 및 보너스 포인트 기록 API
function handleLogMission(sheet, data) {
  var p = getParticipantDetails(data.studentId, data.school);
  var dailySheet = sheet.getSheetByName(p.dailySheet);
  var profileSheet = sheet.getSheetByName(p.profileSheet);
  
  var name = data.name ? data.name.trim() : "";
  var pointsDelta = parseInt(data.points) || 0;
  var height = data.height ? parseFloat(data.height) : "";
  var weight = data.weight ? parseFloat(data.weight) : "";
  var todayStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  var todayDateStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  var currentMonthStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM");
  
  if (!name) {
    var pRows = profileSheet.getDataRange().getValues();
    for (var i = 1; i < pRows.length; i++) {
      var rSch = pRows[i][1] ? pRows[i][1].toString().trim() : "";
      var rId = pRows[i][2] ? pRows[i][2].toString().trim() : "";
      if (rId === p.cleanId && (rSch === p.school || !rSch)) {
        name = pRows[i][3];
        break;
      }
    }
  }
  
  syncDailyRowForToday(dailySheet, p, name, pointsDelta, todayStr, todayDateStr, data.isSetAbsolute);
  
  if (weight !== "" || height !== "") {
    var pRows2 = profileSheet.getDataRange().getValues();
    for (var k = 1; k < pRows2.length; k++) {
      var cellValue = pRows2[k][0];
      var rowMonthStr = formatDateToYYYYMM(cellValue);
      var rSch2 = pRows2[k][1] ? pRows2[k][1].toString().trim() : "";
      var rId2 = pRows2[k][2] ? pRows2[k][2].toString().trim() : "";
      if (rId2 === p.cleanId && (rSch2 === p.school || !rSch2) && rowMonthStr === currentMonthStr) {
        var rowIdx = k + 1;
        profileSheet.getRange(rowIdx, 1).setValue(todayStr);
        profileSheet.getRange(rowIdx, 2).setValue(p.school);
        if (height !== "" && parseFloat(height) > 0) profileSheet.getRange(rowIdx, 5).setValue(parseFloat(height));
        if (weight !== "" && parseFloat(weight) > 0) profileSheet.getRange(rowIdx, 6).setValue(parseFloat(weight));
        var curHeight = parseFloat(profileSheet.getRange(rowIdx, 5).getValue()) || 0;
        var curWeight = parseFloat(profileSheet.getRange(rowIdx, 6).getValue()) || 0;
        var newBmi = (curHeight > 0) ? parseFloat((curWeight / ((curHeight / 100) * (curHeight / 100))).toFixed(1)) : 0;
        profileSheet.getRange(rowIdx, 7).setValue(newBmi);
        break;
      }
    }
  }

  updateProfilePointsRealtime(sheet, p, name, todayStr, currentMonthStr);
  
  return createJsonResponse({
    success: true,
    message: "포인트 정보가 실시간 적립/수정되었습니다."
  });
}

// 헬퍼: 대표 시트(학생기록/교사기록)의 월총포인트, 누적총포인트, 레벨 실시간 동기화
function updateProfilePointsRealtime(sheet, p, name, todayStr, currentMonthStr) {
  var dailySheet = sheet.getSheetByName(p.dailySheet);
  var profileSheet = sheet.getSheetByName(p.profileSheet);
  
  SpreadsheetApp.flush();
  var updatedDRows = dailySheet.getDataRange().getValues();
  var cumulativeStickers = 0;
  var monthlyStickers = 0;
  
  for (var m = 1; m < updatedDRows.length; m++) {
    var dSch = updatedDRows[m][1] ? updatedDRows[m][1].toString().trim() : "";
    var dId = updatedDRows[m][2] ? updatedDRows[m][2].toString().trim() : "";
    if (dId === p.cleanId && (dSch === p.school || !dSch)) {
      var pts = parseInt(updatedDRows[m][4]) || 0;
      var stk = (pts >= 100) ? Math.floor(pts / 100) : pts;
      cumulativeStickers += stk;
      if (formatDateToYYYYMM(updatedDRows[m][0]) === currentMonthStr) {
        monthlyStickers += stk;
      }
    }
  }

  var level = "🥚 알콩이 (0~7개)";
  if (cumulativeStickers >= 45) level = "👑 꼬꼬대장 (45개+)";
  else if (cumulativeStickers >= 24) level = "🐥 튼튼이 (24~44개)";
  else if (cumulativeStickers >= 8) level = "🐣 삐약이 (8~23개)";

  var pRows3 = profileSheet.getDataRange().getValues();
  var foundProfileRow = -1;
  for (var pIdx = 1; pIdx < pRows3.length; pIdx++) {
    var cellVal = pRows3[pIdx][0];
    var pMonthStr = formatDateToYYYYMM(cellVal);
    var pSch = pRows3[pIdx][1] ? pRows3[pIdx][1].toString().trim() : "";
    var pId = pRows3[pIdx][2] ? pRows3[pIdx][2].toString().trim() : "";
    if (pId === p.cleanId && (pSch === p.school || !pSch) && pMonthStr === currentMonthStr) {
      foundProfileRow = pIdx + 1;
      break;
    }
  }

  if (foundProfileRow > -1) {
    profileSheet.getRange(foundProfileRow, 1).setValue(todayStr);            // A: 일시
    profileSheet.getRange(foundProfileRow, 2).setValue(p.school);             // B: 학교
    if (name) profileSheet.getRange(foundProfileRow, 4).setValue(name);      // D: 이름
    profileSheet.getRange(foundProfileRow, 8).setValue(monthlyStickers);     // H: 월총스티커
    profileSheet.getRange(foundProfileRow, 9).setValue(cumulativeStickers);  // I: 누적총스티커
    profileSheet.getRange(foundProfileRow, 10).setValue(level);               // J: 레벨
  } else {
    profileSheet.appendRow([todayStr, p.school, "'" + p.cleanId, name, 0, 0, 0, monthlyStickers, cumulativeStickers, level]);
  }
  
  SpreadsheetApp.flush();
}

// 3. 참여자 신상 정보 및 누적 기록 조회 API
function handleGetStudent(sheet, studentId) {
  if (!studentId) {
    return createJsonResponse({ success: false, message: "개인번호(ID) 파라미터가 누락되었습니다." });
  }
  
  var p = getParticipantDetails(studentId);
  var profileSheet = sheet.getSheetByName(p.profileSheet);
  var pRows = profileSheet.getDataRange().getValues();
  
  var latestRowIndex = -1;
  for (var i = pRows.length - 1; i >= 1; i--) {
    var rSch = pRows[i][1] ? pRows[i][1].toString().trim() : "";
    var rId = pRows[i][2] ? pRows[i][2].toString().trim() : "";
    if (rId === p.cleanId && (rSch === p.school || !rSch)) {
      latestRowIndex = i;
      break;
    }
  }
  
  if (latestRowIndex === -1) {
    return createJsonResponse({ success: false, isRegistered: false, message: "등록되지 않은 개인번호입니다." });
  }
  
  var school = pRows[latestRowIndex][1] || p.school;
  var name = pRows[latestRowIndex][3];
  var height = parseFloat(pRows[latestRowIndex][4]) || 0;
  var weight = parseFloat(pRows[latestRowIndex][5]) || 0;
  
  var totalStickers = 0;
  var monthlyStickers = 0;
  var currentMonthStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM");
  var dailySheet = sheet.getSheetByName(p.dailySheet);
  var dRows = dailySheet.getDataRange().getValues();
  for (var j = 1; j < dRows.length; j++) {
    var dSch = dRows[j][1] ? dRows[j][1].toString().trim() : "";
    var dId = dRows[j][2] ? dRows[j][2].toString().trim() : "";
    if (dId === p.cleanId && (dSch === school || !dSch)) {
      var pts = parseInt(dRows[j][4]) || 0;
      var stk = (pts >= 100) ? Math.floor(pts / 100) : pts;
      totalStickers += stk;
      if (formatDateToYYYYMM(dRows[j][0]) === currentMonthStr) {
        monthlyStickers += stk;
      }
    }
  }
  
  // 사전 설문 조사 참여 여부 체크
  var surveySheet = sheet.getSheetByName(p.surveySheet);
  var sRows = surveySheet.getDataRange().getValues();
  var preSurveyDone = false;
  for (var k = 1; k < sRows.length; k++) {
    var sSch = sRows[k][1] ? sRows[k][1].toString().trim() : "";
    var sId = sRows[k][2] ? sRows[k][2].toString().trim() : "";
    if (sId === p.cleanId && (sSch === school || !sSch)) {
      preSurveyDone = true;
      break;
    }
  }
  
  // 닉네임 자동 빌드
  var nickname = p.isTeacher ? ("[" + school + "] " + name + " 선생님") : getStudentNickname(p.cleanId, name, school);
  
  var studentData = {
    studentId: studentId,
    school: school,
    name: name,
    nickname: nickname,
    height: height,
    weight: weight,
    totalPoints: totalStickers * 100,
    monthlyPoints: monthlyStickers * 100,
    preSurveyDone: preSurveyDone
  };
  
  // 미션 수행 로그(일별 포인트) 변환
  var history = [];
  for (var m = 1; m < dRows.length; m++) {
    var hdSch = dRows[m][1] ? dRows[m][1].toString().trim() : "";
    var hdId = dRows[m][2] ? dRows[m][2].toString().trim() : "";
    if (hdId === p.cleanId && (hdSch === school || !hdSch)) {
      history.push({
        date: formatDateToYYYYMMDDHHMMSS(dRows[m][0]),
        mission: "일일 미션 적립 완료 🐾",
        weight: null,
        points: parseInt(dRows[m][4]) || 0
      });
    }
  }
  history.reverse();
  studentData.history = history;
  
  var weightHistory = [];
  for (var n = 1; n < pRows.length; n++) {
    var hpSch = pRows[n][1] ? pRows[n][1].toString().trim() : "";
    var hpId = pRows[n][2] ? pRows[n][2].toString().trim() : "";
    if (hpId === p.cleanId && (hpSch === school || !hpSch) && pRows[n][5] !== "") {
      weightHistory.push({
        date: formatDateToYYYYMM(pRows[n][0]),
        weight: parseFloat(pRows[n][5])
      });
    }
  }
  studentData.weightHistory = weightHistory;
  studentData.isRegistered = true;
  
  return createJsonResponse({ success: true, student: studentData });
}

// 4. 리더보드 랭킹 통합 API (현재 월에 획득한 포인트 순으로 정렬)
function handleGetLeaderboard(sheet) {
  var studentSheet = sheet.getSheetByName("학생기록");
  var teacherSheet = sheet.getSheetByName("교사기록");
  var dailyStudentSheet = sheet.getSheetByName("학생일별스티커") || sheet.getSheetByName("학생일별포인트");
  var dailyTeacherSheet = sheet.getSheetByName("교사일별스티커") || sheet.getSheetByName("교사일별포인트");
  
  var leaderboard = [];
  var currentMonthStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM");
  
  var realMonthlyPoints = {};
  [dailyStudentSheet, dailyTeacherSheet].forEach(function(ds) {
    if (!ds) return;
    var dRows = ds.getDataRange().getValues();
    for (var i = 1; i < dRows.length; i++) {
      if (formatDateToYYYYMM(dRows[i][0]) === currentMonthStr) {
         var sch = dRows[i][1] ? dRows[i][1].toString().trim() : "A초";
         var id = dRows[i][2] ? dRows[i][2].toString().trim() : "";
         var key = sch + "_" + id;
         realMonthlyPoints[key] = (realMonthlyPoints[key] || 0) + (parseInt(dRows[i][4]) || 0);
      }
    }
  });

  // 학생용 리더보드 빌드
  if (studentSheet) {
    var sRows = studentSheet.getDataRange().getValues();
    var sPointsMap = {};
    var sNamesMap = {};
    var sSchoolsMap = {};
    for (var i = 1; i < sRows.length; i++) {
      var cellValue = sRows[i][0];
      var rowMonthStr = formatDateToYYYYMM(cellValue);
      var sch = sRows[i][1] ? sRows[i][1].toString().trim() : "A초";
      var id = sRows[i][2] ? sRows[i][2].toString().trim() : "";
      var name = sRows[i][3] ? sRows[i][3].toString().trim() : "";
      var key = sch + "_" + id;
      
      if (id && rowMonthStr === currentMonthStr) {
        sPointsMap[key] = realMonthlyPoints[key] || 0;
        sNamesMap[key] = name;
        sSchoolsMap[key] = sch;
      }
    }
    Object.keys(sPointsMap).forEach(function(key) {
      var parts = key.split("_");
      var sch = parts[0];
      var id = parts[1];
      var name = sNamesMap[key];
      var classGroup = sch;
      var grade = id.charAt(0);
      var classNum = id.length === 5 ? parseInt(id.substring(1, 3)) : id.charAt(1);
      if (!isNaN(grade) && !isNaN(classNum)) {
        classGroup = sch + " " + grade + "학년 " + classNum + "반";
      }
      leaderboard.push({
        nickname: getStudentNickname(id, name, sch),
        classGroup: classGroup,
        points: sPointsMap[key]
      });
    });
  }
  
  leaderboard.sort(function(a, b) {
    return b.points - a.points;
  });
  
  return createJsonResponse({
    success: true,
    leaderboard: leaderboard.slice(0, 50)
  });
}

// 5. 설문조사 제출 API
function handleSubmitSurvey(sheet, data) {
  var p = getParticipantDetails(data.studentId, data.school);
  var surveySheet = sheet.getSheetByName(p.surveySheet);
  
  var name = (data.name ? data.name : "").toString().trim();
  var answers = data.answers;
  var todayStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  
  if (surveySheet.getLastRow() === 0) {
    surveySheet.appendRow(["일시", "학교", "개인번호", "이름", "문항1", "문항2", "문항3", "문항4", "문항5", "문항6", "문항7", "문항8", "문항9", "문항10", "문항11", "문항12"]);
  } else {
    var lastCol = surveySheet.getLastColumn();
    if (lastCol < 16) {
      for (var col = lastCol + 1; col <= 16; col++) {
        surveySheet.getRange(1, col).setValue("문항" + (col - 4));
      }
    }
  }
  
  var rowData = [todayStr, p.school, "'" + p.cleanId, name];
  var totalQ = (answers && answers.length) ? answers.length : 12;
  for (var i = 0; i < totalQ; i++) {
    rowData.push(answers[i] !== undefined && answers[i] !== null ? answers[i] : "");
  }
  
  surveySheet.appendRow(rowData);
  SpreadsheetApp.flush();
  
  return createJsonResponse({
    success: true,
    message: "사전 설문조사가 성공적으로 제출되었습니다! 🌟"
  });
}

// 헬퍼: JSON 응답 생성
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

// 헬퍼: 다양한 타입/포맷의 날짜 데이터를 yyyy-MM-dd 문자열로 안전하게 변환
function formatDateToYYYYMMDD(cellValue) {
  if (!cellValue) return "";
  
  // 1. Date 객체이거나 Date 프로토타입을 가지는 경우
  if (typeof cellValue.getTime === 'function' || 
      Object.prototype.toString.call(cellValue) === '[object Date]' || 
      cellValue instanceof Date) {
    try {
      return Utilities.formatDate(cellValue, "Asia/Seoul", "yyyy-MM-dd");
    } catch (e) {
      // 포맷 실패 시 문자열 파싱 시도
    }
  }
  
  // 2. 문자열 형식 파싱
  var str = cellValue.toString().trim();
  if (!str) return "";
  
  // 패턴 A: 2026-06-10 ...
  var matchDash = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (matchDash) {
    return matchDash[1] + "-" + matchDash[2] + "-" + matchDash[3];
  }
  
  // 패턴 B: 2026. 06. 10 ...
  var matchDot = str.match(/^(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  if (matchDot) {
    var y = matchDot[1];
    var m = matchDot[2].length === 1 ? "0" + matchDot[2] : matchDot[2];
    var d = matchDot[3].length === 1 ? "0" + matchDot[3] : matchDot[3];
    return y + "-" + m + "-" + d;
  }
  
  // 패턴 C: 2026/06/10 ...
  var matchSlash = str.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (matchSlash) {
    var y = matchSlash[1];
    var m = matchSlash[2].length === 1 ? "0" + matchSlash[2] : matchSlash[2];
    var d = matchSlash[3].length === 1 ? "0" + matchSlash[3] : matchSlash[3];
    return y + "-" + m + "-" + d;
  }

  // 패턴 D: 영문 요일/월이 포함된 문자열 등 Date 파싱 시도 (예: Wed Jun 10 2026...)
  try {
    var parsedDate = new Date(str);
    if (!isNaN(parsedDate.getTime())) {
      return Utilities.formatDate(parsedDate, "Asia/Seoul", "yyyy-MM-dd");
    }
  } catch (e) {}
  
  return "";
}

// 헬퍼: 다양한 타입/포맷의 날짜 데이터를 yyyy-MM 문자열로 안전하게 변환
function formatDateToYYYYMM(cellValue) {
  var yyyymmdd = formatDateToYYYYMMDD(cellValue);
  if (yyyymmdd.length >= 7) {
    return yyyymmdd.substring(0, 7);
  }
  return "";
}

// 헬퍼: 다양한 타입/포맷의 날짜 데이터를 yyyy-MM-dd HH:mm:ss 문자열로 안전하게 변환
function formatDateToYYYYMMDDHHMMSS(cellValue) {
  if (!cellValue) return "";
  if (typeof cellValue.getTime === 'function' || 
      Object.prototype.toString.call(cellValue) === '[object Date]' || 
      cellValue instanceof Date) {
    try {
      return Utilities.formatDate(cellValue, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    } catch (e) {}
  }
  return cellValue.toString().trim();
}

// -------------------------------------------------------------
// [배치 작업] 매월 30일 자정 결산 트리거 설치용 함수
// -------------------------------------------------------------
function setupMonthlyTrigger() {
  // 기존 중복 트리거 방지를 위해 모두 삭제
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'aggregateMonthlyPoints') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 매월 30일 새벽 1시에 실행되도록 설정
  ScriptApp.newTrigger('aggregateMonthlyPoints')
    .timeBased()
    .onMonthDay(30)
    .atHour(1)
    .create();
    
  Logger.log("매월 30일 배치 처리 트리거가 설정되었습니다.");
}

// -------------------------------------------------------------
// [배치 작업] 매월 30일에 실행되어 일별 포인트를 
// 각 기록 시트의 월총포인트, 누적총포인트, 최종 레벨로 합산 기록
// -------------------------------------------------------------
function aggregateMonthlyPoints() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var configs = [
    { profile: "학생기록", daily: "학생일별포인트" },
    { profile: "교사기록", daily: "교사일별포인트" }
  ];
  
  configs.forEach(function(config) {
    var profileSheet = sheet.getSheetByName(config.profile);
    var dailySheet = sheet.getSheetByName(config.daily);
    if (!profileSheet || !dailySheet) return;
    
    var pRows = profileSheet.getDataRange().getValues();
    var dRows = dailySheet.getDataRange().getValues();
    
    for (var i = 1; i < pRows.length; i++) {
      var dateStr = formatDateToYYYYMM(pRows[i][0]); // 해당 기록행의 연월 (예: "2026-06")
      var cleanId = pRows[i][1].toString().trim();
      if (!cleanId) continue;
      
      var monthlyPoints = 0;
      var cumulativePoints = 0;
      
      // 6개월 이상(영구) 쌓이는 누적 포인트 집계
      for (var j = 1; j < dRows.length; j++) {
        var dId = dRows[j][1].toString().trim();
        if (dId === cleanId) {
          var pts = parseInt(dRows[j][3]) || 0;
          cumulativePoints += pts;
          
          var dDateStr = formatDateToYYYYMM(dRows[j][0]);
          if (dDateStr === dateStr) {
            monthlyPoints += pts;
          }
        }
      }
      
      var totalStickers = Math.floor(cumulativePoints / 100);
      var level = "🥚 알콩이 (0~7개)";
      if (totalStickers >= 45) level = "👑 꼬꼬대장 (45개+)";
      else if (totalStickers >= 24) level = "🐥 튼튼이 (24~44개)";
      else if (totalStickers >= 8) level = "🐣 삐약이 (8~23개)";
      
      profileSheet.getRange(i + 1, 8).setValue(monthlyPoints);
      profileSheet.getRange(i + 1, 9).setValue(cumulativePoints);
      profileSheet.getRange(i + 1, 10).setValue(level);
    }
  });
}

// -------------------------------------------------------------
// [플랜 B - 관리자 메뉴 및 비상 복구 기능]
// -------------------------------------------------------------
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ 튼튼탐험대 관리자')
    .addItem('🔄 전체 포인트 재계산 및 복구', 'recalculateAllPoints')
    .addItem('📁 즉시 시트 백업본 생성', 'createDailyBackup')
    .addItem('⏰ 매일 밤 10시 자동 백업 예약 활성화', 'setupDaily10PMBackupTrigger')
    .addToUi();
}

/**
 * 매일 밤 10시(22:00)에 자동으로 백업본을 생성하는 구글 시간 트리거 설정
 */
function setupDaily10PMBackupTrigger() {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch(e){}
  
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'createDailyBackup') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    ScriptApp.newTrigger('createDailyBackup')
      .timeBased()
      .everyDays(1)
      .atHour(22)
      .nearMinute(0)
      .create();
      
    if (ui) {
      ui.alert("✅ 자동 백업 설정 완료", "매일 밤 10시에 자동으로 구글 드라이브 ['📂 튼튼탐험대_백업모음'] 폴더에 백업본이 저장되도록 예약되었습니다! 🌙", ui.ButtonSet.OK);
    }
  } catch(e) {
    if (ui) {
      ui.alert("❌ 설정 실패", "자동 백업 트리거 생성 중 오류: " + e.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * 플랜 B 복구용: 일별 원본 기록(학생일별포인트/교사일별포인트)을 스캔하여
 * 학생기록/교사기록의 월총포인트, 누적총포인트, 레벨을 100% 원상복구합니다.
 */
function recalculateAllPoints() {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch(e){}
  try {
    aggregateMonthlyPoints();
    if (ui) ui.alert("✅ 포인트 복구 완료", "일별 원본 기록을 바탕으로 모든 학생/교사의 총포인트 및 레벨 재계산이 완료되었습니다!", ui.ButtonSet.OK);
  } catch(e) {
    if (ui) ui.alert("❌ 복구 실패", "재계산 중 오류 발생: " + e.message, ui.ButtonSet.OK);
  }
}

/**
 * 플랜 B 백업용: 현재 구글 시트 전체를 복사하여 구글 드라이브에 시각별 백업본 파일 생성
 */
function createDailyBackup() {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch(e){}
  
  try {
    aggregateMonthlyPoints(); // 백업 전 최신 포인트 & 레벨 자동 최종 재계산!
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var name = ss.getName() + "_백업_" + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd_HHmmss");
    var folderName = "📂 튼튼탐험대_백업모음";
    
    try {
      var folders = DriveApp.getFoldersByName(folderName);
      var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      var file = DriveApp.getFileById(ss.getId());
      file.makeCopy(name, targetFolder);
    } catch(errDrive) {
      // 권한 동의 미완료 시 내 드라이브에 기본 복사
      ss.copy(name);
    }
    
    if (ui) {
      ui.alert("✅ 백업 완료", "최신 데이터 재계산 후 구글 드라이브 ['" + folderName + "'] 전용 폴더에 백업본이 저장되었습니다!", ui.ButtonSet.OK);
    }
  } catch(e) {
    if (ui) {
      ui.alert("❌ 백업 실패", "백업 생성 중 오류 발생: " + e.message, ui.ButtonSet.OK);
    }
    console.error("Backup error: " + e.message);
  }
}
