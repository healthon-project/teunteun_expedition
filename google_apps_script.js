/**
 * 튼튼탐험대 (TeunTeun Expedition) - 구글 앱스 스크립트 (Google Apps Script)
 * 학생용/교사역 분리 및 개인번호(4자리) 기반 누적 합산 스키마 (데이트 파싱 버그 수정 버젼)
 */

function doGet(e) {
  var action = e.parameter.action;
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
    var studentId = e.parameter.studentId;
    return handleGetStudent(sheet, studentId);
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: false, 
    message: "잘못된 action 요청입니다. (get_leaderboard, get_student 필요)"
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
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
    return ContentService.createTextOutput(JSON.stringify({
      success: false, 
      message: "JSON 파싱 에러: " + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  var action = postData.action;
  
  if (action === 'register') {
    return handleRegister(sheet, postData);
  } else if (action === 'log_mission') {
    return handleLogMission(sheet, postData);
  } else if (action === 'submit_survey') {
    return handleSubmitSurvey(sheet, postData);
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: false, 
    message: "잘못된 action 요청입니다."
  })).setMimeType(ContentService.MimeType.JSON);
}

// 헬퍼: ID 구조 파싱 및 매핑 정보 반환 (한국어 시트 매핑)
function getParticipantDetails(studentId) {
  var id = studentId.toString().trim();
  var isTeacher = id.indexOf("T-") === 0;
  var cleanId = isTeacher ? id.substring(2) : id;
  
  return {
    isTeacher: isTeacher,
    cleanId: cleanId,
    profileSheet: isTeacher ? "교사기록" : "학생기록", // 키,몸무게,BMI,월총포인트,레벨 통합 탭
    dailySheet: isTeacher ? "교사일별포인트" : "학생일별포인트", // 매일 적립포인트 탭
    surveySheet: isTeacher ? "교사설문응답" : "학생설문응답" // 설문내용 탭
  };
}

// 헬퍼: 학번/개인번호 파싱 닉네임 생성
function getStudentNickname(cleanId, name) {
  if (cleanId.length === 4) {
    var grade = cleanId.charAt(0);
    var classNum = cleanId.charAt(1);
    var num = parseInt(cleanId.substring(2));
    return grade + "학년 " + classNum + "반 " + name;
  } else if (cleanId.length === 5) {
    var grade = cleanId.charAt(0);
    var classNum = parseInt(cleanId.substring(1, 3));
    var num = parseInt(cleanId.substring(3));
    return grade + "학년 " + classNum + "반 " + name;
  }
  return name;
}

// 시트 초기화 및 헤더 생성
function setupSheets(sheet) {
  if (!sheet) {
    throw new Error("스프레드시트를 활성화할 수 없습니다.");
  }

  // 1. 기존 사용하지 않는 삭제 대상 탭 삭제 (기존 영문 탭들 포함하여 자동 청소)
  var deleteTabs = [
    "시트1", "Sheet1", "MissionLogs", 
    "Students", "Teachers", 
    "StudentDailyPoints", "TeacherDailyPoints", 
    "StudentMonthlySummary", "TeacherMonthlySummary",
    "StudentSurveyResponses", "TeacherSurveyResponses"
  ];
  deleteTabs.forEach(function(name) {
    var target = sheet.getSheetByName(name);
    if (target) {
      try {
        sheet.deleteSheet(target);
      } catch(e) {
        // 스프레드시트에 최소 1개의 시트는 남아있어야 함
      }
    }
  });

  // 2. 신규 6개 탭 구조 초기화 (학생기록/교사기록의 월별포인트를 월총포인트로 변경, 일별포인트의 일일포인트를 일총포인트로 변경)
  var tabConfigs = [
    { name: "학생기록", headers: ["일시", "개인번호", "이름", "키(cm)", "몸무게(kg)", "BMI", "월총포인트", "누적총포인트", "레벨"], color: "#E2F0D9" },
    { name: "교사기록", headers: ["일시", "개인번호", "이름", "키(cm)", "몸무게(kg)", "BMI", "월총포인트", "누적총포인트", "레벨"], color: "#DDEBF7" },
    { name: "학생일별포인트", headers: ["일시", "개인번호", "이름", "일총포인트"], color: "#F2F2F2" },
    { name: "교사일별포인트", headers: ["일시", "개인번호", "이름", "일총포인트"], color: "#FFF2CC" },
    { name: "학생설문응답", headers: ["일시", "개인번호", "이름", "문항1", "문항2", "문항3", "문항4", "문항5", "문항6", "문항7", "문항8", "문항9", "문항10", "문항11"], color: "#E2EFDA" },
    { name: "교사설문응답", headers: ["일시", "개인번호", "이름", "문항1", "문항2", "문항3", "문항4", "문항5", "문항6", "문항7", "문항8", "문항9", "문항10", "문항11"], color: "#F8CBAD" }
  ];

  tabConfigs.forEach(function(config) {
    var targetSheet = sheet.getSheetByName(config.name);
    if (!targetSheet) {
      targetSheet = sheet.insertSheet(config.name);
    }
    
    // 헤더 행 설정 및 디자인 적용
    targetSheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
    targetSheet.getRange(1, 1, 1, config.headers.length).setFontWeight("bold").setBackground(config.color);
    
    // 개인번호 열(B열) 텍스트 포맷 설정으로 자릿수 보존
    targetSheet.getRange("B:B").setNumberFormat("@");
  });
}

// 헬퍼: 오늘자 일일 포인트 행 중복 정리 및 100P 상한 적용
function syncDailyRowForToday(dailySheet, cleanId, name, pointsDelta, todayStr, todayDateStr, isSetAbsolute) {
  var dRows = dailySheet.getDataRange().getValues();
  var matchingRowIndices = [];
  
  for (var j = 1; j < dRows.length; j++) {
    var dDateStr = formatDateToYYYYMMDD(dRows[j][0]);
    if (dDateStr === todayDateStr && dRows[j][1].toString().trim() === cleanId) {
      matchingRowIndices.push(j + 1); // 1-based index
    }
  }
  
  var targetPoints = 0;
  if (matchingRowIndices.length > 0) {
    var primaryRowIndex = matchingRowIndices[0];
    var curPoints = parseInt(dailySheet.getRange(primaryRowIndex, 4).getValue()) || 0;
    targetPoints = isSetAbsolute ? Math.min(100, Math.max(0, pointsDelta)) : Math.min(100, Math.max(0, curPoints + pointsDelta));
    
    dailySheet.getRange(primaryRowIndex, 1).setValue(todayStr);
    if (name) dailySheet.getRange(primaryRowIndex, 3).setValue(name);
    dailySheet.getRange(primaryRowIndex, 4).setValue(targetPoints);
    
    // 중복 생성된 오늘자 행 삭제 (아래 행부터 안전하게 삭제)
    for (var k = matchingRowIndices.length - 1; k > 0; k--) {
      dailySheet.deleteRow(matchingRowIndices[k]);
    }
  } else {
    targetPoints = Math.min(100, Math.max(0, pointsDelta));
    dailySheet.appendRow([todayStr, "'" + cleanId, name, targetPoints]);
  }
  
  SpreadsheetApp.flush();
}

// 1. 참여자 가입/기록 저장 API
function handleRegister(sheet, data) {
  var p = getParticipantDetails(data.studentId);
  var profileSheet = sheet.getSheetByName(p.profileSheet);
  
  var name = data.name.trim();
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
    var rowId = rows[i][1].toString().trim();
    if (rowId === p.cleanId && rowMonthStr === currentMonthStr) {
      foundRowIndex = i + 1;
      break;
    }
  }
  
  if (foundRowIndex > -1) {
    profileSheet.getRange(foundRowIndex, 1).setValue(todayStr);
    if (name) profileSheet.getRange(foundRowIndex, 3).setValue(name);
    if (height > 0) profileSheet.getRange(foundRowIndex, 4).setValue(height);
    if (weight > 0) profileSheet.getRange(foundRowIndex, 5).setValue(weight);
    
    var curH = parseFloat(profileSheet.getRange(foundRowIndex, 4).getValue()) || 0;
    var curW = parseFloat(profileSheet.getRange(foundRowIndex, 5).getValue()) || 0;
    var curBmi = (curH > 0) ? parseFloat((curW / ((curH / 100) * (curH / 100))).toFixed(1)) : 0;
    profileSheet.getRange(foundRowIndex, 6).setValue(curBmi);
  } else {
    var prevH = height;
    var prevW = weight;
    
    if (prevH <= 0 || prevW <= 0) {
      for (var r = rows.length - 1; r >= 1; r--) {
        if (rows[r][1].toString().trim() === p.cleanId) {
          if (prevH <= 0 && parseFloat(rows[r][3]) > 0) prevH = parseFloat(rows[r][3]);
          if (prevW <= 0 && parseFloat(rows[r][4]) > 0) prevW = parseFloat(rows[r][4]);
          if (prevH > 0 && prevW > 0) break;
        }
      }
    }
    
    var calcBmi = (prevH > 0 && prevW > 0) ? parseFloat((prevW / ((prevH / 100) * (prevH / 100))).toFixed(1)) : 0;
    profileSheet.appendRow([todayStr, "'" + p.cleanId, name, prevH, prevW, calcBmi, 0, 0, "알콩이"]);
  }
  
  // 신규 가입/로그인 시 오늘자 일일 행을 0P에서 시작 (미션 수행에 따라 100P까지 누적)
  var dailySheet = sheet.getSheetByName(p.dailySheet);
  syncDailyRowForToday(dailySheet, p.cleanId, name, 0, todayStr, todayDateStr, false);
  
  // 대표 시트 월총포인트, 누적총포인트 실시간 계산 및 갱신
  updateProfilePointsRealtime(sheet, p, name, todayStr, currentMonthStr);
  
  return createJsonResponse({
    success: true,
    message: "새로운 참여자 기록(0P 시작)이 추가되고 정상 등록되었습니다!",
    isNew: true
  });
}

// 2. 미션 및 보너스 포인트 기록 API
function handleLogMission(sheet, data) {
  var p = getParticipantDetails(data.studentId);
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
      if (pRows[i][1].toString().trim() === p.cleanId) {
        name = pRows[i][2];
        break;
      }
    }
  }
  
  // 일일 포인트 100P 상한 및 중복행 정리 후 동기화
  syncDailyRowForToday(dailySheet, p.cleanId, name, pointsDelta, todayStr, todayDateStr, false);
  
  // 키/몸무게 전송 시 신체수치 업데이트
  if (weight !== "" || height !== "") {
    var pRows2 = profileSheet.getDataRange().getValues();
    for (var k = 1; k < pRows2.length; k++) {
      var cellValue = pRows2[k][0];
      var rowMonthStr = formatDateToYYYYMM(cellValue);
      var rowId = pRows2[k][1].toString().trim();
      if (rowId === p.cleanId && rowMonthStr === currentMonthStr) {
        var rowIdx = k + 1;
        profileSheet.getRange(rowIdx, 1).setValue(todayStr);
        if (height !== "" && parseFloat(height) > 0) profileSheet.getRange(rowIdx, 4).setValue(parseFloat(height));
        if (weight !== "" && parseFloat(weight) > 0) profileSheet.getRange(rowIdx, 5).setValue(parseFloat(weight));
        var curHeight = parseFloat(profileSheet.getRange(rowIdx, 4).getValue()) || 0;
        var curWeight = parseFloat(profileSheet.getRange(rowIdx, 5).getValue()) || 0;
        var newBmi = (curHeight > 0) ? parseFloat((curWeight / ((curHeight / 100) * (curHeight / 100))).toFixed(1)) : 0;
        profileSheet.getRange(rowIdx, 6).setValue(newBmi);
        break;
      }
    }
  }

  // 대표 시트 실시간 업데이트
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
  var cumulativePoints = 0;
  var monthlyPoints = 0;
  
  for (var m = 1; m < updatedDRows.length; m++) {
    var dId = updatedDRows[m][1].toString().trim();
    if (dId === p.cleanId) {
      var pts = parseInt(updatedDRows[m][3]) || 0;
      cumulativePoints += pts;
      if (formatDateToYYYYMM(updatedDRows[m][0]) === currentMonthStr) {
        monthlyPoints += pts;
      }
    }
  }

  var level = "알콩이";
  if (cumulativePoints > 3000) level = "꼬꼬대장";
  else if (cumulativePoints > 2000) level = "튼튼이";
  else if (cumulativePoints > 1000) level = "삐약이";

  var pRows3 = profileSheet.getDataRange().getValues();
  var foundProfileRow = -1;
  for (var pIdx = 1; pIdx < pRows3.length; pIdx++) {
    var cellVal = pRows3[pIdx][0];
    var pMonthStr = formatDateToYYYYMM(cellVal);
    var pId = pRows3[pIdx][1].toString().trim();
    if (pId === p.cleanId && pMonthStr === currentMonthStr) {
      foundProfileRow = pIdx + 1;
      break;
    }
  }

  if (foundProfileRow > -1) {
    profileSheet.getRange(foundProfileRow, 1).setValue(todayStr);       // A: 일시
    if (name) profileSheet.getRange(foundProfileRow, 3).setValue(name); // C: 이름
    profileSheet.getRange(foundProfileRow, 7).setValue(monthlyPoints);   // G: 월총포인트
    profileSheet.getRange(foundProfileRow, 8).setValue(cumulativePoints);// H: 누적총포인트
    profileSheet.getRange(foundProfileRow, 9).setValue(level);           // I: 레벨
  } else {
    profileSheet.appendRow([todayStr, "'" + p.cleanId, name, 0, 0, 0, monthlyPoints, cumulativePoints, level]);
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
  
  // 가장 최근의 등록 기록 찾기 (최신 월 데이터)
  var latestRowIndex = -1;
  for (var i = pRows.length - 1; i >= 1; i--) {
    if (pRows[i][1].toString().trim() === p.cleanId) {
      latestRowIndex = i;
      break;
    }
  }
  
  if (latestRowIndex === -1) {
    return createJsonResponse({ success: false, isRegistered: false, message: "등록되지 않은 개인번호입니다." });
  }
  
  var name = pRows[latestRowIndex][2];
  var height = parseFloat(pRows[latestRowIndex][3]) || 0;
  var weight = parseFloat(pRows[latestRowIndex][4]) || 0;
  
  // 총포인트는 모든 일별 포인트 합산, 월총포인트는 당월 합산으로 제공
  var totalPoints = 0;
  var monthlyPoints = 0;
  var currentMonthStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM");
  var dailySheet = sheet.getSheetByName(p.dailySheet);
  var dRows = dailySheet.getDataRange().getValues();
  for (var j = 1; j < dRows.length; j++) {
    if (dRows[j][1].toString().trim() === p.cleanId) {
      var pts = parseInt(dRows[j][3]) || 0;
      totalPoints += pts;
      if (formatDateToYYYYMM(dRows[j][0]) === currentMonthStr) {
        monthlyPoints += pts;
      }
    }
  }
  
  // 사전 설문 조사 참여 여부 체크
  var surveySheet = sheet.getSheetByName(p.surveySheet);
  var sRows = surveySheet.getDataRange().getValues();
  var preSurveyDone = false;
  for (var k = 1; k < sRows.length; k++) {
    if (sRows[k][1].toString().trim() === p.cleanId) {
      preSurveyDone = true;
      break;
    }
  }
  
  // 닉네임 자동 빌드
  var nickname = p.isTeacher ? (name + " 선생님") : getStudentNickname(p.cleanId, name);
  
  var studentData = {
    studentId: studentId,
    name: name,
    nickname: nickname,
    height: height,
    weight: weight,
    totalPoints: totalPoints,
    monthlyPoints: monthlyPoints,
    preSurveyDone: preSurveyDone
  };
  
  // 미션 수행 로그(일별 포인트) 변환
  var history = [];
  for (var m = 1; m < dRows.length; m++) {
    if (dRows[m][1].toString().trim() === p.cleanId) {
      history.push({
        date: formatDateToYYYYMMDDHHMMSS(dRows[m][0]),
        mission: "일일 미션 적립 완료 🐾",
        weight: null,
        points: parseInt(dRows[m][3]) || 0
      });
    }
  }
  history.reverse();
  studentData.history = history;
  
  // 신체 기록 월별 트래킹 목록 추출 (각 월별 행에서 추출)
  var weightHistory = [];
  for (var n = 1; n < pRows.length; n++) {
    if (pRows[n][1].toString().trim() === p.cleanId && pRows[n][4] !== "") {
      weightHistory.push({
        date: formatDateToYYYYMM(pRows[n][0]), // YYYY-MM
        weight: parseFloat(pRows[n][4])
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
  var dailyStudentSheet = sheet.getSheetByName("학생일별포인트");
  var dailyTeacherSheet = sheet.getSheetByName("교사일별포인트");
  
  var leaderboard = [];
  var currentMonthStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM");
  
  // 일별 포인트 시트에서 실시간 월별 포인트를 집계합니다 (기록 시트는 30일에만 업데이트 되므로)
  var realMonthlyPoints = {};
  [dailyStudentSheet, dailyTeacherSheet].forEach(function(ds) {
    if (!ds) return;
    var dRows = ds.getDataRange().getValues();
    for (var i = 1; i < dRows.length; i++) {
      if (formatDateToYYYYMM(dRows[i][0]) === currentMonthStr) {
         var id = dRows[i][1].toString().trim();
         realMonthlyPoints[id] = (realMonthlyPoints[id] || 0) + (parseInt(dRows[i][3]) || 0);
      }
    }
  });

  // 학생용 리더보드 빌드
  if (studentSheet) {
    var sRows = studentSheet.getDataRange().getValues();
    var sPointsMap = {};
    var sNamesMap = {};
    for (var i = 1; i < sRows.length; i++) {
      var cellValue = sRows[i][0];
      var rowMonthStr = formatDateToYYYYMM(cellValue);
      var id = sRows[i][1].toString().trim();
      var name = sRows[i][2].toString().trim();
      
      // 현재 월에 해당하는 행만 필터링하여 순위 반영
      if (id && rowMonthStr === currentMonthStr) {
        sPointsMap[id] = realMonthlyPoints[id] || 0;
        sNamesMap[id] = name;
      }
    }
    Object.keys(sPointsMap).forEach(function(id) {
      var name = sNamesMap[id];
      var classGroup = "새싹반";
      var grade = id.charAt(0);
      var classNum = id.length === 5 ? parseInt(id.substring(1, 3)) : id.charAt(1);
      if (!isNaN(grade) && !isNaN(classNum)) {
        classGroup = grade + "학년 " + classNum + "반";
      }
      leaderboard.push({
        nickname: getStudentNickname(id, name),
        classGroup: classGroup,
        points: sPointsMap[id]
      });
    });
  }
  
  // 교사용 리더보드 빌드
  if (teacherSheet) {
    var tRows = teacherSheet.getDataRange().getValues();
    var tPointsMap = {};
    var tNamesMap = {};
    for (var j = 1; j < tRows.length; j++) {
      var cellValue = tRows[j][0];
      var rowMonthStr = formatDateToYYYYMM(cellValue);
      var id = tRows[j][1].toString().trim();
      var name = tRows[j][2].toString().trim();
      
      if (id && rowMonthStr === currentMonthStr) {
        tPointsMap[id] = realMonthlyPoints[id] || 0;
        tNamesMap[id] = name;
      }
    }
    Object.keys(tPointsMap).forEach(function(id) {
      leaderboard.push({
        nickname: tNamesMap[id] + " 선생님",
        classGroup: "교사",
        points: tPointsMap[id]
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
  var p = getParticipantDetails(data.studentId);
  var surveySheet = sheet.getSheetByName(p.surveySheet);
  
  var name = data.name.trim();
  var answers = data.answers;
  var todayStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  
  // 구글 시트 헤더(1행) 자동 확장 검증 (문항12까지 O열 헤더 자동 생성)
  if (surveySheet.getLastRow() === 0) {
    surveySheet.appendRow(["일시", "개인번호", "이름", "문항1", "문항2", "문항3", "문항4", "문항5", "문항6", "문항7", "문항8", "문항9", "문항10", "문항11", "문항12"]);
  } else {
    var lastCol = surveySheet.getLastColumn();
    if (lastCol < 15) {
      for (var col = lastCol + 1; col <= 15; col++) {
        surveySheet.getRange(1, col).setValue("문항" + (col - 3));
      }
    }
  }
  
  var rowData = [todayStr, "'" + p.cleanId, name];
  var totalQ = (answers && answers.length) ? answers.length : 12;
  for (var i = 0; i < totalQ; i++) {
    rowData.push(answers[i] !== undefined && answers[i] !== null ? answers[i] : "");
  }
  
  surveySheet.appendRow(rowData);
  SpreadsheetApp.flush();
  
  return createJsonResponse({
    success: true,
    message: "사전 설문조사(12문항)가 성공적으로 제출되었습니다! 🌟"
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
      
      var level = "알콩이";
      if (cumulativePoints > 3000) level = "꼬꼬대장";
      else if (cumulativePoints > 2000) level = "튼튼이";
      else if (cumulativePoints > 1000) level = "삐약이";
      
      // 실제 시트에 고정된 값(Value)으로 쓰기 
      // (G: 7=월총포인트, H: 8=누적총포인트, I: 9=레벨)
      profileSheet.getRange(i + 1, 7).setValue(monthlyPoints);
      profileSheet.getRange(i + 1, 8).setValue(cumulativePoints);
      profileSheet.getRange(i + 1, 9).setValue(level);
    }
  });
}
