/**
 * 튼튼탐험대 (TeunTeun Expedition) - 구글 앱스 스크립트 (Google Apps Script)
 * 
 * [설치 방법]
 * 1. 구글 스프레드시트를 하나 새로 만듭니다.
 * 2. 상단 메뉴에서 [확장 프로그램] > [Apps Script]를 클릭합니다.
 * 3. 기존 코드를 모두 지우고 이 파일의 전체 코드를 복사하여 붙여넣습니다.
 * 4. 상단 저장(디스크 아이콘) 버튼을 누릅니다.
 * 5. 우측 상단의 [배포] > [새 배포]를 클릭합니다.
 * 6. 유형 선택(톱니바퀴)에서 [웹 앱]을 선택합니다.
 * 7. 아래와 같이 설정한 후 [배포] 버튼을 누릅니다.
 *    - 설명: 튼튼탐험대 API v1
 *    - 웹 앱을 실행할 사용자: 나 (보건교사 구글 계정)
 *    - 액세스할 수 있는 사용자: 모든 사용자 (가장 중요! 학생들이 로그인 없이 접근할 수 있어야 합니다)
 * 8. 처음 배포 시 [액세스 승인]이 뜨면 승인하고, 구글 경고 창이 뜨면 [고급] > [제목 없는 프로젝트(이동)]를 클릭하여 권한을 허용합니다.
 * 9. 생성된 [웹 앱 URL]을 복사하여 웹앱 소스코드(index.html)의 `GAS_WEB_APP_URL` 부분에 붙여넣습니다.
 */

function doGet(e) {
  var action = e.parameter.action;
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 시트가 없으면 생성
  setupSheets(sheet);
  
  if (action === 'get_leaderboard') {
    return handleGetLeaderboard(sheet);
  } else if (action === 'get_student') {
    var studentId = e.parameter.studentId;
    return handleGetStudent(sheet, studentId);
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: false, 
    message: "잘못된 action 요청입니다. (get_leaderboard 또는 get_student 필요)"
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets(sheet);
  
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
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: false, 
    message: "잘못된 action 요청입니다."
  })).setMimeType(ContentService.MimeType.JSON);
}

// 시트 초기화 및 헤더 생성
function setupSheets(sheet) {
  var studentSheet = sheet.getSheetByName("Students");
  if (!studentSheet) {
    studentSheet = sheet.insertSheet("Students");
    studentSheet.appendRow(["학번", "이름", "닉네임", "가입일", "키(cm)", "몸무게(kg)", "총포인트"]);
    studentSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#E2F0D9");
  }
  
  var missionSheet = sheet.getSheetByName("MissionLogs");
  if (!missionSheet) {
    missionSheet = sheet.insertSheet("MissionLogs");
    missionSheet.appendRow(["일시", "학번", "미션내용", "기록체중(kg)", "획득포인트"]);
    missionSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#FFF2CC");
  }
}

// 1. 학생 등록/수정 API
function handleRegister(sheet, data) {
  var studentSheet = sheet.getSheetByName("Students");
  var rows = studentSheet.getDataRange().getValues();
  var studentId = data.studentId.trim();
  var name = data.name.trim();
  var nickname = data.nickname.trim();
  var height = parseFloat(data.height);
  var weight = parseFloat(data.weight);
  
  var foundRowIndex = -1;
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().trim() === studentId) {
      foundRowIndex = i + 1; // 1-based index
      break;
    }
  }
  
  var todayStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  
  if (foundRowIndex > -1) {
    // 기존 정보 수정 (재등록/수정)
    studentSheet.getRange(foundRowIndex, 2).setValue(name);
    studentSheet.getRange(foundRowIndex, 3).setValue(nickname);
    studentSheet.getRange(foundRowIndex, 5).setValue(height);
    studentSheet.getRange(foundRowIndex, 6).setValue(weight);
    // 포인트는 수식이 채워져 있으므로 건드리지 않음
    
    // 초기 몸무게 기록용 미션 자동 추가 (기존 로그가 없는 경우만)
    logInitialWeightIfNeeded(sheet, studentId, weight, todayStr);
    
    return createJsonResponse({
      success: true,
      message: "정보가 성공적으로 수정되었습니다!",
      isNew: false
    });
  } else {
    // 신규 등록
    var newRow = studentSheet.getLastRow() + 1;
    // G열(7번째 열)에 SUMIF 수식을 적용하여 MissionLogs 시트에서 포인트 자동 계산
    var formula = "=SUMIF(MissionLogs!B:B, A" + newRow + ", MissionLogs!E:E)";
    
    studentSheet.appendRow([studentId, name, nickname, todayStr, height, weight, formula]);
    
    // 최초 가입 보너스 점수(100p) 및 몸무게 기록 로그 생성
    var missionSheet = sheet.getSheetByName("MissionLogs");
    missionSheet.appendRow([todayStr, studentId, "신청서 작성 보너스 🌱", weight, 100]);
    
    return createJsonResponse({
      success: true,
      message: "튼튼탐험대에 가입된 것을 축하합니다! (보너스 100p 지급)",
      isNew: true
    });
  }
}

// 최초 가입 몸무게 로그 생성 지원
function logInitialWeightIfNeeded(sheet, studentId, weight, dateStr) {
  var missionSheet = sheet.getSheetByName("MissionLogs");
  var logs = missionSheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < logs.length; i++) {
    if (logs[i][1].toString().trim() === studentId) {
      found = true;
      break;
    }
  }
  if (!found) {
    missionSheet.appendRow([dateStr, studentId, "신청서 작성 보너스 🌱", weight, 100]);
  }
}

// 2. 미션 인증 저장 API
function handleLogMission(sheet, data) {
  var studentId = data.studentId.trim();
  var missionType = data.missionType;
  var points = parseInt(data.points) || 0;
  var weight = data.weight ? parseFloat(data.weight) : "";
  var todayStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  
  // 학생 존재 여부 체크
  var studentSheet = sheet.getSheetByName("Students");
  var students = studentSheet.getDataRange().getValues();
  var foundStudentRow = -1;
  
  for (var i = 1; i < students.length; i++) {
    if (students[i][0].toString().trim() === studentId) {
      foundStudentRow = i + 1;
      break;
    }
  }
  
  if (foundStudentRow === -1) {
    return createJsonResponse({ success: false, message: "등록되지 않은 학번입니다. 먼저 신청서를 작성해 주세요." });
  }
  
  // 미션 로그 추가
  var missionSheet = sheet.getSheetByName("MissionLogs");
  missionSheet.appendRow([todayStr, studentId, missionType, weight, points]);
  
  // 몸무게 기록이 있으면 학생 정보의 최신 몸무게도 업데이트
  if (weight !== "") {
    studentSheet.getRange(foundStudentRow, 6).setValue(weight);
  }
  
  return createJsonResponse({
    success: true,
    message: "'" + missionType + "' 미션 성공! " + points + "p가 적립되었습니다.",
    pointsEarned: points
  });
}

// 3. 학생 데이터 및 최근 기록 조회 API
function handleGetStudent(sheet, studentId) {
  if (!studentId) {
    return createJsonResponse({ success: false, message: "학번 파라미터가 누락되었습니다." });
  }
  
  var studentSheet = sheet.getSheetByName("Students");
  var students = studentSheet.getDataRange().getValues();
  var studentRow = -1;
  var studentData = null;
  
  studentId = studentId.trim();
  for (var i = 1; i < students.length; i++) {
    if (students[i][0].toString().trim() === studentId) {
      studentRow = i;
      studentData = {
        studentId: students[i][0].toString(),
        name: students[i][1],
        nickname: students[i][2],
        joinDate: students[i][3],
        height: parseFloat(students[i][4]),
        weight: parseFloat(students[i][5]),
        totalPoints: parseInt(students[i][6]) || 0
      };
      break;
    }
  }
  
  if (!studentData) {
    return createJsonResponse({ success: false, isRegistered: false, message: "등록되지 않은 학번입니다." });
  }
  
  // 학생의 미션 수행 히스토리 및 몸무게 트렌드 추출
  var missionSheet = sheet.getSheetByName("MissionLogs");
  var logs = missionSheet.getDataRange().getValues();
  var history = [];
  var weightHistory = [];
  
  for (var j = 1; j < logs.length; j++) {
    if (logs[j][1].toString().trim() === studentId) {
      var logDate = logs[j][0];
      var formattedDate = (logDate instanceof Date) ? Utilities.formatDate(logDate, "Asia/Seoul", "yyyy-MM-dd HH:mm") : logDate.toString();
      
      history.push({
        date: formattedDate,
        mission: logs[j][2],
        weight: logs[j][3] ? parseFloat(logs[j][3]) : null,
        points: parseInt(logs[j][4]) || 0
      });
      
      if (logs[j][3] !== "") {
        weightHistory.push({
          date: formattedDate.split(" ")[0], // 날짜만 추출 (YYYY-MM-DD)
          weight: parseFloat(logs[j][3])
        });
      }
    }
  }
  
  // 히스토리는 최신순으로 정렬
  history.reverse();
  
  studentData.history = history.slice(0, 30); // 최근 30개 기록만 전송
  studentData.weightHistory = weightHistory; // 성장 그래프용 전체 체중 기록
  studentData.isRegistered = true;
  
  return createJsonResponse({ success: true, student: studentData });
}

// 4. 익명 랭킹 조회 API (닉네임으로 조회하여 아동 개인정보 보호)
function handleGetLeaderboard(sheet) {
  var studentSheet = sheet.getSheetByName("Students");
  var students = studentSheet.getDataRange().getValues();
  var leaderboard = [];
  
  for (var i = 1; i < students.length; i++) {
    var nickname = students[i][2].toString().trim();
    var totalPoints = parseInt(students[i][6]) || 0;
    var studentId = students[i][0].toString().trim();
    
    // 학급 구분을 위해 학년-반까지 추출 (예: '5-2-13' -> '5학년 2반')
    var classGroup = "새싹반";
    var match = studentId.match(/^(\d+)-(\d+)-/);
    if (match) {
      classGroup = match[1] + "학년 " + match[2] + "반";
    }
    
    if (nickname) {
      leaderboard.push({
        nickname: nickname,
        classGroup: classGroup,
        points: totalPoints
      });
    }
  }
  
  // 포인트 내림차순 정렬
  leaderboard.sort(function(a, b) {
    return b.points - a.points;
  });
  
  // 상위 50명만 반환
  var topList = leaderboard.slice(0, 50);
  
  return createJsonResponse({
    success: true,
    leaderboard: topList
  });
}

// 헬퍼: JSON 응답 생성
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}
