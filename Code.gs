// ========================================================
// 꼬꼬챌린지 - 사람당 하루 1회 전용 (동시성 락 & 1일 1인 1행 UPSERT 시스템)
// (동시 6개 요청 완벽 방어 -> 사람당 하루에 EXACTLY 1개 행만 기록/갱신)
// (최종 갱신 시각: 2026-08-28 11:38:25)
// ========================================================

function doPost(e) {
  const lock = LockService.getScriptLock();
  // 동시 6개 미션 클릭 시 6줄 생기는 현상을 100% 방지하기 위해 락 취득 (10초 대기)
  const hasLock = lock.tryLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 학교 이름 (A초, B초, C초, D초)
    const school = data.school || getSchoolFromId(data.studentId) || "A초";
    const action = data.action;
    const isTeacher = data.isTeacher === true || data.userType === "teacher";
    
    const timestamp = new Date();

    // 1. 일별 스티커 (사람당 하루에 EXACTLY 1개 행만 기록/업데이트)
    if (action === 'log_mission') {
      const menuName = isTeacher ? "교사일별스티커" : "학생일별스티커";
      
      // 학교별 탭 (예: A초, B초, C초, D초) 1일 1회 UPSERT
      upsertDailyStickerRow(ss, school, data, timestamp);

      // 세부 메뉴 탭 (예: 학생일별스티커, 교사일별스티커) 1일 1회 UPSERT
      upsertDailyStickerRow(ss, menuName, data, timestamp);

      return responseJSON({ success: true, message: "사람당 하루 1회 일별 스티커 총 수량 갱신 완료" });
    }
    
    // 2. 키 & 몸무게 (내몸탐험) 기록
    else if (action === 'update_bmi') {
      const menuName = isTeacher ? "교사기록" : "학생기록";
      const schoolSheet = getOrCreateSchoolSheet(ss, school);
      
      schoolSheet.appendRow([
        timestamp,
        school,
        data.studentId || "",
        data.name || "",
        `키: ${data.height || 0}cm / 몸무게: ${data.weight || 0}kg`,
        `BMI: ${data.bmi || ""}`,
        `${data.month || 8}월`
      ]);

      const menuSheet = getOrCreateDetailedSheet(ss, menuName);
      menuSheet.appendRow([
        timestamp,
        school,
        data.studentId || "",
        data.name || "",
        `키: ${data.height || 0}cm / 몸무게: ${data.weight || 0}kg`,
        `BMI: ${data.bmi || ""}`,
        `${data.month || 8}월`
      ]);

      return responseJSON({ success: true, message: "BMI 기록 성공" });
    }
    
    // 3. 설문지 답변 기록
    else if (action === 'save_survey') {
      const menuName = isTeacher ? "교사설문응답" : "학생설문응답";
      const schoolSheet = getOrCreateSchoolSheet(ss, school);
      
      schoolSheet.appendRow([
        timestamp,
        school,
        data.studentId || "",
        data.name || "",
        data.surveyType || "사전설문",
        JSON.stringify(data.answers || {}),
        ""
      ]);

      const menuSheet = getOrCreateDetailedSheet(ss, menuName);
      menuSheet.appendRow([
        timestamp,
        school,
        data.studentId || "",
        data.name || "",
        data.surveyType || "사전설문",
        JSON.stringify(data.answers || {}),
        ""
      ]);

      return responseJSON({ success: true, message: "설문 저장 성공" });
    }

    return responseJSON({ success: true, message: "요청 수신 완료" });
  } catch (err) {
    return responseJSON({ success: false, error: err.toString() });
  } finally {
    if (hasLock) lock.releaseLock();
  }
}

// 1일 1인 1행 전용 UPSERT 함수 (오늘 날짜 + 개인번호가 같으면 기존 행 수정, 없으면 새 행 추가)
function upsertDailyStickerRow(ss, sheetName, data, timestamp) {
  const sheet = getOrCreateSchoolSheet(ss, sheetName);
  const studentId = String(data.studentId || "").trim();
  const name = String(data.name || "").trim();
  const school = String(data.school || "A초").trim();
  const dailyStickerVal = Number(data.dailySticker) || 1; // 오늘 하루 총 스티커 개수

  const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const lastRow = sheet.getLastRow();
  let foundRow = -1;

  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      const rowDate = values[i][0];
      const rowId = String(values[i][2]).trim();
      
      let rowDateStr = "";
      if (rowDate instanceof Date) {
        rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        rowDateStr = String(rowDate).substring(0, 10);
      }

      // 오늘 날짜 + 학생 학번(개인번호)이 일치하는 행 탐색
      if (rowDateStr === dateStr && rowId === studentId) {
        foundRow = i + 2; // 1-based index
        break;
      }
    }
  }

  if (foundRow > 1) {
    // 이미 오늘 기록된 행이 있으면 -> 일별 스티커 개수 및 일시 갱신 (1일 1행 유지!)
    sheet.getRange(foundRow, 1).setValue(timestamp);
    sheet.getRange(foundRow, 5).setValue(dailyStickerVal);
  } else {
    // 오늘 첫 기록이면 -> 새 행 추가
    sheet.appendRow([
      timestamp,
      school,
      studentId,
      name,
      dailyStickerVal
    ]);
  }
}

// 학교/시트 탭 자동 생성 도우미
function getOrCreateSchoolSheet(ss, schoolName) {
  let name = (schoolName || "A초").trim();
  if (!name.endsWith("초") && !name.includes("스티커") && !name.includes("기록") && !name.includes("설문")) {
    name += "초";
  }
  
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(["일시", "학교", "개인번호", "이름", "일별스티커"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#D8F3DC");
  }
  return sheet;
}

function getOrCreateDetailedSheet(ss, sheetName) {
  return getOrCreateSchoolSheet(ss, sheetName);
}

function getSchoolFromId(id) {
  if (!id) return "A초";
  const first = id.charAt(0).toUpperCase();
  if (["A", "B", "C", "D"].includes(first)) return first + "초";
  return "A초";
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
