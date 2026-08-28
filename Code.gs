// ========================================================
// 꼬꼬챌린지 - Option A: B초 포함 설문응답 탭 정확 매칭 완결판 (Code.gs)
// (학생 ID의 학교 정보를 최우선 판별하여 [B초_설문응답] 탭에 100% 저장)
// (최종 갱신 시각: 2026-08-28 13:53:00)
// ========================================================

function doPost(e) {
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 학생 ID와 학교 데이터에서 최우선으로 학교 추출
    let school = getSchoolFromId(data.studentId || data.school);

    const action = data.action;
    const timestamp = new Date();

    // 1. 일일 출석 및 미션 스티커 ➡️ [A초, B초, C초, D초] 매일 출석부 탭 (5열 1일 1행 고정!)
    if (action === 'log_mission') {
      const dailySheet = getOrCreateDailySheet(ss, school);
      upsertDailySticker5Col(dailySheet, school, data, timestamp);
      return responseJSON({ success: true, message: `${school} 매일출석부 연동 성공` });
    }
    
    // 2. 키, 몸무게 (내몸탐험) ➡️ [A초_월별성장, B초_월별성장 등] 탭
    else if (action === 'update_bmi') {
      const monthlySheet = getOrCreateMonthlySheet(ss, `${school}_월별성장`);
      const cleanId = cleanStudentId(data.studentId);
      const name = String(data.name || cleanId).trim();
      const pts = Number(data.totalPoints || 0);
      const totalStickers = Math.floor(pts / 100);
      const levelName = getLevelNameFromPoints(pts);
      
      monthlySheet.appendRow([
        timestamp,
        school,
        cleanId,
        name,
        `${data.month || (new Date().getMonth() + 1)}월`,
        data.height || 0,
        data.weight || 0,
        data.bmi || "",
        totalStickers,
        levelName
      ]);

      sortSheetStudentsFirst(monthlySheet, 10);
      return responseJSON({ success: true, message: `${school}_월별성장 신체기록 완료` });
    }
    
    // 3. 사전 / 사후 설문조사 ➡️ [A초_설문응답, B초_설문응답 등] 탭
    else if (action === 'save_survey' || action === 'submit_survey') {
      const surveySheet = getOrCreateSurveySheet(ss, `${school}_설문응답`);
      const cleanId = cleanStudentId(data.studentId);
      const name = String(data.name || cleanId).trim();
      
      surveySheet.appendRow([
        timestamp,
        school,
        cleanId,
        name,
        data.surveyType || "사전설문",
        JSON.stringify(data.answers || {})
      ]);

      sortSheetStudentsFirst(surveySheet, 6);
      return responseJSON({ success: true, message: `${school}_설문응답 저장 완료` });
    }

    return responseJSON({ success: true, message: "수신 완료" });
  } catch (err) {
    return responseJSON({ success: false, error: err.toString() });
  } finally {
    if (hasLock) lock.releaseLock();
  }
}

function getLevelNameFromPoints(points) {
  if (points >= 4500) return "4단계 꼬꼬대장";
  if (points >= 2400) return "3단계 튼튼이";
  if (points >= 800) return "2단계 삐약이";
  return "1단계 알콩이";
}

// 스티커 0개/1개 절대기준 적용 (totalPoints >= 100 일 때만 1개, 100점 미만은 무조건 0개!)
function upsertDailySticker5Col(sheet, schoolName, data, timestamp) {
  const rawId = String(data.studentId || "").trim();
  const cleanId = cleanStudentId(rawId);
  const name = String(data.name || cleanId).trim();
  
  const pts = Number(data.totalPoints || 0);
  
  // 총포인트가 100점 이상일 때만 일일 스티커 1개, 그 미만(0P~99P)은 무조건 0개!
  const dailyStickerVal = pts >= 100 ? 1 : 0;

  const today = new Date();
  const curY = today.getFullYear();
  const curM = today.getMonth() + 1;
  const curD = today.getDate();

  const lastRow = sheet.getLastRow();
  let foundRow = -1;

  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      const rowDateVal = values[i][0];
      const rowIdStr = String(values[i][2] || "").trim();
      const rowCleanId = cleanStudentId(rowIdStr);

      let isSameDate = false;
      if (rowDateVal instanceof Date) {
        isSameDate = (rowDateVal.getFullYear() === curY && 
                      (rowDateVal.getMonth() + 1) === curM && 
                      rowDateVal.getDate() === curD);
      } else {
        const str = String(rowDateVal);
        isSameDate = str.includes(`${curY}`) && str.includes(`${curD}`);
      }

      let isSameId = (rowCleanId === cleanId) || (rowIdStr === rawId) || (cleanId.length > 0 && rowCleanId.includes(cleanId));

      if (isSameDate && isSameId) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow > 1) {
    sheet.getRange(foundRow, 1).setValue(timestamp);
    sheet.getRange(foundRow, 3).setValue(cleanId);
    sheet.getRange(foundRow, 4).setValue(name);
    sheet.getRange(foundRow, 5).setValue(dailyStickerVal);
  } else {
    sheet.appendRow([
      timestamp,
      schoolName,
      cleanId,
      name,
      dailyStickerVal
    ]);
  }

  sortSheetStudentsFirst(sheet, 5);
}

function getOrCreateDailySheet(ss, schoolName) {
  let sheet = ss.getSheetByName(schoolName);
  if (!sheet) {
    sheet = ss.insertSheet(schoolName);
    sheet.appendRow(["일시", "학교", "개인번호", "이름", "일별스티커"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#D8F3DC");
  }
  return sheet;
}

function getOrCreateMonthlySheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["측정일시", "학교", "개인번호", "이름", "측정월", "키(cm)", "몸무게(kg)", "BMI", "총스티커(개)", "레벨"]);
    sheet.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#E2ECE9");
  } else {
    sheet.getRange(1, 9).setValue("총스티커(개)").setFontWeight("bold");
    sheet.getRange(1, 10).setValue("레벨").setFontWeight("bold");
  }
  return sheet;
}

function getOrCreateSurveySheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["응답일시", "학교", "개인번호", "이름", "설문구분", "응답내용"]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#FDE2E4");
  }
  return sheet;
}

function cleanStudentId(id) {
  let str = String(id || "").trim();
  str = str.replace(/^[A-D]초_/i, '').replace(/^[A-D]_/i, '');
  return str;
}

function sortSheetStudentsFirst(sheet, numCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;
  
  const range = sheet.getRange(2, 1, lastRow - 1, numCols);
  const values = range.getValues();

  values.sort((a, b) => {
    const idA = String(a[2] || "").trim();
    const idB = String(b[2] || "").trim();

    const isTeacherA = /^t/i.test(idA);
    const isTeacherB = /^t/i.test(idB);

    if (isTeacherA !== isTeacherB) {
      return isTeacherA ? 1 : -1;
    }
    return idA.localeCompare(idB);
  });

  range.setValues(values);
}

function getSchoolFromId(id) {
  if (!id) return "A초";
  const str = String(id).trim().toUpperCase();
  const first = str.charAt(0);
  if (["A", "B", "C", "D"].includes(first)) return first + "초";
  if (str.includes("A초") || str.includes("A_")) return "A초";
  if (str.includes("B초") || str.includes("B_")) return "B초";
  if (str.includes("C초") || str.includes("C_")) return "C초";
  if (str.includes("D초") || str.includes("D_")) return "D초";
  return "A초";
}

function doGet(e) {
  try {
    const rawId = String((e.parameter && e.parameter.studentId) || "").trim();
    const cleanId = cleanStudentId(rawId);
    return responseJSON({
      success: true,
      student: { id: rawId, name: cleanId || "학생", totalPoints: 0 }
    });
  } catch (err) {
    return responseJSON({ success: true, student: { id: "guest", name: "학생", totalPoints: 0 } });
  }
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
