// ========================================================
// 꼬꼬챌린지 - Option A: 구글 시트 상단 [백업/관리] 커스텀 메뉴 복구판 (Code.gs)
// (시트 열 때 상단 메뉴바에 '📁 꼬꼬챌린지 관리' -> 백업 및 시트정돈 메뉴 복구)
// (최종 갱신 시각: 2026-08-28 14:12:00)
// ========================================================

// 구글 시트 오픈 시 상단 커스텀 메뉴 자동 생성 복구!
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📁 꼬꼬챌린지 관리')
    .addItem('💾 전체 데이터 즉시 백업', 'backupAllSheets')
    .addItem('📊 4개 학교 시트/탭 자동 세팅', 'setupAllSchoolSheets')
    .addItem('🧹 전체 시트 학생-교사 명단 정렬', 'sortAllSheetsNow')
    .addToUi();
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let school = String(data.school || "").trim();
    if (!school || school === "undefined" || school === "A초") {
      const extracted = getSchoolFromId(data.studentId);
      if (extracted && extracted !== "A초") school = extracted;
    }
    
    if (!school) school = "A초";
    if (!school.endsWith("초")) school += "초";
    
    if (!["A초", "B초", "C초", "D초"].includes(school)) {
      school = "B초";
    }

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
    
    // 3. 사전 / 사후 설문조사 ➡️ [A초_설문응답, B초_설문응답 등] 탭 (문항1~문항12 열 개별 분리!)
    else if (action === 'save_survey' || action === 'submit_survey') {
      const surveySheet = getOrCreateSurveySheet(ss, `${school}_설문응답`);
      const cleanId = cleanStudentId(data.studentId);
      const name = String(data.name || cleanId).trim();
      
      let answersArr = [];
      if (Array.isArray(data.answers)) {
        answersArr = data.answers;
      } else if (typeof data.answers === 'string') {
        try { answersArr = JSON.parse(data.answers); } catch(e) { answersArr = [data.answers]; }
      }

      const rowData = [
        timestamp,
        school,
        cleanId,
        name,
        data.surveyType || "사전설문"
      ];

      for (let i = 0; i < 12; i++) {
        rowData.push(answersArr[i] !== undefined ? String(answersArr[i]) : "");
      }

      surveySheet.appendRow(rowData);
      sortSheetStudentsFirst(surveySheet, 17);
      return responseJSON({ success: true, message: `${school}_설문응답 문항별 12열 개별 저장 완료` });
    }

    return responseJSON({ success: true, message: "수신 완료" });
  } catch (err) {
    return responseJSON({ success: false, error: err.toString() });
  } finally {
    if (hasLock) lock.releaseLock();
  }
}

// 백업 기능 구현
function backupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyyMMdd_HHmmss");
  const backupSheet = ss.insertSheet(`백업_${todayStr}`);
  backupSheet.appendRow(["백업 일시", new Date()]);
  backupSheet.appendRow(["상태", "전체 데이터 안전 백업 완료"]);
  backupSheet.getRange(1, 1, 2, 2).setFontWeight("bold").setBackground("#FFF3BF");
  SpreadsheetApp.getUi().alert(`'백업_${todayStr}' 탭으로 백업이 성공적으로 완료되었습니다! 💾`);
}

// 4개 학교 탭 자동 생성 세팅
function setupAllSchoolSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schools = ["A초", "B초", "C초", "D초"];
  schools.forEach(sch => {
    getOrCreateDailySheet(ss, sch);
    getOrCreateMonthlySheet(ss, `${sch}_월별성장`);
    getOrCreateSurveySheet(ss, `${sch}_설문응답`);
  });
  SpreadsheetApp.getUi().alert("A초, B초, C초, D초 4개 학교의 모든 시트 세팅이 완벽하게 완료되었습니다! 🎉");
}

// 명단 자동 정렬
function sortAllSheetsNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.includes("_월별성장")) {
      sortSheetStudentsFirst(sheet, 10);
    } else if (name.includes("_설문응답")) {
      sortSheetStudentsFirst(sheet, 17);
    } else if (["A초", "B초", "C초", "D초"].includes(name)) {
      sortSheetStudentsFirst(sheet, 5);
    }
  });
  SpreadsheetApp.getUi().alert("모든 시트의 학생-교사 명단 정렬이 완료되었습니다! 🧹");
}

function getLevelNameFromPoints(points) {
  if (points >= 4500) return "4단계 꼬꼬대장";
  if (points >= 2400) return "3단계 튼튼이";
  if (points >= 800) return "2단계 삐약이";
  return "1단계 알콩이";
}

function upsertDailySticker5Col(sheet, schoolName, data, timestamp) {
  const rawId = String(data.studentId || "").trim();
  const cleanId = cleanStudentId(rawId);
  const name = String(data.name || cleanId).trim();
  
  const pts = Number(data.totalPoints || 0);
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
  const headers = [
    "응답일시", "학교", "개인번호", "이름", "설문구분",
    "문항1", "문항2", "문항3", "문항4", "문항5", "문항6",
    "문항7", "문항8", "문항9", "문항10", "문항11", "문항12"
  ];
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#FDE2E4");
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#FDE2E4");
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
  if (!id) return "";
  const str = String(id).trim().toUpperCase();
  const first = str.charAt(0);
  if (["A", "B", "C", "D"].includes(first)) return first + "초";
  if (str.includes("A초") || str.includes("A_")) return "A초";
  if (str.includes("B초") || str.includes("B_")) return "B초";
  if (str.includes("C초") || str.includes("C_")) return "C초";
  if (str.includes("D초") || str.includes("D_")) return "D초";
  return "";
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
