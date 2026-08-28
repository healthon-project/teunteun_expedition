// ========================================================
// 꼬꼬챌린지 - 로그인 ID 100% 수용 및 5열 순정 완결 백엔드 (Code.gs)
// (doGet rawId 반환으로 학생/교사 로그인 초고속 완성)
// (A열:일시, B열:학교, C열:개인번호, D열:이름, E열:일별스티커)
// (최종 갱신 시각: 2026-08-28 12:41:30)
// ========================================================

function doGet(e) {
  try {
    const action = e.parameter ? e.parameter.action : "";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'get_student') {
      const rawId = String(e.parameter.studentId || "").trim();
      const cleanId = cleanStudentId(rawId);
      
      const sheets = ss.getSheets();
      let totalPts = 0;
      let foundName = cleanId;

      for (let s = 0; s < sheets.length; s++) {
        const sheet = sheets[s];
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          const numCols = Math.min(sheet.getLastColumn(), 11);
          if (numCols >= 3) {
            const values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
            for (let i = 0; i < values.length; i++) {
              const rowIdStr = String(values[i][2] || values[i][4] || "").trim();
              const rowCleanId = cleanStudentId(rowIdStr);
              if (rowCleanId === cleanId || rowIdStr === rawId) {
                if (values[i][3]) foundName = String(values[i][3]).trim();
                const ptsVal = Number(values[i][4] || 0);
                if (!isNaN(ptsVal) && ptsVal > 0) {
                  totalPts = Math.max(totalPts, ptsVal * 100);
                }
              }
            }
          }
        }
      }

      return responseJSON({
        success: true,
        student: {
          id: rawId,
          name: foundName || cleanId,
          totalPoints: totalPts || 0
        }
      });
    }

    return responseJSON({ success: true, message: "Server Ready" });
  } catch (err) {
    return responseJSON({
      success: true,
      student: {
        id: String((e.parameter && e.parameter.studentId) || "guest"),
        name: "학생",
        totalPoints: 0
      }
    });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let school = String(data.school || "").trim();
    if (!school) school = getSchoolFromId(data.studentId);
    if (!school.endsWith("초")) school += "초";
    
    if (!["A초", "B초", "C초", "D초"].includes(school)) {
      const first = String(data.studentId || "").trim().charAt(0).toUpperCase();
      school = ["A", "B", "C", "D"].includes(first) ? first + "초" : "A초";
    }

    const action = data.action;
    const timestamp = new Date();

    const sheet = getOrCreate5ColSheet(ss, school);

    if (action === 'log_mission') {
      upsert5ColDailySticker(sheet, school, data, timestamp);
      return responseJSON({ success: true, message: `${school} 5열 일일출석 연동 성공` });
    }
    
    else if (action === 'update_bmi') {
      const cleanId = cleanStudentId(data.studentId);
      sheet.appendRow([
        timestamp,
        school,
        cleanId,
        data.name || cleanId,
        `키:${data.height||0}cm,몸무게:${data.weight||0}kg(BMI:${data.bmi||""})`
      ]);
      sortSheet5Col(sheet);
      return responseJSON({ success: true, message: `${school} 신체기록 완료` });
    }
    
    else if (action === 'save_survey') {
      const cleanId = cleanStudentId(data.studentId);
      sheet.appendRow([
        timestamp,
        school,
        cleanId,
        data.name || cleanId,
        `[${data.surveyType || "설문"}] ${JSON.stringify(data.answers || {})}`
      ]);
      sortSheet5Col(sheet);
      return responseJSON({ success: true, message: `${school} 설문저장 완료` });
    }

    return responseJSON({ success: true, message: "수신 완료" });
  } catch (err) {
    return responseJSON({ success: false, error: err.toString() });
  } finally {
    if (hasLock) lock.releaseLock();
  }
}

function upsert5ColDailySticker(sheet, schoolName, data, timestamp) {
  const rawId = String(data.studentId || "").trim();
  const cleanId = cleanStudentId(rawId);
  const name = String(data.name || cleanId).trim();
  const dailyStickerVal = Number(data.dailySticker) || 1;

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
        isSameDate = str.includes(`${curY}`) && str.includes(`${curM}`) && str.includes(`${curD}`);
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

  sortSheet5Col(sheet);
}

function cleanStudentId(id) {
  let str = String(id || "").trim();
  str = str.replace(/^[A-D]초_/i, '').replace(/^[A-D]_/i, '');
  return str;
}

function getOrCreate5ColSheet(ss, schoolName) {
  let sheet = ss.getSheetByName(schoolName);
  if (!sheet) {
    sheet = ss.insertSheet(schoolName);
    sheet.appendRow(["일시", "학교", "개인번호", "이름", "일별스티커"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#D8F3DC");
  }
  return sheet;
}

function sortSheet5Col(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;
  
  const range = sheet.getRange(2, 1, lastRow - 1, 5);
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
  const first = String(id).trim().charAt(0).toUpperCase();
  if (["A", "B", "C", "D"].includes(first)) return first + "초";
  return "A초";
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
