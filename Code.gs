/**
 * TKD Theory Quiz - v8.5
 * Based on stable v8.2 logic with minor performance refinements
 */

// Column Mappings for Maintenance
const USER_COLS = { USER: 0, PASS: 1, GRADE: 2, LAST_ACTIVE: 3, STREAK: 4, NAME: 5 };
const QUEST_COLS = { TEXT: 0, ANS: 11, ID: 14, EXAM: 16, LEVEL: 17 };

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('TKD Theory v8.5')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function loginUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  const userData = userSheet.getRange(1, 1, userSheet.getLastRow(), 6).getValues();
  const now = new Date();
  const cleanU = username ? username.toString().trim().toLowerCase() : "";
  const cleanP = password ? password.toString().trim() : "";

  for (let i = 1; i < userData.length; i++) {
    if (userData[i][USER_COLS.USER] && userData[i][USER_COLS.USER].toString().trim().toLowerCase() === cleanU && userData[i][USER_COLS.PASS].toString() === cleanP) {
      let streak = parseInt(userData[i][USER_COLS.STREAK]) || 0;
      let lastActiveStr = userData[i][USER_COLS.LAST_ACTIVE] ? new Date(userData[i][USER_COLS.LAST_ACTIVE]).toDateString() : "";
      
      if (lastActiveStr !== now.toDateString()) {
        const yesterday = new Date();
        yesterday.setDate(now.getDate() - 1);
        streak = (lastActiveStr === yesterday.toDateString()) ? streak + 1 : 1;
        userSheet.getRange(i + 1, 4, 1, 2).setValues([[now, streak]]);
      }
      return { 
        success: true, 
        username: userData[i][USER_COLS.USER], 
        displayName: userData[i][USER_COLS.NAME] || userData[i][USER_COLS.USER], 
        gradeValue: parseInt(userData[i][USER_COLS.GRADE]) || 1, 
        streak: streak 
      };
    }
  }
  return { success: false };
}

function getQuizData(username, isTest) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("Questions");
  const userData = ss.getSheetByName("Users").getDataRange().getValues();
  
  let userGrade = 1;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][USER_COLS.USER] == username) {
      userGrade = parseInt(userData[i][USER_COLS.GRADE]) || 1;
      break;
    }
  }

  // Get data only up to the last populated row
  const qData = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, qSheet.getLastColumn()).getValues();
  
  // Filter by user grade and exclude non-exam questions (Q column)
  const filtered = qData.filter(row => (parseInt(row[QUEST_COLS.LEVEL]) || 1) <= userGrade && row[QUEST_COLS.EXAM] !== "N");
  
  const limit = isTest ? 50 : 10;
  
  // High-performance shuffle
  return filtered.sort(() => Math.random() - 0.5).slice(0, limit).map(row => ({
    question: row[QUEST_COLS.TEXT].toString(),
    answer: (row[QUEST_COLS.ANS] || "").toString().trim(),
    options: [row[4], row[5], row[6], row[7]].filter(String).sort(() => Math.random() - 0.5),
    qId: row[QUEST_COLS.ID].toString()
  }));
}

function updateQuestionScore(username, qId, isCorrect) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName("UserProgress");
  const pData = pSheet.getDataRange().getValues();
  const qIdStr = qId.toString();
  let foundRow = -1;

  for (let i = 0; i < pData.length; i++) {
    if (pData[i][0] == username && pData[i][1] == qIdStr) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow !== -1) {
    // Optimization: Pull current bucket from the local pData array instead of the sheet
    let bucket = parseInt(pData[foundRow - 1][3]) || 1;
    pSheet.getRange(foundRow, 3, 1, 3).setValues([[isCorrect ? 1 : 0, isCorrect ? Math.min(bucket + 1, 4) : 1, new Date()]]);
  } else {
    pSheet.appendRow([username, qIdStr, isCorrect ? 1 : 0, isCorrect ? 2 : 1, new Date()]);
  }
}

function saveGrade(u, g) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const d = s.getDataRange().getValues();
  for(let i=1; i<d.length; i++) { if(d[i][0]==u) { s.getRange(i+1,3).setValue(g); break; } }
}

function updatePass(u, p) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const d = s.getDataRange().getValues();
  for(let i=1; i<d.length; i++) { if(d[i][0]==u) { s.getRange(i+1,2).setValue(p); break; } }
}