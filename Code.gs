/**
 * TKD Theory Quiz - v8.4 (Optimized)
 */

// Column Mappings for easier maintenance
const USER_COLS = { USER: 0, PASS: 1, GRADE: 2, LAST_ACTIVE: 3, STREAK: 4, NAME: 5 };
const QUEST_COLS = { TEXT: 0, OPT1: 4, OPT2: 5, OPT3: 6, OPT4: 7, ANS: 11, ID: 14, EXAM: 16, LEVEL: 17 };
const BUCKET_INTERVALS = { 1: 0, 2: 2, 3: 4, 4: 5 };

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('TKD Theory Practice')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function loginUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 6).getValues();
  const now = new Date();
  
  const cleanU = username ? username.toString().trim().toLowerCase() : "";
  const cleanP = password ? password.toString().trim() : "";

  for (let i = 1; i < data.length; i++) {
    if (data[i][USER_COLS.USER].toString().trim().toLowerCase() === cleanU && 
        data[i][USER_COLS.PASS].toString() === cleanP) {
      
      let streak = parseInt(data[i][USER_COLS.STREAK]) || 0;
      let lastActiveDate = data[i][USER_COLS.LAST_ACTIVE] ? new Date(data[i][USER_COLS.LAST_ACTIVE]) : null;
      let todayStr = now.toDateString();
      
      if (!lastActiveDate || lastActiveDate.toDateString() !== todayStr) {
        const yesterday = new Date();
        yesterday.setDate(now.getDate() - 1);
        streak = (lastActiveDate && lastActiveDate.toDateString() === yesterday.toDateString()) ? streak + 1 : 1;
        sheet.getRange(i + 1, 4, 1, 2).setValues([[now, streak]]);
      }

      return { 
        success: true, 
        username: data[i][USER_COLS.USER].toString().trim(),
        displayName: data[i][USER_COLS.NAME] || data[i][USER_COLS.USER],
        gradeValue: parseInt(data[i][USER_COLS.GRADE]) || 1, 
        streak: streak 
      };
    }
  }
  return { success: false };
}

function getQuizData(username, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qData = ss.getSheetByName("Questions").getDataRange().getValues();
  const pData = ss.getSheetByName("UserProgress").getDataRange().getValues();
  const userData = ss.getSheetByName("Users").getDataRange().getValues();
  
  const cleanUser = username.trim().toLowerCase();
  let userGrade = 1;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0].toString().toLowerCase() === cleanUser) {
      userGrade = parseInt(userData[i][2]) || 1;
      break;
    }
  }

  // Optimize: Create a Map for user progress to avoid nested loops
  const progressMap = pData.reduce((acc, row) => {
    if (row[0].toString().toLowerCase() === cleanUser) {
      acc[row[1].toString()] = { bucket: parseInt(row[3]) || 1, date: new Date(row[4]) };
    }
    return acc;
  }, {});

  const now = new Date();

  let filtered = qData.slice(1).filter(row => {
    const qId = row[QUEST_COLS.ID].toString();
    const beltLevel = parseInt(row[QUEST_COLS.LEVEL]) || 1;
    const isExam = row[QUEST_COLS.EXAM] !== "N";

    if (!row[QUEST_COLS.TEXT] || beltLevel > userGrade || !isExam) return false;
    if (mode === 'test') return true;

    const prog = progressMap[qId];
    if (!prog) return true;
    const diffDays = (now - prog.date) / (1000 * 60 * 60 * 24);
    return diffDays >= (BUCKET_INTERVALS[prog.bucket] || 0);
  });

  // Fallback if Spaced Repetition returns nothing
  if (filtered.length === 0) {
    filtered = qData.slice(1).filter(row => (parseInt(row[QUEST_COLS.LEVEL]) || 1) <= userGrade && row[QUEST_COLS.EXAM] !== "N");
  }

  const limit = (mode === 'test') ? 50 : 10;
  Utilities.shuffle(filtered); // Higher quality randomness than Math.random()

  return filtered.slice(0, limit).map(row => ({
    question: row[QUEST_COLS.TEXT].toString(),
    options: [row[4], row[5], row[6], row[7]].filter(String).sort(() => Math.random() - 0.5),
    answer: row[QUEST_COLS.ANS].toString().trim(),
    qId: row[QUEST_COLS.ID].toString()
  }));
}

function updateQuestionScore(username, qId, isCorrect) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName("UserProgress");
  const pData = pSheet.getDataRange().getValues();
  const qIdStr = qId.toString();
  const cleanU = username.trim().toLowerCase();
  
  // Find row using optimized loop
  let rowIndex = -1;
  for (let i = 0; i < pData.length; i++) {
    if (pData[i][0].toString().toLowerCase() === cleanU && pData[i][1].toString() === qIdStr) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex !== -1) {
    const currentBucket = parseInt(pData[rowIndex-1][3]) || 1;
    const currentScore = parseInt(pData[rowIndex-1][2]) || 0;
    const nextBucket = isCorrect ? Math.min(currentBucket + 1, 4) : 1;
    const nextScore = isCorrect ? currentScore + 1 : currentScore - 1;
    pSheet.getRange(rowIndex, 3, 1, 3).setValues([[nextScore, nextBucket, new Date()]]);
  } else {
    pSheet.appendRow([username, qIdStr, isCorrect ? 1 : -1, isCorrect ? 2 : 1, new Date()]);
  }
}