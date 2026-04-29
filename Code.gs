/**
 * TKD Theory Quiz - v8.2 (Corrected)
 * Users Tab Map: A:User | B:Pass | C:Grade | D:LastActive | E:Streak | F:Name(5)
 * Questions Tab Map: A:Quest(0) | ... | O:qId(14) | Q:Exam(16) | R:BeltLevel(17)
 */

const BUCKET_INTERVALS = { 1: 0, 2: 2, 3: 4, 4: 5 };

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('TKD Master')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function loginUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  const userData = userSheet.getDataRange().getValues();
  const now = new Date();
  
  const cleanU = username ? username.toString().trim().toLowerCase() : "";
  const cleanP = password ? password.toString().trim() : "";

  for (let i = 1; i < userData.length; i++) {
    if (!userData[i][0]) continue;
    
    if (userData[i][0].toString().trim().toLowerCase() === cleanU && 
        userData[i][1].toString() === cleanP) {
      
      let streak = parseInt(userData[i][4]) || 0;
      let lastActiveDate = userData[i][3] ? new Date(userData[i][3]) : null;
      let lastActiveStr = lastActiveDate ? lastActiveDate.toDateString() : "";
      let todayStr = now.toDateString();
      
      if (lastActiveStr !== todayStr) {
        const yesterday = new Date();
        yesterday.setDate(now.getDate() - 1);
        streak = (lastActiveStr === yesterday.toDateString()) ? streak + 1 : 1;
        userSheet.getRange(i + 1, 4, 1, 2).setValues([[now, streak]]);
      }

      return { 
        success: true, 
        username: userData[i][0].toString().trim(),
        displayName: userData[i][5] ? userData[i][5].toString() : userData[i][0].toString(),
        gradeValue: parseInt(userData[i][2]) || 1, 
        streak: streak 
      };
    }
  }
  return { success: false, message: "Invalid credentials" };
}

function getQuizData(username, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("Questions");
  const pSheet = ss.getSheetByName("UserProgress");
  const userSheet = ss.getSheetByName("Users");
  const cleanUser = username ? username.toString().trim() : "";
  
  const userData = userSheet.getDataRange().getValues();
  let userGradeLevel = 1; 
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim() === cleanUser) {
      userGradeLevel = parseInt(userData[i][2]) || 1;
      break;
    }
  }

  const qData = qSheet.getDataRange().getValues();
  const pData = pSheet.getDataRange().getValues() || [];
  
  const progressMap = {};
  pData.forEach(row => {
    if (row[0] && row[0].toString().trim() === cleanUser) {
      progressMap[row[1].toString()] = {
        bucket: parseInt(row[3]) || 1,
        date: row[4] instanceof Date ? row[4] : new Date(0)
      };
    }
  });

  const now = new Date();

  let filtered = qData.slice(1).filter(row => {
    if (!row[0] || !row[14]) return false; 
    const questionBeltLevel = parseInt(row[17]) || 1; 
    if (questionBeltLevel > userGradeLevel) return false;

    // Strict Exam Filter: Column Q must not be "N"
    if (row[16] === "N") return false;

    if (mode === 'test') return true; 

    const qId = row[14].toString();
    const prog = progressMap[qId];
    if (!prog) return true; 

    const diffDays = (now - prog.date) / (1000 * 60 * 60 * 24);
    return diffDays >= (BUCKET_INTERVALS[prog.bucket] || 0);
  });

  if (filtered.length === 0) {
    filtered = qData.slice(1).filter(row => {
      const level = parseInt(row[17]) || 1;
      return row[0] && row[14] && level <= userGradeLevel && row[16] !== "N";
    });
  }

  const limit = (mode === 'test') ? 50 : 10;

  return filtered.map(row => {
    let rawOpts = [row[4], row[5], row[6], row[7]].filter(String);
    return {
      question: row[0].toString(),
      options: rawOpts.sort(() => Math.random() - 0.5).map(s => s.toString().trim()),
      answer: row[11] ? row[11].toString().trim() : "",
      qId: row[14].toString()
    };
  }).sort(() => Math.random() - 0.5).slice(0, limit);
}

function getUserGrade(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const cleanU = username ? username.toString().trim().toLowerCase() : "";

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toLowerCase() === cleanU) {
      return parseInt(data[i][2]) || 1;
    }
  }
  return 1;
}

/**
 * GET MATCHING DATA: Only rows with data in Column N.
 * Returns: qId, matchingTerm, correctAnswer.
 */
function getMatchingData(userBelt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Questions");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 19).getValues(); // Up to Col S (19)

  return data
    .filter(r => r[17] === userBelt && r[13] !== "") // Col R (Belt), Col N (Matching Check)
    .map(r => ({
      qId: r[0],
      matchingTerm: r[18], // Column S
      correctAnswer: r[11] // Column L
    }))
    .sort(() => 0.5 - Math.random());
}

/**
 * GET TRUE/FALSE DATA: Only rows with data in Column T.
 * Excludes "N" in Column Q for Practice/Test logic if applied.
 */
function getTFData(userBelt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Questions");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 20).getValues(); // Up to Col T (20)

  return data
    .filter(r => r[17] === userBelt && r[19] !== "") // Col R (Belt), Col T (TF Check)
    .map(r => {
      const isLying = Math.random() > 0.5;
      return {
        qId: r[0],
        simplifiedDef: r[19], // Column T
        correctAnswer: r[11], // Column L
        isCorrect: !isLying,
        displayDef: isLying ? "DECOY_PLACEHOLDER" : r[11] 
      };
    })
    .sort(() => 0.5 - Math.random());
}

/**
 * GET TUL TRUMPS DATA: Only pulls from the Tuls sheet[cite: 1].
 */
function getTulTrumpsData(activeUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Tuls");
  if (!sheet) return { error: "Tuls sheet not found" };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  // Logic to split hands for Jacob vs CPU
  const shuffled = rows.sort(() => 0.5 - Math.random());
  return {
    playerHand: shuffled.slice(0, 5),
    cpuHand: shuffled.slice(5, 10)
  };
}

function updateQuestionScore(username, qId, isCorrect) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName("UserProgress");
  const pData = pSheet.getDataRange().getValues();
  const now = new Date();
  const qIdStr = qId ? qId.toString() : "";
  const cleanUser = username ? username.toString().trim() : "";
  
  let foundRow = -1;
  for (let i = 0; i < pData.length; i++) {
    if (pData[i][0].toString().trim() === cleanUser && pData[i][1].toString() === qIdStr) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow !== -1) {
    let currentBucket = parseInt(pData[foundRow - 1][3]) || 1;
    let currentScore = parseInt(pData[foundRow - 1][2]) || 0;
    let nextBucket = isCorrect ? Math.min(currentBucket + 1, 4) : 1;
    let nextScore = isCorrect ? currentScore + 1 : currentScore - 1;
    pSheet.getRange(foundRow, 3, 1, 3).setValues([[nextScore, nextBucket, now]]);
  } else {
    pSheet.appendRow([cleanUser, qIdStr, isCorrect ? 1 : -1, isCorrect ? 2 : 1, now]);
  }
}

function getGameData(username, gameType) {
  const userGrade = getUserGrade(username);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("Questions");
  const qData = qSheet.getDataRange().getValues().slice(1);
  
  // Filter by: Grade Match AND NOT "N" in Col Q[cite: 1]
  const filtered = qData.filter(row => {
    const isLevelMatch = (parseInt(row[17]) || 1) <= userGrade;
    const isNotExcluded = row[16] !== "N"; // Column Q[cite: 1]
    
    if (gameType === 'game_match') {
      return isLevelMatch && isNotExcluded && row[13] !== ""; // Value in Col N
    } else {
      return isLevelMatch && isNotExcluded && row[19] !== ""; // Value in Col T
    }
  });

  return filtered.sort(() => 0.5 - Math.random()).slice(0, 10).map(row => ({
    matchingTerm: row[18],  // Column S
    simplifiedDef: row[19], // Column T
    correctAnswer: row[11], // Column L
    qId: row[14].toString()
  }));
}

function saveGrade(username, newGrade) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const cleanUser = username ? username.toString().trim() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === cleanUser) {
      sheet.getRange(i + 1, 3).setValue(newGrade);
      return "Grade Updated!";
    }
  }
}

function updatePass(u, p) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const cleanUser = u ? u.toString().trim() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === cleanUser) { 
      sheet.getRange(i + 1, 2).setValue(p); 
      return "Updated!"; 
    }
  }
}