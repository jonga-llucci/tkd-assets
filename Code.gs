/**
 * TKD Theory Quiz - v8.2 (Corrected)
 * Users Tab Map: A:User | B:Pass | C:Grade | D:LastActive | E:Streak | F:Name(5)
 * Questions Tab Map: A:Quest(0) | ... | O:qId(14) | Q:Exam(16) | R:BeltLevel(17)
 */

const BUCKET_INTERVALS = { 1: 0, 2: 2, 3: 4, 4: 5 };


/** 
 * Keep your existing loginUser, getQuizData, and updateQuestionScore functions exactly as they were.
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('TKD Theory Academy')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) 
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
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
  const cleanUser = username ? username.toString().trim() : "";
  
  const userGradeLevel = getUserGrade_(username);

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

  // Before returning, sort practice questions by SRS priority (bucket ascending, most overdue first)
  // then shuffle within each priority tier so it doesn't feel mechanical
  if (mode !== 'test') {
    filtered.sort((a, b) => {
      const progA = progressMap[a[14]?.toString()] || { bucket: 1, date: new Date(0) };
      const progB = progressMap[b[14]?.toString()] || { bucket: 1, date: new Date(0) };
      if (progA.bucket !== progB.bucket) return progA.bucket - progB.bucket; // Bucket 1 first
      return (progA.date - progB.date); // Most overdue first within same bucket
    });
  }

  return filtered.map(row => {
  let rawOpts = [row[4], row[5], row[6], row[7]].filter(String);
  return {
    question: row[0].toString(),
    options: rawOpts.sort(() => Math.random() - 0.5).map(s => s.toString().trim()),
    answer: row[11] ? row[11].toString().trim() : "",
    qId: row[14].toString(),
    timeLimit: parseInt(row[21]) || 5
  };
}).slice(0, limit);
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("Questions");
  
  const userGradeLevel = getUserGrade_(username);

  const qData = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, qSheet.getLastColumn()).getValues();
  
  const filtered = qData.filter(row => {
    const isLevelMatch = (parseInt(row[17]) || 1) <= userGradeLevel;
    const isNotExcluded = row[16] !== "N";
    
    if (gameType === 'game_match') {
      return isLevelMatch && isNotExcluded && row[18].toString().trim() !== "";
    } else {
      return isLevelMatch && isNotExcluded && row[19].toString().trim() !== "";
    }
  });

  return filtered.sort(() => Math.random() - 0.5).slice(0, 10).map(row => {
    const possibleDecoys = [row[4], row[5], row[6], row[7]].filter(val => 
      val && val.toString().trim() !== "" && val.toString().trim() !== row[11].toString().trim()
    );
    return {
      matchingTerm: row[18].toString().trim(),
      simplifiedDef: row[19].toString().trim(),
      correctAnswer: row[11].toString().trim(),
      decoys: possibleDecoys,
      qId: row[14].toString()
    };
  });
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

function getBeltOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Belts");
  const data = sheet.getDataRange().getValues();
  return data.slice(1)
    .filter(row => row[0] && row[1])
    .map(row => ({ label: row[0].toString().trim(), value: parseInt(row[1]) }));
}

function getSrsStats(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pData = ss.getSheetByName("UserProgress").getDataRange().getValues();
  const qData = ss.getSheetByName("Questions").getDataRange().getValues();
  const cleanUser = username ? username.toString().trim() : "";
  const userGradeLevel = getUserGrade_(username);

  // Total questions eligible for this user's grade
  const eligible = qData.slice(1).filter(row =>
    row[0] && row[14] && (parseInt(row[17]) || 1) <= userGradeLevel && row[16] !== "N"
  ).length;

  // Count user's rows by bucket
  const buckets = { 1: 0, 2: 0, 3: 0, 4: 0 };
  pData.slice(1).forEach(row => {
    if (row[0] && row[0].toString().trim() === cleanUser) {
      const b = parseInt(row[3]) || 1;
      if (buckets[b] !== undefined) buckets[b]++;
    }
  });

  // Questions not yet seen = eligible minus all seen
  const seen = buckets[1] + buckets[2] + buckets[3] + buckets[4];
  const unseen = Math.max(0, eligible - seen);

  return { buckets, unseen, eligible };
}

function getHighScore(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("HighScores");
  if (!sheet) {
    sheet = ss.insertSheet("HighScores");
    sheet.appendRow(["Username", "InfiniteWarrior"]);
  }
  const data = sheet.getDataRange().getValues();
  const clean = username ? username.toString().trim() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === clean)
      return parseInt(data[i][1]) || 0;
  }
  return 0;
}

function saveHighScore(username, newScore) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("HighScores");
  if (!sheet) {
    sheet = ss.insertSheet("HighScores");
    sheet.appendRow(["Username", "InfiniteWarrior"]);
  }
  const data = sheet.getDataRange().getValues();
  const clean = username ? username.toString().trim() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === clean) {
      sheet.getRange(i + 1, 2).setValue(newScore);
      return;
    }
  }
  sheet.appendRow([clean, newScore]);
}

function getAllTimeHighScore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("HighScores");
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const val = parseInt(data[i][1]) || 0;
    if (val > max) max = val;
  }
  return max;
}

function getUserGrade_(username) {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Users").getDataRange().getValues();
  const clean = username ? username.toString().trim().toLowerCase() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === clean)
      return parseInt(data[i][2]) || 1;
  }
  return 1;
}

function getTulTrumpsData(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tulSheet = ss.getSheetByName("Tuls");

    const userGradeLevel = getUserGrade_(username);

    const tulData = tulSheet.getDataRange().getValues();
    
    const deck = tulData.slice(1)
      .filter(row => {
        const tulName = row[0];
        const tulRequiredGrade = Number(row[5]);
        return tulName && !isNaN(tulRequiredGrade) && tulRequiredGrade <= userGradeLevel;
      }) 
      .map(row => ({
        name: row[0],         
        movements: parseInt(row[1]) || 0, 
        stances: parseInt(row[2]) || 0,   
        readyStance: row[3] || "---",  
        difficulty: parseInt(row[4]) || 0, 
        meaning: row[6] || "",      
        img: row[7] || "https://placehold.co/300x200?text=No+Pattern+Image"
      }));

    if (deck.length < 4) {
      return { error: "Not enough Tuls unlocked for this grade. (Found: " + deck.length + ")" };
    }

    const shuffled = deck.sort(() => Math.random() - 0.5);
    const mid = Math.ceil(shuffled.length / 2);
    
    return {
      playerHand: shuffled.slice(0, mid),
      cpuHand: shuffled.slice(mid)
    };
  } catch (e) {
    return { error: e.message };
  }
}