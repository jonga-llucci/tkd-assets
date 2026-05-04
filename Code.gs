/**
 * TKD Theory Quiz - v8.2 (Corrected)
 * Users Tab Map: A:User | B:Pass | C:Grade | D:LastActive | E:Streak | F:Name(5)
 * Questions Tab Map: A:Quest(0) | ... | O:qId(14) | Q:Exam(16) | R:BeltLevel(17)
 */

const BUCKET_INTERVALS = { 1: 0, 2: 2, 3: 4, 4: 5 };
//const TUL_PLACEHOLDER = "https://placehold.co/300x200?text=Pattern+Image";



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

// THIS IS THE BRIDGE: It returns the TulUI content to the main Index file
function getTrumpsHTML() {
  return HtmlService.createHtmlOutputFromFile('TulUI').getContent();
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

/** Helper to get Grade Level without redundant sheet calls */
function _getUserGrade(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const uSheet = ss.getSheetByName("Users");
  const uData = uSheet.getDataRange().getValues();
  
  // Find user row; default to Grade 1 if not found
  const userRow = uData.find(r => r[0] === username);
  return userRow ? parseInt(userRow[1]) : 1; 
}

function getQuizData(username, mode) {
  const userGrade = _getUserGrade(username);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("Questions");
  const qData = qSheet.getDataRange().getValues().slice(1);
  const pSheet = ss.getSheetByName("UserProgress");
  const pData = pSheet && pSheet.getLastRow() > 1 ? pSheet.getDataRange().getValues() : [];
  
  // Map SRS progress for Practice mode
  const progressMap = {};
  pData.filter(r => r[0] == username).forEach(r => {
    progressMap[r[1]] = { bucket: r[3], date: new Date(r[4]) };
  });

  const now = new Date();
  const filtered = qData.filter(row => {
    const qGrade = parseInt(row[17]) || 1; // Column R
    const isExamAllowed = (row[16] === "Y"); // Column Q
    
    // 1. Must be within user's Belt Level
    if (qGrade > userGrade) return false;
    
    // 2. Must be marked for Exams/Practice (Column Q)
    if (!isExamAllowed) return false;

    // 3. SRS Logic for Daily Practice only
    if (mode === 'practice') {
      const prog = progressMap[row[14]]; // Column O (qId)
      if (!prog) return true; 
      const diff = (now - prog.date) / 86400000;
      return diff >= (BUCKET_INTERVALS[prog.bucket] || 0);
    }

    return true;
  });

  const limit = (mode === 'test') ? 50 : 10; //
  
  return filtered.sort(() => 0.5 - Math.random()).slice(0, limit).map(row => {
    // Filter out blank columns to prevent empty UI buttons[cite: 13]
    const rawOptions = [row[4], row[5], row[6], row[7]]; 
    const cleanOptions = rawOptions.filter(opt => opt && opt.toString().trim() !== "");

    return {
      question: row[0],
      options: cleanOptions.sort(() => 0.5 - Math.random()),
      answer: row[11], // Column L
      qId: row[14]   // Column O
    };
  });
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
  const userGrade = _getUserGrade(username);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("Questions");
  const qData = qSheet.getDataRange().getValues().slice(1);

  // Games can use all questions (including Flashcard-only) 
  // or you can add (row[16] === "Y") here if you want them strictly restricted.
  const filtered = qData.filter(row => {
    const qGrade = parseInt(row[17]) || 1;
    return qGrade <= userGrade;
  });

  // Shuffle and return a subset appropriate for games (e.g., 15 items)
  return filtered.sort(() => 0.5 - Math.random()).slice(0, 15).map(row => {
    return {
      term: row[0],       // The Question/Term
      definition: row[11], // The Answer/Definition
      qId: row[14]
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

/**
 * EXPLICIT PLACEMENT: At the end of your grading logic.
 */
function processFinalResults() {
  const score = calculateScore(); // Your existing logic to sum points
  const total = questions.length;
  const percentage = (score / total) * 100;

  // Update the UI elements for the results page
  document.getElementById('display-score').innerText = percentage.toFixed(0) + '%';
  document.getElementById('pass-fail-status').innerText = percentage >= 75 ? "PASSED" : "FAILED";

  // Save the data to the spreadsheet[cite: 13, 14]
  google.script.run
    .withSuccessHandler(() => {
      console.log("Progress Saved");
    })
    .updateUserProgress(currentUser, percentage);

  // --- THE EXPLICIT CALL ---
  // This must be the final action to ensure the UI shifts only after 
  // the variables and elements above are populated.
  switchView('view-results'); 
}