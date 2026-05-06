/**
 * MANDATORY: Serves the HTML file as a Web App
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('TKD Theory Academy')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * AUTHENTICATION:
 * Handles numeric passwords (like 789) and string matching[cite: 2].
 */
function loginUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  const cleanUser = String(username).trim().toLowerCase();
  const cleanPass = String(password).trim();

  for (let i = 1; i < data.length; i++) {
    const dbUser = String(data[i][0]).trim().toLowerCase();
    const dbPass = String(data[i][1]).trim();

    if (dbUser === cleanUser && dbPass === cleanPass) {
      // Column D: Update last active date[cite: 2]
      sheet.getRange(i + 1, 4).setValue(new Date());
      
      return {
        success: true,
        username: data[i][0], 
        displayName: data[i][5], // Column F: Name
        streak: data[i][4],      // Column E: Streak
        gradeValue: data[i][2]   // Column C: Grade
      };
    }
  }
  return { success: false };
}

/**
 * FETCH QUIZ DATA:
 * Filters by belt level and checks the "Exam" flag in Column Q[cite: 1, 2].
 */
function getQuizData(username, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  const questSheet = ss.getSheetByName("Questions");
  const userData = userSheet.getDataRange().getValues();
  const questData = questSheet.getDataRange().getValues();
  
  // 1. Get the current user's grade level
  let userGrade = 1;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === username) {
      userGrade = parseInt(userData[i][2]);
      break;
    }
  }

  let questions = [];
  
  // 2. Filter Questions (Row 0 is header, start at Row 1)
  for (let j = 1; j < questData.length; j++) {
    const qGrade = parseInt(questData[j][1]); // Column B: Grade Level
    const qExamFlag = String(questData[j][16]).trim().toUpperCase(); // Column Q: Exam Mode Flag[cite: 1]
    
    // Rule: Only include questions at or below user's current grade
    if (qGrade <= userGrade) {
      
      // Rule: If Exam Mode (Mock Grading), skip questions where Column Q is "N"[cite: 1, 2]
      if (mode === 'test' && qExamFlag === "N") {
        continue;
      }
      
      questions.push({
        qId: j + 1,
        question: questData[j][2], // Column C
        answer: questData[j][3],   // Column D
        options: [
          questData[j][3], 
          questData[j][4], 
          questData[j][5], 
          questData[j][6]
        ].filter(opt => opt !== "").sort(() => Math.random() - 0.5)
      });
    }
  }
  
  // Return randomized selection (20 for Mock Grading, 10 for Daily Practice)
  const limit = (mode === 'test') ? 20 : 10;
  return questions.sort(() => Math.random() - 0.5).slice(0, limit);
}

/**
 * UPDATE USER DATA:
 * Saves changes from the Account View[cite: 2].
 */
function updateUserData(username, gradeValue, newPassword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      // Update Column C (Index 2) for Grade/Belt level
      sheet.getRange(i + 1, 3).setValue(gradeValue);
      
      // Update Column B (Index 1) if a new password was provided
      if (newPassword && newPassword !== "") {
        sheet.getRange(i + 1, 2).setValue(newPassword);
      }
      return true;
    }
  }
  return false;
}

/**
 * Placeholder for individual question performance tracking
 */
function updateQuestionScore(username, qId, isCorrect) {
  // Logic for SRS (Spaced Repetition System) can be implemented here
  return true;
}