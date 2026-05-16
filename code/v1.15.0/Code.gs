const BUCKET_INTERVALS = { 1: 0, 2: 2, 3: 4, 4: 5 };
const TRIAL_DAYS = 7;
const SUB_DAYS = 30;

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('TKD Theory Academy')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function parseSheetDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const str = val.toString().trim();
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

// ── Single user-context lookup (grade + club) ─────────────────────────────────
function getUserContext_(username) {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Users").getDataRange().getValues();
  const clean = username ? username.toString().trim().toLowerCase() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === clean)
      return {
        grade: parseInt(data[i][2]) || 1,
        club:  data[i][10] ? data[i][10].toString().trim() : ''
      };
  }
  return { grade: 1, club: '' };
}

// Kept for callers that only need grade (avoids fetching club unnecessarily)
function getUserGrade_(username) {
  return getUserContext_(username).grade;
}

// ─────────────────────────────────────────────────────────────────────────────

function loginUser(username, password) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  const userData  = userSheet.getDataRange().getValues();
  const now       = new Date();

  const cleanU = username ? username.toString().trim().toLowerCase() : "";
  const cleanP = password ? password.toString().trim() : "";

  for (let i = 1; i < userData.length; i++) {
    if (!userData[i][0]) continue;
    if (userData[i][0].toString().trim().toLowerCase() !== cleanU) continue;
    if (userData[i][1].toString() !== cleanP) continue;

    // ── Streak reset check ──────────────────────────────────────────────────
    // Col E (index 4)  = streak count
    // Col P (index 15) = streakDate: the last date the daily goal was met
    // Col D (index 3)  = lastActive (general login date)
    let streak          = parseInt(userData[i][4])  || 0;
    const streakDateRaw = userData[i][15];
    const lastStreakDate = streakDateRaw ? new Date(streakDateRaw) : null;

    if (lastStreakDate) {
      const daysSinceStreak = Math.floor((now - lastStreakDate) / (1000 * 60 * 60 * 24));
      // If the user missed yesterday (2+ days since last goal-met), reset streak
      if (daysSinceStreak >= 2) {
        streak = 0;
        userSheet.getRange(i + 1, 5).setValue(0);
      }
    }

    // Update last-active date
    const lastActiveDate = userData[i][3] ? new Date(userData[i][3]) : null;
    if (!lastActiveDate || lastActiveDate.toDateString() !== now.toDateString()) {
      userSheet.getRange(i + 1, 4).setValue(now);
    }

    // ── Access / subscription status ────────────────────────────────────────
    const registeredDate   = parseSheetDate_(userData[i][8]);
    const subscriptionDate = parseSheetDate_(userData[i][9]);
    let accessStatus  = 'active';
    let daysRemaining = null;

    if (registeredDate) {
      const daysSinceRegistered = (now - registeredDate) / (1000 * 60 * 60 * 24);
      const daysSinceSub        = subscriptionDate
        ? (now - subscriptionDate) / (1000 * 60 * 60 * 24)
        : null;

      if (subscriptionDate && daysSinceSub <= SUB_DAYS) {
        daysRemaining = Math.max(0, Math.ceil(SUB_DAYS - daysSinceSub));
        accessStatus  = 'subscribed';
      } else if (daysSinceRegistered <= TRIAL_DAYS) {
        daysRemaining = Math.max(0, Math.ceil(TRIAL_DAYS - daysSinceRegistered));
        accessStatus  = 'trial';
      } else {
        accessStatus = subscriptionDate ? 'subscription_expired' : 'trial_expired';
      }
    }

    return {
      success:          true,
      username:         userData[i][0].toString().trim(),
      displayName:      userData[i][5] ? userData[i][5].toString() : userData[i][0].toString(),
      gradeValue:       parseInt(userData[i][2]) || 1,
      streak:           streak,
      isAdmin:          userData[i][6] && userData[i][6].toString().trim().toUpperCase() === 'Y',
      accessStatus:     accessStatus,
      daysRemaining:    daysRemaining,
      learningGoalMins: parseInt(userData[i][11]) || 0
    };
  }
}

function registerUser(displayName, email, gradeValue, password, club) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const sheet      = ss.getSheetByName("Users");
  const data       = sheet.getDataRange().getValues();
  const cleanEmail = email ? email.toString().trim().toLowerCase() : "";

  if (!cleanEmail || !displayName || !password || !gradeValue)
    return { success: false, message: "All fields are required." };

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === cleanEmail)
      return { success: false, message: "An account with this email already exists." };
  }

  const now = new Date();
  sheet.appendRow([
    cleanEmail,
    password,
    parseInt(gradeValue),
    now,
    1,
    displayName.toString().trim(),
    '', '', now, '',
    club ? club.toString().trim() : ''
  ]);
  return { success: true };
}

function getQuizData(username, mode) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet  = ss.getSheetByName("Questions");
  const pSheet  = ss.getSheetByName("UserProgress");
  const cleanUser = username ? username.toString().trim() : "";
  const { grade: userGradeLevel, club: userClub } = getUserContext_(username);

  const qData = qSheet.getDataRange().getValues();
  const pData = pSheet.getDataRange().getValues() || [];

  const progressMap = {};
  pData.forEach(row => {
    if (row[0] && row[0].toString().trim() === cleanUser) {
      progressMap[row[1].toString()] = {
        bucket: parseInt(row[3]) || 1,
        date:   row[4] instanceof Date ? row[4] : new Date(0)
      };
    }
  });

  const now = new Date();

  const eligible = qData.slice(1).filter(row => {
    if (!row[0] || !row[14]) return false;
    if ((parseInt(row[17]) || 1) > userGradeLevel) return false;
    if (row[16] === "N") return false;
    const questionClub = row[23] ? row[23].toString().trim() : '';
    if (questionClub && questionClub !== userClub) return false;
    return true;
  });

  // ── EXAM MODE ─────────────────────────────────────────────────────────────
  if (mode === 'test') {
    let limit;
    if (userGradeLevel >= 11)     limit = 50;
    else if (userGradeLevel >= 7) limit = 20;
    else                          limit = 10;

    const categories = [
      { name: 'Basics',   pct: 0.10 },
      { name: 'Numbers',  pct: 0.10 },
      { name: 'Belts',    pct: 0.10 },
      { name: 'Korean',   pct: 0.30 },
      { name: 'Patterns', pct: 0.30 }
    ];

    function weightedSample(pool, n) {
      if (pool.length === 0) return [];
      const weights     = pool.map(row => parseInt(row[17]) || 1);
      const totalWeight = weights.reduce((a, b) => a + b, 0);
      const selected = [], used = new Set();
      let attempts = 0;
      while (selected.length < Math.min(n, pool.length) && attempts < pool.length * 10) {
        attempts++;
        let rand = Math.random() * totalWeight;
        for (let i = 0; i < pool.length; i++) {
          if (used.has(i)) continue;
          rand -= weights[i];
          if (rand <= 0) { selected.push(pool[i]); used.add(i); break; }
        }
      }
      return selected;
    }

    const pools = {};
    categories.forEach(c => pools[c.name] = []);
    const uncategorised = [];
    eligible.forEach(row => {
      const cat = row[20] ? row[20].toString().trim() : '';
      if (pools[cat] !== undefined) pools[cat].push(row);
      else uncategorised.push(row);
    });

    let selected = [];
    categories.forEach(c => {
      selected = selected.concat(weightedSample(pools[c.name], Math.round(c.pct * limit)));
    });

    const remaining = limit - selected.length;
    if (remaining > 0) {
      const selectedIds = new Set(selected.map(r => r[14].toString()));
      const fillPool    = [...uncategorised, ...eligible].filter(r => !selectedIds.has(r[14].toString()));
      selected = selected.concat(weightedSample(fillPool, remaining));
    }

    return selected.sort(() => Math.random() - 0.5).map(row => ({
      question: row[0].toString(),
      options:  [row[4], row[5], row[6], row[7]].filter(String).sort(() => Math.random() - 0.5).map(s => s.toString().trim()),
      answer:   row[11] ? row[11].toString().trim() : "",
      qId:      row[14].toString(),
      timeLimit: parseInt(row[21]) || 5
    }));
  }

  // ── PRACTICE MODE ─────────────────────────────────────────────────────────
  let filtered = eligible.filter(row => {
    const qId  = row[14].toString();
    const prog = progressMap[qId];
    if (!prog) return true;
    const diffDays = (now - prog.date) / (1000 * 60 * 60 * 24);
    return diffDays >= (BUCKET_INTERVALS[prog.bucket] || 0);
  });

  if (filtered.length === 0) filtered = [...eligible];

  filtered.sort((a, b) => {
    const progA = progressMap[a[14]?.toString()] || { bucket: 1, date: new Date(0) };
    const progB = progressMap[b[14]?.toString()] || { bucket: 1, date: new Date(0) };
    if (progA.bucket !== progB.bucket) return progA.bucket - progB.bucket;
    return progA.date - progB.date;
  });

  return filtered.slice(0, 10).map(row => ({
    question: row[0].toString(),
    options:  [row[4], row[5], row[6], row[7]].filter(String).sort(() => Math.random() - 0.5).map(s => s.toString().trim()),
    answer:   row[11] ? row[11].toString().trim() : "",
    qId:      row[14].toString(),
    timeLimit: parseInt(row[21]) || 5
  }));
}

function updateQuestionScore(username, qId, isCorrect) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet  = ss.getSheetByName("UserProgress");
  const pData   = pSheet.getDataRange().getValues();
  const now     = new Date();
  const qIdStr  = qId ? qId.toString() : "";
  const cleanUser = username ? username.toString().trim() : "";

  let foundRow = -1;
  for (let i = 0; i < pData.length; i++) {
    if (pData[i][0].toString().trim() === cleanUser && pData[i][1].toString() === qIdStr) {
      foundRow = i + 1; break;
    }
  }

  if (foundRow !== -1) {
    let currentBucket = parseInt(pData[foundRow - 1][3]) || 1;
    let currentScore  = parseInt(pData[foundRow - 1][2]) || 0;
    pSheet.getRange(foundRow, 3, 1, 3).setValues([[
      isCorrect ? currentScore + 1 : currentScore - 1,
      isCorrect ? Math.min(currentBucket + 1, 4) : 1,
      now
    ]]);
  } else {
    pSheet.appendRow([cleanUser, qIdStr, isCorrect ? 1 : -1, isCorrect ? 2 : 1, now]);
  }
}

function getGameData(username, gameType) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("Questions");
  const { grade: userGradeLevel, club: userClub } = getUserContext_(username);

  const qData    = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, qSheet.getLastColumn()).getValues();
  const filtered = qData.filter(row => {
    const isLevelMatch   = (parseInt(row[17]) || 1) <= userGradeLevel;
    const isNotExcluded  = row[16] !== "N";
    const questionClub   = row[23] ? row[23].toString().trim() : '';
    const isClubMatch    = !questionClub || questionClub === userClub;
    if (gameType === 'game_match')
      return isLevelMatch && isNotExcluded && isClubMatch && row[18].toString().trim() !== "";
    return isLevelMatch && isNotExcluded && isClubMatch && row[19].toString().trim() !== "";
  });

  return filtered.sort(() => Math.random() - 0.5).slice(0, 10).map(row => {
    const possibleDecoys = [row[4], row[5], row[6], row[7]].filter(val =>
      val && val.toString().trim() !== "" && val.toString().trim() !== row[11].toString().trim()
    );
    return {
      matchingTerm:  row[18].toString().trim(),
      simplifiedDef: row[19].toString().trim(),
      correctAnswer: row[11].toString().trim(),
      decoys:        possibleDecoys,
      qId:           row[14].toString()
    };
  });
}

function saveGrade(username, newGrade) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data  = sheet.getDataRange().getValues();
  const clean = username ? username.toString().trim() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === clean) {
      sheet.getRange(i + 1, 3).setValue(newGrade);
      return "Grade Updated!";
    }
  }
}

function updatePass(u, p) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data  = sheet.getDataRange().getValues();
  const clean = u ? u.toString().trim() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === clean) {
      sheet.getRange(i + 1, 2).setValue(p);
      return "Updated!";
    }
  }
}

function getBeltOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Belts");
  const data  = sheet.getDataRange().getValues();
  return data.slice(1)
    .filter(row => row[0] && row[1])
    .map(row => ({ label: row[0].toString().trim(), value: parseInt(row[1]) }));
}

function getSrsStats(username) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const pData     = ss.getSheetByName("UserProgress").getDataRange().getValues();
  const qData     = ss.getSheetByName("Questions").getDataRange().getValues();
  const cleanUser = username ? username.toString().trim() : "";
  const { grade: userGradeLevel } = getUserContext_(username);

  const eligible = qData.slice(1).filter(row =>
    row[0] && row[14] && (parseInt(row[17]) || 1) <= userGradeLevel && row[16] !== "N"
  ).length;

  const buckets = { 1: 0, 2: 0, 3: 0, 4: 0 };
  pData.slice(1).forEach(row => {
    if (row[0] && row[0].toString().trim() === cleanUser) {
      const b = parseInt(row[3]) || 1;
      if (buckets[b] !== undefined) buckets[b]++;
    }
  });

  const seen   = buckets[1] + buckets[2] + buckets[3] + buckets[4];
  const unseen = Math.max(0, eligible - seen);
  return { buckets, unseen, eligible };
}

function getHighScore(username) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName("HighScores");
  if (!sheet) { sheet = ss.insertSheet("HighScores"); sheet.appendRow(["Username", "InfiniteWarrior"]); }
  const data  = sheet.getDataRange().getValues();
  const clean = username ? username.toString().trim() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === clean) return parseInt(data[i][1]) || 0;
  }
  return 0;
}

function saveHighScore(username, newScore) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName("HighScores");
  if (!sheet) { sheet = ss.insertSheet("HighScores"); sheet.appendRow(["Username", "InfiniteWarrior"]); }
  const data  = sheet.getDataRange().getValues();
  const clean = username ? username.toString().trim() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === clean) { sheet.getRange(i + 1, 2).setValue(newScore); return; }
  }
  sheet.appendRow([clean, newScore]);
}

function getAllTimeHighScore() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HighScores");
  if (!sheet) return 0;
  const data  = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const val = parseInt(data[i][1]) || 0;
    if (val > max) max = val;
  }
  return max;
}

function getTulTrumpsData(username) {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const tulSheet = ss.getSheetByName("Tuls");
    const { grade: userGradeLevel } = getUserContext_(username);
    const tulData  = tulSheet.getDataRange().getValues();

    const deck = tulData.slice(1)
      .filter(row => {
        const tulName          = row[0];
        const tulRequiredGrade = Number(row[5]);
        return tulName && !isNaN(tulRequiredGrade) && tulRequiredGrade <= userGradeLevel;
      })
      .map(row => ({
        name:        row[0],
        movements:   parseInt(row[1]) || 0,
        stances:     parseInt(row[2]) || 0,
        readyStance: row[3] || "---",
        difficulty:  parseInt(row[4]) || 0,
        meaning:     row[6] || "",
        img:         row[7] || "https://placehold.co/300x200?text=No+Pattern+Image"
      }));

    if (deck.length < 4) return { error: "Not enough Tuls unlocked for this grade. (Found: " + deck.length + ")" };

    const shuffled = deck.sort(() => Math.random() - 0.5);
    const mid      = Math.ceil(shuffled.length / 2);
    return { playerHand: shuffled.slice(0, mid), cpuHand: shuffled.slice(mid) };
  } catch(e) {
    return { error: e.message };
  }
}

// ── Streak: increment once per day, only when daily goal is met ──────────────
// Col E  (index 4)  = streak count
// Col P  (index 15) = streakDate (last date goal was fully met)
function incrementStreak(username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data  = sheet.getDataRange().getValues();
  const clean = username ? username.toString().trim().toLowerCase() : "";
  const today = new Date().toDateString();

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().trim().toLowerCase() !== clean) continue;

    // Idempotency guard: only increment if goal hasn't already been met today
    const streakDateRaw = data[i][15];
    const lastGoalDate  = streakDateRaw ? new Date(streakDateRaw).toDateString() : '';
    if (lastGoalDate === today) {
      return { success: true, streak: parseInt(data[i][4]) || 0, alreadyDone: true };
    }

    const currentStreak = parseInt(data[i][4]) || 0;
    sheet.getRange(i + 1, 5).setValue(currentStreak + 1);   // Col E: streak count
    sheet.getRange(i + 1, 16).setValue(new Date());          // Col P: streakDate (1-based = 16)
    return { success: true, streak: currentStreak + 1 };
  }
  return { success: false };
}

// ── Play-time logging: daily JSON + monthly JSON (auto-resets each month) ────
// Col M (index 12, 1-based 13) = daily:   { date: "...", seconds: N }
// Col N (index 13, 1-based 14) = monthly: { month: "YYYY-M", seconds: N }
function logPlayTime(username, seconds) {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data    = sheet.getDataRange().getValues();
  const clean   = username ? username.toString().trim().toLowerCase() : "";
  const now     = new Date();
  const todayKey = now.toDateString();
  const monthKey = `${now.getFullYear()}-${now.getMonth()}`; // e.g. "2026-4"

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().trim().toLowerCase() !== clean) continue;

    // Daily
    let dailyData = { date: todayKey, seconds: 0 };
    try {
      const ex = data[i][12] ? JSON.parse(data[i][12].toString()) : null;
      dailyData  = (ex && ex.date === todayKey)
        ? { date: todayKey, seconds: (ex.seconds || 0) + seconds }
        : { date: todayKey, seconds };
    } catch(e) { dailyData = { date: todayKey, seconds }; }

    // Monthly — resets automatically when month changes
    let monthlyData = { month: monthKey, seconds: 0 };
    try {
      const exM = data[i][13] ? JSON.parse(data[i][13].toString()) : null;
      monthlyData = (exM && exM.month === monthKey)
        ? { month: monthKey, seconds: (exM.seconds || 0) + seconds }
        : { month: monthKey, seconds };
    } catch(e) { monthlyData = { month: monthKey, seconds }; }

    sheet.getRange(i + 1, 13).setValue(JSON.stringify(dailyData));
    sheet.getRange(i + 1, 14).setValue(JSON.stringify(monthlyData));
    return { success: true };
  }
  return { success: false };
}

function saveLearningGoal(username, mins) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data  = sheet.getDataRange().getValues();
  const clean = username ? username.toString().trim().toLowerCase() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === clean) {
      sheet.getRange(i + 1, 12).setValue(parseInt(mins) || 0); // Col L
      return { success: true };
    }
  }
  return { success: false };
}

// ── Flashcard data — includes belt labels to eliminate getBeltOptions call ────
function getFlashcardData(username) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Flashcards");
  if (!sheet) return { error: "Flashcards sheet not found." };

  const { grade: userGradeLevel } = getUserContext_(username);
  const data = sheet.getDataRange().getValues();

  // Load belt labels once
  const beltSheet = ss.getSheetByName("Belts");
  const beltRaw   = beltSheet ? beltSheet.getDataRange().getValues() : [];
  const beltMap   = {};
  beltRaw.slice(1).forEach(r => { if (r[0] && r[1]) beltMap[parseInt(r[1])] = r[0].toString().trim(); });

  const cards = [];
  for (let i = 1; i < data.length; i++) {
    const row       = data[i];
    const beltLevel = parseInt(row[0]) || 0;
    if (!beltLevel || beltLevel > userGradeLevel) continue;

    const category = row[1] ? row[1].toString().trim() : "";
    const term     = row[2] ? row[2].toString().trim() : "";
    const meaning  = row[3] ? row[3].toString().trim() : "";
    if (!term && !meaning) continue;

    const card = { beltLevel, category, term, meaning };

    if (category === "Patterns") {
      const movesNum      = row[4]  ? row[4].toString().trim()  : "";
      const movesKorean   = row[5]  ? row[5].toString().trim()  : "";
      const stancesNum    = row[6]  ? row[6].toString().trim()  : "";
      const stancesKorean = row[7]  ? row[7].toString().trim()  : "";
      const readyEn       = row[8]  ? row[8].toString().trim()  : "";
      const readyKo       = row[9]  ? row[9].toString().trim()  : "";
      const imgUrl        = row[10] ? row[10].toString().trim() : "";
      if (movesNum)      card.movesNum      = movesNum;
      if (movesKorean)   card.movesKorean   = movesKorean;
      if (stancesNum)    card.stancesNum    = stancesNum;
      if (stancesKorean) card.stancesKorean = stancesKorean;
      if (readyEn)       card.readyEn       = readyEn;
      if (readyKo)       card.readyKo       = readyKo;
      if (imgUrl)        card.img           = imgUrl;
    }
    cards.push(card);
  }

  // beltLevels now includes labels — front end needs no separate getBeltOptions call
  const beltLevels = [...new Set(cards.map(c => c.beltLevel))].sort((a, b) => a - b)
    .map(lv => ({ value: lv, label: beltMap[lv] || `Level ${lv}` }));

  const categories = [...new Set(cards.map(c => c.category).filter(Boolean))].sort();
  return { cards, beltLevels, categories };
}

// ── Admin ─────────────────────────────────────────────────────────────────────

function getAdminPageData(username) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName("Users");
  const userData   = usersSheet.getDataRange().getValues();
  const cleanCaller = username ? username.toString().trim().toLowerCase() : "";

  let callerRow = null;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === cleanCaller) {
      callerRow = userData[i]; break;
    }
  }
  if (!callerRow || callerRow[6].toString().trim().toUpperCase() !== 'Y') return { error: "Unauthorised" };

  const callerClub     = callerRow[10] ? callerRow[10].toString().trim() : '';
  const progressSheet  = ss.getSheetByName("UserProgress");
  const highScoreSheet = ss.getSheetByName("HighScores");
  const beltsSheet     = ss.getSheetByName("Belts");
  const questionsSheet = ss.getSheetByName("Questions");

  const progressData  = progressSheet  ? progressSheet.getDataRange().getValues()  : [];
  const highScoreData = highScoreSheet ? highScoreSheet.getDataRange().getValues() : [];
  const beltData      = beltsSheet     ? beltsSheet.getDataRange().getValues()     : [];
  const questionData  = questionsSheet ? questionsSheet.getDataRange().getValues() : [];

  const beltMap = {};
  beltData.slice(1).forEach(row => { if (row[0] && row[1]) beltMap[parseInt(row[1])] = row[0].toString().trim(); });

  const progressMap = {};
  progressData.slice(1).forEach(row => {
    const u = row[0] ? row[0].toString().trim() : "";
    if (!u) return;
    if (!progressMap[u]) progressMap[u] = { total: 0, buckets: { 1:0, 2:0, 3:0, 4:0 } };
    progressMap[u].total++;
    const b = parseInt(row[3]) || 1;
    if (progressMap[u].buckets[b] !== undefined) progressMap[u].buckets[b]++;
  });

  const highScoreMap = {};
  highScoreData.slice(1).forEach(row => {
    if (row[0]) highScoreMap[row[0].toString().trim()] = parseInt(row[1]) || 0;
  });

  const now      = new Date();
  const monthKey = `${now.getFullYear()}-${now.getMonth()}`;

  const users = userData.slice(1)
    .filter(row => {
      if (!row[0]) return false;
      if (callerClub) {
        const memberClub = row[10] ? row[10].toString().trim() : '';
        return memberClub === callerClub;
      }
      return true;
    })
    .map(row => {
      const uName      = row[0].toString().trim();
      const lastActive = row[3] ? new Date(row[3]) : null;
      const daysSince  = lastActive ? Math.floor((now - lastActive) / (1000 * 60 * 60 * 24)) : null;
      const gradeVal   = parseInt(row[2]) || 1;
      const prog       = progressMap[uName] || { total: 0, buckets: { 1:0, 2:0, 3:0, 4:0 } };
      const eligible   = questionData.slice(1).filter(q =>
        q[0] && q[14] && (parseInt(q[17]) || 1) <= gradeVal && q[16] !== "N"
      ).length;

      const registeredDate   = parseSheetDate_(row[8]);
      const subscriptionDate = parseSheetDate_(row[9]);
      let subStatus = 'active', subEndDate = null;

      if (registeredDate) {
        const daysSinceReg = (now - registeredDate) / (1000 * 60 * 60 * 24);
        const daysSinceSub = subscriptionDate ? (now - subscriptionDate) / (1000 * 60 * 60 * 24) : null;
        if (subscriptionDate && daysSinceSub <= SUB_DAYS) {
          subStatus = 'subscribed';
          const endDate = new Date(subscriptionDate); endDate.setDate(endDate.getDate() + SUB_DAYS);
          subEndDate = endDate.toLocaleDateString('en-GB');
        } else if (daysSinceReg <= TRIAL_DAYS) {
          subStatus = 'trial';
          const endDate = new Date(registeredDate); endDate.setDate(endDate.getDate() + TRIAL_DAYS);
          subEndDate = endDate.toLocaleDateString('en-GB');
        } else {
          subStatus = subscriptionDate ? 'sub_expired' : 'trial_expired';
        }
      }

      // Daily play: JSON { date, seconds }
      const dailyPlaySeconds = (() => {
        try {
          const d = row[12] ? JSON.parse(row[12].toString()) : null;
          return (d && d.date === now.toDateString()) ? parseInt(d.seconds) || 0 : 0;
        } catch(e) { return 0; }
      })();

      // Monthly play: JSON { month, seconds } — auto-resets when month changes
      const monthlyPlaySeconds = (() => {
        try {
          const m = row[13] ? JSON.parse(row[13].toString()) : null;
          return (m && m.month === monthKey) ? parseInt(m.seconds) || 0 : 0;
        } catch(e) { return 0; }
      })();

      return {
        username:            uName,
        displayName:         row[5] ? row[5].toString() : uName,
        grade:               gradeVal,
        gradeName:           beltMap[gradeVal] || `Level ${gradeVal}`,
        streak:              parseInt(row[4]) || 0,
        lastActive:          lastActive ? lastActive.toLocaleString('en-GB') : 'Never',
        daysSince:           daysSince,
        isAdmin:             row[6] && row[6].toString().trim().toUpperCase() === 'Y',
        club:                row[10] ? row[10].toString().trim() : '',
        totalAnswered:       prog.total,
        buckets:             prog.buckets,
        highScore:           highScoreMap[uName] || 0,
        eligible:            eligible,
        subStatus:           subStatus,
        subEndDate:          subEndDate,
        learningGoalMins:    parseInt(row[11]) || 0,
        dailyPlaySeconds:    dailyPlaySeconds,
        monthlyPlaySeconds:  monthlyPlaySeconds
      };
    });

  const belts = beltData.slice(1)
    .filter(row => row[0] && row[1])
    .map(row => ({ label: row[0].toString().trim(), value: parseInt(row[1]) }));

  const categories = [...new Set(
    questionData.slice(1).map(row => row[20] ? row[20].toString().trim() : '').filter(Boolean)
  )].sort();

  return { users, belts, categories, callerClub };
}

function getAdminQuestions(username) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName("Users");
  const userData   = usersSheet.getDataRange().getValues();
  const cleanCaller = username ? username.toString().trim().toLowerCase() : "";

  let callerRow = null;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === cleanCaller) {
      callerRow = userData[i]; break;
    }
  }
  if (!callerRow || callerRow[6].toString().trim().toUpperCase() !== 'Y') return { error: "Unauthorised" };

  const callerClub   = callerRow[10] ? callerRow[10].toString().trim() : '';
  const qSheet       = ss.getSheetByName("Questions");
  const qData        = qSheet.getDataRange().getValues();
  const beltsSheet   = ss.getSheetByName("Belts");
  const beltData     = beltsSheet ? beltsSheet.getDataRange().getValues() : [];
  const beltMap      = {};
  beltData.slice(1).forEach(row => { if (row[0] && row[1]) beltMap[parseInt(row[1])] = row[0].toString().trim(); });

  const questions = qData.slice(1)
    .filter(row => row[0] && row[14])
    .map(row => ({
      qId:       row[14] ? row[14].toString() : '',
      question:  row[0]  ? row[0].toString()  : '',
      answer:    row[11] ? row[11].toString()  : '',
      opt1:      row[5]  ? row[5].toString()   : '',
      opt2:      row[6]  ? row[6].toString()   : '',
      opt3:      row[7]  ? row[7].toString()   : '',
      beltLevel: parseInt(row[17]) || 1,
      beltName:  beltMap[parseInt(row[17])] || `Level ${parseInt(row[17]) || 1}`,
      category:  row[20] ? row[20].toString().trim() : '',
      club:      row[23] ? row[23].toString().trim() : '',
      addedByMe: row[23] ? row[23].toString().trim() === callerClub : false,
      examFlag:  row[16] ? row[16].toString().trim() : ''
    }));

  return { questions, callerClub };
}

function submitQuestion(username, questionData) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName("Users");
  const userData   = usersSheet.getDataRange().getValues();
  const cleanUser  = username ? username.toString().trim().toLowerCase() : "";

  let isAdmin = false, callerClub = '';
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === cleanUser) {
      isAdmin    = userData[i][6] && userData[i][6].toString().trim().toUpperCase() === 'Y';
      callerClub = userData[i][10] ? userData[i][10].toString().trim() : '';
      break;
    }
  }
  if (!isAdmin) return { success: false, message: "Unauthorised" };

  const qSheet = ss.getSheetByName("Questions");
  const qId    = 'q_' + Date.now();
  const newRow = new Array(24).fill('');
  newRow[0]  = questionData.question;
  newRow[4]  = questionData.correctAnswer;
  newRow[5]  = questionData.opt1;
  newRow[6]  = questionData.opt2;
  if (questionData.opt3) newRow[7] = questionData.opt3;
  newRow[11] = questionData.correctAnswer;
  newRow[14] = qId;
  newRow[16] = 'Y';
  newRow[17] = parseInt(questionData.beltLevel);
  newRow[20] = questionData.category;
  newRow[23] = questionData.clubOnly ? callerClub : '';
  qSheet.appendRow(newRow);
  return { success: true, qId };
}

function submitContact(username, contactData) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName("Contact");
  if (!sheet) { sheet = ss.insertSheet("Contact"); sheet.appendRow(["Timestamp","Username","DisplayName","Club","Type","Message"]); }

  sheet.appendRow([new Date(), username, contactData.displayName, contactData.club, contactData.type, contactData.message]);

  try {
    MailApp.sendEmail({
      to:      "jonathan_gallucci@hotmail.com",
      subject: "TKD Academy contact",
      body:    `New contact from TKD Academy app:\n\nName: ${contactData.displayName}\nClub: ${contactData.club || 'N/A'}\nType: ${contactData.type}\n\nMessage:\n${contactData.message}\n\nFrom: ${username}\nTime: ${new Date().toLocaleString('en-GB')}`
    });
  } catch(e) { Logger.log("Email failed: " + e.message); }

  return { success: true };
}

function getClubOptions() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users").getDataRange().getValues();
  return [...new Set(
    data.slice(1).map(row => row[10] ? row[10].toString().trim() : '').filter(Boolean)
  )].sort();
}

function doPost(e) {
  const STRIPE_WEBHOOK_SECRET = PropertiesService.getScriptProperties().getProperty('STRIPE_WEBHOOK_SECRET');
  try {
    if (!STRIPE_WEBHOOK_SECRET) throw new Error('Webhook secret is not configured.');

    const payload   = e.postData.contents;
    const sigHeader = e.parameter['stripe-signature'] || (e.headers && e.headers['Stripe-Signature']) || '';

    if (STRIPE_WEBHOOK_SECRET && sigHeader) {
      const timestamp = sigHeader.split(',').find(p => p.startsWith('t='))?.split('=')[1];
      const sig       = sigHeader.split(',').find(p => p.startsWith('v1='))?.split('=')[1];

      if (timestamp && sig) {
        const computedSig = Utilities.computeHmacSha256Signature(`${timestamp}.${payload}`, STRIPE_WEBHOOK_SECRET);
        const computedHex = computedSig.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
        if (computedHex !== sig)
          return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid signature' })).setMimeType(ContentService.MimeType.JSON);
        if (Math.abs(Math.floor(Date.now() / 1000) - parseInt(timestamp)) > 300)
          return ContentService.createTextOutput(JSON.stringify({ error: 'Timestamp too old' })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    const parsed = JSON.parse(payload);
    if (parsed.type === 'checkout.session.completed' || parsed.type === 'invoice.payment_succeeded') {
      const email = parsed.data.object.customer_email || parsed.data.object.customer_details?.email;
      if (email) updateSubscriptionDate(email.toString().trim().toLowerCase());
    }

    return ContentService.createTextOutput(JSON.stringify({ received: true })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function updateSubscriptionDate(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === email) {
      sheet.getRange(i + 1, 10).setValue(new Date()); return;
    }
  }
}
