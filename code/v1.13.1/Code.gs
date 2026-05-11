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

function parseSheetDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  // Handle ISO string YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS
  const str = val.toString().trim();
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
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

      const registeredDate = parseSheetDate_(userData[i][8]);
      const subscriptionDate = parseSheetDate_(userData[i][9]);
      const TRIAL_DAYS = 7;
      const SUB_DAYS = 30;
      let accessStatus = 'active';
      let daysRemaining = null;

      if (registeredDate) {
        const daysSinceRegistered = (now - registeredDate) / (1000 * 60 * 60 * 24);
        const daysSinceSub = subscriptionDate
          ? (now - subscriptionDate) / (1000 * 60 * 60 * 24)
          : null;

        // Check subscription first — active sub overrides trial status
        if (subscriptionDate && daysSinceSub <= SUB_DAYS) {
          daysRemaining = Math.max(0, Math.ceil(SUB_DAYS - daysSinceSub));
          accessStatus = 'subscribed';
        } else if (daysSinceRegistered <= TRIAL_DAYS) {
          // Within trial and no active subscription
          daysRemaining = Math.max(0, Math.ceil(TRIAL_DAYS - daysSinceRegistered));
          accessStatus = 'trial';
        } else {
          // Trial expired — check subscription
          if (!subscriptionDate || daysSinceSub > SUB_DAYS) {
            accessStatus = subscriptionDate ? 'subscription_expired' : 'trial_expired';
          }
        }
      }
      
      return {
        success: true,
        username: userData[i][0].toString().trim(),
        displayName: userData[i][5] ? userData[i][5].toString() : userData[i][0].toString(),
        gradeValue: parseInt(userData[i][2]) || 1,
        streak: streak,
        isAdmin: userData[i][6] && userData[i][6].toString().trim().toUpperCase() === 'Y',
        accessStatus: accessStatus,
        daysRemaining: daysRemaining
      };

    }
  }
  return { success: false, message: "Invalid credentials" };
}

function registerUser(displayName, email, gradeValue, password, club) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const cleanEmail = email ? email.toString().trim().toLowerCase() : "";

  if (!cleanEmail || !displayName || !password || !gradeValue) {
    return { success: false, message: "All fields are required." };
  }

  // Check email not already registered
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === cleanEmail) {
      return { success: false, message: "An account with this email already exists." };
    }
  }

  const now = new Date();

  // Col A:Username(email) | B:Password | C:BeltLevel | D:LastActive | E:Streak | 
  // F:DisplayName | G:isAdmin | H:FCMToken | I:RegisteredDate | J:SubscriptionDate
  sheet.appendRow([
    cleanEmail,
    password,
    parseInt(gradeValue),
    now,
    1,
    displayName.toString().trim(),
    '',
    '',
    now,
    '',
    club ? club.toString().trim() : ''  // Col K — Club
  ]);
  return { success: true };
}

function getQuizData(username, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("Questions");
  const pSheet = ss.getSheetByName("UserProgress");
  const cleanUser = username ? username.toString().trim() : "";
  const userGradeLevel = getUserGrade_(username);
  const userClub = getUserClub_(username);

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

  // Base eligibility filter — shared by both modes
    const eligible = qData.slice(1).filter(row => {
    if (!row[0] || !row[14]) return false;
    if ((parseInt(row[17]) || 1) > userGradeLevel) return false;
    if (row[16] === "N") return false;
    // Col X (index 23) — club-restricted questions
    const questionClub = row[23] ? row[23].toString().trim() : '';
    if (questionClub && questionClub !== userClub) return false;
    return true;
  });

  // ── EXAM MODE ──────────────────────────────────────────────────────────────
  if (mode === 'test') {
    // Determine limit by grade
    let limit;
    if (userGradeLevel >= 11) limit = 50;
    else if (userGradeLevel >= 7) limit = 20;
    else limit = 10;

    // Category allocations as proportions (must sum to 1.0)
    const categories = [
      { name: 'Basics',   pct: 0.10 },
      { name: 'Numbers',  pct: 0.10 },
      { name: 'Belts',    pct: 0.10 },
      { name: 'Korean',   pct: 0.30 },
      { name: 'Patterns', pct: 0.30 }
      // remaining 10% filled from any category below
    ];

    // Weight questions by belt level — higher level = higher weight
    // Weight = beltLevel / sum of all beltLevels in pool
    function weightedSample(pool, n) {
      if (pool.length === 0) return [];
      const weights = pool.map(row => parseInt(row[17]) || 1);
      const totalWeight = weights.reduce((a, b) => a + b, 0);
      const selected = [];
      const used = new Set();
      let attempts = 0;
      while (selected.length < Math.min(n, pool.length) && attempts < pool.length * 10) {
        attempts++;
        let rand = Math.random() * totalWeight;
        for (let i = 0; i < pool.length; i++) {
          if (used.has(i)) continue;
          rand -= weights[i];
          if (rand <= 0) {
            selected.push(pool[i]);
            used.add(i);
            break;
          }
        }
      }
      return selected;
    }

    // Build category pools
    const pools = {};
    categories.forEach(c => pools[c.name] = []);
    const uncategorised = [];
    eligible.forEach(row => {
      const cat = row[20] ? row[20].toString().trim() : '';
      if (pools[cat] !== undefined) pools[cat].push(row);
      else uncategorised.push(row);
    });

    // Sample from each category
    let selected = [];
    categories.forEach(c => {
      const n = Math.round(c.pct * limit);
      selected = selected.concat(weightedSample(pools[c.name], n));
    });

    // Fill remaining slots (target 10% + any shortfall) from uncategorised or any pool
    const remaining = limit - selected.length;
    if (remaining > 0) {
      const selectedIds = new Set(selected.map(r => r[14].toString()));
      const fillPool = [...uncategorised, ...eligible].filter(r => !selectedIds.has(r[14].toString()));
      selected = selected.concat(weightedSample(fillPool, remaining));
    }

    // Shuffle final selection so category groups aren't obvious
    selected = selected.sort(() => Math.random() - 0.5);

    return selected.map(row => ({
      question: row[0].toString(),
      options: [row[4], row[5], row[6], row[7]].filter(String)
        .sort(() => Math.random() - 0.5).map(s => s.toString().trim()),
      answer: row[11] ? row[11].toString().trim() : "",
      qId: row[14].toString(),
      timeLimit: parseInt(row[21]) || 5
    }));
  }

  // ── PRACTICE MODE ──────────────────────────────────────────────────────────
  let filtered = eligible.filter(row => {
    const qId = row[14].toString();
    const prog = progressMap[qId];
    if (!prog) return true;
    const diffDays = (now - prog.date) / (1000 * 60 * 60 * 24);
    return diffDays >= (BUCKET_INTERVALS[prog.bucket] || 0);
  });

  // SRS fallback
  if (filtered.length === 0) filtered = [...eligible];

  // SRS sort: bucket ascending, most overdue first
  filtered.sort((a, b) => {
    const progA = progressMap[a[14]?.toString()] || { bucket: 1, date: new Date(0) };
    const progB = progressMap[b[14]?.toString()] || { bucket: 1, date: new Date(0) };
    if (progA.bucket !== progB.bucket) return progA.bucket - progB.bucket;
    return progA.date - progB.date;
  });

  return filtered.slice(0, 10).map(row => ({
    question: row[0].toString(),
    options: [row[4], row[5], row[6], row[7]].filter(String)
      .sort(() => Math.random() - 0.5).map(s => s.toString().trim()),
    answer: row[11] ? row[11].toString().trim() : "",
    qId: row[14].toString(),
    timeLimit: parseInt(row[21]) || 5
  }));
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
  const userClub = getUserClub_(username);

  const qData = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, qSheet.getLastColumn()).getValues();
  
  const filtered = qData.filter(row => {
    const isLevelMatch = (parseInt(row[17]) || 1) <= userGradeLevel;
    const isNotExcluded = row[16] !== "N";
    const questionClub = row[23] ? row[23].toString().trim() : '';
    const isClubMatch = !questionClub || questionClub === userClub;

    if (gameType === 'game_match') {
      return isLevelMatch && isNotExcluded && isClubMatch && row[18].toString().trim() !== "";
    } else {
      return isLevelMatch && isNotExcluded && isClubMatch && row[19].toString().trim() !== "";
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

function getUserClub_(username) {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Users").getDataRange().getValues();
  const clean = username ? username.toString().trim().toLowerCase() : "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === clean)
      return data[i][10] ? data[i][10].toString().trim() : '';
  }
  return '';
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

function getAdminData(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Verify requester is admin
  const callerGrade = getUserGrade_(username);
  const usersSheet = ss.getSheetByName("Users");
  const userData = usersSheet.getDataRange().getValues();
  const cleanCaller = username ? username.toString().trim().toLowerCase() : "";
  let isAdmin = false;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === cleanCaller) {
      isAdmin = (userData[i][6] && userData[i][6].toString().trim().toUpperCase() === 'Y');
      break;
    }
  }
  if (!isAdmin) return { error: "Unauthorised" };

  const progressSheet = ss.getSheetByName("UserProgress");
  const highScoreSheet = ss.getSheetByName("HighScores");
  const beltsSheet = ss.getSheetByName("Belts");
  const questionsSheet = ss.getSheetByName("Questions");

  const progressData = progressSheet ? progressSheet.getDataRange().getValues() : [];
  const highScoreData = highScoreSheet ? highScoreSheet.getDataRange().getValues() : [];
  const beltData = beltsSheet ? beltsSheet.getDataRange().getValues() : [];
  const questionData = questionsSheet ? questionsSheet.getDataRange().getValues() : [];

  // Belt label lookup
  const beltMap = {};
  beltData.slice(1).forEach(row => {
    if (row[0] && row[1]) beltMap[parseInt(row[1])] = row[0].toString().trim();
  });

  // Progress stats per user
  const progressMap = {};
  progressData.slice(1).forEach(row => {
    const u = row[0] ? row[0].toString().trim() : "";
    if (!u) return;
    if (!progressMap[u]) progressMap[u] = { total: 0, buckets: { 1:0, 2:0, 3:0, 4:0 } };
    progressMap[u].total++;
    const b = parseInt(row[3]) || 1;
    if (progressMap[u].buckets[b] !== undefined) progressMap[u].buckets[b]++;
  });

  // High score per user
  const highScoreMap = {};
  highScoreData.slice(1).forEach(row => {
    if (row[0]) highScoreMap[row[0].toString().trim()] = parseInt(row[1]) || 0;
  });

  const now = new Date();
  const users = userData.slice(1)
    .filter(row => row[0])
    .map(row => {
      const uName = row[0].toString().trim();
      const lastActive = row[3] ? new Date(row[3]) : null;
      const daysSince = lastActive ? Math.floor((now - lastActive) / (1000 * 60 * 60 * 24)) : null;
      const gradeVal = parseInt(row[2]) || 1;
      const prog = progressMap[uName] || { total: 0, buckets: { 1:0, 2:0, 3:0, 4:0 } };
      const eligible = questionData.slice(1).filter(q =>
        q[0] && q[14] && (parseInt(q[17]) || 1) <= gradeVal && q[16] !== "N"
      ).length;

      return {
        username: uName,
        displayName: row[5] ? row[5].toString() : uName,
        grade: gradeVal,
        gradeName: beltMap[gradeVal] || `Level ${gradeVal}`,
        streak: parseInt(row[4]) || 0,
        lastActive: lastActive ? lastActive.toLocaleString('en-GB') : 'Never',
        daysSince: daysSince,
        isAdmin: row[6] && row[6].toString().trim().toUpperCase() === 'Y',
        totalAnswered: prog.total,
        buckets: prog.buckets,
        highScore: highScoreMap[uName] || 0,
        eligible: eligible
      };
    });

  return { users };
}

function doPost(e) {
  // Retrieve the secret from Script Properties
  const STRIPE_WEBHOOK_SECRET = PropertiesService.getScriptProperties().getProperty('STRIPE_WEBHOOK_SECRET');

  try {
    if (!STRIPE_WEBHOOK_SECRET) {
      throw new Error('Webhook secret is not configured.');
    }
    
    const payload = e.postData.contents;
    const sigHeader = e.parameter['stripe-signature'] || 
                      (e.headers && e.headers['Stripe-Signature']) || '';

    // Stripe signature verification
    // Note: GAS doesn't support the full HMAC timing-safe comparison Stripe recommends,
    // but this provides a meaningful integrity check for low-risk subscription data.
    if (STRIPE_WEBHOOK_SECRET && sigHeader) {
      const timestamp = sigHeader.split(',')
        .find(p => p.startsWith('t='))?.split('=')[1];
      const sig = sigHeader.split(',')
        .find(p => p.startsWith('v1='))?.split('=')[1];

      if (timestamp && sig) {
        const signedPayload = `${timestamp}.${payload}`;
        const computedSig = Utilities.computeHmacSha256Signature(
          signedPayload, STRIPE_WEBHOOK_SECRET
        );
        const computedHex = computedSig.map(b => 
          ('0' + (b & 0xFF).toString(16)).slice(-2)
        ).join('');

        if (computedHex !== sig) {
          return ContentService.createTextOutput(
            JSON.stringify({ error: 'Invalid signature' })
          ).setMimeType(ContentService.MimeType.JSON);
        }

        // Reject requests older than 5 minutes
        const fiveMinutes = 5 * 60;
        const nowSeconds = Math.floor(Date.now() / 1000);
        if (Math.abs(nowSeconds - parseInt(timestamp)) > fiveMinutes) {
          return ContentService.createTextOutput(
            JSON.stringify({ error: 'Timestamp too old' })
          ).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    const parsed = JSON.parse(payload);

    if (parsed.type === 'checkout.session.completed' ||
        parsed.type === 'invoice.payment_succeeded') {
      const email = parsed.data.object.customer_email ||
                    parsed.data.object.customer_details?.email;
      if (email) {
        updateSubscriptionDate(email.toString().trim().toLowerCase());
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ received: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function updateSubscriptionDate(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === email) {
      sheet.getRange(i + 1, 10).setValue(now); // Col J = SubscriptionDate
      return;
    }
  }
}

function getAdminPageData(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName("Users");
  const userData = usersSheet.getDataRange().getValues();
  const cleanCaller = username ? username.toString().trim().toLowerCase() : "";

  // Find caller and verify admin
  let callerRow = null;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === cleanCaller) {
      callerRow = userData[i];
      break;
    }
  }
  if (!callerRow || callerRow[6].toString().trim().toUpperCase() !== 'Y') {
    return { error: "Unauthorised" };
  }

  const callerClub = callerRow[10] ? callerRow[10].toString().trim() : '';

  const progressSheet = ss.getSheetByName("UserProgress");
  const highScoreSheet = ss.getSheetByName("HighScores");
  const beltsSheet = ss.getSheetByName("Belts");
  const questionsSheet = ss.getSheetByName("Questions");

  const progressData = progressSheet ? progressSheet.getDataRange().getValues() : [];
  const highScoreData = highScoreSheet ? highScoreSheet.getDataRange().getValues() : [];
  const beltData = beltsSheet ? beltsSheet.getDataRange().getValues() : [];
  const questionData = questionsSheet ? questionsSheet.getDataRange().getValues() : [];

  const beltMap = {};
  beltData.slice(1).forEach(row => {
    if (row[0] && row[1]) beltMap[parseInt(row[1])] = row[0].toString().trim();
  });

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

  const now = new Date();

  const users = userData.slice(1)
    .filter(row => {
      if (!row[0]) return false;
      // Club filter — if admin has a club, only show same club members
      if (callerClub) {
        const memberClub = row[10] ? row[10].toString().trim() : '';
        return memberClub === callerClub;
      }
      return true; // Super admin (no club) sees all
    })
    .map(row => {
      const uName = row[0].toString().trim();
      const lastActive = row[3] ? new Date(row[3]) : null;
      const daysSince = lastActive ? Math.floor((now - lastActive) / (1000 * 60 * 60 * 24)) : null;
      const gradeVal = parseInt(row[2]) || 1;
      const prog = progressMap[uName] || { total: 0, buckets: { 1:0, 2:0, 3:0, 4:0 } };
      const eligible = questionData.slice(1).filter(q =>
        q[0] && q[14] && (parseInt(q[17]) || 1) <= gradeVal && q[16] !== "N"
      ).length;

      const registeredDate = parseSheetDate_(row[8]);
      const subscriptionDate = parseSheetDate_(row[9]);
      const TRIAL_DAYS = 7;
      const SUB_DAYS = 30;
      let subStatus = 'active';
      let subEndDate = null;

      if (registeredDate) {
        const daysSinceReg = (now - registeredDate) / (1000 * 60 * 60 * 24);
        const daysSinceSub = subscriptionDate ? (now - subscriptionDate) / (1000 * 60 * 60 * 24) : null;
        if (subscriptionDate && daysSinceSub <= SUB_DAYS) {
          subStatus = 'subscribed';
          const endDate = new Date(subscriptionDate);
          endDate.setDate(endDate.getDate() + SUB_DAYS);
          subEndDate = endDate.toLocaleDateString('en-GB');
        } else if (daysSinceReg <= TRIAL_DAYS) {
          subStatus = 'trial';
          const endDate = new Date(registeredDate);
          endDate.setDate(endDate.getDate() + TRIAL_DAYS);
          subEndDate = endDate.toLocaleDateString('en-GB');
        } else {
          subStatus = subscriptionDate ? 'sub_expired' : 'trial_expired';
        }
      }

      return {
        username: uName,
        displayName: row[5] ? row[5].toString() : uName,
        grade: gradeVal,
        gradeName: beltMap[gradeVal] || `Level ${gradeVal}`,
        streak: parseInt(row[4]) || 0,
        lastActive: lastActive ? lastActive.toLocaleString('en-GB') : 'Never',
        daysSince: daysSince,
        isAdmin: row[6] && row[6].toString().trim().toUpperCase() === 'Y',
        club: row[10] ? row[10].toString().trim() : '',
        totalAnswered: prog.total,
        buckets: prog.buckets,
        highScore: highScoreMap[uName] || 0,
        eligible: eligible,
        subStatus: subStatus,
        subEndDate: subEndDate
      };
    });

  // Belt and category options for question submission
  const belts = beltData.slice(1)
    .filter(row => row[0] && row[1])
    .map(row => ({ label: row[0].toString().trim(), value: parseInt(row[1]) }));

  const categories = [...new Set(
    questionData.slice(1)
      .map(row => row[20] ? row[20].toString().trim() : '')
      .filter(Boolean)
  )].sort();

  return { users, belts, categories, callerClub };
}

function submitQuestion(username, questionData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName("Users");
  const userData = usersSheet.getDataRange().getValues();
  const cleanUser = username ? username.toString().trim().toLowerCase() : "";

  // Verify admin
  let isAdmin = false;
  let callerClub = '';
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === cleanUser) {
      isAdmin = userData[i][6] && userData[i][6].toString().trim().toUpperCase() === 'Y';
      callerClub = userData[i][10] ? userData[i][10].toString().trim() : '';
      break;
    }
  }
  if (!isAdmin) return { success: false, message: "Unauthorised" };

  const qSheet = ss.getSheetByName("Questions");
  const lastRow = qSheet.getLastRow();

  // Generate a unique qId
  const qId = 'q_' + Date.now();

  // Build the row — columns A through X (24 cols)
  // A:Question | E:Opt1 | F:Opt2 | L:CorrectAnswer | O:qId | Q:ExamFlag | R:BeltLevel | U:Category | X:Club
  const newRow = new Array(24).fill('');
  newRow[0]  = questionData.question;          // Col A
  newRow[4]  = questionData.opt1;              // Col E
  newRow[5]  = questionData.opt2;              // Col F
  newRow[11] = questionData.correctAnswer;     // Col L
  newRow[14] = qId;                            // Col O
  newRow[16] = 'Y';                            // Col Q — exam eligible by default
  newRow[17] = parseInt(questionData.beltLevel); // Col R
  newRow[20] = questionData.category;          // Col U
  newRow[23] = questionData.clubOnly ? callerClub : ''; // Col X

  qSheet.appendRow(newRow);
  return { success: true, qId: qId };
}

function submitContact(username, contactData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Write to Contact sheet — auto-create if missing
  let sheet = ss.getSheetByName("Contact");
  if (!sheet) {
    sheet = ss.insertSheet("Contact");
    sheet.appendRow(["Timestamp","Username","DisplayName","Club","Type","Message"]);
  }

  const now = new Date();
  sheet.appendRow([
    now,
    username,
    contactData.displayName,
    contactData.club,
    contactData.type,
    contactData.message
  ]);

  // Send email notification
  try {
    MailApp.sendEmail({
      to: "jonathan_gallucci@hotmail.com",
      subject: "TKD Academy contact",
      body: `New contact from TKD Academy app:\n\nName: ${contactData.displayName}\nClub: ${contactData.club || 'N/A'}\nType: ${contactData.type}\n\nMessage:\n${contactData.message}\n\nFrom: ${username}\nTime: ${now.toLocaleString('en-GB')}`
    });
  } catch(e) {
    // Email failure shouldn't block the submission
    Logger.log("Email failed: " + e.message);
  }

  return { success: true };
}

function getAdminFormOptions(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const beltData = ss.getSheetByName("Belts").getDataRange().getValues();
  const qData = ss.getSheetByName("Questions").getDataRange().getValues();

  const belts = beltData.slice(1)
    .filter(row => row[0] && row[1])
    .map(row => ({ label: row[0].toString().trim(), value: parseInt(row[1]) }));

  const categories = [...new Set(
    qData.slice(1)
      .map(row => row[20] ? row[20].toString().trim() : '')
      .filter(Boolean)
  )].sort();

  // Get caller's club
  const userData = ss.getSheetByName("Users").getDataRange().getValues();
  const clean = username ? username.toString().trim().toLowerCase() : "";
  let callerClub = '';
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === clean) {
      callerClub = userData[i][10] ? userData[i][10].toString().trim() : '';
      break;
    }
  }

  return { belts, categories, callerClub };
}

function getClubOptions() {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Users").getDataRange().getValues();
  const clubs = [...new Set(
    data.slice(1)
      .map(row => row[10] ? row[10].toString().trim() : '')
      .filter(Boolean)
  )].sort();
  return clubs;
}
