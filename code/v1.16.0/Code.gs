// ── Constants ─────────────────────────────────────────────────────────────────
const BUCKET_INTERVALS = { 1: 0, 2: 2, 3: 4, 4: 5 };
const TRIAL_DAYS       = 7;
const TRIAL_GRACE_DAYS = 7;   // extra days after trial_expired before hard lock
const SUB_DAYS         = 30;

const COLS = {
  USERS:         22,  // A–V
  QUESTIONS:     24,  // A–X
  BELTS:          2,  // A–B
  FLASHCARDS:    11,  // A–K
  TULS:           8,  // A–H
  PROGRESS:       5,  // A–E
  HIGHSCORES:     2,  // A–B
  SIMPLEDEFS:     6,  // A–F
  NOTIFICATIONS:  6   // A–F
};

// Users sheet column indices (0-based) and 1-based positions
// A=0  email
// B=1  password
// C=2  beltLevel
// D=3  lastActive
// E=4  streak
// F=5  displayName
// G=6  isAdmin
// H=7  FCMToken
// I=8  registeredDate
// J=9  subscriptionDate
// K=10 club
// L=11 learningGoalMins
// M=12 dailyPlayLog (JSON)
// N=13 monthlyPlayLog (JSON)
// O=14 (unused)
// P=15 streakDate
// Q=16 clubStatus  (member | pending | pending_other | '')
// R=17 gradingDate
// S=18 newClubDetails
// T=19 pendingClub
// U=20 subscriptionDaysStored
// V=21 pendingClubRequestDate

const CACHE_TTL = {
  QUESTIONS:  21600,
  BELTS:      21600,
  FLASHCARDS: 21600,
  TULS:       21600,
  SIMPLEDEFS: 21600,
  CLUBS:       3600
};

// ── Spreadsheet helpers ───────────────────────────────────────────────────────

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetValues_(sheetName, numCols) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  return sheet.getRange(1, 1, lastRow, numCols).getValues();
}

function parseSheetDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const d = new Date(val.toString().trim());
  return isNaN(d.getTime()) ? null : d;
}

function formatDateYMD_(val) {
  if (!val) return '';
  const d = val instanceof Date ? val : parseSheetDate_(val);
  if (!d) return '';
  try {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch(e) {
    return d.toISOString().split('T')[0];
  }
}

// ── CacheService layer ────────────────────────────────────────────────────────

function getCached_(key) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch(e) { return null; }
}

function putCached_(key, data, ttl) {
  try {
    const cache = CacheService.getScriptCache();
    const json  = JSON.stringify(data);
    if (json.length > 98000) {
      const chunkSize = 90000;
      const chunks    = Math.ceil(json.length / chunkSize);
      cache.put(key + '_meta', JSON.stringify({ chunks }), ttl);
      for (let c = 0; c < chunks; c++) {
        cache.put(key + '_c' + c, json.slice(c * chunkSize, (c + 1) * chunkSize), ttl);
      }
    } else {
      cache.put(key, json, ttl);
    }
  } catch(e) {}
}

function getCachedChunked_(key) {
  try {
    const cache   = CacheService.getScriptCache();
    const metaRaw = cache.get(key + '_meta');
    if (!metaRaw) return getCached_(key);
    const { chunks } = JSON.parse(metaRaw);
    const keys = Array.from({ length: chunks }, (_, i) => key + '_c' + i);
    const parts = cache.getAll(keys);
    const json  = keys.map(k => parts[k] || '').join('');
    return json ? JSON.parse(json) : null;
  } catch(e) { return null; }
}

function bustCache_(key) {
  try {
    const cache   = CacheService.getScriptCache();
    const metaRaw = cache.get(key + '_meta');
    if (metaRaw) {
      const { chunks } = JSON.parse(metaRaw);
      const keys = [key + '_meta', ...Array.from({ length: chunks }, (_, i) => key + '_c' + i)];
      cache.removeAll(keys);
    } else {
      cache.remove(key);
    }
  } catch(e) {}
}

// ── Cached data accessors ─────────────────────────────────────────────────────

function getQuestionsParsed_() {
  const KEY    = 'questions_v2';
  const cached = getCachedChunked_(KEY);
  if (cached) return cached;
  const rows = getSheetValues_('Questions', COLS.QUESTIONS);
  const data = rows.slice(1)
    .filter(r => r[0] && r[14])
    .map(r => ({
      q:    r[0]  ? r[0].toString()  : '',
      a:    r[11] ? r[11].toString() : '',
      o1:   r[4]  ? r[4].toString()  : '',
      o2:   r[5]  ? r[5].toString()  : '',
      o3:   r[6]  ? r[6].toString()  : '',
      o4:   r[7]  ? r[7].toString()  : '',
      id:   r[14] ? r[14].toString() : '',
      flag: r[16] ? r[16].toString() : '',
      lv:   parseInt(r[17]) || 1,
      tl:   parseInt(r[21]) || 5,
      mt:   r[18] ? r[18].toString() : '',
      sd:   r[19] ? r[19].toString() : '',
      cat:  r[20] ? r[20].toString() : '',
      club: r[23] ? r[23].toString() : ''
    }));
  putCached_(KEY, data, CACHE_TTL.QUESTIONS);
  return data;
}

function getBeltsParsed_() {
  const KEY    = 'belts_v2';
  const cached = getCached_(KEY);
  if (cached) return cached;
  const rows = getSheetValues_('Belts', COLS.BELTS);
  const data = rows.slice(1)
    .filter(r => r[0] && r[1])
    .map(r => ({ label: r[0].toString().trim(), value: parseInt(r[1]) }));
  putCached_(KEY, data, CACHE_TTL.BELTS);
  return data;
}

function getFlashcardsParsed_() {
  const KEY    = 'flashcards_v2';
  const cached = getCachedChunked_(KEY);
  if (cached) return cached;
  const rows = getSheetValues_('Flashcards', COLS.FLASHCARDS);
  const data = rows.slice(1)
    .filter(r => (r[2] || r[3]) && parseInt(r[0]))
    .map(r => {
      const card = {
        lv:  parseInt(r[0]) || 0,
        cat: r[1]  ? r[1].toString().trim()  : '',
        t:   r[2]  ? r[2].toString().trim()  : '',
        m:   r[3]  ? r[3].toString().trim()  : ''
      };
      if (r[1] && r[1].toString().trim() === 'Patterns') {
        if (r[4])  card.mn  = r[4].toString().trim();
        if (r[5])  card.mk  = r[5].toString().trim();
        if (r[6])  card.sn  = r[6].toString().trim();
        if (r[7])  card.sk  = r[7].toString().trim();
        if (r[8])  card.re  = r[8].toString().trim();
        if (r[9])  card.rk  = r[9].toString().trim();
        if (r[10]) card.img = r[10].toString().trim();
      }
      return card;
    });
  putCached_(KEY, data, CACHE_TTL.FLASHCARDS);
  return data;
}

function getTulsParsed_() {
  const KEY    = 'tuls_v2';
  const cached = getCached_(KEY);
  if (cached) return cached;
  const rows = getSheetValues_('Tuls', COLS.TULS);
  const data = rows.slice(1)
    .filter(r => r[0] && !isNaN(Number(r[5])))
    .map(r => ({
      name: r[0].toString(),
      mov:  parseInt(r[1]) || 0,
      sta:  parseInt(r[2]) || 0,
      rs:   r[3] ? r[3].toString() : '---',
      dif:  parseInt(r[4]) || 0,
      minGrade: Number(r[5]),
      meaning: r[6] ? r[6].toString() : '',
      img:  r[7] ? r[7].toString() : 'https://placehold.co/300x200?text=No+Pattern+Image'
    }));
  putCached_(KEY, data, CACHE_TTL.TULS);
  return data;
}

function getSimpleDefsParsed_() {
  const KEY    = 'simpledefs_v1';
  const cached = getCachedChunked_(KEY);
  if (cached) return cached;
  const rows = getSheetValues_('SimpleDefinition', COLS.SIMPLEDEFS);
  const data = rows.slice(1)
    .filter(r => r[0] && r[4])
    .map(r => ({
      q:   r[0] ? r[0].toString().trim() : '',
      a:   r[1] ? r[1].toString().trim() : '',
      w:   r[2] ? r[2].toString().trim() : '',
      lv:  parseInt(r[3]) || 1,
      id:  r[4] ? r[4].toString().trim() : '',
      cat: r[5] ? r[5].toString().trim() : ''
    }));
  putCached_(KEY, data, CACHE_TTL.SIMPLEDEFS);
  return data;
}

function getClubList_() {
  const KEY    = 'club_list';
  const cached = getCached_(KEY);
  if (cached) return cached;
  const data  = getSheetValues_('Users', COLS.USERS);
  const clubs = [...new Set(
    data.slice(1).map(row => row[10] ? row[10].toString().trim() : '').filter(Boolean)
  )].sort();
  putCached_(KEY, clubs, CACHE_TTL.CLUBS);
  return clubs;
}

// ── Derived helpers ───────────────────────────────────────────────────────────

function getBeltMap_() {
  const map = {};
  getBeltsParsed_().forEach(b => { map[b.value] = b.label; });
  return map;
}

function getBeltOptions_() {
  return getBeltsParsed_();
}

function getUserContext_(username) {
  const data  = getSheetValues_('Users', COLS.USERS);
  const clean = username ? username.toString().trim().toLowerCase() : '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === clean)
      return {
        grade:      parseInt(data[i][2]) || 1,
        club:       data[i][10] ? data[i][10].toString().trim() : '',
        clubStatus: data[i][16] ? data[i][16].toString().trim() : ''
      };
  }
  return { grade: 1, club: '', clubStatus: '' };
}

function getUserGrade_(username) {
  return getUserContext_(username).grade;
}

// ── Notifications ─────────────────────────────────────────────────────────────
// Notifications sheet columns (0-based):
// A=0 toUsername  B=1 message  C=2 timestamp  D=3 isRead  E=4 type  F=5 fromClub

function writeNotification_(toUsername, message, type, fromClub) {
  const ss    = getSpreadsheet_();
  let sheet   = ss.getSheetByName('Notifications');
  if (!sheet) {
    sheet = ss.insertSheet('Notifications');
    sheet.appendRow(['toUsername', 'message', 'timestamp', 'isRead', 'type', 'fromClub']);
  }
  sheet.appendRow([
    toUsername.toString().trim().toLowerCase(),
    message,
    new Date(),
    'N',
    type || 'general',
    fromClub || ''
  ]);
}

function getNotifications(username) {
  const clean = username ? username.toString().trim().toLowerCase() : '';
  const sheet = getSpreadsheet_().getSheetByName('Notifications');
  if (!sheet || sheet.getLastRow() < 2) return { notifications: [], unreadCount: 0 };

  const data = sheet.getRange(1, 1, sheet.getLastRow(), COLS.NOTIFICATIONS).getValues();
  const notifications = [];
  let unreadCount = 0;

  data.slice(1).forEach((row, idx) => {
    if (!row[0] || row[0].toString().trim().toLowerCase() !== clean) return;
    const isRead = row[3] && row[3].toString().trim().toUpperCase() === 'Y';
    if (!isRead) unreadCount++;
    notifications.push({
      id:        idx + 2, // 1-based sheet row
      message:   row[1] ? row[1].toString() : '',
      timestamp: row[2] ? new Date(row[2]).toLocaleString('en-GB') : '',
      isRead,
      type:      row[4] ? row[4].toString() : 'general',
      fromClub:  row[5] ? row[5].toString() : ''
    });
  });

  // Return most recent first
  notifications.reverse();
  return { notifications, unreadCount };
}

function markNotificationsRead(username) {
  const clean = username ? username.toString().trim().toLowerCase() : '';
  const sheet = getSpreadsheet_().getSheetByName('Notifications');
  if (!sheet || sheet.getLastRow() < 2) return { success: true };

  const data = sheet.getRange(1, 1, sheet.getLastRow(), COLS.NOTIFICATIONS).getValues();
  data.slice(1).forEach((row, idx) => {
    if (!row[0] || row[0].toString().trim().toLowerCase() !== clean) return;
    if (row[3] && row[3].toString().trim().toUpperCase() !== 'Y') {
      sheet.getRange(idx + 2, 4).setValue('Y'); // Col D, 1-based
    }
  });
  return { success: true };
}

// ── doGet ─────────────────────────────────────────────────────────────────────

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('TKD Theory Academy')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ── loginUser ─────────────────────────────────────────────────────────────────

function loginUser(username, password) {
  const ss        = getSpreadsheet_();
  const userSheet = ss.getSheetByName('Users');
  const lastRow   = userSheet.getLastRow();
  const userData  = lastRow > 1
    ? userSheet.getRange(1, 1, lastRow, COLS.USERS).getValues()
    : [];
  const now = new Date();

  const cleanU = username ? username.toString().trim().toLowerCase() : '';
  const cleanP = password ? password.toString().trim() : '';

  for (let i = 1; i < userData.length; i++) {
    if (!userData[i][0]) continue;
    if (userData[i][0].toString().trim().toLowerCase() !== cleanU) continue;
    if (userData[i][1].toString() !== cleanP) continue;

    // ── Streak reset check ──────────────────────────────────────────────────
    let streak           = parseInt(userData[i][4])  || 0;
    const streakDateRaw  = userData[i][15];
    const lastStreakDate = streakDateRaw ? new Date(streakDateRaw) : null;

    if (lastStreakDate) {
      const daysSinceStreak = Math.floor((now - lastStreakDate) / (1000 * 60 * 60 * 24));
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

    // ── Club status ─────────────────────────────────────────────────────────
    const clubStatus  = userData[i][16] ? userData[i][16].toString().trim() : '';
    const pendingClub = userData[i][19] ? userData[i][19].toString().trim() : '';

    // ── Access / subscription status ────────────────────────────────────────
    const registeredDate   = parseSheetDate_(userData[i][8]);
    const subscriptionDate = parseSheetDate_(userData[i][9]);
    const gradeValue       = parseInt(userData[i][2]) || 1;

    // Club members are always fully active — no subscription gate
    if (clubStatus === 'member') {
      const { notifications, unreadCount } = getNotifications(cleanU);
      return {
        success:          true,
        username:         userData[i][0].toString().trim(),
        displayName:      userData[i][5] ? userData[i][5].toString() : userData[i][0].toString(),
        gradeValue,
        streak,
        isAdmin:          userData[i][6] && userData[i][6].toString().trim().toUpperCase() === 'Y',
        accessStatus:     'active',
        daysRemaining:    null,
        learningGoalMins: parseInt(userData[i][11]) || 0,
        clubStatus,
        pendingClub,
        gradingDate:      formatDateYMD_(userData[i][17]),
        unreadCount,
        bootstrap: {
          belts: getBeltOptions_(),
          clubs: getClubList_()
        }
      };
    }

    let accessStatus  = 'active';
    let daysRemaining = null;
    let graceDaysRemaining = null;

    if (registeredDate) {
      const daysSinceRegistered = (now - registeredDate) / (1000 * 60 * 60 * 24);
      const daysSinceSub        = subscriptionDate
        ? (now - subscriptionDate) / (1000 * 60 * 60 * 24) : null;

      if (subscriptionDate && daysSinceSub <= SUB_DAYS) {
        daysRemaining = Math.max(0, Math.ceil(SUB_DAYS - daysSinceSub));
        accessStatus  = 'subscribed';
      } else if (daysSinceRegistered <= TRIAL_DAYS) {
        daysRemaining = Math.max(0, Math.ceil(TRIAL_DAYS - daysSinceRegistered));
        accessStatus  = 'trial';
      } else if (daysSinceRegistered <= TRIAL_DAYS + TRIAL_GRACE_DAYS) {
        // Within grace period — allow login with warning
        graceDaysRemaining = Math.max(0, Math.ceil(TRIAL_DAYS + TRIAL_GRACE_DAYS - daysSinceRegistered));
        accessStatus = 'trial_expired';
      } else {
        // Beyond grace period — hard lock
        accessStatus = subscriptionDate ? 'subscription_expired' : 'trial_hard_locked';
      }
    }

    // Hard lock: return minimal payload, no bootstrap
    if (accessStatus === 'trial_hard_locked') {
      return {
        success:      true,
        accessStatus: 'trial_hard_locked',
        username:     userData[i][0].toString().trim()
      };
    }

    const { notifications, unreadCount } = getNotifications(cleanU);

    return {
      success:             true,
      username:            userData[i][0].toString().trim(),
      displayName:         userData[i][5] ? userData[i][5].toString() : userData[i][0].toString(),
      gradeValue,
      streak,
      isAdmin:             userData[i][6] && userData[i][6].toString().trim().toUpperCase() === 'Y',
      accessStatus,
      daysRemaining,
      graceDaysRemaining,
      learningGoalMins:    parseInt(userData[i][11]) || 0,
      clubStatus,
      pendingClub,
      gradingDate:         formatDateYMD_(userData[i][17]),
      unreadCount,
      bootstrap: {
        belts: getBeltOptions_(),
        clubs: getClubList_()
      }
    };
  }
  return { success: false };
}

// ── Private helpers ───────────────────────────────────────────────────────────

function getSrsStatsForUser_(cleanUser, userGradeLevel) {
  const ss     = getSpreadsheet_();
  const pSheet = ss.getSheetByName('UserProgress');
  const pLastRow = pSheet.getLastRow();
  const pData    = pLastRow > 1
    ? pSheet.getRange(1, 1, pLastRow, COLS.PROGRESS).getValues()
    : [];

  const eligible = getQuestionsParsed_().filter(q =>
    q.id && q.q && q.lv <= userGradeLevel
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

function getHighScoreForUser_(cleanUser) {
  const ss      = getSpreadsheet_();
  const sheet   = ss.getSheetByName('HighScores');
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const data = sheet.getRange(1, 1, lastRow, COLS.HIGHSCORES).getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === cleanUser) return parseInt(data[i][1]) || 0;
  }
  return 0;
}

// ── registerUser ──────────────────────────────────────────────────────────────

function registerUser(displayName, email, gradeValue, password, club, newClubDetails) {
  const sheet      = getSpreadsheet_().getSheetByName('Users');
  const lastRow    = sheet.getLastRow();
  const data       = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  const cleanEmail = email ? email.toString().trim().toLowerCase() : '';

  if (!cleanEmail || !displayName || !password || !gradeValue)
    return { success: false, message: 'All fields are required.' };

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === cleanEmail)
      return { success: false, message: 'An account with this email already exists.' };
  }

  const now        = new Date();
  const cleanClub  = club ? club.toString().trim() : '';
  const isOther    = cleanClub.toUpperCase() === 'OTHER';
  const clubToStore = isOther ? '' : cleanClub;

  // Determine initial clubStatus
  let clubStatus = '';
  let pendingClub = '';
  if (!isOther && cleanClub) {
    clubStatus  = 'pending';
    pendingClub = cleanClub;
  } else if (isOther) {
    clubStatus = 'pending_other';
  }

  // Build the row — 22 columns (A–V)
  const row = new Array(COLS.USERS).fill('');
  row[0]  = cleanEmail;
  row[1]  = password;
  row[2]  = parseInt(gradeValue);
  row[3]  = now;
  row[4]  = 1;
  row[5]  = displayName.toString().trim();
  row[6]  = '';
  row[7]  = '';
  row[8]  = now;
  row[9]  = '';
  row[10] = clubToStore;
  row[11] = 0;
  row[16] = clubStatus;
  row[17] = '';
  row[18] = isOther ? (newClubDetails ? newClubDetails.toString().trim() : '') : '';
  row[19] = pendingClub;
  row[20] = '';
  row[21] = pendingClub ? now : '';

  sheet.appendRow(row);
  bustCache_('club_list');

  // Notify the club admin if a licensed club was nominated
  if (clubStatus === 'pending' && pendingClub) {
    notifyClubAdmins_(pendingClub,
      `New member request: ${displayName.toString().trim()} has requested to join your club.`,
      'club_request'
    );
  }

  return { success: true };
}

// Notify all admins belonging to a given club
function notifyClubAdmins_(clubName, message, type) {
  const data = getSheetValues_('Users', COLS.USERS);
  data.slice(1).forEach(row => {
    const isAdmin    = row[6] && row[6].toString().trim().toUpperCase() === 'Y';
    const adminClub  = row[10] ? row[10].toString().trim() : '';
    const adminEmail = row[0] ? row[0].toString().trim().toLowerCase() : '';
    if (isAdmin && adminClub === clubName && adminEmail) {
      writeNotification_(adminEmail, message, type, clubName);
    }
  });
}

// ── Club join request (existing user) ─────────────────────────────────────────

function requestClubJoin(username, clubName, newClubDetails) {
  const sheet   = getSpreadsheet_().getSheetByName('Users');
  const lastRow = sheet.getLastRow();
  const data    = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  const clean   = username ? username.toString().trim().toLowerCase() : '';
  const now     = new Date();

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().trim().toLowerCase() !== clean) continue;

    const currentStatus = data[i][16] ? data[i][16].toString().trim() : '';

    // Block if already pending
    if (currentStatus === 'pending' || currentStatus === 'pending_other')
      return { success: false, message: 'You already have a pending club request.' };

    // Block if already a member
    if (currentStatus === 'member')
      return { success: false, message: 'You are already linked to a club.' };

    // Rate limit: 1 request per 7 days
    const lastRequestDate = data[i][21] ? new Date(data[i][21]) : null;
    if (lastRequestDate) {
      const daysSince = (now - lastRequestDate) / (1000 * 60 * 60 * 24);
      if (daysSince < 7)
        return { success: false, message: 'You can only submit one club request per 7 days. Please try again later.' };
    }

    const cleanClub = clubName ? clubName.toString().trim() : '';
    const isOther   = cleanClub.toUpperCase() === 'OTHER';

    const row = i + 1;
    if (isOther) {
      sheet.getRange(row, 17).setValue('pending_other');                                             // Col Q
      sheet.getRange(row, 19).setValue(newClubDetails ? newClubDetails.toString().trim() : '');     // Col S
      sheet.getRange(row, 22).setValue(now);                                                        // Col V
    } else {
      sheet.getRange(row, 17).setValue('pending');                                                  // Col Q
      sheet.getRange(row, 20).setValue(cleanClub);                                                  // Col T
      sheet.getRange(row, 22).setValue(now);                                                        // Col V
      // Notify club admins
      const displayName = data[i][5] ? data[i][5].toString() : clean;
      notifyClubAdmins_(cleanClub,
        `New member request: ${displayName} has requested to join your club.`,
        'club_request'
      );
    }

    return { success: true };
  }
  return { success: false, message: 'User not found.' };
}

// ── Club request approval / decline ──────────────────────────────────────────

function approveClubRequest(adminUsername, targetUsername) {
  const sheet   = getSpreadsheet_().getSheetByName('Users');
  const lastRow = sheet.getLastRow();
  const data    = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  const cleanAdmin  = adminUsername  ? adminUsername.toString().trim().toLowerCase()  : '';
  const cleanTarget = targetUsername ? targetUsername.toString().trim().toLowerCase() : '';

  // Verify caller is admin
  let callerIsAdmin = false, callerClub = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === cleanAdmin) {
      callerIsAdmin = data[i][6] && data[i][6].toString().trim().toUpperCase() === 'Y';
      callerClub    = data[i][10] ? data[i][10].toString().trim() : '';
      break;
    }
  }
  if (!callerIsAdmin) return { success: false, message: 'Unauthorised' };

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().trim().toLowerCase() !== cleanTarget) continue;

    const pendingClub = data[i][19] ? data[i][19].toString().trim() : '';

    // Club-scoped admin can only approve users for their own club
    if (callerClub && pendingClub !== callerClub)
      return { success: false, message: 'Unauthorised' };

    const row = i + 1;
    const clubToJoin = pendingClub || callerClub;

    // Store remaining subscription days if user is currently subscribed
    const subscriptionDate = parseSheetDate_(data[i][9]);
    if (subscriptionDate) {
      const now          = new Date();
      const daysSinceSub = (now - subscriptionDate) / (1000 * 60 * 60 * 24);
      const daysRemaining = Math.max(0, Math.ceil(SUB_DAYS - daysSinceSub));
      if (daysRemaining > 0) {
        sheet.getRange(row, 21).setValue(daysRemaining); // Col U
      }
    }

    sheet.getRange(row, 11).setValue(clubToJoin);  // Col K — set actual club
    sheet.getRange(row, 17).setValue('member');     // Col Q — clubStatus
    sheet.getRange(row, 20).setValue('');           // Col T — clear pendingClub

    const displayName = data[i][5] ? data[i][5].toString() : cleanTarget;
    writeNotification_(cleanTarget,
      `Your application to link to ${clubToJoin} has been approved.`,
      'club_approved',
      clubToJoin
    );

    bustCache_('club_list');
    return { success: true };
  }
  return { success: false, message: 'User not found.' };
}

function declineClubRequest(adminUsername, targetUsername) {
  const sheet   = getSpreadsheet_().getSheetByName('Users');
  const lastRow = sheet.getLastRow();
  const data    = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  const cleanAdmin  = adminUsername  ? adminUsername.toString().trim().toLowerCase()  : '';
  const cleanTarget = targetUsername ? targetUsername.toString().trim().toLowerCase() : '';

  let callerIsAdmin = false, callerClub = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === cleanAdmin) {
      callerIsAdmin = data[i][6] && data[i][6].toString().trim().toUpperCase() === 'Y';
      callerClub    = data[i][10] ? data[i][10].toString().trim() : '';
      break;
    }
  }
  if (!callerIsAdmin) return { success: false, message: 'Unauthorised' };

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().trim().toLowerCase() !== cleanTarget) continue;

    const pendingClub = data[i][19] ? data[i][19].toString().trim() : '';
    if (callerClub && pendingClub !== callerClub)
      return { success: false, message: 'Unauthorised' };

    const row         = i + 1;
    const declinedClub = pendingClub || callerClub;

    sheet.getRange(row, 17).setValue('');  // Col Q — clear clubStatus
    sheet.getRange(row, 20).setValue('');  // Col T — clear pendingClub

    writeNotification_(cleanTarget,
      `Sorry, your application to link to ${declinedClub} has not been approved. Please speak to the club administrator.`,
      'club_declined',
      declinedClub
    );

    return { success: true };
  }
  return { success: false, message: 'User not found.' };
}

// ── getClubRequests — pending requests for admin action centre ────────────────

function getClubRequests(adminUsername) {
  const data       = getSheetValues_('Users', COLS.USERS);
  const cleanAdmin = adminUsername ? adminUsername.toString().trim().toLowerCase() : '';

  let callerIsAdmin = false, callerClub = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === cleanAdmin) {
      callerIsAdmin = data[i][6] && data[i][6].toString().trim().toUpperCase() === 'Y';
      callerClub    = data[i][10] ? data[i][10].toString().trim() : '';
      break;
    }
  }
  if (!callerIsAdmin) return { error: 'Unauthorised' };

  const beltMap  = getBeltMap_();
  const requests = [];

  data.slice(1).forEach(row => {
    const status      = row[16] ? row[16].toString().trim() : '';
    const pendingClub = row[19] ? row[19].toString().trim() : '';
    if (status !== 'pending') return;
    // Club-scoped admin only sees requests for their club
    if (callerClub && pendingClub !== callerClub) return;
    const registeredDate = parseSheetDate_(row[8]);
    requests.push({
      username:     row[0] ? row[0].toString().trim() : '',
      displayName:  row[5] ? row[5].toString() : '',
      grade:        parseInt(row[2]) || 1,
      gradeName:    beltMap[parseInt(row[2])] || `Level ${parseInt(row[2]) || 1}`,
      pendingClub,
      registeredDate: registeredDate ? registeredDate.toLocaleDateString('en-GB') : ''
    });
  });

  return { requests, callerClub };
}

// ── getQuizData ───────────────────────────────────────────────────────────────

function getQuizData(username, mode) {
  const ss        = getSpreadsheet_();
  const pSheet    = ss.getSheetByName('UserProgress');
  const cleanUser = username ? username.toString().trim() : '';
  const { grade: userGradeLevel, club: userClub } = getUserContext_(username);

  const questions = getQuestionsParsed_();

  const pLastRow = pSheet.getLastRow();
  const pData    = pLastRow > 1
    ? pSheet.getRange(1, 1, pLastRow, COLS.PROGRESS).getValues()
    : [];

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

  const eligible = questions.filter(q => {
    if (!q.id || !q.q) return false;
    if (q.lv > userGradeLevel) return false;
    if (q.club && q.club !== userClub) return false;
    return true;
  });

  function buildOptions(q) {
    return [q.o1, q.o2, q.o3, q.o4].filter(Boolean).sort(() => Math.random() - 0.5);
  }

  if (mode === 'test') {
    let limit;
    if (userGradeLevel >= 11)     limit = 50;
    else if (userGradeLevel >= 7) limit = 20;
    else                          limit = 10;

    const examCategories = [
      { name: 'Basics',   pct: 0.10 },
      { name: 'Numbers',  pct: 0.15 },
      { name: 'Belts',    pct: 0.15 },
      { name: 'Korean',   pct: 0.30 },
      { name: 'Patterns', pct: 0.30 }
    ];

    function weightedSample(pool, n) {
      if (!pool.length) return [];
      const weights     = pool.map(q => q.lv);
      const totalWeight = weights.reduce((a, b) => a + b, 0);
      const selected    = [], used = new Set();
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
    examCategories.forEach(c => pools[c.name] = []);
    const uncategorised = [];
    eligible.forEach(q => {
      if (pools[q.cat] !== undefined) pools[q.cat].push(q);
      else uncategorised.push(q);
    });

    let selected = [];
    examCategories.forEach(c => {
      selected = selected.concat(weightedSample(pools[c.name], Math.round(c.pct * limit)));
    });

    const remaining = limit - selected.length;
    if (remaining > 0) {
      const selectedIds = new Set(selected.map(q => q.id));
      const fillPool    = [...uncategorised, ...eligible].filter(q => !selectedIds.has(q.id));
      selected = selected.concat(weightedSample(fillPool, remaining));
    }

    return selected.sort(() => Math.random() - 0.5).map(q => ({
      question:  q.q,
      options:   buildOptions(q),
      answer:    q.a,
      qId:       q.id,
      timeLimit: q.tl
    }));
  }

  let filtered = eligible.filter(q => {
    const prog = progressMap[q.id];
    if (!prog) return true;
    return (now - prog.date) / (1000 * 60 * 60 * 24) >= (BUCKET_INTERVALS[prog.bucket] || 0);
  });

  if (!filtered.length) filtered = [...eligible];

  filtered.sort((a, b) => {
    const pA = progressMap[a.id] || { bucket: 1, date: new Date(0) };
    const pB = progressMap[b.id] || { bucket: 1, date: new Date(0) };
    return pA.bucket !== pB.bucket ? pA.bucket - pB.bucket : pA.date - pB.date;
  });

  return filtered.slice(0, 10).map(q => ({
    question:  q.q,
    options:   buildOptions(q),
    answer:    q.a,
    qId:       q.id,
    timeLimit: q.tl
  }));
}

// ── updateQuestionScore ───────────────────────────────────────────────────────

function updateQuestionScore(username, qId, isCorrect) {
  const pSheet    = getSpreadsheet_().getSheetByName('UserProgress');
  const pLastRow  = pSheet.getLastRow();
  const pData     = pLastRow > 1
    ? pSheet.getRange(1, 1, pLastRow, COLS.PROGRESS).getValues()
    : [];
  const now       = new Date();
  const qIdStr    = qId ? qId.toString() : '';
  const cleanUser = username ? username.toString().trim() : '';

  let foundRow = -1;
  for (let i = 0; i < pData.length; i++) {
    if (pData[i][0].toString().trim() === cleanUser && pData[i][1].toString() === qIdStr) {
      foundRow = i + 1; break;
    }
  }

  if (foundRow !== -1) {
    const currentBucket = parseInt(pData[foundRow - 1][3]) || 1;
    const currentScore  = parseInt(pData[foundRow - 1][2]) || 0;
    pSheet.getRange(foundRow, 3, 1, 3).setValues([[
      isCorrect ? currentScore + 1 : currentScore - 1,
      isCorrect ? Math.min(currentBucket + 1, 4) : 1,
      now
    ]]);
  } else {
    pSheet.appendRow([cleanUser, qIdStr, isCorrect ? 1 : -1, isCorrect ? 2 : 1, now]);
  }
}

// ── getGameData ───────────────────────────────────────────────────────────────

function getGameData(username, gameType) {
  const { grade: userGradeLevel } = getUserContext_(username);
  const cards = getSimpleDefsParsed_();
  const eligible = cards.filter(c => c.lv <= userGradeLevel && c.id && c.q);

  if (gameType === 'game_match') {
    const pool = eligible.filter(c => c.q !== '' && c.a !== '');
    return pool.sort(() => Math.random() - 0.5).slice(0, 10).map(c => ({
      matchingTerm:  c.q,
      correctAnswer: c.a,
      qId:           c.id
    }));
  }

  const pool = eligible.filter(c => c.q !== '' && c.a !== '' && c.w !== '');
  return pool.sort(() => Math.random() - 0.5).slice(0, 10).map(c => ({
    simplifiedDef: c.q,
    correctAnswer: c.a,
    decoys:        [c.w],
    qId:           c.id
  }));
}

// ── getBeltOptions / getClubOptions (public) ──────────────────────────────────

function getBeltOptions() {
  return getBeltOptions_();
}

function getClubOptions() {
  return getClubList_();
}

// ── getSrsStats ───────────────────────────────────────────────────────────────

function getSrsStats(username) {
  const { grade } = getUserContext_(username);
  return getSrsStatsForUser_(username ? username.toString().trim() : '', grade);
}

// ── saveGrade ─────────────────────────────────────────────────────────────────

function saveGrade(username, newGrade) {
  const sheet   = getSpreadsheet_().getSheetByName('Users');
  const lastRow = sheet.getLastRow();
  const data    = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  const clean   = username ? username.toString().trim() : '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === clean) {
      sheet.getRange(i + 1, 3).setValue(newGrade);
      return 'Grade Updated!';
    }
  }
}

// ── updatePass ────────────────────────────────────────────────────────────────

function updatePass(u, p) {
  const sheet   = getSpreadsheet_().getSheetByName('Users');
  const lastRow = sheet.getLastRow();
  const data    = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  const clean   = u ? u.toString().trim() : '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === clean) {
      sheet.getRange(i + 1, 2).setValue(p);
      return 'Updated!';
    }
  }
}

// ── Tul Trumps ────────────────────────────────────────────────────────────────

function getTulTrumpsData(username) {
  try {
    const { grade: userGradeLevel } = getUserContext_(username);
    const deck = getTulsParsed_()
      .filter(t => t.minGrade <= userGradeLevel)
      .map(t => ({
        name:        t.name,
        movements:   t.mov,
        stances:     t.sta,
        readyStance: t.rs,
        difficulty:  t.dif,
        meaning:     t.meaning,
        img:         t.img
      }));

    if (deck.length < 4)
      return { error: `Not enough Tuls unlocked for this grade. (Found: ${deck.length})` };

    const shuffled = deck.sort(() => Math.random() - 0.5);
    const mid      = Math.ceil(shuffled.length / 2);
    return { playerHand: shuffled.slice(0, mid), cpuHand: shuffled.slice(mid) };
  } catch(e) {
    return { error: e.message };
  }
}

// ── Flashcards ────────────────────────────────────────────────────────────────

function getFlashcardData(username) {
  const { grade: userGradeLevel } = getUserContext_(username);
  const beltMap  = getBeltMap_();
  const rawCards = getFlashcardsParsed_().filter(c => c.lv && c.lv <= userGradeLevel);

  const cards = rawCards.map(c => {
    const card = { beltLevel: c.lv, category: c.cat, term: c.t, meaning: c.m };
    if (c.cat === 'Patterns') {
      if (c.mn)  card.movesNum      = c.mn;
      if (c.mk)  card.movesKorean   = c.mk;
      if (c.sn)  card.stancesNum    = c.sn;
      if (c.sk)  card.stancesKorean = c.sk;
      if (c.re)  card.readyEn       = c.re;
      if (c.rk)  card.readyKo       = c.rk;
      if (c.img) card.img           = c.img;
    }
    return card;
  });

  const beltLevels = [...new Set(cards.map(c => c.beltLevel))].sort((a, b) => a - b)
    .map(lv => ({ value: lv, label: beltMap[lv] || `Level ${lv}` }));
  const categories = [...new Set(cards.map(c => c.category).filter(Boolean))].sort();

  return { cards, beltLevels, categories };
}

// ── High scores ───────────────────────────────────────────────────────────────

function getHighScore(username) {
  return getHighScoreForUser_(username ? username.toString().trim() : '');
}

function saveHighScore(username, newScore) {
  const ss    = getSpreadsheet_();
  let sheet   = ss.getSheetByName('HighScores');
  if (!sheet) { sheet = ss.insertSheet('HighScores'); sheet.appendRow(['Username', 'InfiniteWarrior']); }
  const lastRow = sheet.getLastRow();
  const data    = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.HIGHSCORES).getValues() : [];
  const clean   = username ? username.toString().trim() : '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === clean) {
      sheet.getRange(i + 1, 2).setValue(newScore); return;
    }
  }
  sheet.appendRow([clean, newScore]);
}

function getAllTimeHighScore() {
  const sheet   = getSpreadsheet_().getSheetByName('HighScores');
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const data = sheet.getRange(1, 1, lastRow, COLS.HIGHSCORES).getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const val = parseInt(data[i][1]) || 0;
    if (val > max) max = val;
  }
  return max;
}

// ── Streak ────────────────────────────────────────────────────────────────────

function incrementStreak(username) {
  const sheet   = getSpreadsheet_().getSheetByName('Users');
  const lastRow = sheet.getLastRow();
  const data    = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  const clean   = username ? username.toString().trim().toLowerCase() : '';
  const today   = new Date().toDateString();

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().trim().toLowerCase() !== clean) continue;
    const streakDateRaw = data[i][15];
    const lastGoalDate  = streakDateRaw ? new Date(streakDateRaw).toDateString() : '';
    if (lastGoalDate === today)
      return { success: true, streak: parseInt(data[i][4]) || 0, alreadyDone: true };
    const currentStreak = parseInt(data[i][4]) || 0;
    sheet.getRange(i + 1, 5).setValue(currentStreak + 1);
    sheet.getRange(i + 1, 16).setValue(new Date());
    return { success: true, streak: currentStreak + 1 };
  }
  return { success: false };
}

// ── Play time logging ─────────────────────────────────────────────────────────

function logPlayTime(username, seconds) {
  const sheet   = getSpreadsheet_().getSheetByName('Users');
  const lastRow = sheet.getLastRow();
  const data    = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  const clean   = username ? username.toString().trim().toLowerCase() : '';
  const now     = new Date();
  const todayKey  = now.toDateString();
  const monthKey  = `${now.getFullYear()}-${now.getMonth()}`;

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().trim().toLowerCase() !== clean) continue;

    let dailyData = { date: todayKey, seconds: 0 };
    try {
      const ex = data[i][12] ? JSON.parse(data[i][12].toString()) : null;
      dailyData = (ex && ex.date === todayKey)
        ? { date: todayKey, seconds: (ex.seconds || 0) + seconds }
        : { date: todayKey, seconds };
    } catch(e) { dailyData = { date: todayKey, seconds }; }

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

// ── Learning goal ─────────────────────────────────────────────────────────────

function saveLearningGoal(username, mins) {
  const sheet   = getSpreadsheet_().getSheetByName('Users');
  const lastRow = sheet.getLastRow();
  const data    = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  const clean   = username ? username.toString().trim().toLowerCase() : '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === clean) {
      sheet.getRange(i + 1, 12).setValue(parseInt(mins) || 0);
      return { success: true };
    }
  }
  return { success: false };
}

// ── saveUserGradingDate ───────────────────────────────────────────────────────

function saveUserGradingDate(username, dateStr) {
  const sheet   = getSpreadsheet_().getSheetByName('Users');
  const lastRow = sheet.getLastRow();
  const data    = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  const clean   = username ? username.toString().trim().toLowerCase() : '';
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().trim().toLowerCase() !== clean) continue;
    if (dateStr) {
      const d = new Date(dateStr);
      sheet.getRange(i + 1, 18).setValue(isNaN(d.getTime()) ? '' : d); // Col R
    } else {
      sheet.getRange(i + 1, 18).setValue('');
    }
    return { success: true };
  }
  return { success: false };
}

// ── updateUserConfig ──────────────────────────────────────────────────────────
// clubMember boolean replaces suspended.
// Unchecking clubMember removes club link and restores stored subscription days.

function updateUserConfig(adminUsername, targetUsername, config) {
  const sheet     = getSpreadsheet_().getSheetByName('Users');
  const lastRow   = sheet.getLastRow();
  const data      = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  const cleanAdmin  = adminUsername  ? adminUsername.toString().trim().toLowerCase()  : '';
  const cleanTarget = targetUsername ? targetUsername.toString().trim().toLowerCase() : '';

  let callerIsAdmin = false, callerClub = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === cleanAdmin) {
      callerIsAdmin = data[i][6] && data[i][6].toString().trim().toUpperCase() === 'Y';
      callerClub    = data[i][10] ? data[i][10].toString().trim() : '';
      break;
    }
  }
  if (!callerIsAdmin) return { success: false, message: 'Unauthorised' };

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().trim().toLowerCase() !== cleanTarget) continue;

    const targetClub = data[i][10] ? data[i][10].toString().trim() : '';
    if (callerClub && targetClub !== callerClub) return { success: false, message: 'Unauthorised' };

    const row = i + 1;

    if (config.grade !== undefined && config.grade !== null)
      sheet.getRange(row, 3).setValue(parseInt(config.grade) || 1);

    if (config.gradingDate !== undefined) {
      if (config.gradingDate) {
        const d = new Date(config.gradingDate);
        sheet.getRange(row, 18).setValue(isNaN(d.getTime()) ? '' : d);
      } else {
        sheet.getRange(row, 18).setValue('');
      }
    }

    // clubMember tick box — replaces suspended
    if (config.clubMember !== undefined) {
      if (config.clubMember) {
        // Checking the box: confirm as member (used for manual override / new club admins)
        sheet.getRange(row, 17).setValue('member');  // Col Q
      } else {
        // Unchecking: remove from club, restore stored subscription days
        const storedDays = parseInt(data[i][20]) || 0;  // Col U
        sheet.getRange(row, 11).setValue('');            // Col K — clear club
        sheet.getRange(row, 17).setValue('');            // Col Q — clear clubStatus
        sheet.getRange(row, 21).setValue('');            // Col U — clear stored days

        // If there were stored subscription days, write a new subscription date
        // calculated as (now - (SUB_DAYS - storedDays)) so remaining days are correct
        if (storedDays > 0) {
          const now         = new Date();
          const newSubStart = new Date(now.getTime() - ((SUB_DAYS - storedDays) * 24 * 60 * 60 * 1000));
          sheet.getRange(row, 10).setValue(newSubStart); // Col J
        }

        bustCache_('club_list');
      }
    }

    return { success: true };
  }
  return { success: false, message: 'User not found' };
}

// ── Admin page data ───────────────────────────────────────────────────────────

function getAdminPageData(adminUsername) {
  const ss          = getSpreadsheet_();
  const usersSheet  = ss.getSheetByName('Users');
  const uLastRow    = usersSheet.getLastRow();
  const userData    = uLastRow > 1 ? usersSheet.getRange(1, 1, uLastRow, COLS.USERS).getValues() : [];
  const cleanCaller = adminUsername ? adminUsername.toString().trim().toLowerCase() : '';

  let callerRow = null;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === cleanCaller) {
      callerRow = userData[i]; break;
    }
  }
  if (!callerRow || callerRow[6].toString().trim().toUpperCase() !== 'Y')
    return { error: 'Unauthorised' };

  const callerClub = callerRow[10] ? callerRow[10].toString().trim() : '';

  const progressSheet  = ss.getSheetByName('UserProgress');
  const highScoreSheet = ss.getSheetByName('HighScores');
  const pLastRow       = progressSheet  ? progressSheet.getLastRow()  : 0;
  const hsLastRow      = highScoreSheet ? highScoreSheet.getLastRow() : 0;
  const progressData   = pLastRow  > 1 ? progressSheet.getRange(1, 1, pLastRow, COLS.PROGRESS).getValues()    : [];
  const highScoreData  = hsLastRow > 1 ? highScoreSheet.getRange(1, 1, hsLastRow, COLS.HIGHSCORES).getValues() : [];

  const questionData = getQuestionsParsed_();
  const beltMap      = getBeltMap_();

  const progressMap = {};
  progressData.slice(1).forEach(row => {
    const u = row[0] ? row[0].toString().trim() : '';
    if (!u) return;
    if (!progressMap[u]) progressMap[u] = { total: 0, buckets: { 1:0,2:0,3:0,4:0 } };
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
      if (callerClub) return (row[10] ? row[10].toString().trim() : '') === callerClub;
      return true;
    })
    .map(row => {
      const uName      = row[0].toString().trim();
      const lastActive = row[3] ? new Date(row[3]) : null;
      const daysSince  = lastActive ? Math.floor((now - lastActive) / (1000 * 60 * 60 * 24)) : null;
      const gradeVal   = parseInt(row[2]) || 1;
      const prog       = progressMap[uName] || { total: 0, buckets: { 1:0,2:0,3:0,4:0 } };

      const eligible = questionData.filter(q => q.id && q.q && q.lv <= gradeVal).length;

      const registeredDate   = parseSheetDate_(row[8]);
      const subscriptionDate = parseSheetDate_(row[9]);
      const clubStatus       = row[16] ? row[16].toString().trim() : '';
      let subStatus = 'active', subEndDate = null;

      if (clubStatus === 'member') {
        subStatus = 'club_member';
      } else if (registeredDate) {
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
        } else if (daysSinceReg <= TRIAL_DAYS + TRIAL_GRACE_DAYS) {
          subStatus = 'trial_grace';
        } else {
          subStatus = subscriptionDate ? 'sub_expired' : 'trial_expired';
        }
      } else {
        subStatus = 'club_member'; // manually created, no registered date
      }

      const dailyPlaySeconds = (() => {
        try { const d = row[12] ? JSON.parse(row[12].toString()) : null; return (d && d.date === now.toDateString()) ? parseInt(d.seconds) || 0 : 0; } catch(e) { return 0; }
      })();
      const monthlyPlaySeconds = (() => {
        try { const m = row[13] ? JSON.parse(row[13].toString()) : null; return (m && m.month === monthKey) ? parseInt(m.seconds) || 0 : 0; } catch(e) { return 0; }
      })();

      return {
        username:      uName,
        displayName:   row[5] ? row[5].toString() : uName,
        grade:         gradeVal,
        gradeName:     beltMap[gradeVal] || `Level ${gradeVal}`,
        streak:        parseInt(row[4]) || 0,
        lastActive:    lastActive ? lastActive.toLocaleString('en-GB') : 'Never',
        daysSince,
        isAdmin:       row[6] && row[6].toString().trim().toUpperCase() === 'Y',
        club:          row[10] ? row[10].toString().trim() : '',
        clubStatus,
        clubMember:    clubStatus === 'member',  // boolean for tick box
        totalAnswered: prog.total,
        buckets:       prog.buckets,
        highScore:     highScoreMap[uName] || 0,
        eligible,
        subStatus,
        subEndDate,
        learningGoalMins:    parseInt(row[11]) || 0,
        dailyPlaySeconds,
        monthlyPlaySeconds,
        gradingDate:   formatDateYMD_(row[17]),
      };
    });

  // Pending request count — for the badge on the Members tile
  const pendingCount = users.filter(u => u.clubStatus === 'pending').length;

  const belts      = getBeltOptions_();
  const categories = [...new Set(questionData.map(q => q.cat).filter(Boolean))].sort();

  return { users, belts, categories, callerClub, pendingCount, memberCount: users.length };
}

// ── getAdminQuestions ─────────────────────────────────────────────────────────

function getAdminQuestions(adminUsername) {
  const ss          = getSpreadsheet_();
  const usersSheet  = ss.getSheetByName('Users');
  const uLastRow    = usersSheet.getLastRow();
  const userData    = uLastRow > 1 ? usersSheet.getRange(1, 1, uLastRow, COLS.USERS).getValues() : [];
  const cleanCaller = adminUsername ? adminUsername.toString().trim().toLowerCase() : '';

  let callerRow = null;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === cleanCaller) {
      callerRow = userData[i]; break;
    }
  }
  if (!callerRow || callerRow[6].toString().trim().toUpperCase() !== 'Y')
    return { error: 'Unauthorised' };

  const callerClub = callerRow[10] ? callerRow[10].toString().trim() : '';
  const beltMap    = getBeltMap_();

  const questions = getQuestionsParsed_()
    .filter(q => q.id && q.q)
    .map(q => ({
      qId:       q.id,
      question:  q.q,
      answer:    q.a,
      opt1:      q.o2,
      opt2:      q.o3,
      opt3:      q.o4,
      beltLevel: q.lv,
      beltName:  beltMap[q.lv] || `Level ${q.lv}`,
      category:  q.cat,
      club:      q.club,
      addedByMe: q.club ? q.club === callerClub : false,
      examFlag:  q.flag
    }));

  return { questions, callerClub };
}

// ── submitQuestion ────────────────────────────────────────────────────────────

function submitQuestion(adminUsername, questionData) {
  const ss         = getSpreadsheet_();
  const usersSheet = ss.getSheetByName('Users');
  const uLastRow   = usersSheet.getLastRow();
  const userData   = uLastRow > 1 ? usersSheet.getRange(1, 1, uLastRow, COLS.USERS).getValues() : [];
  const cleanUser  = adminUsername ? adminUsername.toString().trim().toLowerCase() : '';

  let isAdmin = false, callerClub = '';
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === cleanUser) {
      isAdmin    = userData[i][6] && userData[i][6].toString().trim().toUpperCase() === 'Y';
      callerClub = userData[i][10] ? userData[i][10].toString().trim() : '';
      break;
    }
  }
  if (!isAdmin) return { success: false, message: 'Unauthorised' };

  const qSheet = ss.getSheetByName('Questions');
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
  bustCache_('questions_v2');
  return { success: true, qId };
}

// ── submitContact ─────────────────────────────────────────────────────────────

function submitContact(username, contactData) {
  const ss    = getSpreadsheet_();
  let sheet   = ss.getSheetByName('Contact');
  if (!sheet) { sheet = ss.insertSheet('Contact'); sheet.appendRow(['Timestamp','Username','DisplayName','Club','Type','Message']); }
  sheet.appendRow([new Date(), username, contactData.displayName, contactData.club, contactData.type, contactData.message]);
  try {
    MailApp.sendEmail({
      to:      'jonathan_gallucci@hotmail.com',
      subject: 'TKD Academy contact',
      body:    `New contact from TKD Academy app:\n\nName: ${contactData.displayName}\nClub: ${contactData.club || 'N/A'}\nType: ${contactData.type}\n\nMessage:\n${contactData.message}\n\nFrom: ${username}\nTime: ${new Date().toLocaleString('en-GB')}`
    });
  } catch(e) { Logger.log('Email failed: ' + e.message); }
  return { success: true };
}

// ── Stripe webhook ────────────────────────────────────────────────────────────

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
      if (email) updateSubscriptionDate_(email.toString().trim().toLowerCase());
    }
    return ContentService.createTextOutput(JSON.stringify({ received: true })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function updateSubscriptionDate_(email) {
  const sheet   = getSpreadsheet_().getSheetByName('Users');
  const lastRow = sheet.getLastRow();
  const data    = lastRow > 1 ? sheet.getRange(1, 1, lastRow, COLS.USERS).getValues() : [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === email) {
      sheet.getRange(i + 1, 10).setValue(new Date()); return;
    }
  }
}
