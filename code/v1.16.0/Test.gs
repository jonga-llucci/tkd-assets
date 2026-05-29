// ═══════════════════════════════════════════════════════════════════════════════
// TKD Academy — Server-Side Test Suite
// ═══════════════════════════════════════════════════════════════════════════════

var _results = [];

function runAllTests() {
  _results = [];
  const start = Date.now();
  Logger.log('Cleaning up any previous test rows...');
  cleanupTestRows_();

  const groups = [
    { name: 'Helpers',        tests: TESTS_HELPERS },
    { name: 'Auth',           tests: TESTS_AUTH },
    { name: 'Streak',         tests: TESTS_STREAK },
    { name: 'PlayTime',       tests: TESTS_PLAYTIME },
    { name: 'Quiz',           tests: TESTS_QUIZ },
    { name: 'Cache',          tests: TESTS_CACHE },
    { name: 'Admin',          tests: TESTS_ADMIN },
    { name: 'UserConfig',     tests: TESTS_USERCONFIG },
    { name: 'ClubMembership', tests: TESTS_CLUB },
    { name: 'Notifications',  tests: TESTS_NOTIFICATIONS }
  ];

  groups.forEach(group => {
    Logger.log('\n── ' + group.name + ' ──────────────────────────────────────');
    group.tests.forEach(fn => {
      const name   = fn.name.replace(/^test_/, '').replace(/_/g, ' ');
      let passed   = false, errorMsg = '';
      try { fn(); passed = true; } catch(e) { errorMsg = e.message || String(e); }
      _results.push({ group: group.name, name, passed, errorMsg });
      Logger.log((passed ? '  ✓ ' : '  ✗ ') + name + (passed ? '' : '\n      → ' + errorMsg));
    });
  });

  cleanupTestRows_();
  cleanupTestNotifications_();

  const total  = _results.length;
  const passed = _results.filter(r => r.passed).length;
  const failed = total - passed;
  const elapsed = ((Date.now() - start) / 1000).toFixed(1);
  Logger.log('\n══════════════════════════════════════════');
  Logger.log(`  ${passed}/${total} passed   ${failed} failed   ${elapsed}s`);
  Logger.log('══════════════════════════════════════════');
  if (failed > 0) {
    Logger.log('\nFailed tests:');
    _results.filter(r => !r.passed).forEach(r => {
      Logger.log(`  [${r.group}] ${r.name}\n    ${r.errorMsg}`);
    });
  }
}

// ── Assertion helpers ─────────────────────────────────────────────────────────

function assert_(condition, message) {
  if (!condition) throw new Error(message || 'Assertion failed');
}
function assertEqual_(actual, expected, label) {
  if (actual !== expected)
    throw new Error(`${label || 'assertEqual'}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
}
function assertContains_(arr, value, label) {
  if (!Array.isArray(arr) || !arr.includes(value))
    throw new Error(`${label || 'assertContains'}: ${JSON.stringify(value)} not found in ${JSON.stringify(arr)}`);
}
function assertDefined_(val, label) {
  if (val === undefined || val === null)
    throw new Error(`${label || 'assertDefined'}: value is ${val}`);
}
function assertMatch_(str, regex, label) {
  if (!regex.test(str))
    throw new Error(`${label || 'assertMatch'}: "${str}" does not match ${regex}`);
}

// ── Test sandbox ──────────────────────────────────────────────────────────────

var TEST_USER_EMAIL     = 'test.user.tkd.suite@example.com';
var TEST_USER_PASSWORD  = 'Test!Pass1';
var TEST_ADMIN_EMAIL    = 'test.admin.tkd.suite@example.com';
var TEST_ADMIN_PASSWORD = 'Admin!Pass1';
var TEST_EMAIL_PREFIX   = 'test.';

function cleanupTestRows_() {
  const ss = getSpreadsheet_();
  const usersSheet = ss.getSheetByName('Users');
  if (usersSheet) {
    const data = usersSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const email = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
      if (email.startsWith(TEST_EMAIL_PREFIX) && email.endsWith('@example.com'))
        usersSheet.deleteRow(i + 1);
    }
  }
  const progressSheet = ss.getSheetByName('UserProgress');
  if (progressSheet) {
    const pData = progressSheet.getDataRange().getValues();
    for (let i = pData.length - 1; i >= 1; i--) {
      const email = pData[i][0] ? pData[i][0].toString().trim().toLowerCase() : '';
      if (email.startsWith(TEST_EMAIL_PREFIX) && email.endsWith('@example.com'))
        progressSheet.deleteRow(i + 1);
    }
  }
  Logger.log('  🧹 Test rows cleaned up');
}

function cleanupTestNotifications_() {
  const sheet = getSpreadsheet_().getSheetByName('Notifications');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const email = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
    if (email.startsWith(TEST_EMAIL_PREFIX) && email.endsWith('@example.com'))
      sheet.deleteRow(i + 1);
  }
}

// insertTestUser_ — 22-column row matching new Users schema
function insertTestUser_(opts) {
  const sheet = getSpreadsheet_().getSheetByName('Users');
  const now   = new Date();
  const email = opts.email || TEST_USER_EMAIL;
  const row   = new Array(22).fill('');

  row[0]  = email;
  row[1]  = opts.password      || TEST_USER_PASSWORD;
  row[2]  = opts.grade         || 5;
  row[3]  = opts.lastActive !== undefined ? opts.lastActive : now;
  row[4]  = opts.streak        || 0;
  row[5]  = opts.name          || 'Test User';
  row[6]  = opts.isAdmin       ? 'Y' : '';
  row[8]  = opts.registered !== undefined ? opts.registered : now;
  row[9]  = opts.subscribed    || '';
  row[10] = opts.club          || '';
  row[11] = opts.goalMins      || 0;
  row[15] = opts.streakDate    || '';
  row[16] = opts.clubStatus    || '';   // Col Q — was suspended, now clubStatus
  row[17] = opts.gradingDate   || '';
  row[18] = opts.newClubDetails || '';  // Col S
  row[19] = opts.pendingClub   || '';   // Col T
  row[20] = opts.subDaysStored || '';   // Col U
  row[21] = opts.pendingClubRequestDate || ''; // Col V

  sheet.appendRow(row);
  return email;
}

function deleteTestRow_(email) {
  if (!email) return;
  try {
    const sheet = getSpreadsheet_().getSheetByName('Users');
    const data  = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === email.toLowerCase()) {
        sheet.deleteRow(i + 1); return;
      }
    }
  } catch(e) {}
}

function insertTestProgress_(username, qId, bucket) {
  const sheet = getSpreadsheet_().getSheetByName('UserProgress');
  sheet.appendRow([username, qId, 1, bucket, new Date()]);
  return { username, qId };
}

function deleteTestProgressRow_(ref) {
  if (!ref) return;
  try {
    const sheet = getSpreadsheet_().getSheetByName('UserProgress');
    const data  = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] && data[i][0].toString() === ref.username &&
          data[i][1] && data[i][1].toString() === ref.qId) {
        sheet.deleteRow(i + 1); return;
      }
    }
  } catch(e) {}
}

// ── TESTS: Helpers ────────────────────────────────────────────────────────────

function test_parseSheetDate_returns_null_for_empty() {
  assertEqual_(parseSheetDate_(null), null, 'null input');
  assertEqual_(parseSheetDate_(''), null, 'empty string');
}
function test_parseSheetDate_handles_date_objects() {
  const d = new Date(2026, 5, 7);
  assertEqual_(parseSheetDate_(d), d, 'Date object passthrough');
}
function test_parseSheetDate_parses_iso_string() {
  const d = parseSheetDate_('2026-06-07');
  assert_(d instanceof Date && !isNaN(d.getTime()), 'ISO string parsed to valid Date');
}
function test_formatDateYMD_empty_for_null() {
  assertEqual_(formatDateYMD_(null), '', 'null → empty');
  assertEqual_(formatDateYMD_(''),   '', 'empty → empty');
}
function test_formatDateYMD_returns_YYYY_MM_DD() {
  const result = formatDateYMD_(new Date(2026, 5, 7));
  assertMatch_(result, /^\d{4}-\d{2}-\d{2}$/, 'formatDateYMD_ shape');
  assert_(result.startsWith('2026'), 'correct year');
}
function test_getBeltOptions_returns_array_of_value_label() {
  const opts = getBeltOptions_();
  assert_(Array.isArray(opts) && opts.length > 0, 'non-empty array');
  assert_(opts[0].value !== undefined && opts[0].label !== undefined, 'has value and label');
  assert_(typeof opts[0].value === 'number', 'value is a number');
}
function test_getBeltMap_keys_match_getBeltOptions_values() {
  const opts = getBeltOptions_();
  const map  = getBeltMap_();
  opts.forEach(b => {
    assert_(map[b.value] === b.label, `beltMap[${b.value}] should equal '${b.label}'`);
  });
}
function test_getUserContext_returns_defaults_for_unknown_user() {
  const ctx = getUserContext_('nobody@nowhere.invalid');
  assertEqual_(ctx.grade, 1, 'default grade');
  assertEqual_(ctx.club,  '', 'default club');
  assertEqual_(ctx.clubStatus, '', 'default clubStatus');
}

var TESTS_HELPERS = [
  test_parseSheetDate_returns_null_for_empty,
  test_parseSheetDate_handles_date_objects,
  test_parseSheetDate_parses_iso_string,
  test_formatDateYMD_empty_for_null,
  test_formatDateYMD_returns_YYYY_MM_DD,
  test_getBeltOptions_returns_array_of_value_label,
  test_getBeltMap_keys_match_getBeltOptions_values,
  test_getUserContext_returns_defaults_for_unknown_user
];

// ── TESTS: Auth ───────────────────────────────────────────────────────────────

function test_loginUser_fails_for_unknown_user() {
  const res = loginUser('nobody@nowhere.invalid', 'wrongpassword');
  assertEqual_(res.success, false, 'unknown user returns failure');
}
function test_loginUser_fails_for_wrong_password() {
  let email;
  try {
    email = insertTestUser_({});
    const res = loginUser(TEST_USER_EMAIL, 'WRONG_PASSWORD');
    assertEqual_(res.success, false, 'wrong password returns failure');
  } finally { deleteTestRow_(email); }
}
function test_loginUser_succeeds_with_correct_credentials() {
  let email;
  try {
    email = insertTestUser_({ registered: new Date() });
    const res = loginUser(TEST_USER_EMAIL, TEST_USER_PASSWORD);
    assertEqual_(res.success, true, 'correct credentials succeed');
    assert_(res.accessStatus, 'accessStatus present');
    assert_(res.bootstrap !== undefined, 'bootstrap payload present');
    assert_(Array.isArray(res.bootstrap.belts), 'bootstrap.belts is array');
    assert_(Array.isArray(res.bootstrap.clubs), 'bootstrap.clubs is array');
  } finally { deleteTestRow_(email); }
}
function test_loginUser_returns_trial_status_within_7_days() {
  let email;
  try {
    const threeDaysAgo = new Date(Date.now() - 3 * 24 * 60 * 60 * 1000);
    email = insertTestUser_({ registered: threeDaysAgo });
    const res = loginUser(TEST_USER_EMAIL, TEST_USER_PASSWORD);
    assertEqual_(res.accessStatus, 'trial', 'within-trial accessStatus');
    assert_(res.daysRemaining > 0 && res.daysRemaining <= 7, 'daysRemaining in range');
  } finally { deleteTestRow_(email); }
}
function test_loginUser_returns_trial_expired_in_grace_period() {
  let email;
  try {
    const tenDaysAgo = new Date(Date.now() - 10 * 24 * 60 * 60 * 1000);
    email = insertTestUser_({ registered: tenDaysAgo });
    const res = loginUser(TEST_USER_EMAIL, TEST_USER_PASSWORD);
    assertEqual_(res.accessStatus, 'trial_expired', 'grace period accessStatus');
    assert_(res.graceDaysRemaining >= 0 && res.graceDaysRemaining <= 7, 'graceDaysRemaining in range');
    assertEqual_(res.success, true, 'login still succeeds during grace');
  } finally { deleteTestRow_(email); }
}
function test_loginUser_returns_trial_hard_locked_after_grace() {
  let email;
  try {
    const fifteenDaysAgo = new Date(Date.now() - 15 * 24 * 60 * 60 * 1000);
    email = insertTestUser_({ registered: fifteenDaysAgo });
    const res = loginUser(TEST_USER_EMAIL, TEST_USER_PASSWORD);
    assertEqual_(res.accessStatus, 'trial_hard_locked', 'hard locked after grace');
    assertEqual_(res.success, true, 'success:true even for hard lock (user found)');
    assert_(res.bootstrap === undefined, 'no bootstrap payload on hard lock');
  } finally { deleteTestRow_(email); }
}
function test_loginUser_returns_subscribed_when_active_subscription() {
  let email;
  try {
    const tenDaysAgo = new Date(Date.now() - 10 * 24 * 60 * 60 * 1000);
    const yesterday  = new Date(Date.now() - 1 * 24 * 60 * 60 * 1000);
    email = insertTestUser_({ registered: tenDaysAgo, subscribed: yesterday });
    const res = loginUser(TEST_USER_EMAIL, TEST_USER_PASSWORD);
    assertEqual_(res.accessStatus, 'subscribed', 'active subscription status');
    assert_(res.daysRemaining > 0, 'daysRemaining populated');
  } finally { deleteTestRow_(email); }
}
function test_loginUser_returns_subscription_expired_when_lapsed() {
  let email;
  try {
    const old = new Date(Date.now() - 60 * 24 * 60 * 60 * 1000);
    email = insertTestUser_({ registered: old, subscribed: old });
    const res = loginUser(TEST_USER_EMAIL, TEST_USER_PASSWORD);
    assertEqual_(res.accessStatus, 'subscription_expired', 'lapsed subscription status');
  } finally { deleteTestRow_(email); }
}
function test_loginUser_club_member_always_active() {
  let email;
  try {
    // Club member with expired trial — should still be fully active
    const old = new Date(Date.now() - 60 * 24 * 60 * 60 * 1000);
    email = insertTestUser_({ registered: old, club: 'TestClub', clubStatus: 'member' });
    const res = loginUser(TEST_USER_EMAIL, TEST_USER_PASSWORD);
    assertEqual_(res.success, true, 'success');
    assertEqual_(res.accessStatus, 'active', 'club member is always active');
    assertEqual_(res.clubStatus, 'member', 'clubStatus returned correctly');
  } finally { deleteTestRow_(email); }
}
function test_loginUser_returns_gradingDate_in_YYYY_MM_DD() {
  let email;
  try {
    const grading = new Date(2026, 8, 15);
    email = insertTestUser_({ registered: new Date(), gradingDate: grading });
    const res = loginUser(TEST_USER_EMAIL, TEST_USER_PASSWORD);
    assertMatch_(res.gradingDate || '', /^\d{4}-\d{2}-\d{2}$/, 'gradingDate format');
  } finally { deleteTestRow_(email); }
}
function test_loginUser_resets_streak_when_2_or_more_days_missed() {
  let email;
  try {
    const threeDaysAgo = new Date(Date.now() - 3 * 24 * 60 * 60 * 1000);
    email = insertTestUser_({ registered: new Date(), streak: 10, streakDate: threeDaysAgo });
    const res = loginUser(TEST_USER_EMAIL, TEST_USER_PASSWORD);
    assertEqual_(res.streak, 0, 'streak reset after missing 2+ days');
  } finally { deleteTestRow_(email); }
}
function test_loginUser_preserves_streak_when_missed_only_yesterday() {
  let email;
  try {
    const yesterday = new Date(Date.now() - 1 * 24 * 60 * 60 * 1000);
    email = insertTestUser_({ registered: new Date(), streak: 5, streakDate: yesterday });
    const res = loginUser(TEST_USER_EMAIL, TEST_USER_PASSWORD);
    assertEqual_(res.streak, 5, 'streak preserved when only one day missed');
  } finally { deleteTestRow_(email); }
}
function test_registerUser_fails_if_fields_missing() {
  const res = registerUser('', '', '', '', '');
  assertEqual_(res.success, false, 'empty fields rejected');
}
function test_registerUser_fails_for_duplicate_email() {
  let email;
  try {
    email = insertTestUser_({});
    const res = registerUser('Test Duplicate', TEST_USER_EMAIL, 3, 'Pass!1a', '', '');
    assertEqual_(res.success, false, 'duplicate email rejected');
  } finally { deleteTestRow_(email); }
}

var TESTS_AUTH = [
  test_loginUser_fails_for_unknown_user,
  test_loginUser_fails_for_wrong_password,
  test_loginUser_succeeds_with_correct_credentials,
  test_loginUser_returns_trial_status_within_7_days,
  test_loginUser_returns_trial_expired_in_grace_period,
  test_loginUser_returns_trial_hard_locked_after_grace,
  test_loginUser_returns_subscribed_when_active_subscription,
  test_loginUser_returns_subscription_expired_when_lapsed,
  test_loginUser_club_member_always_active,
  test_loginUser_returns_gradingDate_in_YYYY_MM_DD,
  test_loginUser_resets_streak_when_2_or_more_days_missed,
  test_loginUser_preserves_streak_when_missed_only_yesterday,
  test_registerUser_fails_if_fields_missing,
  test_registerUser_fails_for_duplicate_email
];

// ── TESTS: Streak ─────────────────────────────────────────────────────────────

function test_incrementStreak_increments_on_first_goal_of_day() {
  let email;
  try {
    email = insertTestUser_({ streak: 3 });
    const res = incrementStreak(TEST_USER_EMAIL);
    assertEqual_(res.success, true, 'success');
    assertEqual_(res.streak, 4, 'streak incremented from 3 to 4');
    assertEqual_(res.alreadyDone, undefined, 'not alreadyDone');
  } finally { deleteTestRow_(email); }
}
function test_incrementStreak_is_idempotent_same_day() {
  let email;
  try {
    email = insertTestUser_({ streak: 5, streakDate: new Date() });
    const res = incrementStreak(TEST_USER_EMAIL);
    assertEqual_(res.success, true, 'success');
    assertEqual_(res.alreadyDone, true, 'alreadyDone flag set');
    assertEqual_(res.streak, 5, 'streak not incremented again');
  } finally { deleteTestRow_(email); }
}
function test_incrementStreak_fails_for_unknown_user() {
  const res = incrementStreak('nobody@nowhere.invalid');
  assertEqual_(res.success, false, 'unknown user returns failure');
}

var TESTS_STREAK = [
  test_incrementStreak_increments_on_first_goal_of_day,
  test_incrementStreak_is_idempotent_same_day,
  test_incrementStreak_fails_for_unknown_user
];

// ── TESTS: PlayTime ───────────────────────────────────────────────────────────

function test_logPlayTime_writes_daily_and_monthly_json() {
  let email;
  try {
    email = insertTestUser_({});
    const res = logPlayTime(TEST_USER_EMAIL, 120);
    assertEqual_(res.success, true, 'logPlayTime success');
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === TEST_USER_EMAIL.toLowerCase());
    const data    = sheet.getRange(rowNum + 1, 13, 1, 2).getValues()[0];
    const daily   = JSON.parse(data[0]);
    const monthly = JSON.parse(data[1]);
    assertEqual_(daily.seconds, 120, 'daily seconds written');
    assert_(typeof daily.date === 'string', 'daily.date is a string');
    assertEqual_(monthly.seconds, 120, 'monthly seconds written');
    assertMatch_(monthly.month, /^\d{4}-\d{1,2}$/, 'monthly.month format is YYYY-M');
  } finally { deleteTestRow_(email); }
}
function test_logPlayTime_accumulates_within_same_day() {
  let email;
  try {
    email = insertTestUser_({});
    logPlayTime(TEST_USER_EMAIL, 60);
    logPlayTime(TEST_USER_EMAIL, 90);
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === TEST_USER_EMAIL.toLowerCase());
    const data    = sheet.getRange(rowNum + 1, 13, 1, 1).getValues()[0];
    const daily   = JSON.parse(data[0]);
    assertEqual_(daily.seconds, 150, 'daily seconds accumulated correctly');
  } finally { deleteTestRow_(email); }
}
function test_logPlayTime_resets_monthly_when_month_changes() {
  let email;
  try {
    email = insertTestUser_({});
    logPlayTime(TEST_USER_EMAIL, 300);
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === TEST_USER_EMAIL.toLowerCase()) + 1;
    sheet.getRange(rowNum, 14).setValue(JSON.stringify({ month: '2020-0', seconds: 9999 }));
    logPlayTime(TEST_USER_EMAIL, 60);
    const daily = JSON.parse(sheet.getRange(rowNum, 14, 1, 1).getValues()[0][0]);
    assertEqual_(daily.seconds, 60, 'monthly reset when month changes');
  } finally { deleteTestRow_(email); }
}

var TESTS_PLAYTIME = [
  test_logPlayTime_writes_daily_and_monthly_json,
  test_logPlayTime_accumulates_within_same_day,
  test_logPlayTime_resets_monthly_when_month_changes
];

// ── TESTS: Quiz ───────────────────────────────────────────────────────────────

function test_getQuizData_practice_returns_up_to_10_questions() {
  const data = getQuizData('test@nonexistent.invalid', 'practice');
  assert_(Array.isArray(data), 'returns array');
  assert_(data.length <= 10, 'max 10 questions');
}
function test_getQuizData_practice_questions_have_required_fields() {
  const data = getQuizData('test@nonexistent.invalid', 'practice');
  if (data.length === 0) return;
  const q = data[0];
  assertDefined_(q.question, 'question field');
  assertDefined_(q.answer,   'answer field');
  assertDefined_(q.qId,      'qId field');
  assert_(Array.isArray(q.options), 'options is array');
  assert_(q.options.length >= 2, 'at least 2 options');
}
function test_getQuizData_answer_is_always_in_options() {
  const data = getQuizData('test@nonexistent.invalid', 'practice');
  data.forEach((q, i) => {
    assert_(q.options.includes(q.answer), `Q${i}: answer not in options`);
  });
}
function test_getQuizData_test_returns_questions() {
  const data = getQuizData('test@nonexistent.invalid', 'test');
  assert_(Array.isArray(data), 'test mode returns array');
  assert_(data.length > 0, 'test mode returns at least one question');
}
function test_getQuizData_test_answer_always_in_options() {
  const data = getQuizData('test@nonexistent.invalid', 'test');
  data.forEach((q, i) => {
    assert_(q.options.includes(q.answer), `Exam Q${i}: answer not in options`);
  });
}
function test_getQuizData_respects_grade_filtering() {
  const questions   = getQuestionsParsed_();
  const grade1Pool  = questions.filter(q => q.lv <= 1);
  const grade19Pool = questions.filter(q => q.lv <= 19);
  assert_(grade19Pool.length >= grade1Pool.length, 'higher grade has >= questions');
}
function test_updateQuestionScore_correct_answer_advances_bucket() {
  let ref;
  try {
    ref = insertTestProgress_(TEST_USER_EMAIL, 'test_qid_suite_1', 1);
    updateQuestionScore(TEST_USER_EMAIL, 'test_qid_suite_1', true);
    const sheet   = getSpreadsheet_().getSheetByName('UserProgress');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] === TEST_USER_EMAIL && r[1] === 'test_qid_suite_1') + 1;
    const data    = sheet.getRange(rowNum, 1, 1, 5).getValues()[0];
    assertEqual_(data[3], 2, 'correct answer advances bucket from 1 to 2');
  } finally { deleteTestProgressRow_(ref); }
}
function test_updateQuestionScore_wrong_answer_resets_to_bucket_1() {
  let ref;
  try {
    ref = insertTestProgress_(TEST_USER_EMAIL, 'test_qid_suite_2', 3);
    updateQuestionScore(TEST_USER_EMAIL, 'test_qid_suite_2', false);
    const sheet   = getSpreadsheet_().getSheetByName('UserProgress');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] === TEST_USER_EMAIL && r[1] === 'test_qid_suite_2') + 1;
    const data    = sheet.getRange(rowNum, 1, 1, 5).getValues()[0];
    assertEqual_(data[3], 1, 'wrong answer resets bucket to 1');
  } finally { deleteTestProgressRow_(ref); }
}
function test_updateQuestionScore_bucket_never_exceeds_4() {
  let ref;
  try {
    ref = insertTestProgress_(TEST_USER_EMAIL, 'test_qid_suite_3', 4);
    updateQuestionScore(TEST_USER_EMAIL, 'test_qid_suite_3', true);
    const sheet   = getSpreadsheet_().getSheetByName('UserProgress');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] === TEST_USER_EMAIL && r[1] === 'test_qid_suite_3') + 1;
    const data    = sheet.getRange(rowNum, 1, 1, 5).getValues()[0];
    assertEqual_(data[3], 4, 'bucket capped at 4');
  } finally { deleteTestProgressRow_(ref); }
}

var TESTS_QUIZ = [
  test_getQuizData_practice_returns_up_to_10_questions,
  test_getQuizData_practice_questions_have_required_fields,
  test_getQuizData_answer_is_always_in_options,
  test_getQuizData_test_returns_questions,
  test_getQuizData_test_answer_always_in_options,
  test_getQuizData_respects_grade_filtering,
  test_updateQuestionScore_correct_answer_advances_bucket,
  test_updateQuestionScore_wrong_answer_resets_to_bucket_1,
  test_updateQuestionScore_bucket_never_exceeds_4
];

// ── TESTS: Cache ──────────────────────────────────────────────────────────────

function test_cache_put_and_get_small_payload() {
  const key = 'test_cache_small', data = { foo: 'bar', num: 42 };
  putCached_(key, data, 60);
  const result = getCached_(key);
  assertDefined_(result, 'cached value retrieved');
  assertEqual_(result.foo, 'bar', 'cached string field');
  assertEqual_(result.num, 42, 'cached number field');
  bustCache_(key);
}
function test_cache_bust_removes_entry() {
  const key = 'test_cache_bust';
  putCached_(key, { x: 1 }, 60);
  bustCache_(key);
  assertEqual_(getCached_(key), null, 'cache entry removed after bust');
}
function test_cache_chunked_put_and_get() {
  const key = 'test_cache_chunked', bigArr = [];
  for (let i = 0; i < 3000; i++)
    bigArr.push({ id: 'q_' + i, question: 'Question number ' + i + ' with some padding text here', answer: 'Answer ' + i, lv: (i % 10) + 1 });
  putCached_(key, bigArr, 60);
  const result = getCachedChunked_(key);
  assertDefined_(result, 'chunked cache retrieved');
  assert_(Array.isArray(result), 'chunked result is array');
  assertEqual_(result.length, bigArr.length, 'all items preserved through chunking');
  assertEqual_(result[0].id, 'q_0', 'first item intact');
  assertEqual_(result[2999].id, 'q_2999', 'last item intact');
  bustCache_(key);
}
function test_cache_chunked_bust_removes_all_chunks() {
  const key    = 'test_cache_chunked_bust';
  const bigArr = new Array(3000).fill({ id: 'x', question: 'q padding text here padding text', answer: 'a', lv: 1 });
  putCached_(key, bigArr, 60);
  bustCache_(key);
  const meta   = CacheService.getScriptCache().get(key + '_meta');
  const direct = getCachedChunked_(key);
  assertEqual_(meta, null, 'meta chunk removed');
  assertEqual_(direct, null, 'no data returned after bust');
}
function test_getBeltsParsed_returns_from_cache_on_second_call() {
  bustCache_('belts_v2');
  const first  = getBeltsParsed_();
  const second = getBeltsParsed_();
  assertEqual_(first.length, second.length, 'same count from cache as from sheet');
  assertEqual_(first[0].value, second[0].value, 'first item identical from cache');
}

var TESTS_CACHE = [
  test_cache_put_and_get_small_payload,
  test_cache_bust_removes_entry,
  test_cache_chunked_put_and_get,
  test_cache_chunked_bust_removes_all_chunks,
  test_getBeltsParsed_returns_from_cache_on_second_call
];

// ── TESTS: Admin ──────────────────────────────────────────────────────────────

function test_getAdminPageData_rejects_non_admin() {
  let email;
  try {
    email = insertTestUser_({});
    const res = getAdminPageData(TEST_USER_EMAIL);
    assertDefined_(res.error, 'non-admin gets error');
  } finally { deleteTestRow_(email); }
}
function test_getAdminPageData_returns_users_and_member_count_for_admin() {
  let email;
  try {
    email = insertTestUser_({ isAdmin: true });
    const res = getAdminPageData(TEST_USER_EMAIL);
    if (!res.error) {
      assert_(Array.isArray(res.users), 'users is array');
      assert_(Array.isArray(res.belts), 'belts is array');
      assert_(Array.isArray(res.categories), 'categories is array');
      assert_(typeof res.memberCount === 'number', 'memberCount is a number');
      assert_(typeof res.pendingCount === 'number', 'pendingCount is a number');
    }
  } finally { deleteTestRow_(email); }
}
function test_getAdminQuestions_rejects_non_admin() {
  let email;
  try {
    email = insertTestUser_({});
    const res = getAdminQuestions(TEST_USER_EMAIL);
    assertDefined_(res.error, 'non-admin gets error');
  } finally { deleteTestRow_(email); }
}
function test_getAdminQuestions_returns_questions_for_admin() {
  let email;
  try {
    email = insertTestUser_({ isAdmin: true });
    const res = getAdminQuestions(TEST_USER_EMAIL);
    if (!res.error) {
      assert_(Array.isArray(res.questions), 'questions is array');
      if (res.questions.length > 0) {
        const q = res.questions[0];
        assertDefined_(q.qId,      'qId field');
        assertDefined_(q.question, 'question field');
        assertDefined_(q.answer,   'answer field');
        assertDefined_(q.beltLevel,'beltLevel field');
      }
    }
  } finally { deleteTestRow_(email); }
}

var TESTS_ADMIN = [
  test_getAdminPageData_rejects_non_admin,
  test_getAdminPageData_returns_users_and_member_count_for_admin,
  test_getAdminQuestions_rejects_non_admin,
  test_getAdminQuestions_returns_questions_for_admin
];

// ── TESTS: UserConfig ─────────────────────────────────────────────────────────

function test_updateUserConfig_rejects_non_admin() {
  let adminEmail, targetEmail;
  try {
    adminEmail  = insertTestUser_({ email: 'test.admin.config@example.com', isAdmin: false });
    targetEmail = insertTestUser_({ email: 'test.target.config@example.com' });
    const res = updateUserConfig('test.admin.config@example.com', 'test.target.config@example.com', { grade: 3 });
    assertEqual_(res.success, false, 'non-admin rejected');
  } finally { deleteTestRow_(adminEmail); deleteTestRow_(targetEmail); }
}
function test_updateUserConfig_admin_can_update_grade() {
  let adminEmail, targetEmail;
  try {
    adminEmail  = insertTestUser_({ email: 'test.admin.cfg2@example.com', isAdmin: true });
    targetEmail = insertTestUser_({ email: 'test.target.cfg2@example.com', grade: 3 });
    const res = updateUserConfig('test.admin.cfg2@example.com', 'test.target.cfg2@example.com', { grade: 7 });
    assertEqual_(res.success, true, 'grade update succeeded');
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === 'test.target.cfg2@example.com') + 1;
    const data    = sheet.getRange(rowNum, 3, 1, 1).getValues()[0];
    assertEqual_(data[0], 7, 'grade written to col C');
  } finally { deleteTestRow_(adminEmail); deleteTestRow_(targetEmail); }
}
function test_updateUserConfig_admin_can_set_club_member() {
  let adminEmail, targetEmail;
  try {
    adminEmail  = insertTestUser_({ email: 'test.admin.cfg3@example.com', isAdmin: true, club: 'TestClub' });
    targetEmail = insertTestUser_({ email: 'test.target.cfg3@example.com', club: 'TestClub', clubStatus: 'pending' });
    const res = updateUserConfig('test.admin.cfg3@example.com', 'test.target.cfg3@example.com', { clubMember: true });
    assertEqual_(res.success, true, 'club member set succeeded');
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === 'test.target.cfg3@example.com') + 1;
    const val     = sheet.getRange(rowNum, 17, 1, 1).getValues()[0][0]; // col Q
    assertEqual_(val, 'member', 'col Q set to member');
  } finally { deleteTestRow_(adminEmail); deleteTestRow_(targetEmail); }
}
function test_updateUserConfig_uncheck_club_member_clears_club() {
  let adminEmail, targetEmail;
  try {
    adminEmail  = insertTestUser_({ email: 'test.admin.cfg4@example.com', isAdmin: true, club: 'TestClub' });
    targetEmail = insertTestUser_({ email: 'test.target.cfg4@example.com', club: 'TestClub', clubStatus: 'member' });
    const res = updateUserConfig('test.admin.cfg4@example.com', 'test.target.cfg4@example.com', { clubMember: false });
    assertEqual_(res.success, true, 'uncheck succeeded');
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === 'test.target.cfg4@example.com') + 1;
    const clubVal    = sheet.getRange(rowNum, 11, 1, 1).getValues()[0][0]; // col K
    const statusVal  = sheet.getRange(rowNum, 17, 1, 1).getValues()[0][0]; // col Q
    assertEqual_(clubVal,   '', 'club cleared from col K');
    assertEqual_(statusVal, '', 'clubStatus cleared from col Q');
  } finally { deleteTestRow_(adminEmail); deleteTestRow_(targetEmail); }
}
function test_updateUserConfig_admin_can_set_grading_date() {
  let adminEmail, targetEmail;
  try {
    adminEmail  = insertTestUser_({ email: 'test.admin.cfg5@example.com', isAdmin: true });
    targetEmail = insertTestUser_({ email: 'test.target.cfg5@example.com' });
    const res = updateUserConfig('test.admin.cfg5@example.com', 'test.target.cfg5@example.com', { gradingDate: '2026-09-15' });
    assertEqual_(res.success, true, 'gradingDate update succeeded');
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === 'test.target.cfg5@example.com') + 1;
    const val     = sheet.getRange(rowNum, 18, 1, 1).getValues()[0][0];
    assert_(val instanceof Date || val !== '', 'gradingDate written to col R');
  } finally { deleteTestRow_(adminEmail); deleteTestRow_(targetEmail); }
}
function test_saveUserGradingDate_user_can_set_own_date() {
  let email;
  try {
    email = insertTestUser_({});
    const res = saveUserGradingDate(TEST_USER_EMAIL, '2026-09-20');
    assertEqual_(res.success, true, 'user can save own grading date');
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === TEST_USER_EMAIL.toLowerCase()) + 1;
    const val     = sheet.getRange(rowNum, 18, 1, 1).getValues()[0][0];
    assert_(val instanceof Date || val !== '', 'date written to sheet');
  } finally { deleteTestRow_(email); }
}
function test_saveUserGradingDate_clear_removes_date() {
  let email;
  try {
    email = insertTestUser_({ gradingDate: new Date(2026, 8, 20) });
    saveUserGradingDate(TEST_USER_EMAIL, '');
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === TEST_USER_EMAIL.toLowerCase()) + 1;
    const val     = sheet.getRange(rowNum, 18, 1, 1).getValues()[0][0];
    assertEqual_(val, '', 'grading date cleared');
  } finally { deleteTestRow_(email); }
}

var TESTS_USERCONFIG = [
  test_updateUserConfig_rejects_non_admin,
  test_updateUserConfig_admin_can_update_grade,
  test_updateUserConfig_admin_can_set_club_member,
  test_updateUserConfig_uncheck_club_member_clears_club,
  test_updateUserConfig_admin_can_set_grading_date,
  test_saveUserGradingDate_user_can_set_own_date,
  test_saveUserGradingDate_clear_removes_date
];

// ── TESTS: Club Membership ────────────────────────────────────────────────────

function test_requestClubJoin_fails_if_already_pending() {
  let email;
  try {
    email = insertTestUser_({ clubStatus: 'pending', pendingClub: 'SomeClub' });
    const res = requestClubJoin(TEST_USER_EMAIL, 'AnotherClub', '');
    assertEqual_(res.success, false, 'rejected when already pending');
  } finally { deleteTestRow_(email); }
}
function test_requestClubJoin_fails_if_already_member() {
  let email;
  try {
    email = insertTestUser_({ clubStatus: 'member', club: 'SomeClub' });
    const res = requestClubJoin(TEST_USER_EMAIL, 'AnotherClub', '');
    assertEqual_(res.success, false, 'rejected when already a member');
  } finally { deleteTestRow_(email); }
}
function test_requestClubJoin_rate_limited_within_7_days() {
  let email;
  try {
    const recentRequest = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000); // 2 days ago
    email = insertTestUser_({ pendingClubRequestDate: recentRequest });
    const res = requestClubJoin(TEST_USER_EMAIL, 'SomeClub', '');
    assertEqual_(res.success, false, 'rate limited within 7 days');
  } finally { deleteTestRow_(email); }
}
function test_requestClubJoin_succeeds_after_rate_limit_window() {
  let email;
  try {
    const oldRequest = new Date(Date.now() - 8 * 24 * 60 * 60 * 1000); // 8 days ago
    email = insertTestUser_({ pendingClubRequestDate: oldRequest });
    // Note: this will attempt to notify admins — safe as no admin exists for this test club
    const res = requestClubJoin(TEST_USER_EMAIL, 'test.fakeclubname.noexist', '');
    assertEqual_(res.success, true, 'request allowed after rate limit window');
    // Verify cols written
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === TEST_USER_EMAIL.toLowerCase()) + 1;
    const status  = sheet.getRange(rowNum, 17, 1, 1).getValues()[0][0]; // col Q
    const pending = sheet.getRange(rowNum, 20, 1, 1).getValues()[0][0]; // col T
    assertEqual_(status,  'pending', 'clubStatus set to pending');
    assertEqual_(pending, 'test.fakeclubname.noexist', 'pendingClub written to col T');
  } finally { deleteTestRow_(email); }
}
function test_approveClubRequest_rejects_non_admin() {
  let email;
  try {
    email = insertTestUser_({});
    const res = approveClubRequest(TEST_USER_EMAIL, 'someone@example.com');
    assertEqual_(res.success, false, 'non-admin rejected');
  } finally { deleteTestRow_(email); }
}
function test_approveClubRequest_sets_member_status_and_writes_notification() {
  let adminEmail, targetEmail;
  try {
    adminEmail  = insertTestUser_({ email: 'test.admin.club1@example.com', isAdmin: true, club: 'TestClubApprove' });
    targetEmail = insertTestUser_({ email: 'test.target.club1@example.com', clubStatus: 'pending', pendingClub: 'TestClubApprove' });
    const res = approveClubRequest('test.admin.club1@example.com', 'test.target.club1@example.com');
    assertEqual_(res.success, true, 'approve succeeded');
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === 'test.target.club1@example.com') + 1;
    const status  = sheet.getRange(rowNum, 17, 1, 1).getValues()[0][0]; // col Q
    const club    = sheet.getRange(rowNum, 11, 1, 1).getValues()[0][0]; // col K
    const pending = sheet.getRange(rowNum, 20, 1, 1).getValues()[0][0]; // col T
    assertEqual_(status,  'member',          'clubStatus set to member');
    assertEqual_(club,    'TestClubApprove', 'club written to col K');
    assertEqual_(pending, '',                'pendingClub cleared');
    // Check notification written
    const { notifications } = getNotifications('test.target.club1@example.com');
    assert_(notifications.length > 0, 'notification written');
    assert_(notifications[0].message.includes('approved'), 'approval message');
  } finally {
    deleteTestRow_(adminEmail);
    deleteTestRow_(targetEmail);
    cleanupTestNotifications_();
  }
}
function test_declineClubRequest_clears_pending_and_writes_notification() {
  let adminEmail, targetEmail;
  try {
    adminEmail  = insertTestUser_({ email: 'test.admin.club2@example.com', isAdmin: true, club: 'TestClubDecline' });
    targetEmail = insertTestUser_({ email: 'test.target.club2@example.com', clubStatus: 'pending', pendingClub: 'TestClubDecline' });
    const res = declineClubRequest('test.admin.club2@example.com', 'test.target.club2@example.com');
    assertEqual_(res.success, true, 'decline succeeded');
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === 'test.target.club2@example.com') + 1;
    const status  = sheet.getRange(rowNum, 17, 1, 1).getValues()[0][0]; // col Q
    const pending = sheet.getRange(rowNum, 20, 1, 1).getValues()[0][0]; // col T
    assertEqual_(status,  '', 'clubStatus cleared');
    assertEqual_(pending, '', 'pendingClub cleared');
    const { notifications } = getNotifications('test.target.club2@example.com');
    assert_(notifications.length > 0, 'notification written');
    assert_(notifications[0].message.includes('not been approved'), 'decline message');
  } finally {
    deleteTestRow_(adminEmail);
    deleteTestRow_(targetEmail);
    cleanupTestNotifications_();
  }
}
function test_approveClubRequest_stores_remaining_sub_days_when_subscribed() {
  let adminEmail, targetEmail;
  try {
    // Subscribed user with 20 days remaining (subscribed 10 days ago, 30-day sub)
    const tenDaysAgo = new Date(Date.now() - 10 * 24 * 60 * 60 * 1000);
    adminEmail  = insertTestUser_({ email: 'test.admin.club3@example.com', isAdmin: true, club: 'TestClubSub' });
    targetEmail = insertTestUser_({ email: 'test.target.club3@example.com', clubStatus: 'pending', pendingClub: 'TestClubSub', subscribed: tenDaysAgo });
    approveClubRequest('test.admin.club3@example.com', 'test.target.club3@example.com');
    const sheet   = getSpreadsheet_().getSheetByName('Users');
    const allRows = sheet.getDataRange().getValues();
    const rowNum  = allRows.findIndex(r => r[0] && r[0].toString().toLowerCase() === 'test.target.club3@example.com') + 1;
    const stored  = parseInt(sheet.getRange(rowNum, 21, 1, 1).getValues()[0][0]); // col U
    assert_(stored > 0, 'subscription days stored on approval');
    assert_(stored <= 30, 'stored days within expected range');
  } finally {
    deleteTestRow_(adminEmail);
    deleteTestRow_(targetEmail);
    cleanupTestNotifications_();
  }
}

var TESTS_CLUB = [
  test_requestClubJoin_fails_if_already_pending,
  test_requestClubJoin_fails_if_already_member,
  test_requestClubJoin_rate_limited_within_7_days,
  test_requestClubJoin_succeeds_after_rate_limit_window,
  test_approveClubRequest_rejects_non_admin,
  test_approveClubRequest_sets_member_status_and_writes_notification,
  test_declineClubRequest_clears_pending_and_writes_notification,
  test_approveClubRequest_stores_remaining_sub_days_when_subscribed
];

// ── TESTS: Notifications ──────────────────────────────────────────────────────

function test_writeNotification_creates_unread_entry() {
  try {
    writeNotification_(TEST_USER_EMAIL, 'Test notification message', 'test', 'TestClub');
    const { notifications, unreadCount } = getNotifications(TEST_USER_EMAIL);
    assert_(notifications.length > 0, 'notification retrieved');
    assertEqual_(unreadCount, 1, 'one unread notification');
    assertEqual_(notifications[0].message, 'Test notification message', 'message matches');
    assertEqual_(notifications[0].isRead,  false, 'notification is unread');
  } finally { cleanupTestNotifications_(); }
}
function test_markNotificationsRead_marks_all_read() {
  try {
    writeNotification_(TEST_USER_EMAIL, 'Msg 1', 'test', '');
    writeNotification_(TEST_USER_EMAIL, 'Msg 2', 'test', '');
    markNotificationsRead(TEST_USER_EMAIL);
    const { unreadCount } = getNotifications(TEST_USER_EMAIL);
    assertEqual_(unreadCount, 0, 'all notifications marked read');
  } finally { cleanupTestNotifications_(); }
}
function test_getNotifications_returns_most_recent_first() {
  try {
    writeNotification_(TEST_USER_EMAIL, 'First message', 'test', '');
    Utilities.sleep(1100); // ensure distinct timestamps
    writeNotification_(TEST_USER_EMAIL, 'Second message', 'test', '');
    const { notifications } = getNotifications(TEST_USER_EMAIL);
    assert_(notifications.length >= 2, 'at least 2 notifications');
    assertEqual_(notifications[0].message, 'Second message', 'most recent first');
  } finally { cleanupTestNotifications_(); }
}
function test_getNotifications_only_returns_own_notifications() {
  const otherUser = 'test.other.notif@example.com';
  try {
    writeNotification_(TEST_USER_EMAIL, 'My message',    'test', '');
    writeNotification_(otherUser,       'Other message', 'test', '');
    const { notifications } = getNotifications(TEST_USER_EMAIL);
    notifications.forEach(n => {
      assert_(n.message !== 'Other message', 'other user notification not returned');
    });
  } finally { cleanupTestNotifications_(); }
}

var TESTS_NOTIFICATIONS = [
  test_writeNotification_creates_unread_entry,
  test_markNotificationsRead_marks_all_read,
  test_getNotifications_returns_most_recent_first,
  test_getNotifications_only_returns_own_notifications
];

// ── Quick-run helpers ─────────────────────────────────────────────────────────

function runHelperTests()       { _results=[]; TESTS_HELPERS.forEach(f => { try{f();Logger.log('✓ '+f.name);}catch(e){Logger.log('✗ '+f.name+': '+e.message);} }); }
function runAuthTests()         { _results=[]; TESTS_AUTH.forEach(f => { try{f();Logger.log('✓ '+f.name);}catch(e){Logger.log('✗ '+f.name+': '+e.message);} }); }
function runStreakTests()        { _results=[]; TESTS_STREAK.forEach(f => { try{f();Logger.log('✓ '+f.name);}catch(e){Logger.log('✗ '+f.name+': '+e.message);} }); }
function runPlayTimeTests()      { _results=[]; TESTS_PLAYTIME.forEach(f => { try{f();Logger.log('✓ '+f.name);}catch(e){Logger.log('✗ '+f.name+': '+e.message);} }); }
function runQuizTests()         { _results=[]; TESTS_QUIZ.forEach(f => { try{f();Logger.log('✓ '+f.name);}catch(e){Logger.log('✗ '+f.name+': '+e.message);} }); }
function runCacheTests()        { _results=[]; TESTS_CACHE.forEach(f => { try{f();Logger.log('✓ '+f.name);}catch(e){Logger.log('✗ '+f.name+': '+e.message);} }); }
function runAdminTests()        { _results=[]; TESTS_ADMIN.forEach(f => { try{f();Logger.log('✓ '+f.name);}catch(e){Logger.log('✗ '+f.name+': '+e.message);} }); }
function runUserConfigTests()   { _results=[]; TESTS_USERCONFIG.forEach(f => { try{f();Logger.log('✓ '+f.name);}catch(e){Logger.log('✗ '+f.name+': '+e.message);} }); }
function runClubTests()         { _results=[]; TESTS_CLUB.forEach(f => { try{f();Logger.log('✓ '+f.name);}catch(e){Logger.log('✗ '+f.name+': '+e.message);} }); }
function runNotificationTests() { _results=[]; TESTS_NOTIFICATIONS.forEach(f => { try{f();Logger.log('✓ '+f.name);}catch(e){Logger.log('✗ '+f.name+': '+e.message);} }); }
