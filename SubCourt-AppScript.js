// ══════════════════════════════════════════════════
// SUBCOURT — Apps Script Web App
// MWF Tennis League
// ══════════════════════════════════════════════════

const SHEET_ID = '1VjFuq63KLEgZpYvCVi2bJrWEgMxDP6hXygYwjDpUmRE';

// Email enabled state is stored in Config B20 and toggled from the Admin UI.
// Do not hardcode this — use isEmailEnabled() instead.
function isEmailEnabled() {
  try {
    var v = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config).getRange('B27').getValue();
    return v === true || v.toString().toUpperCase() === 'TRUE';
  } catch(e) { return false; }
}

function getEmailSettings() {
  return { emailEnabled: isEmailEnabled() };
}

function setEmailEnabled(params) {
  var enabled = params.enabled === 'true' || params.enabled === true;
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);
  sheet.getRange('A27').setValue('Email Notifications Enabled');
  sheet.getRange('B27').setValue(enabled);
  return { success: true, emailEnabled: enabled };
}

const TABS = {
  players:      'Players',
  requests:     'SubRequests',
  volunteers:   'Volunteers',
  config:       'Config',
  availability: 'Availability',
  matchGroups:  'MatchGroups'
};

const TIMES = ['08:00','09:30','11:00','12:30'];
const TIME_LABELS = {
  '08:00': '8:00 AM',
  '09:30': '9:30 AM',
  '11:00': '11:00 AM',
  '12:30': '12:30 PM'
};

// ──────────────────────────────────────────────────
// CONFIG
// ──────────────────────────────────────────────────

function getConfig() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);
    return {
      // Matching engine — rows 4–7
      skillWindowPreSchedule:   parseFloat(sheet.getRange('B4').getValue())  || 0.5,
      skillWindowPostSchedule:  parseFloat(sheet.getRange('B5').getValue())  || 2.0,
      preScheduleThresholdHrs:  parseInt(sheet.getRange('B6').getValue())    || 48,
      lastMinuteThresholdHrs:   parseInt(sheet.getRange('B7').getValue())    || 24,
      // Volunteer calendar — row 10
      calendarLookaheadDays:    parseInt(sheet.getRange('B10').getValue())   || 30,
      // Dispatch automation — rows 13–14
      autoDispatchEnabled:      (function() { var v = sheet.getRange('B13').getValue(); return v === true || v.toString().toUpperCase() === 'TRUE'; })(),
      autoDispatchTimeET:       formatSheetTime(sheet.getRange('B14').getValue()) || '08:00',
      // Match time reminder — rows 28–29
      matchTimeReminderEnabled: (function() { var v = sheet.getRange('B28').getValue(); return v === true || v.toString().toUpperCase() === 'TRUE'; })(),
      matchTimeReminderTimeET:  formatSheetTime(sheet.getRange('B29').getValue()) || '10:00',
      // Availability window — rows 16–18
      availWindowOpenDate:      (function() { var v = sheet.getRange('B16').getValue(); return v instanceof Date ? formatSheetDate(v) : (v ? v.toString() : ''); })(),
      availWindowCloseDate:     (function() { var v = sheet.getRange('B17').getValue(); return v instanceof Date ? formatSheetDate(v) : (v ? v.toString() : ''); })(),
      availWindowActive:        (function() { var v = sheet.getRange('B18').getValue(); return v === true || v.toString().toUpperCase() === 'TRUE'; })(),
    };
  } catch(e) {
    // If Config tab is missing or unreadable, return safe defaults
    return {
      skillWindowPreSchedule:  0.5,
      skillWindowPostSchedule: 2.0,
      preScheduleThresholdHrs: 48,
      lastMinuteThresholdHrs:  24,
      calendarLookaheadDays:   30,
      autoDispatchEnabled:      false,
      autoDispatchTimeET:       '08:00',
      matchTimeReminderEnabled: false,
      matchTimeReminderTimeET:  '10:00',
      availWindowOpenDate:     '',
      availWindowCloseDate:    '',
      availWindowActive:       false,
    };
  }
}

// ──────────────────────────────────────────────────
// DISPATCH AUTOMATION TRIGGER
// Run this function manually from the Apps Script
// editor whenever you change the dispatch time in
// the Config tab.
// ──────────────────────────────────────────────────

// ──────────────────────────────────────────────────
// ONE-TIME SETUP
// Run setupTriggers() once from the Apps Script editor
// to install the auto-dispatch schedule and the
// config watcher. Re-runs automatically when B13/B14
// are edited thereafter.
// ──────────────────────────────────────────────────

function setupTriggers() {
  // Remove existing managed triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'runAutoDispatch' || fn === 'onConfigEdit' || fn === 'cleanupOldAvailability' || fn === 'checkAvailabilityWindow') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Install onEdit watcher for Config tab changes
  ScriptApp.newTrigger('onConfigEdit')
    .forSpreadsheet(SHEET_ID)
    .onEdit()
    .create();
  // Monthly cleanup of old availability records (runs on the 1st of each month)
  ScriptApp.newTrigger('cleanupOldAvailability')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();
  // Daily check to auto-close availability window when close date passes
  ScriptApp.newTrigger('checkAvailabilityWindow')
    .timeBased()
    .atHour(1)   // 1 AM — runs before anyone opens the app
    .everyDays(1)
    .create();
  // Set up the dispatch time trigger
  updateDispatchTrigger();
  var config = getConfig();
  Logger.log('onConfigEdit watcher installed. Dispatch trigger: ' +
    (config.autoDispatchEnabled
      ? 'ACTIVE — runs daily at ' + config.autoDispatchTimeET + ' ET'
      : 'disabled (Config B13=FALSE) — set B13 to TRUE to activate') + '.');
}

function onConfigEdit(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== TABS.config) return;
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (col === 2 && (row === 13 || row === 14)) {
    updateDispatchTrigger();
    Logger.log('Config changed — dispatch trigger updated.');
  }
}

function getOrCreateDispatchLog() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('DispatchLog');
  if (!sheet) {
    sheet = ss.insertSheet('DispatchLog');
    sheet.getRange(1, 1, 1, 9).setValues([[
      'Timestamp','RequestID','RequestorName','MatchDate','MatchTime','Result','SubName','SubEmail','Notes'
    ]]);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
  }
  return sheet;
}

function updateDispatchTrigger(enabledOverride, timeOverride) {
  var enabled, timeET;
  if (enabledOverride !== undefined && timeOverride !== undefined) {
    enabled = enabledOverride;
    timeET  = timeOverride;
  } else {
    var config = getConfig();
    enabled = config.autoDispatchEnabled;
    timeET  = config.autoDispatchTimeET;
  }

  // Delete any existing dispatch triggers
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runAutoDispatch') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  if (!enabled) {
    Logger.log('Auto-dispatch is disabled. No trigger set.');
    return;
  }

  // Parse the ET time string (HH:MM)
  const parts = timeET.split(':');
  const hourET = parseInt(parts[0]);
  const minET  = parseInt(parts[1]) || 0;

  // Convert ET to UTC (ET = UTC-5 standard, UTC-4 daylight)
  // Apps Script runs in the script timezone — set project timezone to America/New_York
  // and use the hour directly
  ScriptApp.newTrigger('runAutoDispatch')
    .timeBased()
    .atHour(hourET)
    .nearMinute(minET)
    .everyDays(1)
    .inTimezone('America/New_York')
    .create();

  Logger.log('Dispatch trigger set for ' + config.autoDispatchTimeET + ' ET daily.');
}

function runAutoDispatch() {
  var config = getConfig();
  if (!config.autoDispatchEnabled) {
    Logger.log('runAutoDispatch: disabled, exiting.');
    return { skipped: 'disabled' };
  }

  // Step 1: expire all sub requests and volunteer records on or before today
  expireUpToToday();

  // Step 2: fetch open requests (after expiry, so already-expired ones are excluded)
  var requests  = getRequests();
  var open      = requests.filter(function(r) { return r.status === 'open'; });
  var logSheet  = getOrCreateDispatchLog();
  var timestamp = new Date().toISOString();

  Logger.log('runAutoDispatch: started at ' + timestamp + ', ' + open.length + ' open request(s).');
  if (!open.length) return { dispatched: 0 };

  // Track volunteers assigned during this run to prevent double-booking
  // (sheet-read cache within one execution can return stale data after confirmSub writes)
  var assignedThisRun = {}; // key: email|matchDate → true

  open.forEach(function(req) {
    try {
      var result = runMatch({ requestId: req.id });
      if (result.candidates && result.candidates.length > 0) {
        // Filter out anyone already assigned in this run
        var eligible = result.candidates.filter(function(c) {
          return !assignedThisRun[c.email.toLowerCase() + '|' + req.matchDate];
        });
        if (!eligible.length) {
          logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, 'no_candidates', '', '', 'all candidates already assigned this run']);
          Logger.log('No eligible candidates (all assigned this run): ' + req.name);
          return;
        }
        var best = eligible[0];
        confirmSub({
          requestId:         req.id,
          requestRowIndex:   req.rowIndex,
          subEmail:          best.email,
          subName:           best.name,
          requestorName:     req.name,
          requestorEmail:    req.email,
          matchDate:         req.matchDate,
          matchTime:         req.matchTime,
          volunteerRowIndex: best.rowIndex || null,
          groupPlayers:      JSON.stringify(req.groupPlayers || [])
        });
        assignedThisRun[best.email.toLowerCase() + '|' + req.matchDate] = true;
        logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, 'matched', best.name, best.email, '']);
        Logger.log('Auto-dispatched: ' + req.name + ' → ' + best.name);
      } else {
        // No match found — notify requestor if within 24 hours
        if (isLastMinute(req, config.lastMinuteThresholdHrs)) {
          sendRetirementEmail(req);
          logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, 'no_candidates', '', '', 'notified — last-minute, no candidates']);
          Logger.log('No candidates (last-minute, notified): ' + req.name);
        } else {
          logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, 'no_candidates', '', '', '']);
          Logger.log('No candidates for: ' + req.name + ' (' + req.id + ')');
        }
      }
    } catch(err) {
      logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, 'error', '', '', err.message]);
      Logger.log('Auto-dispatch error for ' + req.id + ': ' + err.message);
    }
  });

  return { dispatched: open.length };
}

// ──────────────────────────────────────────────────
// ROUTING
// ──────────────────────────────────────────────────

function doGet(e) {
  const action   = e.parameter.action;
  const callback = e.parameter.callback;
  let result;

  try {
    if (action === 'getRequests')          result = getRequests();
    else if (action === 'getVolunteers')   result = getVolunteers();
    else if (action === 'getPlayers')      result = getPlayers();
    else if (action === 'getHomeData')     result = getHomeData();
    else if (action === 'getConfig')       result = getConfig();
    else if (action === 'submitRequest')   result = submitRequest(e.parameter);
    else if (action === 'submitVolunteer') result = submitVolunteer(e.parameter);
    else if (action === 'confirmSub')      result = confirmSub(e.parameter);
    else if (action === 'runMatch')        result = runMatch(e.parameter);
    else if (action === 'updateVolunteer')  result = updateVolunteer(e.parameter);
    else if (action === 'deleteVolunteer')  result = deleteVolunteer(e.parameter);
    else if (action === 'getDispatchLog')    result = getDispatchLog();
    else if (action === 'expireToday')       result = expireToday();
    else if (action === 'retireRequest')          result = retireRequest(e.parameter);
    else if (action === 'saveAutoDispatchSettings')      result = saveAutoDispatchSettings(e.parameter);
    else if (action === 'runAutoDispatchNow')             result = runAutoDispatch();
    else if (action === 'saveMatchTimeReminderSettings') result = saveMatchTimeReminderSettings(e.parameter);
    else if (action === 'runMatchTimeReminderNow')        result = runMatchTimeReminder();
    else if (action === 'updateRequestTime') result = updateRequestTime(e.parameter);
    else if (action === 'sendAdminCode')          result = sendAdminCode(e.parameter);
    else if (action === 'verifyAdminCode')         result = verifyAdminCode(e.parameter);
    else if (action === 'debugAdmin')              result = debugAdmin(e.parameter);
    else if (action === 'getCoordinatorRatings')   result = getCoordinatorRatings(e.parameter);
    else if (action === 'saveCoordinatorRatings')  result = saveCoordinatorRatings(e.parameter);
    else if (action === 'getEmailSettings')         result = getEmailSettings();
    else if (action === 'setEmailEnabled')          result = setEmailEnabled(e.parameter);
    else if (action === 'getAvailabilityConfig')   result = getAvailabilityConfig();
    else if (action === 'openAvailabilityWindow')  result = openAvailabilityWindow(e.parameter);
    else if (action === 'closeAvailabilityWindow') result = closeAvailabilityWindow();
    else if (action === 'submitAvailability')       result = submitAvailability(e.parameter);
    else if (action === 'getMyAvailability')        result = getMyAvailability(e.parameter);
    else if (action === 'getAvailabilityData')      result = getAvailabilityData(e.parameter);
    else if (action === 'getSchedulerSettings')     result = getSchedulerSettings();
    else if (action === 'getSchedulerDashboard')   result = getSchedulerDashboard();
    else if (action === 'generateSchedule')         result = generateSchedule(e.parameter);
    else if (action === 'publishScheduleStart')     result = publishScheduleStart(e.parameter);
    else if (action === 'publishScheduleSlot')      result = publishScheduleSlot(e.parameter);
    else if (action === 'getPublishedSchedule')     result = getPublishedSchedule();
    else if (action === 'ping')            result = { version: 'V36', ts: new Date().toISOString() };
    else if (action === 'debugMatch') {
      const requestId = e.parameter.requestId;
      const reqs      = getRequests();
      const vols      = getVolunteers();
      const players   = getPlayersWithRatings();
      const config    = getConfig();
      const req = reqs.find(r => r.id === requestId);
      if (!req) {
        result = { error: 'Request not found' };
      } else {
        const reqPlayer       = players.find(p => p.email === req.email.toLowerCase());
        const reqRating       = reqPlayer ? reqPlayer.rating : null;
        const matchDate       = req.matchDate;
        const matchTime       = req.matchTime;
        const lastMinute      = isLastMinute(req, config.lastMinuteThresholdHrs);
        const urgent          = isUrgent(req, config.preScheduleThresholdHrs);
        const skillWindow     = lastMinute ? Infinity : (!urgent ? config.skillWindowPreSchedule : config.skillWindowPostSchedule);
        const hasTBDTime      = !matchTime;
        const effectiveTime   = (matchTime || '08:00').trim();
        const requireAllTimes = !lastMinute && !urgent && !hasTBDTime;
        const trace = vols.map(v => {
          const volTimes     = v.times.map(t => t.trim());
          const dateMatch    = v.date.trim() === matchDate.trim();
          const notRequestor = v.email.toLowerCase() !== req.email.toLowerCase();
          const timeMatch    = requireAllTimes
                                 ? TIMES.every(t => volTimes.includes(t))
                                 : volTimes.includes(effectiveTime);
          const skillOk      = skillWindow === Infinity ? true : (() => {
            const p = players.find(p => p.email.toLowerCase() === v.email.toLowerCase());
            return p ? Math.abs(p.rating - reqRating) <= skillWindow : false;
          })();
          const notAssigned  = !reqs.some(r =>
            r.assignedSub && r.assignedSub.toLowerCase() === v.email.toLowerCase() &&
            r.matchDate === matchDate && r.status === 'filled'
          );
          const notPlaying   = !reqs.some(r =>
            r.email.toLowerCase() === v.email.toLowerCase() &&
            r.matchDate === matchDate && r.status !== 'open'
          );
          const playingTriggers = reqs.filter(r =>
            r.email.toLowerCase() === v.email.toLowerCase() &&
            r.matchDate === matchDate && r.status !== 'open'
          ).map(r => ({ id: r.id, status: r.status, matchTime: r.matchTime }));
          return {
            name: v.name, email: v.email,
            volDate: v.date, reqDate: matchDate,
            volTimes: v.times, reqTime: matchTime,
            dateMatch, notRequestor, timeMatch, skillOk, notAssigned, notPlaying,
            passes: dateMatch && notRequestor && timeMatch && skillOk && notAssigned && notPlaying,
            playingTriggers
          };
        });
        result = {
          req: { id: req.id, matchDate, matchTime, email: req.email },
          lastMinute, requireAllTimes,
          skillWindow: skillWindow === Infinity ? 'none' : skillWindow,
          trace
        };
      }
    }
    else result = { error: 'Unknown action: ' + action };
  } catch (err) {
    result = { error: err.message };
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // Kept for backwards compatibility but all actions now use doGet
  return doGet(e);
}

// ──────────────────────────────────────────────────
// READS
// ──────────────────────────────────────────────────

function getRequests() {
  const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(1, 1, lastRow, 9).getValues();
  rows.shift();
  return rows.map((r, i) => ({
    rowIndex:     i + 2,
    id:           r[0] || '',
    timestamp:    r[1] ? new Date(r[1]).toISOString() : '',
    name:         r[2] || '',
    email:        r[3] || '',
    matchDate:    formatSheetDate(r[4]),
    matchTime:    formatSheetTime(r[5]),
    status:       r[6] || 'open',
    assignedSub:  r[7] || '',
    groupPlayers: (function() { try { return JSON.parse(r[8] || '[]'); } catch(e) { return []; } })()
  }));
}

function formatVolTimes(val) {
  if (!val && val !== 0) return [];
  // If it's a Date object (Sheets stored a single time value)
  if (val instanceof Date) {
    const h = String(val.getHours()).padStart(2, '0');
    const m = String(val.getMinutes()).padStart(2, '0');
    return [h + ':' + m];
  }
  // If it's a number (time serial: fraction of a day)
  if (typeof val === 'number') {
    const totalMins = Math.round(val * 24 * 60);
    const h = String(Math.floor(totalMins / 60)).padStart(2, '0');
    const m = String(totalMins % 60).padStart(2, '0');
    return [h + ':' + m];
  }
  // Plain text — decode underscore format (08_00 → 08:00) and normalize
  return val.toString().split(',').map(t => {
    const s = t.trim().replace('_', ':');
    return /^\d:\d{2}$/.test(s) ? '0' + s : s;
  }).filter(Boolean);
}

function formatSheetDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return val.toString().trim();
}

function formatSheetTime(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const h = String(val.getHours()).padStart(2, '0');
    const m = String(val.getMinutes()).padStart(2, '0');
    return h + ':' + m;
  }
  const s = val.toString().trim();
  if (/^\d:\d{2}$/.test(s)) return '0' + s;
  return s;
}

function getVolunteers() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.volunteers);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(1, 1, lastRow, 7).getValues();
  rows.shift();
  return rows.map((r, i) => ({
    rowIndex:  i + 2,
    id:        r[0] || '',
    timestamp: r[1] ? new Date(r[1]).toISOString() : '',
    name:      r[2] || '',
    email:     r[3] || '',
    date:      formatSheetDate(r[4]),
    times:     formatVolTimes(r[5]),
    status:    r[6] || 'pending'
  }));
}

function getPlayers() {
  // Returns players WITHOUT ratings — ratings are used internally only
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  const rows  = sheet.getDataRange().getValues();
  rows.shift();
  return rows.map(r => ({
    name:    r[0] || '',
    email:   (r[1] || '').toLowerCase(),
    phone:   r[2] || '',
    isAdmin: r[5] === true || String(r[5]).toUpperCase() === 'TRUE'
    // rating intentionally excluded from public response
  }));
}

// Combined home-page bootstrap call — returns players + availConfig in one round trip.
function getHomeData() {
  return {
    players:     getPlayers(),
    availConfig: getAvailabilityConfig()
  };
}

function getDispatchLog() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DispatchLog');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  // Return last 30 rows, most recent first
  return rows.slice(-30).reverse().map(function(r) {
    return {
      timestamp:     r[0] ? new Date(r[0]).toISOString() : '',
      requestId:     r[1] || '',
      requestorName: r[2] || '',
      matchDate:     r[3] ? (r[3] instanceof Date ? formatSheetDate(r[3]) : r[3].toString()) : '',
      matchTime:     r[4] ? (r[4] instanceof Date ? formatSheetTime(r[4]) : r[4].toString()) : '',
      result:        r[5] || '',
      subName:       r[6] || '',
      subEmail:      r[7] || '',
      notes:         r[8] || ''
    };
  });
}

function expireToday() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var today = formatSheetDate(new Date());
  var expired = { requests: 0, volunteers: 0 };

  // Expire open sub requests for today
  var reqSheet = ss.getSheetByName(TABS.requests);
  if (reqSheet && reqSheet.getLastRow() >= 2) {
    var reqRows = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < reqRows.length; i++) {
      var matchDate = formatSheetDate(reqRows[i][4]);
      var status    = (reqRows[i][6] || '').toString();
      if (matchDate === today && status === 'open') {
        reqSheet.getRange(i + 2, 7).setValue('expired');
        expired.requests++;
      }
    }
  }

  // Expire pending volunteer records for today
  var volSheet = ss.getSheetByName(TABS.volunteers);
  if (volSheet && volSheet.getLastRow() >= 2) {
    var volRows = volSheet.getRange(2, 1, volSheet.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < volRows.length; i++) {
      var volDate = formatSheetDate(volRows[i][4]);
      var status  = (volRows[i][6] || '').toString();
      if (volDate === today && status === 'pending') {
        volSheet.getRange(i + 2, 7).setValue('expired');
        expired.volunteers++;
      }
    }
  }

  return { success: true, expired: expired };
}

function getPlayersWithRatings() {
  // Internal use only — never sent to browser
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  const rows  = sheet.getDataRange().getValues();
  rows.shift();
  // Deduplicate by email — first row wins; duplicate emails cause identity collisions
  const seen = {};
  return rows.reduce(function(acc, r) {
    const email = (r[1] || '').toLowerCase();
    if (email && !seen[email]) {
      seen[email] = true;
      acc.push({ name: r[0] || '', email: email, rating: parseFloat(r[3]) || 0 });
    } else if (email && seen[email]) {
      Logger.log('WARNING: duplicate email in Players sheet: ' + email);
    }
    return acc;
  }, []);
}

// ──────────────────────────────────────────────────
// WRITES
// ──────────────────────────────────────────────────

function submitRequest(params) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
  const groupPlayers = params.groupPlayers
    ? (typeof params.groupPlayers === 'string' ? params.groupPlayers : JSON.stringify(params.groupPlayers))
    : '[]';
  const row = [
    uid(),
    new Date().toISOString(),
    params.name,
    params.email,
    params.matchDate ? params.matchDate.toString() : '',
    params.matchTime ? params.matchTime.toString() : '',
    'open',
    '',
    groupPlayers
  ];
  sheet.appendRow(row);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 5).setNumberFormat('@');
  sheet.getRange(lastRow, 6).setNumberFormat('@');
  sheet.getRange(lastRow, 9).setNumberFormat('@');
  return { success: true };
}

function submitVolunteer(params) {
  const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.volunteers);
  const entries = JSON.parse(params.entries);
  entries.forEach(entry => {
    const nextRow = sheet.getLastRow() + 1;
    const range   = sheet.getRange(nextRow, 1, 1, 7);
    // Set number format first to prevent auto-conversion
    range.setNumberFormats([['@','@','@','@','@','@','@']]);
    range.setValues([[
      uid(),
      new Date().toISOString(),
      params.name,
      params.email,
      entry.date,
      entry.times.join(','),  // stored as 08_00,09_30 etc to prevent Sheets time conversion
      'pending'
    ]]);
  });
  return { success: true };
}

function updateVolunteer(params) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.volunteers);
  const times = JSON.parse(params.times); // e.g. ["08:00","09:30"]
  const encoded = times.map(t => t.replace(':', '_')).join(','); // 08_00,09_30
  const cell = sheet.getRange(parseInt(params.rowIndex), 6);
  cell.setNumberFormat('@');
  cell.setValue(encoded);
  return { success: true };
}

function deleteVolunteer(params) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.volunteers);
  sheet.getRange(parseInt(params.rowIndex), 7).setValue('cancelled');
  return { success: true };
}

// ──────────────────────────────────────────────────
// ADMIN AUTH
// ──────────────────────────────────────────────────

function testAdminAuth() {
  var testEmail = 'brianna.biesecker@gmail.com'; // change if needed
  var ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log('Spreadsheet name: ' + ss.getName() + ' | ID: ' + ss.getId());
  var sheet = ss.getSheetByName(TABS.players);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  Logger.log('lastRow=' + lastRow + ' lastCol=' + lastCol);
  // Read direct cell values
  Logger.log('E1 direct: [' + sheet.getRange('E1').getValue() + '] type=' + typeof sheet.getRange('E1').getValue());
  Logger.log('E2 direct: [' + sheet.getRange('E2').getValue() + '] type=' + typeof sheet.getRange('E2').getValue());
  Logger.log('E3 direct: [' + sheet.getRange('E3').getValue() + '] type=' + typeof sheet.getRange('E3').getValue());
  // Also log what getLastColumn sees
  var rows = sheet.getRange(1, 1, lastRow, 5).getValues();
  rows.forEach(function(r, i) {
    Logger.log('Row ' + i + ': r[4]=[' + r[4] + '] type=' + typeof r[4]);
  });
  Logger.log('isAdminEmail result: ' + isAdminEmail(testEmail));
}

function debugAdmin(params) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  var rows  = sheet.getDataRange().getValues();
  var email = (params.email || '').toLowerCase().trim();
  return {
    rangeAddress: sheet.getDataRange().getA1Notation(),
    totalRows: rows.length,
    rows: rows.map(function(r) {
      return {
        col_A: r[0], col_B: r[1], col_C: r[2], col_D: r[3], col_E: r[4], col_F: r[5],
        col_F_type: typeof r[5],
        emailMatch: (r[1] || '').toLowerCase().trim() === email,
        flagCheck: r[5] === true || String(r[5]).toUpperCase() === 'TRUE'
      };
    })
  };
}

function isAdminEmail(email) {
  const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  // Read A:F explicitly — getDataRange() misses col F when booleans are stored as checkboxes
  const rows = sheet.getRange(1, 1, lastRow, 6).getValues();
  rows.shift(); // remove header
  return rows.some(function(r) {
    const rowEmail = (r[1] || '').toLowerCase().trim();
    const flag     = r[5]; // column F = isAdmin
    return rowEmail === email.toLowerCase().trim() &&
           (flag === true || String(flag).toUpperCase() === 'TRUE');
  });
}

function sendAdminCode(params) {
  var email = (params.email || '').toLowerCase().trim();
  if (!email) return { success: false, error: 'Email required.' };
  if (!isAdminEmail(email)) return { success: false, error: 'Not authorized.' };

  var code   = Math.floor(100000 + Math.random() * 900000).toString();
  var expiry = new Date(Date.now() + 10 * 60 * 1000).toISOString();

  PropertiesService.getScriptProperties()
    .setProperty('admin_code_' + email, JSON.stringify({ code: code, expiry: expiry }));

  // Admin OTP always sends regardless of EMAIL_ENABLED (testing flag)
  MailApp.sendEmail({
    to: email,
    subject: 'Rally — Your Admin Access Code',
    name: 'MWF Tennis League',
    body: 'Your Rally admin access code is: ' + code +
          '\n\nThis code expires in 10 minutes.' +
          '\n\nIf you did not request this, please ignore this email.'
  });

  return { success: true };
}

function verifyAdminCode(params) {
  var email = (params.email || '').toLowerCase().trim();
  var code  = (params.code  || '').trim();
  if (!email || !code) return { success: false, error: 'Email and code required.' };

  var props  = PropertiesService.getScriptProperties();
  var stored = props.getProperty('admin_code_' + email);
  if (!stored) return { success: false, error: 'No code found. Please request a new one.' };

  var data = JSON.parse(stored);
  if (new Date() > new Date(data.expiry)) {
    props.deleteProperty('admin_code_' + email);
    return { success: false, error: 'Code expired. Please request a new one.' };
  }
  if (code !== data.code) return { success: false, error: 'Incorrect code. Please try again.' };

  props.deleteProperty('admin_code_' + email);
  return { success: true };
}

// ──────────────────────────────────────────────────
// COORDINATOR RATINGS
// ──────────────────────────────────────────────────

function getCoordinatorRatings(params) {
  var coordEmail = (params.email || '').toLowerCase().trim();
  var sheet      = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  var lastRow    = sheet.getLastRow();
  if (lastRow < 2) return { players: [] };

  var lastCol  = Math.max(sheet.getLastColumn(), 10); // ensure we read through col J (coordinators)
  var allData  = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers  = allData[0];

  // Find this coordinator's column (cols F–J = index 5–9)
  var coordColIdx = -1;
  for (var i = 5; i <= 9; i++) {
    if ((headers[i] || '').toString().toLowerCase().trim() === coordEmail) {
      coordColIdx = i;
      break;
    }
  }

  if (coordColIdx === -1) return { players: [], notAssigned: true };

  var players = [];
  for (var r = 1; r < allData.length; r++) {
    var row = allData[r];
    if (!row[0]) continue;
    var no8amVal = row[4];
    players.push({
      name:     row[0] || '',
      email:    (row[1] || '').toLowerCase(),
      myRating: row[coordColIdx] !== '' ? row[coordColIdx] : '',
      no8am:    no8amVal === true || (no8amVal && no8amVal.toString().toUpperCase() === 'TRUE')
    });
  }
  return { players: players, notAssigned: false };
}

function saveCoordinatorRatings(params) {
  var coordEmail = (params.coordEmail || '').toLowerCase().trim();
  var ratings    = JSON.parse(params.ratings || '[]'); // [{playerEmail, rating, no8am}]
  var sheet      = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  var lastRow    = sheet.getLastRow();
  var lastCol    = Math.max(sheet.getLastColumn(), 10);
  var allData    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers    = allData[0];

  // Find coordinator column — must be pre-assigned in sheet header (cols F–J)
  var coordColIdx = -1;
  for (var i = 5; i <= 9; i++) {
    if ((headers[i] || '').toString().toLowerCase().trim() === coordEmail) {
      coordColIdx = i; break;
    }
  }
  if (coordColIdx === -1) return { success: false, error: 'not_assigned' };

  // Build player email → row number map
  var emailToRow = {};
  for (var r = 1; r < allData.length; r++) {
    var e = (allData[r][1] || '').toLowerCase().trim();
    if (e) emailToRow[e] = r + 1; // 1-indexed sheet row
  }

  // Build rating + no8am lookup from input
  var ratingMap = {};
  var no8amMap = {};
  ratings.forEach(function(item) {
    var pe = (item.playerEmail || '').toLowerCase().trim();
    if (pe) {
      ratingMap[pe] = item.rating !== '' && item.rating !== null ? parseFloat(item.rating) : '';
      no8amMap[pe] = item.no8am === true || item.no8am === 'true';
    }
  });

  // Find all coordinator columns with data for average calculation
  var coordCols = [];
  for (var k = 5; k <= 9; k++) {
    if (headers[k]) coordCols.push(k);
  }

  // Update allData in memory, then batch-write ratings column + averages column
  for (var row = 1; row < allData.length; row++) {
    var pe = (allData[row][1] || '').toLowerCase().trim();
    if (pe && ratingMap.hasOwnProperty(pe)) {
      allData[row][coordColIdx] = ratingMap[pe];
      allData[row][4] = no8amMap[pe] ? true : false;
    }
    // Recalculate average from all coordinator columns
    if (!allData[row][0]) continue;
    var vals = coordCols.map(function(ci) {
      var v = allData[row][ci];
      return (v !== '' && !isNaN(parseFloat(v))) ? parseFloat(v) : null;
    }).filter(function(v) { return v !== null; });
    allData[row][3] = vals.length ? Math.round((vals.reduce(function(a,b){return a+b;},0) / vals.length) * 10) / 10 : '';
  }

  // Batch write: ratings column
  var ratingsCol = allData.slice(1).map(function(r) { return [r[coordColIdx]]; });
  var ratingRange = sheet.getRange(2, coordColIdx + 1, ratingsCol.length, 1);
  ratingRange.setNumberFormat('0.0');
  ratingRange.setValues(ratingsCol);

  // Batch write: averages column (D)
  var avgsCol = allData.slice(1).map(function(r) { return [r[3]]; });
  sheet.getRange(2, 4, avgsCol.length, 1).setValues(avgsCol);

  // Batch write: No 8am column (E)
  if (!headers[4] || headers[4].toString() !== 'No8am') {
    sheet.getRange(1, 5).setValue('No8am');
  }
  var no8amCol = allData.slice(1).map(function(r) { return [r[4] === true]; });
  sheet.getRange(2, 5, no8amCol.length, 1).setValues(no8amCol);

  return { success: true };
}

function updateRequestTime(params) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
  const cell  = sheet.getRange(parseInt(params.rowIndex), 6); // column F = matchTime
  cell.setNumberFormat('@');
  cell.setValue(params.matchTime || '');
  return { success: true };
}

function confirmSub(params) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 1. Update SubRequests tab
  const reqSheet = ss.getSheetByName(TABS.requests);
  reqSheet.getRange(parseInt(params.requestRowIndex), 7).setValue('filled');
  reqSheet.getRange(parseInt(params.requestRowIndex), 8).setValue(params.subEmail);

  // 2. Update Volunteers tab if rowIndex provided
  if (params.volunteerRowIndex) {
    const volSheet = ss.getSheetByName(TABS.volunteers);
    volSheet.getRange(parseInt(params.volunteerRowIndex), 7).setValue('matched');
  }

  // 3. Parse group players
  var groupPlayers = [];
  try { groupPlayers = JSON.parse(params.groupPlayers || '[]'); } catch(e) {}

  // 4. Send email
  sendConfirmationEmails(params, groupPlayers);

  return { success: true };
}

// ──────────────────────────────────────────────────
// MATCHING ENGINE (server-side)
// ──────────────────────────────────────────────────

function runMatch(params) {
  const config     = getConfig();
  const requests   = getRequests();
  const volunteers = getVolunteers();
  const players    = getPlayersWithRatings();

  const req = requests.find(r => r.id === params.requestId);
  if (!req) return { error: 'Request not found' };

  const reqPlayer = players.find(p => p.email === req.email.toLowerCase());
  if (!reqPlayer) return { error: 'Requestor not found in Players sheet' };

  const reqRating    = reqPlayer.rating;
  const matchDate    = req.matchDate;
  const matchTime    = req.matchTime;
  const lastMinute   = isLastMinute(req, config.lastMinuteThresholdHrs);
  const urgent       = isUrgent(req, config.preScheduleThresholdHrs);
  const hasTBDTime      = !matchTime;
  const effectiveTime   = (matchTime || '08:00').trim();

  const skillWindow     = lastMinute ? Infinity : (!urgent ? config.skillWindowPreSchedule : config.skillWindowPostSchedule);
  const requireAllTimes = !lastMinute && !urgent && !hasTBDTime;
  const phase           = lastMinute ? 'last-minute' : (!urgent ? 'pre-schedule' : 'post-schedule');

  let candidates = volunteers.filter(v => {
    if (v.date.trim() !== matchDate.trim()) return false;
    if (v.email.toLowerCase() === req.email.toLowerCase()) return false;
    if (v.status === 'matched' || v.status === 'cancelled' || v.status === 'expired') return false;
    const volTimes = v.times.map(t => t.trim());
    if (requireAllTimes) {
      if (!TIMES.every(t => volTimes.includes(t))) return false;
    } else {
      if (!volTimes.includes(effectiveTime)) return false;
    }
    if (skillWindow !== Infinity) {
      const vol = players.find(p => p.email.toLowerCase() === v.email.toLowerCase());
      if (!vol) return false;
      if (Math.abs(vol.rating - reqRating) > skillWindow) return false;
    }
    const alreadyAssigned = requests.some(r =>
      r.assignedSub && r.assignedSub.toLowerCase() === v.email.toLowerCase() &&
      r.matchDate === matchDate && r.status === 'filled' &&
      (!matchTime || !r.matchTime || r.matchTime === matchTime)
    );
    if (alreadyAssigned) return false;
    const alreadyPlaying = requests.some(r =>
      r.email.toLowerCase() === v.email.toLowerCase() &&
      r.matchDate === matchDate && r.status !== 'open' &&
      (!matchTime || !r.matchTime || r.matchTime === matchTime)
    );
    if (alreadyPlaying) return false;
    return true;
  });

  // Deduplicate by email, keep earliest submission
  const seen = new Map();
  candidates.forEach(c => {
    if (!seen.has(c.email) || c.timestamp < seen.get(c.email).timestamp) {
      seen.set(c.email, c);
    }
  });
  candidates = Array.from(seen.values());

  // Enrich with rating diff
  candidates = candidates.map(c => {
    const p = players.find(p => p.email === c.email.toLowerCase());
    return {
      ...c,
      ratingDiff: p ? Math.abs(p.rating - reqRating) : 99
    };
    // Note: rating itself is NOT included — only the diff
  });

  // Sort: closest rating first, then earliest submission
  candidates.sort((a, b) => {
    if (a.ratingDiff !== b.ratingDiff) return a.ratingDiff - b.ratingDiff;
    return a.timestamp.localeCompare(b.timestamp);
  });

  return {
    candidates: candidates.slice(0, 5),
    skillWindow: skillWindow === Infinity ? 'none' : skillWindow,
    requireAllTimes,
    phase,
    matchTime: matchTime ? TIME_LABELS[matchTime] : null
  };
}

// ──────────────────────────────────────────────────
// EMAIL
// ──────────────────────────────────────────────────

function sendConfirmationEmails(data, groupPlayers) {
  groupPlayers = groupPlayers || [];
  const dateStr    = formatDate(data.matchDate);
  const timeStr    = data.matchTime ? TIME_LABELS[data.matchTime] : 'TBD';
  const senderName = 'MWF Tennis League';

  // To: requestor + sub   CC: group partners
  const toAddresses = [data.requestorEmail, data.subEmail].filter(Boolean).join(', ');
  const ccList      = groupPlayers.map(function(p) { return p.email; }).filter(Boolean);
  const ccAddresses = ccList.join(', ');

  const subject =
    'MWF Tennis League — Substitute confirmed: ' + data.subName + ' for ' + data.requestorName;

  const body =
    'Hi team,\n\n' +
    data.subName + ' will be substituting for ' + data.requestorName +
    ' on ' + dateStr + ' at ' + timeStr + '.\n\n' +
    'Make updates in Chelsea as required.\n\n' +
    'See you on the court!\n\n' +
    'MWF Tennis League';

  var emailParams = {
    to:      toAddresses,
    subject: subject,
    body:    body,
    name:    senderName
  };
  if (ccAddresses) emailParams.cc = ccAddresses;

  if (isEmailEnabled()) MailApp.sendEmail(emailParams);
}

function saveMatchTimeReminderSettings(params) {
  var sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);
  var enabled = params.enabled === 'true' || params.enabled === true;
  var time    = (params.time || '10:00').trim();

  sheet.getRange('B28').setValue(enabled);
  var timeCell = sheet.getRange('B29');
  timeCell.setNumberFormat('@');
  timeCell.setValue(time);
  SpreadsheetApp.flush();

  try { updateMatchTimeReminderTrigger(enabled, time); } catch(e) { Logger.log('updateMatchTimeReminderTrigger error: ' + e.message); }

  return { success: true, matchTimeReminderEnabled: enabled, matchTimeReminderTimeET: time };
}

function updateMatchTimeReminderTrigger(enabled, time) {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runMatchTimeReminder') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  if (!enabled) return;
  var parts  = time.split(':');
  var hourET = parseInt(parts[0]);
  var minET  = parseInt(parts[1]) || 0;
  ScriptApp.newTrigger('runMatchTimeReminder')
    .timeBased().atHour(hourET).nearMinute(minET).everyDays(1)
    .inTimezone('America/New_York').create();
}

function runMatchTimeReminder() {
  var config = getConfig();
  if (!config.matchTimeReminderEnabled) return { skipped: 'disabled' };

  var requests = getRequests();
  var now      = new Date();
  var siteUrl  = 'https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-prod.html#request';
  var notified = 0;

  requests.forEach(function(req) {
    if (req.status !== 'open') return;
    if (req.matchTime) return; // already has a time

    // Check if match date is within 60 hours (use 8:00 AM for TBD times)
    var matchDT = new Date(req.matchDate + 'T08:00:00');
    var diffHrs = (matchDT - now) / 36e5;
    if (diffHrs <= 0 || diffHrs > 60) return;

    var dateStr = formatDate(req.matchDate);
    var subject = 'MWF Tennis League — Court time needed for your sub request: ' + dateStr;

    var body =
      'Hi ' + req.name + ',\n\n' +
      'You have an open sub request for ' + dateStr + ' and no court time has been assigned yet.\n\n' +
      'Once Chelsea has scheduled a court, please add the court time to your request at:\n' + siteUrl + '\n\n' +
      'If you are on Overflow, do nothing. Rally will still try to find a sub.\n\n' +
      'Note: Non 8am players are ineligible to fill a sub request without a court time assigned.\n\n' +
      'MWF Tennis League';

    var htmlBody =
      'Hi ' + req.name + ',<br><br>' +
      'You have an open sub request for <strong>' + dateStr + '</strong> and no court time has been assigned yet.<br><br>' +
      'Once Chelsea has scheduled a court, please <a href="' + siteUrl + '">add the court time to your request</a>.<br><br>' +
      '<em>If you are on Overflow, do nothing. Rally will still try to find a sub.</em><br><br>' +
      '<em>Note: Non 8am players are ineligible to fill a sub request without a court time assigned.</em><br><br>' +
      'MWF Tennis League';

    var groupPlayers = req.groupPlayers || [];
    var ccList = groupPlayers.map(function(p) { return p.email; }).filter(Boolean);
    var emailParams = {
      to:       req.email,
      subject:  subject,
      body:     body,
      htmlBody: htmlBody,
      name:     'MWF Tennis League'
    };
    if (ccList.length) emailParams.cc = ccList.join(', ');
    if (isEmailEnabled()) MailApp.sendEmail(emailParams);
    notified++;
  });

  Logger.log('runMatchTimeReminder: notified ' + notified + ' requestor(s).');
  return { success: true, notified: notified };
}

function saveAutoDispatchSettings(params) {
  var sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);
  var enabled = params.enabled === 'true' || params.enabled === true;
  var time    = (params.time || '08:00').trim();

  sheet.getRange('B13').setValue(enabled);
  var timeCell = sheet.getRange('B14');
  timeCell.setNumberFormat('@');
  timeCell.setValue(time);
  SpreadsheetApp.flush();

  try { updateDispatchTrigger(enabled, time); } catch(e) { Logger.log('updateDispatchTrigger error: ' + e.message); }

  return { success: true, autoDispatchEnabled: enabled, autoDispatchTimeET: time };
}

function sendRetirementEmail(req) {
  var dateStr      = formatDate(req.matchDate);
  var timeStr      = req.matchTime ? TIME_LABELS[req.matchTime] : 'TBD';
  var subject      = 'MWF Tennis League — Unable to find substitute: ' + dateStr + (req.matchTime ? ' at ' + timeStr : '');
  var directoryUrl = 'https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-prod.html#directory';
  var body =
    'Hi ' + req.name + ',\n\n' +
    'Unfortunately, we were unable to find a volunteer to fill the sub request for your match:\n\n' +
    '  Date: ' + dateStr + '\n' +
    '  Time: ' + timeStr + '\n\n' +
    'Player email addresses and phone numbers can be found on the Directory page: ' + directoryUrl + '\n\n' +
    'MWF Tennis League';
  var htmlBody =
    'Hi ' + req.name + ',<br><br>' +
    'Unfortunately, we were unable to find a volunteer to fill the sub request for your match:<br><br>' +
    '&nbsp;&nbsp;Date: ' + dateStr + '<br>' +
    '&nbsp;&nbsp;Time: ' + timeStr + '<br><br>' +
    'Player email addresses and phone numbers can be found on the <a href="' + directoryUrl + '">Directory</a> page.<br><br>' +
    'MWF Tennis League';
  var groupPlayers = req.groupPlayers || [];
  var ccList = groupPlayers.map(function(p) { return p.email; }).filter(Boolean);
  var emailParams = { to: req.email, subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' };
  if (ccList.length) emailParams.cc = ccList.join(', ');
  if (isEmailEnabled()) MailApp.sendEmail(emailParams);
}

function retireRequest(params) {
  var requests = getRequests();
  var req = requests.find(function(r) { return r.id === params.requestId; });
  if (!req) return { success: false, error: 'Request not found' };

  var reqSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
  reqSheet.getRange(parseInt(req.rowIndex), 7).setValue('expired');
  sendRetirementEmail(req);

  return { success: true };
}

function expireUpToToday() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var today = formatSheetDate(new Date());

  var reqSheet = ss.getSheetByName(TABS.requests);
  if (reqSheet && reqSheet.getLastRow() >= 2) {
    var reqRows = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < reqRows.length; i++) {
      var matchDate = formatSheetDate(reqRows[i][4]);
      var status    = (reqRows[i][6] || '').toString();
      if (matchDate && matchDate <= today && status === 'open') {
        reqSheet.getRange(i + 2, 7).setValue('expired');
      }
    }
  }

  var volSheet = ss.getSheetByName(TABS.volunteers);
  if (volSheet && volSheet.getLastRow() >= 2) {
    var volRows = volSheet.getRange(2, 1, volSheet.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < volRows.length; i++) {
      var volDate = formatSheetDate(volRows[i][4]);
      var status  = (volRows[i][6] || '').toString();
      if (volDate && volDate <= today && status === 'pending') {
        volSheet.getRange(i + 2, 7).setValue('expired');
      }
    }
  }
}

// ──────────────────────────────────────────────────
// HELPERS
// ──────────────────────────────────────────────────

function isUrgent(req, thresholdHrs) {
  if (!req.matchDate) return false;
  const hrs     = thresholdHrs || 48;
  const timeStr = req.matchTime || '08:00'; // TBD: treat as 8:00 AM
  const matchDT = new Date(req.matchDate + 'T' + timeStr + ':00');
  const now     = new Date();
  const diffHrs = (matchDT - now) / 36e5;
  return diffHrs <= hrs && diffHrs > 0;
}

function isLastMinute(req, thresholdHrs) {
  if (!req.matchDate) return false;
  const hrs     = thresholdHrs || 24;
  const timeStr = req.matchTime || '08:00'; // TBD: treat as 8:00 AM
  const matchDT = new Date(req.matchDate + 'T' + timeStr + ':00');
  const now     = new Date();
  const diffHrs = (matchDT - now) / 36e5;
  // Past matches (diffHrs <= 0) are treated as last-minute so open requests
  // remain matchable even after the scheduled time
  return diffHrs <= hrs;
}

function isDST() {
  const now = new Date();
  const jan = new Date(now.getFullYear(), 0, 1);
  const jul = new Date(now.getFullYear(), 6, 1);
  const stdOffset = Math.max(jan.getTimezoneOffset(), jul.getTimezoneOffset());
  return now.getTimezoneOffset() < stdOffset;
}

function formatDate(str) {
  if (!str) return '—';
  const d = new Date(str + 'T12:00:00');
  return d.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
}

function getDayOfWeek(str) {
  if (!str) return 'day';
  const d = new Date(str + 'T12:00:00');
  return d.toLocaleDateString('en-US', { weekday: 'long' });
}

function uid() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

// ──────────────────────────────────────────────────
// AVAILABILITY
// ──────────────────────────────────────────────────

function getAvailabilityConfig() {
  const config   = getConfig();
  const today    = new Date();
  today.setHours(0, 0, 0, 0);

  const openDate  = config.availWindowOpenDate  ? new Date(config.availWindowOpenDate  + 'T00:00:00') : null;
  const closeDate = config.availWindowCloseDate ? new Date(config.availWindowCloseDate + 'T00:00:00') : null;

  // Auto-close if close date has passed
  let isOpen = config.availWindowActive;
  if (isOpen && closeDate && today > closeDate) {
    isOpen = false;
    // Write FALSE back to sheet to keep it in sync
    try {
      SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config).getRange('B18').setValue(false);
    } catch(e) {}
  }

  // Derive target month from the open date (or next month if no date set)
  let targetMonth, targetMonthLabel;
  if (openDate) {
    // Target month = month after the open date's month (the month players are scheduling for)
    const t = new Date(openDate.getFullYear(), openDate.getMonth() + 1, 1);
    targetMonth      = t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0');
    targetMonthLabel = t.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  } else {
    const t = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    targetMonth      = t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0');
    targetMonthLabel = t.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  }

  return {
    isOpen:           isOpen,
    openDate:         config.availWindowOpenDate  || '',
    closeDate:        config.availWindowCloseDate || '',
    targetMonth:      targetMonth,
    targetMonthLabel: targetMonthLabel
  };
}

// Returns players from the Players sheet who have NOT submitted availability
// for the given month (e.g. "2026-05").
function getPlayersWithoutSubmission(month) {
  var players = getPlayers(); // [{name, email, ...}]
  if (!players.length) return [];

  var avSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.availability);
  var submitted = {};
  if (avSheet && avSheet.getLastRow() >= 2) {
    var rows = avSheet.getRange(2, 1, avSheet.getLastRow() - 1, 6).getValues();
    rows.forEach(function(r) {
      if (normalizeMonth(r[3]) === month) {
        var em = (r[2] || '').toLowerCase();
        if (em) submitted[em] = true;
      }
    });
  }

  return players.filter(function(p) {
    return p.email && !submitted[p.email.toLowerCase()];
  });
}

// Runs daily at 1 AM to enforce the close date and send T-2 / T-1 reminders.
function checkAvailabilityWindow() {
  var config = getAvailabilityConfig();
  // getAvailabilityConfig already writes B18=false when past close date
  Logger.log('checkAvailabilityWindow: isOpen=' + config.isOpen + ' closeDate=' + config.closeDate);

  // Only send reminders while the window is open and a close date is set
  if (!config.isOpen || !config.closeDate) return;

  var today     = new Date();
  today.setHours(0, 0, 0, 0);
  var closeDate = new Date(config.closeDate + 'T00:00:00');
  var daysUntilClose = Math.round((closeDate - today) / 864e5);

  if (daysUntilClose !== 2 && daysUntilClose !== 1) return;

  var missing = getPlayersWithoutSubmission(config.targetMonth);
  if (!missing.length) {
    Logger.log('checkAvailabilityWindow: T-' + daysUntilClose + ' reminder — all players already submitted');
    return;
  }

  var closeDateLabel = closeDate.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
  var urgency        = daysUntilClose === 1 ? 'tomorrow' : 'in 2 days';
  var subject        = 'Reminder: Submit your availability for ' + config.targetMonthLabel + ' — closes ' + urgency;
  var body =
    'Hi,\n\n' +
    'Just a reminder — the availability window for ' + config.targetMonthLabel + ' closes ' + urgency + ' (' + closeDateLabel + ').\n\n' +
    'We haven\'t received your availability yet. Please submit before the window closes so we can include you in the schedule.\n\n' +
    'Open the Rally app to submit:\n' +
    'https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-prod.html\n\n' +
    'See you on the court!\n' +
    'MWF Tennis League';

  var emails = missing.map(function(p) { return p.email; }).filter(Boolean);
  Logger.log('checkAvailabilityWindow: T-' + daysUntilClose + ' reminder → ' + emails.length + ' player(s): ' + emails.join(', '));
  if (isEmailEnabled()) {
    MailApp.sendEmail({ to: emails.join(', '), subject: subject, body: body, name: 'MWF Tennis League' });
  }
}

function testAvailabilityEmail() {
  var config = getAvailabilityConfig();
  var closeDateLabel = 'Friday, April 25';
  var subject = '[TEST] MWF League - Submit your availability for ' + config.targetMonthLabel;
  var body =
    'Hi,\n\n' +
    'It\'s time to submit your availability for ' + config.targetMonthLabel + '.\n\n' +
    'Please submit your available dates by ' + closeDateLabel + '.\n\n' +
    'Open the Rally app to get started:\n' +
    'https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-prod.html\n\n' +
    'See you on the court!\n' +
    'MWF Tennis League';
  MailApp.sendEmail({ to: 'brianna.biesecker@gmail.com, marobria@gmail.com', subject: subject, body: body, name: 'MWF Tennis League' });
  return { success: true, sent: 'brianna.biesecker@gmail.com, marobria@gmail.com' };
}

function openAvailabilityWindow(params) {
  const closeDate = params.closeDate;
  if (!closeDate) return { success: false, error: 'A close date is required.' };

  // Open date = today (the day the coordinator clicks the button)
  const today = new Date();
  const openDate = today.getFullYear() + '-' +
    String(today.getMonth() + 1).padStart(2, '0') + '-' +
    String(today.getDate()).padStart(2, '0');

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);
  sheet.getRange('B16').setValue(openDate);
  sheet.getRange('B17').setValue(closeDate);
  sheet.getRange('B18').setValue(true);

  // Send email blast to all players
  const players = getPlayers();
  const emails  = players.map(function(p) { return p.email; }).filter(Boolean);
  if (emails.length) {
    const config          = getAvailabilityConfig();
    const closeDateLabel  = new Date(closeDate + 'T12:00:00').toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
    const subject         = 'MWF League - Submit your availability for ' + config.targetMonthLabel;
    const body =
      'Hi,\n\n' +
      'It\'s time to submit your availability for ' + config.targetMonthLabel + '.\n\n' +
      'Please submit your available dates by ' + closeDateLabel + '.\n\n' +
      'Open the Rally app to get started:\n' +
      'https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-prod.html\n\n' +
      'See you on the court!\n' +
      'MWF Tennis League';
    if (isEmailEnabled()) MailApp.sendEmail({ to: emails.join(', '), subject: subject, body: body, name: 'MWF Tennis League' });
  }

  return { success: true, playerCount: emails.length };
}

function closeAvailabilityWindow() {
  SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config).getRange('B18').setValue(false);
  return { success: true };
}

function getOrCreateAvailabilitySheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.availability);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.availability);
    sheet.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Name', 'Email', 'Month', 'AvailableDates', 'Notes']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Normalize a Sheets cell value to "YYYY-MM" string regardless of how Sheets stored it
// Parses the AvailableDates cell and always returns ["YYYY-MM-DD", ...].
// Handles both the legacy [{date, times}] object format and the current string-array format.
function parseDatesField(jsonStr) {
  var parsed = [];
  try { parsed = JSON.parse(jsonStr || '[]'); } catch(e) { return []; }
  if (!Array.isArray(parsed) || !parsed.length) return [];
  if (typeof parsed[0] === 'object' && parsed[0] !== null) {
    // Legacy format: [{date: "YYYY-MM-DD", times: [...]}]
    return parsed.map(function(d) { return d.date || ''; }).filter(Boolean);
  }
  // Current format: ["YYYY-MM-DD", ...]
  return parsed.filter(function(d) { return typeof d === 'string' && d.length === 10; });
}

function normalizeMonth(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '-' + String(val.getMonth() + 1).padStart(2, '0');
  }
  return String(val).trim().slice(0, 7); // take first 7 chars of "YYYY-MM..." just in case
}

function submitAvailability(params) {
  const name           = params.name           || '';
  const email          = (params.email         || '').toLowerCase();
  const month          = params.month          || '';
  const availableDates = params.availableDates || '[]';
  const notes          = params.notes          || '';

  Logger.log('submitAvailability called: name=%s email=%s month=%s dates=%s', name, email, month, availableDates);

  if (!name || !email || !month) return { success: false, error: 'Missing required fields.' };

  // Validate window is still open
  const avConfig = getAvailabilityConfig();
  if (!avConfig.isOpen) return { success: false, error: 'The availability window is currently closed.' };

  const sheet   = getOrCreateAvailabilitySheet();
  const lastRow = sheet.getLastRow();

  // Upsert: find existing row for this email + month
  let targetRow = -1;
  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    rows.forEach(function(r, i) {
      if ((r[2] || '').toLowerCase() === email && normalizeMonth(r[3]) === month) {
        targetRow = i + 2;
      }
    });
  }

  const timestamp = new Date().toISOString();
  const rowData   = [timestamp, name, email, month, availableDates, notes];

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, 6).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  // Confirmation email to the player
  try {
    const dates     = parseDatesField(availableDates);
    const dateLines = dates.map(function(d) {
      return '  ' + new Date(d + 'T12:00:00').toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
    }).join('\n');

    const subject = 'MWF League - Your availability for ' + avConfig.targetMonthLabel + ' is confirmed';
    const body =
      'Hi ' + name + ',\n\n' +
      'We received your availability for ' + avConfig.targetMonthLabel + '.\n\n' +
      'Your selected dates:\n' + (dateLines || '  (none selected)') + '\n\n' +
      (notes ? 'Notes: ' + notes + '\n\n' : '') +
      'If you need to make changes, you can re-submit before the window closes.\n\n' +
      'See you on the court!\n' +
      'MWF Tennis League';

    if (isEmailEnabled()) MailApp.sendEmail({ to: email, subject: subject, body: body, name: 'MWF Tennis League' });
  } catch(err) {
    Logger.log('Confirmation email failed: ' + err.message);
  }

  return { success: true };
}

function getMyAvailability(params) {
  const email = (params.email || '').toLowerCase();
  const month = params.month  || '';
  if (!email || !month) return null;

  const sheet   = getOrCreateAvailabilitySheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const row  = rows.find(function(r) {
    return (r[2] || '').toLowerCase() === email && normalizeMonth(r[3]) === month;
  });

  if (!row) return null;

  return {
    timestamp:      row[0] ? new Date(row[0]).toISOString() : '',
    name:           row[1] || '',
    email:          row[2] || '',
    month:          row[3] || '',
    availableDates: parseDatesField(row[4]),
    notes:          row[5] || ''
  };
}

// Combined fetch: returns availability config + optional existing submission in one call.
// Pass email= to also get the player's submission for the target month.
function getAvailabilityData(params) {
  const config = getAvailabilityConfig();
  const result = { config: config };
  const email  = (params.email || '').toLowerCase();
  if (email && config.targetMonth) {
    result.submission = getMyAvailability({ email: email, month: config.targetMonth });
  }
  return result;
}

function cleanupOldAvailability() {
  const sheet   = getOrCreateAvailabilitySheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const now       = new Date();
  const cutoff    = new Date(now.getTime() - 60 * 24 * 60 * 60 * 1000); // 60 days ago
  const rows      = sheet.getRange(2, 1, lastRow - 1, 1).getValues();   // timestamp column only
  var   deleted   = 0;

  // Delete bottom-up to avoid row index shifting
  for (var i = rows.length - 1; i >= 0; i--) {
    var ts = rows[i][0];
    if (!ts) continue;
    var submitted = (ts instanceof Date) ? ts : new Date(ts);
    if (submitted < cutoff) {
      sheet.deleteRow(i + 2);
      deleted++;
    }
  }
  Logger.log('cleanupOldAvailability: deleted ' + deleted + ' row(s) older than 60 days.');
}

// ══════════════════════════════════════════════════
// SCHEDULER
// ══════════════════════════════════════════════════

// ── Settings ──────────────────────────────────────
// Reads scheduler weight rows from Config tab (B20–B25).
// Coordinators can tune these directly in the sheet.
function getSchedulerSettings() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var configSheet = ss.getSheetByName(TABS.config);
    var raw = configSheet.getRange('B20:B25').getValues();
    var wTV   = parseFloat(raw[0][0]);
    var wGV   = parseFloat(raw[1][0]);
    var wSV   = parseFloat(raw[2][0]);
    var wRec  = parseFloat(raw[3][0]);
    var iters = parseInt(raw[4][0]);
    var rests = parseInt(raw[5][0]);
    var settings = {
      weightTeamVariance:  isNaN(wTV)   ? 1.0 : wTV,
      weightGroupVariance: isNaN(wGV)   ? 0.5 : wGV,
      weightSocialVariety: isNaN(wSV)   ? 2.0 : wSV,
      weightRecency:       isNaN(wRec)  ? 1.5 : wRec,
      solverIterations:    isNaN(iters) ? 800  : iters,
      solverRestarts:      isNaN(rests) ? 10   : rests
    };

    var availConfig = getAvailabilityConfig();
    var targetMonth = availConfig.targetMonth;
    var submissionCount = 0;
    if (targetMonth) {
      var avSheet = ss.getSheetByName(TABS.availability);
      if (avSheet && avSheet.getLastRow() >= 2) {
        var avRows = avSheet.getRange(2, 3, avSheet.getLastRow() - 1, 2).getValues();
        var seen = {};
        avRows.forEach(function(r) {
          var email = (r[0] || '').toLowerCase();
          var mon   = normalizeMonth(r[1]);
          if (email && mon === targetMonth && !seen[email]) {
            seen[email] = true;
            submissionCount++;
          }
        });
      }
    }
    // Count total roster size
    var playersSheet = ss.getSheetByName(TABS.players);
    var rosterCount = 0;
    if (playersSheet && playersSheet.getLastRow() >= 2) {
      rosterCount = playersSheet.getLastRow() - 1;
    }

    settings.targetMonth      = targetMonth;
    settings.targetMonthLabel = availConfig.targetMonthLabel;
    settings.submissionCount  = submissionCount;
    settings.rosterCount      = rosterCount;
    return settings;
  } catch(e) {
    return {
      weightTeamVariance:  1.0,
      weightGroupVariance: 0.5,
      weightSocialVariety: 2.0,
      weightRecency:       1.5,
      solverIterations:    800,
      solverRestarts:      10,
      targetMonth:         '',
      targetMonthLabel:    '',
      submissionCount:     0
    };
  }
}

// ── Combined Scheduler Dashboard ──────────────────
// Single endpoint returning both availability config and scheduler settings.
// Eliminates redundant getConfig() calls from separate endpoints.
function getSchedulerDashboard() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var configSheet = ss.getSheetByName(TABS.config);

    // Read availability window state (B16–B18)
    var openDateRaw  = configSheet.getRange('B16').getValue();
    var closeDateRaw = configSheet.getRange('B17').getValue();
    var activeRaw    = configSheet.getRange('B18').getValue();

    var openDate  = openDateRaw instanceof Date ? formatSheetDate(openDateRaw) : (openDateRaw ? openDateRaw.toString() : '');
    var closeDate = closeDateRaw instanceof Date ? formatSheetDate(closeDateRaw) : (closeDateRaw ? closeDateRaw.toString() : '');
    var isOpen    = activeRaw === true || activeRaw.toString().toUpperCase() === 'TRUE';

    // Auto-close if past close date
    var today = new Date(); today.setHours(0,0,0,0);
    var closeDateObj = closeDate ? new Date(closeDate + 'T00:00:00') : null;
    if (isOpen && closeDateObj && today > closeDateObj) {
      isOpen = false;
      configSheet.getRange('B18').setValue(false);
    }

    // Target month
    var targetMonth, targetMonthLabel;
    var openDateObj = openDate ? new Date(openDate + 'T00:00:00') : null;
    if (openDateObj) {
      var t = new Date(openDateObj.getFullYear(), openDateObj.getMonth() + 1, 1);
      targetMonth      = t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0');
      targetMonthLabel = t.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    } else {
      var t = new Date(today.getFullYear(), today.getMonth() + 1, 1);
      targetMonth      = t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0');
      targetMonthLabel = t.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    }

    // Scheduler weights (B20–B25)
    var raw = configSheet.getRange('B20:B25').getValues();
    var wTV   = parseFloat(raw[0][0]);
    var wGV   = parseFloat(raw[1][0]);
    var wSV   = parseFloat(raw[2][0]);
    var wRec  = parseFloat(raw[3][0]);
    var iters = parseInt(raw[4][0]);
    var rests = parseInt(raw[5][0]);

    // Submission count
    var submissionCount = 0;
    if (targetMonth) {
      var avSheet = ss.getSheetByName(TABS.availability);
      if (avSheet && avSheet.getLastRow() >= 2) {
        var avRows = avSheet.getRange(2, 3, avSheet.getLastRow() - 1, 2).getValues();
        var seen = {};
        avRows.forEach(function(r) {
          var email = (r[0] || '').toLowerCase();
          var mon   = normalizeMonth(r[1]);
          if (email && mon === targetMonth && !seen[email]) {
            seen[email] = true;
            submissionCount++;
          }
        });
      }
    }

    // Roster count + no8am emails
    var playersSheet = ss.getSheetByName(TABS.players);
    var rosterCount = 0;
    var no8amEmails = [];
    if (playersSheet && playersSheet.getLastRow() >= 2) {
      rosterCount = playersSheet.getLastRow() - 1;
      var pRows = playersSheet.getRange(2, 1, rosterCount, 5).getValues();
      pRows.forEach(function(r) {
        var email = (r[1] || '').toLowerCase().trim();
        var flag  = r[4];
        if (email && (flag === true || (flag && flag.toString().toUpperCase() === 'TRUE'))) {
          no8amEmails.push(email);
        }
      });
    }

    return {
      isOpen: isOpen,
      openDate: openDate,
      closeDate: closeDate,
      targetMonth: targetMonth,
      targetMonthLabel: targetMonthLabel,
      submissionCount: submissionCount,
      rosterCount: rosterCount,
      no8amEmails: no8amEmails,
      weightTeamVariance:  isNaN(wTV)   ? 1.0 : wTV,
      weightGroupVariance: isNaN(wGV)   ? 0.5 : wGV,
      weightSocialVariety: isNaN(wSV)   ? 2.0 : wSV,
      weightRecency:       isNaN(wRec)  ? 1.5 : wRec,
      solverIterations:    isNaN(iters) ? 800  : iters,
      solverRestarts:      isNaN(rests) ? 10   : rests
    };
  } catch(e) {
    return { error: 'Could not load scheduler dashboard.' };
  }
}

// ── Generate ───────────────────────────────────────
// Reads availability submissions for targetMonth, joins player ratings,
// and runs the local-search optimizer for each date+time slot that has
// enough available players (≥3). Returns an array of slot results.
//
// params.month     — "YYYY-MM" to schedule (defaults to next month)
// params.pairCounts — JSON string of { "email|email": N } from prior sessions (optional)
// params.sitOutCounts — JSON string of { "email": N } (optional)
function generateSchedule(params) {
  var month = params.month || '';
  var pairCounts   = safeParseJSON(params.pairCounts,   {});
  var sitOutCounts = safeParseJSON(params.sitOutCounts, {});

  // Fall back to target month from availability config
  if (!month) {
    month = getAvailabilityConfig().targetMonth;
  }
  if (!month) return { error: 'No target month available.' };

  // Load players with ratings (internal)
  var players = getPlayersWithRatings(); // [{ name, email, rating }]
  var playerMap = {};
  players.forEach(function(p) { playerMap[p.email.toLowerCase()] = p; });

  // Load availability submissions for this month
  var avSheet  = getOrCreateAvailabilitySheet();
  var lastRow  = avSheet.getLastRow();
  if (lastRow < 2) return { error: 'No availability submissions found for ' + month + '.' };

  var avRows = avSheet.getRange(2, 1, lastRow - 1, 6).getValues();

  // Group submissions by month, keyed by email
  // Each row: [timestamp, name, email, month, availableDatesJSON, notes]
  var submissionsByEmail = {};
  avRows.forEach(function(r) {
    var rowMonth = normalizeMonth(r[3]);
    if (rowMonth !== month) return;
    var email = (r[2] || '').toLowerCase();
    if (!email) return;
    submissionsByEmail[email] = {
      name:   r[1] || '',
      email:  email,
      rating: playerMap[email] ? playerMap[email].rating : 0,
      dates:  parseDatesField(r[4])  // ["YYYY-MM-DD", ...]
    };
  });

  var emailList = Object.keys(submissionsByEmail);
  if (!emailList.length) return { error: 'No submissions found for ' + month + '.' };

  // Build a map of { "YYYY-MM-DD": [player, ...] } for each available date
  var slotMap = {};
  emailList.forEach(function(email) {
    var sub = submissionsByEmail[email];
    (sub.dates || []).forEach(function(date) {
      if (!slotMap[date]) slotMap[date] = [];
      slotMap[date].push(sub);
    });
  });

  // Sort slots chronologically
  var slotKeys = Object.keys(slotMap).sort();

  var settings = getSchedulerSettings();

  var slotResults = [];
  slotKeys.forEach(function(slotKey) {
    var date      = slotKey;
    var available = slotMap[slotKey];

    if (available.length < 3) {
      // Not enough for even one group of 3 — skip
      slotResults.push({
        date: date,
        skipped: true,
        reason: 'Only ' + available.length + ' player(s) available — need at least 3.'
      });
      return;
    }

    var result = optimizeSlot(available, settings, pairCounts, sitOutCounts);
    slotResults.push({ date: date, skipped: false, groups: result.groups, sitOut: result.sitOut });

    // Update running pairCounts and sitOutCounts for subsequent slots in the same run
    result.groups.forEach(function(group) {
      for (var i = 0; i < group.length; i++) {
        for (var j = i + 1; j < group.length; j++) {
          var key = pairKey(group[i].email, group[j].email);
          pairCounts[key] = (pairCounts[key] || 0) + 1;
        }
      }
    });
    if (result.sitOut) {
      sitOutCounts[result.sitOut.email] = (sitOutCounts[result.sitOut.email] || 0) + 1;
    }
  });

  assignCaptains(slotResults);

  return {
    month:          month,
    submissionCount: emailList.length,
    slots:          slotResults,
    pairCounts:     pairCounts,
    sitOutCounts:   sitOutCounts
  };
}

// ── Captain Assignment ──────────────────────────────
// Assigns one captain per group so each player is captain ~25% of their scheduled dates.
// Adds slot.captains = [emailForGroupA, emailForGroupB, ...] to every active slot.
function assignCaptains(slotResults) {
  // Count total appearances per player across all slots
  var appearanceCounts = {};
  slotResults.forEach(function(slot) {
    if (slot.skipped) return;
    slot.groups.forEach(function(group) {
      group.forEach(function(p) {
        if (p.email) appearanceCounts[p.email] = (appearanceCounts[p.email] || 0) + 1;
      });
    });
  });

  // Greedy assignment: for each group pick the player with the lowest captaincy ratio
  var captainCounts = {};
  slotResults.forEach(function(slot) {
    if (slot.skipped) return;
    var captains = [];
    slot.groups.forEach(function(group) {
      var best = null;
      var bestRatio = Infinity;
      group.forEach(function(p) {
        if (!p.email) return;
        var ratio = (captainCounts[p.email] || 0) / (appearanceCounts[p.email] || 1);
        if (ratio < bestRatio) { bestRatio = ratio; best = p; }
      });
      if (best) {
        captains.push(best.email);
        captainCounts[best.email] = (captainCounts[best.email] || 0) + 1;
      } else {
        captains.push('');
      }
    });
    slot.captains = captains;
  });
}

// ── Core Optimizer ─────────────────────────────────
// Runs local search with random restarts for one date slot.
// Returns { groups: [[player,...], ...], sitOut: player|null }
function optimizeSlot(available, settings, pairCounts, sitOutCounts) {
  var n         = available.length;
  var remainder = n % 4;

  // Decide group structure
  var groupSizes;
  if (remainder === 0) {
    groupSizes = fillArray(n / 4, 4);
  } else if (remainder === 1) {
    groupSizes = fillArray(Math.floor((n - 1) / 4), 4);
  } else if (remainder === 2) {
    groupSizes = fillArray(Math.floor(n / 4) - 1, 4).concat([3, 3]);
  } else {
    groupSizes = fillArray(Math.floor(n / 4), 4).concat([3]);
  }

  var sitOutPlayer = null;
  var pool = available.slice();

  if (remainder === 1) {
    var minSitOuts = Infinity;
    var sitOutIdx  = 0;
    pool.forEach(function(p, i) {
      var count = sitOutCounts[p.email] || 0;
      if (count < minSitOuts) { minSitOuts = count; sitOutIdx = i; }
    });
    sitOutPlayer = pool.splice(sitOutIdx, 1)[0];
  }

  var iters    = settings.solverIterations || 800;
  var restarts = settings.solverRestarts   || 10;
  var wTV = settings.weightTeamVariance  || 1.0;
  var wGV = settings.weightGroupVariance || 0.5;
  var wSV = settings.weightSocialVariety || 2.0;

  var N = pool.length;

  // Tag each pool player with an integer index for O(1) pair-penalty lookup
  for (var idx = 0; idx < N; idx++) { pool[idx]._idx = idx; }

  // Pre-compute social pair-penalty table (triangular: a < b only)
  // Avoids string operations (pairKey + hash lookup) inside the hot loop
  var pairPen = [];
  for (var a = 0; a < N; a++) {
    pairPen[a] = [];
    for (var b = a + 1; b < N; b++) {
      var hist = pairCounts[pairKey(pool[a].email, pool[b].email)] || 0;
      pairPen[a][b] = wSV * hist * hist;
    }
  }

  // totalGroupVar is constant within a restart (same pool, same ratings) — compute once
  var allRatings = [];
  for (var ri = 0; ri < N; ri++) allRatings.push(pool[ri].rating);
  var totalGroupVarPenalty = variance(allRatings) * wGV;

  // Per-group penalty — inlines variance math to avoid array allocations in the hot loop
  function groupPenalty(group) {
    var sz = group.length;
    var social = 0;
    for (var i = 0; i < sz; i++) {
      for (var j = i + 1; j < sz; j++) {
        var ai = group[i]._idx, bi = group[j]._idx;
        social += ai < bi ? pairPen[ai][bi] : pairPen[bi][ai];
      }
    }
    var r0 = group[0].rating, r1 = group[1].rating, r2 = group[2].rating;
    var gv, tv;
    if (sz === 4) {
      var r3 = group[3].rating;
      var m4 = (r0 + r1 + r2 + r3) * 0.25;
      gv = ((r0-m4)*(r0-m4) + (r1-m4)*(r1-m4) + (r2-m4)*(r2-m4) + (r3-m4)*(r3-m4)) * 0.25;
      var d01 = r0 - r1, d23 = r2 - r3;
      tv = (d01*d01 + d23*d23) * 0.25;
    } else {
      var m3 = (r0 + r1 + r2) / 3;
      gv = ((r0-m3)*(r0-m3) + (r1-m3)*(r1-m3) + (r2-m3)*(r2-m3)) / 3;
      tv = gv;
    }
    return tv * wTV + gv * wGV + social;
  }

  var bestGroups  = null;
  var bestPenalty = Infinity;

  for (var r = 0; r < restarts; r++) {
    var shuffled = shuffleArray(pool.slice());
    var groups   = buildGroupsFromSizes(shuffled, groupSizes);

    // Initialize per-group penalty cache
    var gPen = [];
    var penalty = totalGroupVarPenalty;
    for (var g = 0; g < groups.length; g++) {
      gPen[g] = groupPenalty(groups[g]);
      penalty += gPen[g];
    }

    for (var iter = 0; iter < iters; iter++) {
      var gi = Math.floor(Math.random() * groups.length);
      var gj = Math.floor(Math.random() * groups.length);
      if (gi === gj) continue;

      var pi = Math.floor(Math.random() * groups[gi].length);
      var pj = Math.floor(Math.random() * groups[gj].length);

      // Perform swap
      var tmp = groups[gi][pi];
      groups[gi][pi] = groups[gj][pj];
      groups[gj][pj] = tmp;

      // Incremental delta: recompute only the 2 affected groups (not all groups)
      var newGiPen = groupPenalty(groups[gi]);
      var newGjPen = groupPenalty(groups[gj]);
      var delta = (newGiPen + newGjPen) - (gPen[gi] + gPen[gj]);

      if (delta < 0) {
        gPen[gi] = newGiPen;
        gPen[gj] = newGjPen;
        penalty  += delta;
      } else {
        // Revert swap
        groups[gj][pj] = groups[gi][pi];
        groups[gi][pi] = tmp;
      }
    }

    if (penalty < bestPenalty) {
      bestPenalty = penalty;
      bestGroups  = groups.map(function(g) { return g.slice(); });
    }
  }

  // Output clean player objects (strip _idx)
  var outputGroups = bestGroups.map(function(group) {
    return group.map(function(p) {
      return { name: p.name, email: p.email, rating: p.rating };
    });
  });

  return { groups: outputGroups, sitOut: sitOutPlayer };
}

// ── Chunked Publish Helpers ─────────────────────────
// Step 1: clear existing rows for the month.
function clearAnitaRecords() {
  var ss           = SpreadsheetApp.openById(SHEET_ID); // single open — reused for both sheets
  var anitaPattern = /^Anita Sub\d+$/;

  var pSheet = ss.getSheetByName(TABS.players);
  if (pSheet && pSheet.getLastRow() >= 2) {
    var pRows = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 1).getValues();
    for (var i = pRows.length - 1; i >= 0; i--) {
      if (anitaPattern.test((pRows[i][0] || '').toString().trim())) pSheet.deleteRow(i + 2);
    }
  }

  var rSheet = ss.getSheetByName(TABS.requests);
  if (rSheet && rSheet.getLastRow() >= 2) {
    var rRows = rSheet.getRange(2, 3, rSheet.getLastRow() - 1, 1).getValues();
    for (var i = rRows.length - 1; i >= 0; i--) {
      if (anitaPattern.test((rRows[i][0] || '').toString().trim())) rSheet.deleteRow(i + 2);
    }
  }
}

function publishScheduleStart(params) {
  var month = params.month || '';
  if (!month) return { error: 'Month required.' };

  // Clear fictitious Anita Sub players and their requests before committing new schedule
  clearAnitaRecords();

  var sheet = getOrCreateMatchGroupsSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var monthVals = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = monthVals.length - 1; i >= 0; i--) {
      if (normalizeMonth(monthVals[i][0]) === month) sheet.deleteRow(i + 2);
    }
  }
  return { success: true };
}

// Step 2: append one date's groups (called once per date slot).
function publishScheduleSlot(params) {
  var month = params.month || '';
  var slot  = safeParseJSON(params.slot, null);
  if (!slot || !slot.date) return { error: 'Invalid slot.' };

  // Open spreadsheet once — reused for all writes in this call
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getOrCreateMatchGroupsSheet();
  var saved = 0;
  var sitOutName  = slot.sitOut ? slot.sitOut.name  : '';
  var sitOutEmail = slot.sitOut ? slot.sitOut.email : '';

  // Resources for Anita creation — loaded lazily on first 3-player group, then reused
  var playerRatings = null;
  var pSheet        = null;
  var rSheet        = null;
  var anitaBase     = -1; // count of existing Anita players (loaded once, then incremented)

  (slot.groups || []).forEach(function(group, gi) {
    var captainEmail = (slot.captains || [])[gi] || '';
    var workingGroup = group.slice();

    // If only 3 players, create a fictitious Anita Sub to fill the 4th spot
    if (workingGroup.length === 3) {

      // Lazy-load everything needed — once per publishScheduleSlot call
      if (!playerRatings) {
        playerRatings = getPlayersWithRatings();
        pSheet        = ss.getSheetByName(TABS.players);
        rSheet        = ss.getSheetByName(TABS.requests);
        // Count existing Anita Sub players once; increment in-memory for subsequent groups
        anitaBase = 0;
        if (pSheet && pSheet.getLastRow() >= 2) {
          var names = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 1).getValues();
          anitaBase = names.filter(function(r) {
            return /^Anita Sub\d+$/.test((r[0] || '').toString().trim());
          }).length;
        }
      }

      var n          = anitaBase + 1;
      anitaBase++;   // increment in memory — avoids re-reading the sheet for each group
      var anitaName  = 'Anita Sub' + n;
      var anitaEmail = 'anita.sub' + n + '@xgmail.com';

      // Anita's rating = (partnerRating + avgOf3) / 2
      // partnerRating: adjacent pairing [P0+P1 vs P2+P3] → Anita is P3, paired with P2 (3rd-highest)
      // avgOf3: average of the 3 real players (group-level balance)
      // Fallback: overall pool average when individual ratings are absent (rating = 0 means unrated)
      var ratedGroup = workingGroup.map(function(p) {
        var pr = playerRatings.find(function(r) { return r.email === p.email.toLowerCase(); });
        return (pr && pr.rating > 0) ? pr.rating : null;
      }).filter(function(v) { return v !== null; });
      ratedGroup.sort(function(a, b) { return b - a; }); // descending

      var partnerRating, avgOf3;
      if (ratedGroup.length >= 3) {
        partnerRating = ratedGroup[2]; // P2's rating (3rd-highest = Anita's adjacent partner)
        avgOf3        = (ratedGroup[0] + ratedGroup[1] + ratedGroup[2]) / 3;
      } else if (ratedGroup.length > 0) {
        // Partial ratings — use what's available for both terms
        var partialAvg = ratedGroup.reduce(function(s,v){return s+v;},0) / ratedGroup.length;
        partnerRating  = ratedGroup[ratedGroup.length - 1]; // lowest rated available
        avgOf3         = partialAvg;
      } else {
        // No individual ratings — fall back to pool average
        var poolRated = playerRatings.filter(function(p) { return p.rating > 0; });
        var poolAvg   = poolRated.length > 0
          ? poolRated.reduce(function(s,p){return s+p.rating;},0) / poolRated.length
          : 3.0;
        partnerRating = poolAvg;
        avgOf3        = poolAvg;
      }
      var anitaRating = Math.round(((partnerRating + avgOf3) / 2) * 10) / 10;

      // Add Anita to Players sheet
      pSheet.appendRow([anitaName, anitaEmail, '', anitaRating, false, false]);
      pSheet.getRange(pSheet.getLastRow(), 4).setNumberFormat('0.0');

      // Create Sub Request for Anita with the 3 real players as group
      var groupPlayersJSON = JSON.stringify(workingGroup.map(function(p) {
        return { name: p.name, email: p.email };
      }));
      rSheet.appendRow([
        uid(), new Date().toISOString(),
        anitaName, anitaEmail,
        slot.date, '', 'open', '', groupPlayersJSON
      ]);
      var lastReqRow = rSheet.getLastRow();
      rSheet.getRange(lastReqRow, 5).setNumberFormat('@');
      rSheet.getRange(lastReqRow, 6).setNumberFormat('@');
      rSheet.getRange(lastReqRow, 9).setNumberFormat('@');

      Logger.log('Created ' + anitaName + ' (rating ' + anitaRating + ') for ' + slot.date + ' group ' + String.fromCharCode(65 + gi));
      workingGroup.push({ name: anitaName, email: anitaEmail });
    }

    var ordered = workingGroup.slice().sort(function(a, b) {
      return a.email === captainEmail ? -1 : b.email === captainEmail ? 1 : 0;
    });
    var p = ordered.concat([{name:'',email:''},{name:'',email:''},{name:'',email:''},{name:'',email:''}]);
    sheet.appendRow([
      new Date().toISOString(), month, slot.date,
      String.fromCharCode(65 + gi),
      p[0].name, p[0].email, p[1].name, p[1].email,
      p[2].name, p[2].email, p[3].name, p[3].email,
      sitOutName, sitOutEmail
    ]);
    saved++;
  });
  return { success: true, groupsWritten: saved };
}

// ── Get Published Schedule ──────────────────────────
// Returns the most recently published month's schedule
// grouped by date → groups.
function getPublishedSchedule() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.matchGroups);
  if (!sheet || sheet.getLastRow() < 2) return { month: null, dates: [] };

  // Load no8am flags from Players sheet (column E = index 4)
  var no8amEmails = [];
  var playersSheet = ss.getSheetByName(TABS.players);
  if (playersSheet && playersSheet.getLastRow() >= 2) {
    var pRows = playersSheet.getRange(2, 1, playersSheet.getLastRow() - 1, 5).getValues();
    pRows.forEach(function(r) {
      var email = (r[1] || '').toLowerCase().trim();
      var flag  = r[4];
      if (email && (flag === true || (flag && flag.toString().toUpperCase() === 'TRUE'))) {
        no8amEmails.push(email);
      }
    });
  }

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();

  // Single pass: find latest month and build dateMap simultaneously
  var latestMonth = '';
  var dateMap = {};
  rows.forEach(function(r) {
    var m = normalizeMonth(r[1]);
    if (!m) return;
    if (m > latestMonth) {
      latestMonth = m;
      dateMap = {}; // reset — new latest month found
    }
    if (m !== latestMonth) return;

    var date = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (r[2] ? r[2].toString() : '');
    var letter = r[3] ? r[3].toString() : '';
    var sitOutName  = r[12] ? r[12].toString() : '';
    var sitOutEmail = r[13] ? r[13].toString() : '';

    if (!date) return;
    if (!dateMap[date]) dateMap[date] = {};
    var players = [];
    for (var pi = 0; pi < 4; pi++) {
      var nm = r[4 + pi*2]     ? r[4 + pi*2].toString()     : '';
      var em = r[4 + pi*2 + 1] ? r[4 + pi*2 + 1].toString() : '';
      if (nm) players.push({ name: nm, email: em, isCaptain: pi === 0 });
    }
    dateMap[date][letter] = {
      players: players,
      sitOut: sitOutName ? { name: sitOutName, email: sitOutEmail } : null
    };
  });

  if (!latestMonth) return { month: null, dates: [] };

  var sortedDates = Object.keys(dateMap).sort();
  var dates = sortedDates.map(function(date) {
    var groupLetters = Object.keys(dateMap[date]).sort();
    return {
      date: date,
      groups: groupLetters.map(function(letter) {
        return {
          letter: letter,
          players: dateMap[date][letter].players,
          sitOut:  dateMap[date][letter].sitOut
        };
      })
    };
  });

  return { month: latestMonth, dates: dates, no8amEmails: no8amEmails };
}

// ── Sheet helper ────────────────────────────────────
function getOrCreateMatchGroupsSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.matchGroups);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.matchGroups);
    sheet.getRange(1, 1, 1, 14).setValues([[
      'Timestamp','Month','Date','Group',
      'P1 Name','P1 Email','P2 Name','P2 Email',
      'P3 Name','P3 Email','P4 Name','P4 Email',
      'SitOut Name','SitOut Email'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Scheduler Utilities ─────────────────────────────
function pairKey(emailA, emailB) {
  return emailA < emailB ? emailA + '|' + emailB : emailB + '|' + emailA;
}

function variance(arr) {
  if (!arr || arr.length < 2) return 0;
  var mean = arr.reduce(function(s, v) { return s + v; }, 0) / arr.length;
  return arr.reduce(function(s, v) { return s + (v - mean) * (v - mean); }, 0) / arr.length;
}

function shuffleArray(arr) {
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = arr[i]; arr[i] = arr[j]; arr[j] = tmp;
  }
  return arr;
}

function fillArray(len, val) {
  var out = [];
  for (var i = 0; i < len; i++) out.push(val);
  return out;
}

function buildGroupsFromSizes(players, sizes) {
  var groups = [];
  var idx    = 0;
  sizes.forEach(function(sz) {
    groups.push(players.slice(idx, idx + sz));
    idx += sz;
  });
  return groups;
}

function safeParseJSON(str, fallback) {
  if (!str) return fallback;
  if (typeof str === 'object') return str;
  try { return JSON.parse(str); } catch(e) { return fallback; }
}
