// ══════════════════════════════════════════════════
// SUBCOURT — Apps Script Web App
// MTC Tennis Team
// ══════════════════════════════════════════════════

const SHEET_ID = '1X2oM9GwH206qzFHBqi-oIPZ2wPdn43tIfcMRtvOw9wM';

const TABS = {
  players:    'Players',
  requests:   'SubRequests',
  volunteers: 'Volunteers',
  config:     'Config'
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
      // Dispatch automation — rows 13–15
      autoDispatchEnabled:      sheet.getRange('B13').getValue() === true,
      autoDispatchTimeET:       formatSheetTime(sheet.getRange('B14').getValue()) || '08:00',
    };
  } catch(e) {
    // If Config tab is missing or unreadable, return safe defaults
    return {
      skillWindowPreSchedule:  0.5,
      skillWindowPostSchedule: 2.0,
      preScheduleThresholdHrs: 48,
      lastMinuteThresholdHrs:  24,
      calendarLookaheadDays:   30,
      autoDispatchEnabled:     false,
      autoDispatchTimeET:      '08:00',
    };
  }
}

// ──────────────────────────────────────────────────
// DISPATCH AUTOMATION TRIGGER
// Run this function manually from the Apps Script
// editor whenever you change the dispatch time in
// the Config tab.
// ──────────────────────────────────────────────────

function updateDispatchTrigger() {
  const config = getConfig();

  // Delete any existing dispatch triggers
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runAutoDispatch') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  if (!config.autoDispatchEnabled) {
    Logger.log('Auto-dispatch is disabled. No trigger set.');
    return;
  }

  // Parse the ET time string (HH:MM)
  const parts = config.autoDispatchTimeET.split(':');
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
  const config   = getConfig();
  if (!config.autoDispatchEnabled) return;

  const requests = getRequests();
  const open     = requests.filter(r => r.status === 'open');

  open.forEach(req => {
    try {
      const result = runMatch({ requestId: req.id });
      if (result.candidates && result.candidates.length > 0) {
        const best = result.candidates[0];
        // Auto-confirm the top candidate
        confirmSub({
          requestId:         req.id,
          requestRowIndex:   req.rowIndex,
          subEmail:          best.email,
          subName:           best.name,
          requestorName:     req.name,
          requestorEmail:    req.email,
          matchDate:         req.matchDate,
          matchTime:         req.matchTime,
          volunteerRowIndex: best.rowIndex || null
        });
        Logger.log('Auto-dispatched: ' + req.name + ' → ' + best.name);
      }
    } catch(err) {
      Logger.log('Auto-dispatch error for request ' + req.id + ': ' + err.message);
    }
  });
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
    else if (action === 'getConfig')       result = getConfig();
    else if (action === 'submitRequest')   result = submitRequest(e.parameter);
    else if (action === 'submitVolunteer') result = submitVolunteer(e.parameter);
    else if (action === 'confirmSub')      result = confirmSub(e.parameter);
    else if (action === 'runMatch')        result = runMatch(e.parameter);
    else if (action === 'updateVolunteer') result = updateVolunteer(e.parameter);
    else if (action === 'deleteVolunteer') result = deleteVolunteer(e.parameter);
    else if (action === 'ping')            result = { version: 'V31', ts: new Date().toISOString() };
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
        const requireAllTimes = !lastMinute && (!matchTime || !urgent);
        const trace = vols.map(v => {
          const volTimes     = v.times.map(t => t.trim());
          const reqTime      = matchTime ? matchTime.trim() : '';
          const dateMatch    = v.date.trim() === matchDate.trim();
          const notRequestor = v.email.toLowerCase() !== req.email.toLowerCase();
          const timeMatch    = requireAllTimes
                                 ? TIMES.every(t => volTimes.includes(t))
                                 : (!!reqTime && volTimes.includes(reqTime));
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
    name:  r[0] || '',
    email: (r[1] || '').toLowerCase()
    // rating intentionally excluded from public response
  }));
}

function getPlayersWithRatings() {
  // Internal use only — never sent to browser
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  const rows  = sheet.getDataRange().getValues();
  rows.shift();
  return rows.map(r => ({
    name:   r[0] || '',
    email:  (r[1] || '').toLowerCase(),
    rating: parseFloat(r[2]) || 0
  }));
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
  const hasTBDTime   = !matchTime;

  const skillWindow     = lastMinute ? Infinity : (!urgent ? config.skillWindowPreSchedule : config.skillWindowPostSchedule);
  const requireAllTimes = !lastMinute && (hasTBDTime || !urgent);
  const phase           = lastMinute ? 'last-minute' : (!urgent ? 'pre-schedule' : 'post-schedule');

  let candidates = volunteers.filter(v => {
    if (v.date.trim() !== matchDate.trim()) return false;
    if (v.email.toLowerCase() === req.email.toLowerCase()) return false;
    const volTimes = v.times.map(t => t.trim());
    const reqTime  = matchTime ? matchTime.trim() : '';
    if (requireAllTimes) {
      if (!TIMES.every(t => volTimes.includes(t))) return false;
    } else {
      if (!reqTime || !volTimes.includes(reqTime)) return false;
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
  const senderName = 'MTC Tennis Team';

  // To: requestor + sub   CC: group partners
  const toAddresses = [data.requestorEmail, data.subEmail].filter(Boolean).join(', ');
  const ccList      = groupPlayers.map(function(p) { return p.email; }).filter(Boolean);
  const ccAddresses = ccList.join(', ');

  const subject =
    'MTC Tennis — Substitute confirmed: ' + data.subName + ' for ' + data.requestorName;

  const body =
    'Hi team,\n\n' +
    data.subName + ' will be substituting for ' + data.requestorName +
    ' on ' + dateStr + ' at ' + timeStr + '.\n\n' +
    'No further action needed. Please plan to arrive at least 10 minutes early.\n\n' +
    'See you on the court!\n\n' +
    'MTC Tennis Team';

  var emailParams = {
    to:      toAddresses,
    subject: subject,
    body:    body,
    name:    senderName
  };
  if (ccAddresses) emailParams.cc = ccAddresses;

  MailApp.sendEmail(emailParams);
}

// ──────────────────────────────────────────────────
// HELPERS
// ──────────────────────────────────────────────────

function isUrgent(req, thresholdHrs) {
  if (!req.matchDate || !req.matchTime) return false;
  const hrs     = thresholdHrs || 48;
  const matchDT = new Date(req.matchDate + 'T' + req.matchTime + ':00');
  const now     = new Date();
  const diffHrs = (matchDT - now) / 36e5;
  return diffHrs <= hrs && diffHrs > 0;
}

function isLastMinute(req, thresholdHrs) {
  if (!req.matchDate || !req.matchTime) return false;
  const hrs     = thresholdHrs || 24;
  const matchDT = new Date(req.matchDate + 'T' + req.matchTime + ':00');
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
