// ══════════════════════════════════════════════════
// SUBCOURT — Apps Script Web App
// MWF Tennis League
// ══════════════════════════════════════════════════

const SHEET_ID = '1VjFuq63KLEgZpYvCVi2bJrWEgMxDP6hXygYwjDpUmRE';

// deploy.sh replaces 'rally-tennis-dev.html' with 'rally-tennis-prod.html' when pushing to prod.
const APP_BASE_URL = 'https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-dev.html';

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

// Run once from Apps Script editor after deploying the 4-window dispatch update.
// Updates Config B4-B9 with correct labels and default values for the new system.
function setupDispatchConfig() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);
  var rows = [
    ['B4', 'A4', 'Skill Window >72 hrs',         0.5],
    ['B5', 'A5', 'Skill Window 48-72 hrs',        1.0],
    ['B6', 'A6', 'Skill Window 24-48 hrs',        2.0],
    ['B7', 'A7', 'Last-Minute Threshold (hrs)',    24 ],
    ['B8', 'A8', 'Urgent Threshold (hrs)',         48 ],
    ['B9', 'A9', 'Pre-Schedule Threshold (hrs)',   72 ],
  ];
  rows.forEach(function(r) {
    sheet.getRange(r[1]).setValue(r[2]);
    sheet.getRange(r[0]).setValue(r[3]);
  });
  Logger.log('setupDispatchConfig: Config B4-B9 updated for 4-window dispatch.');
  return { success: true };
}

function saveSenderEmail(params) {
  var email = (params.email || '').toString().trim();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);
  sheet.getRange('A30').setValue('Sender Email');
  sheet.getRange('B30').setValue(email);
  return { success: true, senderEmail: email };
}

function saveGroupEmail(params) {
  var email = (params.email || '').toString().trim().toLowerCase();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);
  sheet.getRange('A33').setValue('Players Email Group');
  sheet.getRange('B33').setValue(email);
  return { success: true, playersGroupEmail: email };
}

// Returns the admin emails (Players sheet isAdmin=true) for internal sync notifications.
function getAdminEmails() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  var col   = getColMap(sheet);
  var rows  = sheet.getDataRange().getValues();
  rows.shift();
  return rows
    .filter(function(r) { return r[col.isAdmin] === true || String(r[col.isAdmin] || '').toUpperCase() === 'TRUE'; })
    .map(function(r) { return (r[col.email] || '').toString().trim(); })
    .filter(function(e) { return e; });
}

// Notifies admins that the Players Email Group needs a manual membership update.
// changes: { add: [{name, email}], remove: [{name, email}] }
function notifyGroupRosterChange(changes) {
  if (!isEmailEnabled()) return;
  var add    = changes.add    || [];
  var remove = changes.remove || [];
  if (!add.length && !remove.length) return;

  var config   = getConfig();
  var groupEmail = config.playersGroupEmail || '';
  var manageLink = groupEmail
    ? 'https://groups.google.com/g/' + groupEmail.split('@')[0] + '/members'
    : '';

  var lines = ['The Players list changed — update the Players Email Group membership:', ''];
  add.forEach(function(p)    { lines.push('Add:    ' + p.name + ' <' + p.email + '>'); });
  remove.forEach(function(p) { lines.push('Remove: ' + p.name + ' <' + p.email + '>'); });
  if (manageLink) {
    lines.push('', 'Manage members: ' + manageLink);
  }

  var admins = getAdminEmails();
  if (!admins.length) return;
  sendLeagueEmail({
    to: admins.join(', '),
    subject: 'Rally — Players Email Group update needed',
    body: lines.join('\n'),
    name: 'MWF Tennis League'
  });
}

// Unified email sender — uses GmailApp with configured from: address (requires Gmail alias setup),
// falls back to MailApp if alias is not configured or not verified.
function sendBrevoEmail(params) {
  // params: { apiKey, recipients: [{email, name}], subject, htmlContent, textContent, attachments }
  var payload = {
    sender: { name: 'MWF Tennis League', email: 'noreply@mtctennis.com' },
    to: params.recipients,
    subject: params.subject
  };
  if (params.htmlContent)  payload.htmlContent = params.htmlContent;
  if (params.textContent)  payload.textContent = params.textContent;
  if (params.attachments)  payload.attachment  = params.attachments;
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'api-key': params.apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  Logger.log('Brevo request — to: ' + JSON.stringify(params.recipients) + ', subject: ' + params.subject);
  var response = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', options);
  var code = response.getResponseCode();
  var body = response.getContentText();
  Logger.log('Brevo response — HTTP ' + code + ': ' + body.substring(0, 500));
  if (code < 200 || code >= 300) {
    throw new Error('Brevo error ' + code + ': ' + body.substring(0, 300));
  }
  return JSON.parse(body);
}

function sendLeagueEmail(params) {
  var config      = getConfig();
  var senderEmail = config.senderEmail || '';
  var options = { name: params.name || 'MWF Tennis League' };
  if (params.htmlBody)  options.htmlBody = params.htmlBody;
  if (params.cc)        options.cc       = params.cc;
  if (params.bcc)       options.bcc      = params.bcc;
  if (params.replyTo)   options.replyTo  = params.replyTo;

  if (senderEmail) {
    try {
      options.from    = senderEmail;
      options.replyTo = senderEmail;
      GmailApp.sendEmail(params.to, params.subject, params.body, options);
      return;
    } catch(e) {
      Logger.log('GmailApp from ' + senderEmail + ' failed, using MailApp: ' + e.message);
    }
  }
  MailApp.sendEmail(params);
}

// Sent to the captain of a 3-player group when their Anita Sub request is auto-created at publish.
function sendCaptainThreePlayerNotification(captainName, captainEmail, matchDate, anitaSubName) {
  if (!captainEmail || !isEmailEnabled()) return;
  var reqUrl    = APP_BASE_URL + '#request';
  var dateStr   = formatDate(matchDate);
  var d         = new Date(matchDate + 'T12:00:00');
  d.setDate(d.getDate() - 1);
  var dayBefore = d.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
  var subject   = 'MWF Tennis League — 3-player group on ' + dateStr;
  var body =
    'Hi ' + captainName + ',\n\n' +
    'You are the captain of a 3-player group on ' + dateStr + ' and therefore a sub request has automatically been created for ' + anitaSubName + '.\n\n' +
    'When Chelsea assigns a court time, update the sub request on the Request a Sub page:\n' +
    reqUrl + '\n\n' +
    'If Rally is unable to fill the request on ' + dayBefore + ', you will be notified by email. At that time, you should use the email/phone process to find a 4th player.\n\n' +
    'MWF Tennis League';
  var htmlBody =
    'Hi ' + captainName + ',<br><br>' +
    'You are the captain of a 3-player group on ' + dateStr + ' and therefore a sub request has automatically been created for ' + anitaSubName + '.<br><br>' +
    'When Chelsea assigns a court time, update the sub request on the <a href="' + reqUrl + '">Request a Sub</a> page.<br><br>' +
    'If Rally is unable to fill the request on ' + dayBefore + ', you will be notified by email. At that time, you should use the email/phone process to find a 4th player.<br><br>' +
    'MWF Tennis League';
  sendLeagueEmail({ to: captainEmail, subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
}

// Sent to a player who was automatically made an alternate when publishing the schedule.
function sendSitOutNotification(playerName, playerEmail, matchDate) {
  if (!playerEmail || !isEmailEnabled()) return;
  var volUrl   = APP_BASE_URL + '#volunteer';
  var dateStr  = formatDate(matchDate);
  var subject  = 'MWF Tennis League — Volunteer to Sub record created for ' + dateStr;
  var body =
    'Hi ' + playerName + ',\n\n' +
    'There was an odd number of players on ' + dateStr + ' and therefore a Volunteer to Sub record has been automatically created for you. ' +
    'You can edit this record on the Volunteer to Sub page:\n' +
    volUrl + '\n\n' +
    'MWF Tennis League';
  var htmlBody =
    'Hi ' + playerName + ',<br><br>' +
    'There was an odd number of players on ' + dateStr + ' and therefore a Volunteer to Sub record has been automatically created for you. ' +
    'You can edit this record on the <a href="' + volUrl + '">Volunteer to Sub</a> page.<br><br>' +
    'MWF Tennis League';
  sendLeagueEmail({ to: playerEmail, subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
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
    // Write labels/defaults for B31–B32 on first use (cells empty)
    var b31 = sheet.getRange('B31').getValue();
    var b32 = sheet.getRange('B32').getValue();
    if (b31 === '' || b31 === null) { sheet.getRange('A31').setValue('Rating Range Limit');      sheet.getRange('B31').setValue(2.0); }
    if (b32 === '' || b32 === null) { sheet.getRange('A32').setValue('Weight Maximum Rating Range'); sheet.getRange('B32').setValue(0.0); }
    // Brevo email section — auto-init on first use
    var b36 = sheet.getRange('B36').getValue();
    if (b36 === '' || b36 === null) {
      sheet.getRange('A34').setValue('── Brevo Email ──');
      sheet.getRange('A35').setValue('Brevo API Key');
      sheet.getRange('A36').setValue('Use Brevo: Availability Notification');
      sheet.getRange('B36').setValue('No');
      sheet.getRange('A37').setValue('Use Brevo: Schedule Email');
      sheet.getRange('B37').setValue('No');
    }
    return {
      // Matching engine — rows 4-7, Timing (hrs) in col B, Window (rating) in col C
      // Row 4: Pre-schedule, Row 5: A little urgent, Row 6: Urgent, Row 7: Last minute (no timing)
      skillWindowFarOut:        parseFloat(sheet.getRange('C4').getValue())  || 0.5,
      skillWindowMid:           parseFloat(sheet.getRange('C5').getValue())  || 1.0,
      skillWindowUrgent:        parseFloat(sheet.getRange('C6').getValue())  || 2.0,
      skillWindowLastMinute:    parseFloat(sheet.getRange('C7').getValue())  || 2.8,
      lastMinuteThresholdHrs:   parseInt(sheet.getRange('B6').getValue())    || 24,
      urgentThresholdHrs:       parseInt(sheet.getRange('B5').getValue())    || 48,
      preScheduleThresholdHrs:  parseInt(sheet.getRange('B4').getValue())    || 72,
      // Volunteer calendar — row 10
      calendarLookaheadDays:    parseInt(sheet.getRange('B10').getValue())   || 30,
      // Dispatch automation — rows 13–14
      autoDispatchEnabled:      (function() { var v = sheet.getRange('B13').getValue(); return v === true || v.toString().toUpperCase() === 'TRUE'; })(),
      autoDispatchTimeET:       formatSheetTime(sheet.getRange('B14').getValue()) || '08:00',
      // Match time reminder — rows 28–29
      matchTimeReminderEnabled: (function() { var v = sheet.getRange('B28').getValue(); return v === true || v.toString().toUpperCase() === 'TRUE'; })(),
      matchTimeReminderTimeET:  formatSheetTime(sheet.getRange('B29').getValue()) || '10:00',
      // Sender email — row 30
      senderEmail: (sheet.getRange('B30').getValue() || '').toString().trim(),
      // Players Email Group — row 33
      playersGroupEmail: (sheet.getRange('B33').getValue() || '').toString().trim(),
      // Brevo — rows 35–37
      brevoApiKey:            (sheet.getRange('B35').getValue() || '').toString().trim(),
      brevoAvailNotification: (function() { var v = sheet.getRange('B36').getValue(); return v === 'Yes' || v === true; })(),
      brevoScheduleEmail:     (function() { var v = sheet.getRange('B37').getValue(); return v === 'Yes' || v === true; })(),
      // Availability window — rows 16–18
      availWindowOpenDate:      (function() { var v = sheet.getRange('B16').getValue(); return v instanceof Date ? formatSheetDate(v) : (v ? v.toString() : ''); })(),
      availWindowCloseDate:     (function() { var v = sheet.getRange('B17').getValue(); return v instanceof Date ? formatSheetDate(v) : (v ? v.toString() : ''); })(),
      availWindowActive:        (function() { var v = sheet.getRange('B18').getValue(); return v === true || v.toString().toUpperCase() === 'TRUE'; })(),
    };
  } catch(e) {
    // If Config tab is missing or unreadable, return safe defaults
    return {
      skillWindowFarOut:       0.5,
      skillWindowMid:          1.0,
      skillWindowUrgent:       2.0,
      skillWindowLastMinute:   2.8,
      lastMinuteThresholdHrs:  24,
      urgentThresholdHrs:      48,
      preScheduleThresholdHrs: 72,
      calendarLookaheadDays:   30,
      autoDispatchEnabled:      false,
      autoDispatchTimeET:       '08:00',
      matchTimeReminderEnabled: false,
      matchTimeReminderTimeET:  '10:00',
      senderEmail: '',
      playersGroupEmail: '',
      brevoApiKey: '',
      brevoAvailNotification: false,
      brevoScheduleEmail: false,
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

function scheduleImmediateDispatch() {
  // Delete any existing runAutoDispatch triggers to avoid stacking
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runAutoDispatch' &&
        t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
      try { ScriptApp.deleteTrigger(t); } catch(e) {}
    }
  });
  // Restore the daily recurring trigger before adding the one-shot,
  // so "Run Dispatch Now" never accidentally removes the scheduled trigger.
  updateDispatchTrigger();
  ScriptApp.newTrigger('runAutoDispatch').timeBased().after(60 * 1000).create();
  return { success: true, scheduled: true };
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
  var reqSheet  = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
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
        // No match found — broadcast to all players if within 24 hours and cancel the request
        if (isLastMinute(req, config.lastMinuteThresholdHrs)) {
          var emailNote = 'broadcast sent — last-minute, no candidates, cancelled';
          try {
            sendSubNeededTomorrowEmail(req);
          } catch(emailErr) {
            emailNote = 'email failed (' + emailErr.message + ') — last-minute, no candidates, cancelled';
            Logger.log('sendSubNeededTomorrowEmail failed for ' + req.id + ': ' + emailErr.message);
          }
          if (reqSheet) reqSheet.getRange(req.rowIndex, 7).setValue('cancelled');
          logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, 'no_candidates', '', '', emailNote]);
          Logger.log('No candidates (last-minute, cancelled): ' + req.name + ' — ' + emailNote);
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
// INSTRUCTIONS PAGE
// ──────────────────────────────────────────────────

var INSTRUCTIONS_URL = 'https://docs.google.com/document/d/e/2PACX-1vT_pn_Shq81mAAObYLF4-PJ1yQdg-OEyzexiTD3Wp59I_dcvBrHSaTc8uqThyWRLK0JHYnVtL1TIU5p/pub';

function serveInstructions() {
  try {
    // Use cache to avoid fetching on every request (5-minute TTL)
    var cache  = CacheService.getScriptCache();
    var html   = cache.get('instructions_html');
    if (!html) {
      html = UrlFetchApp.fetch(INSTRUCTIONS_URL).getContentText();
      try { cache.put('instructions_html', html, 300); } catch(e) {}
    }

    // Strip Google redirect wrappers: href="https://www.google.com/url?q=ENCODED_URL&..."
    // Replace with the decoded target URL directly, opening in a new tab.
    html = html.replace(
      /href="https:\/\/www\.google\.com\/url\?q=([^&"]+)[^"]*"/g,
      function(_, encoded) {
        try { return 'href="' + decodeURIComponent(encoded) + '" target="_blank" rel="noopener"'; }
        catch(e) { return 'href="#"'; }
      }
    );

    return HtmlService.createHtmlOutput(html).setTitle('Rally — Instructions');
  } catch(err) {
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:2rem;color:#c00;">Could not load instructions: ' +
      err.message + '</p>');
  }
}

// ──────────────────────────────────────────────────
// ROUTING
// ──────────────────────────────────────────────────

function doGet(e) {
  // Serve the instructions page as a plain HTML response (not JSONP)
  if (e && e.parameter && e.parameter.page === 'instructions') {
    return serveInstructions();
  }

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
    else if (action === 'cancelRequest')          result = cancelRequest(e.parameter);
    else if (action === 'manuallyAssignSub')      result = manuallyAssignSub(e.parameter);
    else if (action === 'saveAutoDispatchSettings')      result = saveAutoDispatchSettings(e.parameter);
    else if (action === 'runAutoDispatchNow')             result = scheduleImmediateDispatch();
    else if (action === 'saveMatchTimeReminderSettings') result = saveMatchTimeReminderSettings(e.parameter);
    else if (action === 'runMatchTimeReminderNow')        result = runMatchTimeReminder();
    else if (action === 'updateRequestTime')          result = updateRequestTime(e.parameter);
    else if (action === 'recalculateAnitaRatings')    result = recalculateAnitaRatings();
    else if (action === 'sendAdminCode')          result = sendAdminCode(e.parameter);
    else if (action === 'verifyAdminCode')         result = verifyAdminCode(e.parameter);
    else if (action === 'debugAdmin')              result = debugAdmin(e.parameter);
    else if (action === 'getCoordinatorRatings')   result = getCoordinatorRatings(e.parameter);
    else if (action === 'getPlayersForAdmin')       result = getPlayersForAdmin();
    else if (action === 'addPlayer')               result = addPlayer(e.parameter);
    else if (action === 'updatePlayer')            result = updatePlayer(e.parameter);
    else if (action === 'deletePlayer')            result = deletePlayer(e.parameter);
    else if (action === 'saveCoordinatorRatings')  result = saveCoordinatorRatings(e.parameter);
    else if (action === 'getEmailSettings')         result = getEmailSettings();
    else if (action === 'setEmailEnabled')          result = setEmailEnabled(e.parameter);
    else if (action === 'saveSenderEmail')          result = saveSenderEmail(e.parameter);
    else if (action === 'saveGroupEmail')           result = saveGroupEmail(e.parameter);
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
    else if (action === 'sendScheduleEmails')        result = sendScheduleEmails(e.parameter);
    else if (action === 'sendTestScheduleEmail')     result = sendTestScheduleEmail();
    else if (action === 'updateRequest')             result = updateRequest(e.parameter);
    else if (action === 'editRequestPlayers')         result = editRequestPlayers(e.parameter);
    else if (action === 'getMatchSlot')               result = getMatchSlot(e.parameter);
    else if (action === 'createScheduleDraft')         result = createScheduleDraft(e.parameter);
    else if (action === 'sendTestEmail')             result = sendTestEmail();
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
        const hasTBDTime      = !matchTime;
        const effectiveTime   = (matchTime || '08:00').trim();
        const { phase: _phase, skillWindow } = getDispatchPhase(req, config);
        const lastMinute      = _phase === 'last-minute';
        const requireAllTimes = _phase === 'pre-schedule' && !hasTBDTime;
        const trace = vols.map(v => {
          const volTimes     = v.times.map(t => t.trim());
          const dateMatch    = v.date.trim() === matchDate.trim();
          const notRequestor = v.email.toLowerCase() !== req.email.toLowerCase();
          const timeMatch    = requireAllTimes
                                 ? TIMES.every(t => volTimes.includes(t))
                                 : volTimes.includes(effectiveTime);
          const skillOk      = (() => {
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
          skillWindow: skillWindow,
          trace
        };
      }
    }
    else result = { error: 'Unknown action: ' + action };
  } catch (err) {
    result = { error: err.message };
  }

  try {
    var body = JSON.stringify(result);
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + body + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(body)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (serr) {
    var fallback = callback
      ? callback + '({"error":"Serialization error: ' + serr.message.replace(/"/g, "'") + '"})'
      : '{"error":"Serialization error"}';
    return ContentService
      .createTextOutput(fallback)
      .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  }
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

// Detects whether the Players sheet has a Phone column at C (new layout) or not (classic).
// Returns 0-indexed column positions so all functions stay in sync across both layouts.
//   New:     A=Name B=Email C=Phone D=Rating E=No8am F=isAdmin G-K=CoordRatings
//   Classic: A=Name B=Email          C=Rating D=No8am E=isAdmin F-J=CoordRatings
function getColMap(sheet) {
  try {
    var maxCols  = sheet.getMaxColumns();
    // Read at least 14 columns so we can detect coordinator columns beyond the default 5 slots
    var readCols = Math.min(Math.max(sheet.getLastColumn(), 14), maxCols);
    var hdr      = sheet.getRange(1, 1, 1, readCols).getValues()[0];
    var hasPhone = (hdr[2] || '').toString().toLowerCase().trim() === 'phone';
    var coordStart = hasPhone ? 6 : 5;

    // Detect actual coordEnd by finding the last column from coordStart with an @-email header.
    // This handles sheets with more or fewer than the default 5 coordinator columns.
    var coordEnd = coordStart - 1; // default: none found
    var testCol  = -1;
    for (var i = coordStart; i < hdr.length; i++) {
      var h = (hdr[i] || '').toString().trim();
      if (h.indexOf('@') > 0) {
        coordEnd = i;                         // coordinator column
      } else if (h.toLowerCase() === 'test') {
        testCol = i;                          // Test column already exists
        break;
      } else if (h) {
        break;                                // non-empty, non-coordinator header — stop
      }
    }
    if (coordEnd < coordStart) coordEnd = hasPhone ? 10 : 9; // fallback to default 5-slot end
    if (testCol === -1) testCol = coordEnd + 1;              // place Test right after last coordinator

    return hasPhone ? {
      name: 0, email: 1, phone: 2, rating: 3, no8am: 4, isAdmin: 5,
      coordStart: 6, coordEnd: coordEnd, testCol: testCol,
      totalCols: Math.min(testCol + 1, maxCols)
    } : {
      name: 0, email: 1, phone: -1, rating: 2, no8am: 3, isAdmin: 4,
      coordStart: 5, coordEnd: coordEnd, testCol: testCol,
      totalCols: Math.min(testCol + 1, maxCols)
    };
  } catch(e) {
    // Safe fallback: classic layout with Test at column L
    return { name: 0, email: 1, phone: -1, rating: 2, no8am: 3, isAdmin: 4,
             coordStart: 5, coordEnd: 9, testCol: 11, totalCols: 12 };
  }
}

function getPlayers() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  if (!sheet) return [];
  const col  = getColMap(sheet);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  rows.shift(); // remove header
  return rows.map(r => ({
    name:    r[col.name]  || '',
    email:   (r[col.email] || '').toLowerCase(),
    phone:   col.phone >= 0 ? (r[col.phone] || '') : '',
    isAdmin: r[col.isAdmin] === true || String(r[col.isAdmin] || '').toUpperCase() === 'TRUE'
  })).filter(p => p.name || p.email);
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
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  const col   = getColMap(sheet);
  const rows  = sheet.getDataRange().getValues();
  // Auto-init Test column header if missing
  if (rows.length > 0 && (rows[0].length <= col.testCol || !rows[0][col.testCol])) {
    sheet.getRange(1, col.testCol + 1).setValue('Test');
  }
  rows.shift();
  const seen = {};
  return rows.reduce(function(acc, r) {
    const email = (r[col.email] || '').toLowerCase();
    if (email && !seen[email]) {
      seen[email] = true;
      acc.push({
        name:   r[col.name] || '',
        email:  email,
        rating: parseFloat(r[col.rating]) || 0,
        no8am:  r[col.no8am] === true || (r[col.no8am] && r[col.no8am].toString().toUpperCase() === 'TRUE'),
        isTest: r[col.testCol] === true || String(r[col.testCol] || '').toUpperCase() === 'YES'
      });
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
  // Block requests for tomorrow if auto-dispatch has already run today.
  var matchDate = (params.matchDate || '').toString().trim();
  if (matchDate) {
    var tz       = Session.getScriptTimeZone();
    var now      = new Date();
    var tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
    var tomorrowStr = Utilities.formatDate(tomorrow, tz, 'yyyy-MM-dd');
    if (matchDate === tomorrowStr) {
      var config = getConfig();
      if (config.autoDispatchEnabled) {
        var nowTimeET      = Utilities.formatDate(now, 'America/New_York', 'HH:mm');
        var dispatchTimeET = config.autoDispatchTimeET || '08:00';
        if (nowTimeET >= dispatchTimeET) {
          return {
            success:     false,
            dispatchRan: true,
            message:     'The Dispatch routine has already run for tomorrow\'s match date. ' +
                         'Please use email or phone to find a sub.'
          };
        }
      }
    }
  }

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

  // Confirmation email to requester
  if (params.email && isEmailEnabled()) {
    var dateStr = formatDate(params.matchDate || '');
    var timeStr = params.matchTime ? (TIME_LABELS[params.matchTime] || params.matchTime) : 'TBD';
    var reqUrl  = APP_BASE_URL + '#request';
    var subject = 'MWF Tennis League — Sub request received for ' + dateStr;
    var body =
      'Hi ' + params.name + ',\n\n' +
      'Your sub request has been received for ' + dateStr + ' at ' + timeStr + '.\n\n' +
      'Rally will notify you when a sub has been found. You can view or delete your request on the Request a Sub page:\n' +
      reqUrl + '\n\n' +
      'MWF Tennis League';
    var htmlBody =
      'Hi ' + params.name + ',<br><br>' +
      'Your sub request has been received for <strong>' + dateStr + '</strong> at <strong>' + timeStr + '</strong>.<br><br>' +
      'Rally will notify you when a sub has been found. You can view or delete your request on the <a href="' + reqUrl + '">Request a Sub</a> page at any time.<br><br>' +
      'MWF Tennis League';
    sendLeagueEmail({ to: params.email, subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
  }

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

  // Confirmation email to volunteer
  if (params.email && isEmailEnabled() && entries.length > 0) {
    var volUrl  = APP_BASE_URL + '#volunteer';
    var subject = 'MWF Tennis League — Volunteer to sub confirmed';
    var dateLines = entries.map(function(entry) {
      var times = entry.times.map(function(t) { return TIME_LABELS[t] || t; }).join(', ');
      return '  ' + formatDate(entry.date) + ' — ' + times;
    });
    var body =
      'Hi ' + params.name + ',\n\n' +
      'Thank you for volunteering to sub! Your availability has been recorded for the following date' +
      (entries.length > 1 ? 's' : '') + ':\n\n' +
      dateLines.join('\n') + '\n\n' +
      'Rally will notify you if you are selected as a sub. You can view or update your availability on the Volunteer to Sub page:\n' +
      volUrl + '\n\n' +
      'MWF Tennis League';
    var htmlDateRows = entries.map(function(entry) {
      var times = entry.times.map(function(t) { return TIME_LABELS[t] || t; }).join(', ');
      return '<tr><td style="padding:3px 12px 3px 0;font-weight:600;">' + formatDate(entry.date) +
             '</td><td style="padding:3px 0;">' + times + '</td></tr>';
    }).join('');
    var htmlBody =
      'Hi ' + params.name + ',<br><br>' +
      'Thank you for volunteering to sub! Your availability has been recorded for the following date' +
      (entries.length > 1 ? 's' : '') + ':<br><br>' +
      '<table style="font-family:Arial,sans-serif;font-size:14px;border-collapse:collapse;">' +
      htmlDateRows + '</table><br>' +
      'Rally will notify you if you are selected as a sub. You can view or update your availability on the <a href="' + volUrl +
      '">Volunteer to Sub</a> page at any time.<br><br>' +
      'MWF Tennis League';
    sendLeagueEmail({ to: params.email, subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
  }

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
  const col     = getColMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const rows = sheet.getRange(1, 1, lastRow, col.totalCols).getValues();
  rows.shift();
  return rows.some(function(r) {
    const rowEmail = (r[col.email] || '').toLowerCase().trim();
    const flag     = r[col.isAdmin];
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
  var col        = getColMap(sheet);
  var lastRow    = sheet.getLastRow();
  if (lastRow < 2) return { players: [] };

  var lastCol  = Math.max(sheet.getLastColumn(), col.totalCols);
  var allData  = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers  = allData[0];

  var coordColIdx = -1;
  for (var i = col.coordStart; i <= col.coordEnd; i++) {
    if ((headers[i] || '').toString().toLowerCase().trim() === coordEmail) {
      coordColIdx = i; break;
    }
  }
  if (coordColIdx === -1) return { players: [], notAssigned: true };

  var players = [];
  for (var r = 1; r < allData.length; r++) {
    var row = allData[r];
    if (!row[col.name]) continue;
    var no8amVal = row[col.no8am];
    players.push({
      name:     row[col.name] || '',
      email:    (row[col.email] || '').toLowerCase(),
      myRating: row[coordColIdx] !== '' ? row[coordColIdx] : '',
      no8am:    no8amVal === true || (no8amVal && no8amVal.toString().toUpperCase() === 'TRUE')
    });
  }
  return { players: players, notAssigned: false };
}

function recalculateAnitaRatings() {
  var ss     = SpreadsheetApp.openById(SHEET_ID);
  var pSheet = ss.getSheetByName(TABS.players);
  if (!pSheet) return { success: false, error: 'Players sheet not found' };

  var col     = getColMap(pSheet);
  var lastRow = pSheet.getLastRow();
  if (lastRow < 2) return { success: true, updated: 0 };

  var allData     = pSheet.getRange(1, 1, lastRow, col.totalCols).getValues();
  var playerRatings = getPlayersWithRatings(); // already excludes Anita rows
  var ratingMap   = {};
  playerRatings.forEach(function(p) { ratingMap[p.email.toLowerCase()] = p.rating; });

  // Build lookup: anita email → groupPlayers from sub request
  var requests   = getRequests();
  var requestMap = {};
  requests.forEach(function(req) {
    if (/^anita\.sub\d+@xgmail\.com$/i.test(req.email || '')) {
      requestMap[(req.email || '').toLowerCase()] = req;
    }
  });

  var updated = 0;
  for (var row = 1; row < allData.length; row++) {
    var pe = (allData[row][col.email] || '').toLowerCase().trim();
    if (!/^anita\.sub\d+@xgmail\.com$/i.test(pe)) continue;

    var existing = allData[row][col.rating];
    if (existing !== '' && !isNaN(parseFloat(existing))) continue; // already has a rating

    var req = requestMap[pe];
    if (!req || !req.groupPlayers || !req.groupPlayers.length) continue;

    var ratedGroup = req.groupPlayers.map(function(p) {
      return ratingMap[(p.email || '').toLowerCase()] || null;
    }).filter(function(v) { return v !== null && v > 0; });
    ratedGroup.sort(function(a, b) { return b - a; });

    var partnerRating, avgOf3;
    if (ratedGroup.length >= 3) {
      partnerRating = ratedGroup[2];
      avgOf3        = (ratedGroup[0] + ratedGroup[1] + ratedGroup[2]) / 3;
    } else if (ratedGroup.length > 0) {
      var partialAvg = ratedGroup.reduce(function(s, v) { return s + v; }, 0) / ratedGroup.length;
      partnerRating  = ratedGroup[ratedGroup.length - 1];
      avgOf3         = partialAvg;
    } else {
      var poolRated = playerRatings.filter(function(p) { return p.rating > 0; });
      var poolAvg   = poolRated.length > 0
        ? poolRated.reduce(function(s, p) { return s + p.rating; }, 0) / poolRated.length
        : 3.0;
      partnerRating = poolAvg;
      avgOf3        = poolAvg;
    }

    var anitaRating = Math.round(((partnerRating + avgOf3) / 2) * 100) / 100;
    var cell = pSheet.getRange(row + 1, col.rating + 1);
    cell.setNumberFormat('0.0');
    cell.setValue(anitaRating);
    updated++;
  }

  SpreadsheetApp.flush();
  Logger.log('recalculateAnitaRatings: updated ' + updated + ' Anita Sub rating(s).');
  return { success: true, updated: updated };
}

function saveCoordinatorRatings(params) {
  var coordEmail = (params.coordEmail || '').toLowerCase().trim();
  var ratings    = JSON.parse(params.ratings || '[]');
  var sheet      = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  var col        = getColMap(sheet);
  var lastRow    = sheet.getLastRow();
  var lastCol    = Math.max(sheet.getLastColumn(), col.totalCols);
  var allData    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers    = allData[0];

  var coordColIdx = -1;
  for (var i = col.coordStart; i <= col.coordEnd; i++) {
    if ((headers[i] || '').toString().toLowerCase().trim() === coordEmail) {
      coordColIdx = i; break;
    }
  }
  if (coordColIdx === -1) return { success: false, error: 'not_assigned' };

  var ratingMap = {};
  ratings.forEach(function(item) {
    var pe = (item.playerEmail || '').toLowerCase().trim();
    if (pe) ratingMap[pe] = item.rating !== '' && item.rating !== null ? parseFloat(item.rating) : '';
  });

  var coordCols = [];
  for (var k = col.coordStart; k <= col.coordEnd; k++) {
    if (headers[k]) coordCols.push(k);
  }

  for (var row = 1; row < allData.length; row++) {
    var pe = (allData[row][col.email] || '').toLowerCase().trim();
    if (/^anita\.sub\d+@xgmail\.com$/i.test(pe)) continue; // scheduler-managed rating — don't touch
    if (pe && ratingMap.hasOwnProperty(pe)) {
      allData[row][coordColIdx] = ratingMap[pe];
    }
    if (!allData[row][col.name]) continue;
    var vals = coordCols.map(function(ci) {
      var v = allData[row][ci];
      return (v !== '' && !isNaN(parseFloat(v))) ? parseFloat(v) : null;
    }).filter(function(v) { return v !== null; });
    if (vals.length) {
      // Weighted average: min and max × 1, all middle values × 2
      var sorted = vals.slice().sort(function(a, b) { return a - b; });
      var wSum = 0, wTotal = 0;
      sorted.forEach(function(v, i) {
        var w = (sorted.length === 1 || i === 0 || i === sorted.length - 1) ? 1 : 2;
        wSum += v * w; wTotal += w;
      });
      allData[row][col.rating] = Math.round((wSum / wTotal) * 100) / 100;
    } else {
      allData[row][col.rating] = '';
    }
  }

  var dataRows   = allData.slice(1);
  var ratingsCol = dataRows.map(function(r) { return [r[coordColIdx]]; });
  var avgsCol    = dataRows.map(function(r) { return [r[col.rating]]; });
  sheet.getRange(2, coordColIdx + 1, ratingsCol.length, 1).setValues(ratingsCol);
  try { sheet.getRange(2, col.rating + 1, avgsCol.length, 1).setValues(avgsCol); } catch(e) {}
  SpreadsheetApp.flush();

  return { success: true };
}

function getPlayersForAdmin() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var col  = getColMap(sheet);
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return rows.map(function(r, i) {
    return {
      rowIndex: i + 2,
      name:  r[col.name]  || '',
      email: (r[col.email] || '').toLowerCase(),
      phone: col.phone >= 0 ? (r[col.phone] || '') : '',
      no8am: r[col.no8am] === true || (r[col.no8am] || '').toString().toUpperCase() === 'TRUE'
    };
  }).filter(function(p) {
    return (p.name || p.email) && !/^anita\.sub\d+@xgmail\.com$/i.test(p.email);
  });
}

function sortPlayersSheet(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort({ column: 1, ascending: true });
}

function addPlayer(params) {
  var name  = (params.name  || '').trim();
  var email = (params.email || '').toLowerCase().trim();
  var phone = (params.phone || '').trim();
  var no8am = params.no8am === 'true' || params.no8am === true;
  if (!name || !email) return { success: false, error: 'Name and email are required.' };
  var sheet  = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  var col    = getColMap(sheet);
  var newRow = col.phone >= 0
    ? [name, email, phone, '', no8am, false]   // new layout: name,email,phone,rating,no8am,isAdmin
    : [name, email, '', no8am, false];          // classic:    name,email,rating,no8am,isAdmin
  sheet.appendRow(newRow);
  sortPlayersSheet(sheet);
  notifyGroupRosterChange({ add: [{ name: name, email: email }] });
  return { success: true };
}

function updatePlayer(params) {
  var rowIndex = parseInt(params.rowIndex);
  var name     = (params.name  || '').trim();
  var email    = (params.email || '').toLowerCase().trim();
  var phone    = (params.phone || '').trim();
  var no8am    = params.no8am === 'true' || params.no8am === true;
  if (!name || !email) return { success: false, error: 'Name and email are required.' };
  if (isNaN(rowIndex) || rowIndex < 2) return { success: false, error: 'Invalid row.' };
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  if (rowIndex > sheet.getLastRow()) return { success: false, error: 'Row not found.' };
  var col = getColMap(sheet);
  var oldName  = sheet.getRange(rowIndex, col.name  + 1).getValue();
  var oldEmail = (sheet.getRange(rowIndex, col.email + 1).getValue() || '').toString().toLowerCase().trim();
  sheet.getRange(rowIndex, col.name  + 1).setValue(name);
  sheet.getRange(rowIndex, col.email + 1).setValue(email);
  if (col.phone >= 0) sheet.getRange(rowIndex, col.phone + 1).setValue(phone);
  sheet.getRange(rowIndex, col.no8am + 1).setValue(no8am);
  sortPlayersSheet(sheet);
  if (oldEmail && oldEmail !== email) {
    notifyGroupRosterChange({
      remove: [{ name: oldName, email: oldEmail }],
      add:    [{ name: name, email: email }]
    });
  }
  return { success: true };
}

function deletePlayer(params) {
  var rowIndex = parseInt(params.rowIndex);
  if (isNaN(rowIndex) || rowIndex < 2) return { success: false, error: 'Invalid row.' };
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.players);
  if (rowIndex > sheet.getLastRow()) return { success: false, error: 'Row not found.' };
  var col   = getColMap(sheet);
  var name  = sheet.getRange(rowIndex, col.name  + 1).getValue();
  var email = (sheet.getRange(rowIndex, col.email + 1).getValue() || '').toString().toLowerCase().trim();
  sheet.deleteRow(rowIndex);
  if (email) notifyGroupRosterChange({ remove: [{ name: name, email: email }] });
  return { success: true };
}

function updateRequest(params) {
  const sheet    = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
  const rowIndex = parseInt(params.rowIndex);
  if (isNaN(rowIndex) || rowIndex < 2) return { success: false, error: 'Invalid row.' };
  if (!params.matchDate) return { success: false, error: 'Date required.' };
  const dateCell = sheet.getRange(rowIndex, 5); // col E = matchDate
  dateCell.setNumberFormat('@');
  dateCell.setValue(params.matchDate);
  const timeCell = sheet.getRange(rowIndex, 6); // col F = matchTime
  timeCell.setNumberFormat('@');
  timeCell.setValue(params.matchTime || '');
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

  // 3. Replace requestor's slot in MatchGroups with the sub's name/email
  updateScheduleForSub(ss, params);

  // 4. Parse group players
  var groupPlayers = [];
  try { groupPlayers = JSON.parse(params.groupPlayers || '[]'); } catch(e) {}

  // 5. Send email
  sendConfirmationEmails(params, groupPlayers);

  return { success: true };
}

// Replaces the requestor's player slot in MatchGroups with the confirmed sub.
function updateScheduleForSub(ss, params) {
  var matchDate      = (params.matchDate      || '').toString().trim();
  var requestorEmail = (params.requestorEmail || '').toLowerCase().trim();
  var subName        = (params.subName        || '').toString().trim();
  var subEmail       = (params.subEmail       || '').toString().trim();
  if (!matchDate || !requestorEmail || !subName || !subEmail) return;

  var sheet = ss.getSheetByName(TABS.matchGroups);
  if (!sheet || sheet.getLastRow() < 2) return;

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var rowDate = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (r[2] ? r[2].toString() : '');
    if (rowDate !== matchDate) continue;

    // Player slots: pi=0→cols 5,6  pi=1→cols 7,8  pi=2→cols 9,10  pi=3→cols 11,12
    for (var pi = 0; pi < 4; pi++) {
      var em = (r[5 + pi * 2] || '').toString().toLowerCase().trim();
      if (em === requestorEmail) {
        sheet.getRange(i + 2, 5 + pi * 2, 1, 2).setValues([[subName, subEmail]]);
        return;
      }
    }
  }
}

// Marks a volunteer's pending row for the given email/date as 'matched'.
// Used by manual sub-assignment paths so the Volunteers tab stays in sync
// with the automated dispatch path (confirmSub), which does this already.
function markVolunteerMatched(ss, email, matchDate) {
  var emailLower = (email || '').toString().toLowerCase().trim();
  matchDate = (matchDate || '').toString().trim();
  if (!emailLower || !matchDate) return;
  var sheet = ss.getSheetByName(TABS.volunteers);
  if (!sheet || sheet.getLastRow() < 2) return;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var rowEmail = (r[3] || '').toString().toLowerCase().trim();
    var rowDate  = formatSheetDate(r[4]);
    if (rowEmail === emailLower && rowDate === matchDate && r[6] === 'pending') {
      sheet.getRange(i + 2, 7).setValue('matched');
      return;
    }
  }
}

// Replaces any player slot in MatchGroups that matches oldEmail on matchDate.
function replaceSchedulePlayer(ss, matchDate, oldEmail, newName, newEmail) {
  var sheet = ss.getSheetByName(TABS.matchGroups);
  if (!sheet || sheet.getLastRow() < 2) return;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var rowDate = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (r[2] ? r[2].toString() : '');
    if (rowDate !== matchDate) continue;
    for (var pi = 0; pi < 4; pi++) {
      var em = (r[5 + pi * 2] || '').toString().toLowerCase().trim();
      if (em === oldEmail.toLowerCase().trim()) {
        sheet.getRange(i + 2, 5 + pi * 2, 1, 2).setValues([[newName, newEmail]]);
        return;
      }
    }
  }
}

// Looks up a player's scheduled match group for a given date.
// Returns all 4 players in the group plus any known match time.
function getMatchSlot(params) {
  var playerEmail = (params.playerEmail || '').toLowerCase().trim();
  var matchDate   = (params.matchDate   || '').toString().trim();
  if (!playerEmail || !matchDate) return { found: false };

  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.matchGroups);
  if (!sheet || sheet.getLastRow() < 2) return { found: false };

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var rowDate = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (r[2] ? r[2].toString() : '');
    if (rowDate !== matchDate) continue;

    for (var pi = 0; pi < 4; pi++) {
      var em = (r[5 + pi * 2] || '').toString().toLowerCase().trim();
      if (em !== playerEmail) continue;

      // Found the player — collect all 4 slots (skip empty)
      var players = [];
      for (var pj = 0; pj < 4; pj++) {
        var nm = (r[4 + pj * 2] || '').toString().trim();
        var ev = (r[5 + pj * 2] || '').toString().toLowerCase().trim();
        if (nm) players.push({ name: nm, email: ev });
      }

      // Try to find a match time from any existing sub request on this date
      // by any player in this group who already submitted with a known time.
      var matchTime = '';
      var groupEmails = players.map(function(p) { return p.email; });
      var reqSheet = ss.getSheetByName(TABS.requests);
      if (reqSheet && reqSheet.getLastRow() >= 2) {
        var reqRows = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, 6).getValues();
        for (var j = 0; j < reqRows.length; j++) {
          var rr = reqRows[j];
          var reqDate  = formatSheetDate(rr[4]);
          var reqTime  = (rr[5] ? rr[5].toString().trim() : '');
          var reqEmail = (rr[3] || '').toString().toLowerCase().trim();
          if (reqDate === matchDate && reqTime && groupEmails.indexOf(reqEmail) !== -1) {
            matchTime = reqTime;
            break;
          }
        }
      }

      return { found: true, matchTime: matchTime, players: players };
    }
  }

  return { found: false };
}

// Handles edits from the "Edit Request" modal on the Request a Sub page.
// Supports three cases: date/time update, requestor replacement (fill sub),
// and non-requestor player replacement (schedule update).
function editRequestPlayers(params) {
  var ss       = SpreadsheetApp.openById(SHEET_ID);
  var rowIndex = parseInt(params.rowIndex);
  if (isNaN(rowIndex) || rowIndex < 2) return { success: false, error: 'Invalid row.' };

  var matchDate          = (params.matchDate          || '').toString().trim();
  var matchTime          = (params.matchTime          || '').toString().trim();
  var origRequestorEmail = (params.origRequestorEmail || '').toLowerCase().trim();
  var newP1Email         = (params.newP1Email         || '').toLowerCase().trim();
  var newP1Name          = (params.newP1Name          || '').toString().trim();

  var origGroupPlayers = [];
  var newGroupPlayers  = [];
  try { origGroupPlayers = JSON.parse(params.origGroupPlayers || '[]'); } catch(e) {}
  try { newGroupPlayers  = JSON.parse(params.newGroupPlayers  || '[]'); } catch(e) {}

  var reqSheet = ss.getSheetByName(TABS.requests);

  // 1. Always update date and time
  var dateCell = reqSheet.getRange(rowIndex, 5);
  dateCell.setNumberFormat('@');
  dateCell.setValue(matchDate);
  var timeCell = reqSheet.getRange(rowIndex, 6);
  timeCell.setNumberFormat('@');
  timeCell.setValue(matchTime);

  // 2. Check if requestor was replaced (fill-sub case)
  var requestorReplaced = newP1Email && newP1Email !== origRequestorEmail;

  if (requestorReplaced) {
    reqSheet.getRange(rowIndex, 7).setValue('filled');
    reqSheet.getRange(rowIndex, 8).setValue(newP1Email);

    // Anita's email is stored directly in the MatchGroups P4 slot, so origRequestorEmail
    // is always the correct slot to replace (works for both regular and Anita requests).
    if (origRequestorEmail) {
      replaceSchedulePlayer(ss, matchDate, origRequestorEmail, newP1Name, newP1Email);
    }

    markVolunteerMatched(ss, newP1Email, matchDate);

    // Send confirmation email (requestorName is P1 original)
    var requestorName = (params.origRequestorName || '').toString().trim();
    sendConfirmationEmails({
      requestorName:  requestorName,
      requestorEmail: origRequestorEmail,
      subName:        newP1Name,
      subEmail:       newP1Email,
      matchDate:      matchDate,
      matchTime:      matchTime
    }, newGroupPlayers);
  }

  // 3. Update groupPlayers JSON (always, reflects any player changes)
  var groupPlayersCell = reqSheet.getRange(rowIndex, 9);
  groupPlayersCell.setValue(JSON.stringify(newGroupPlayers));

  // 4. Update MatchGroups for changed non-requestor players
  for (var i = 0; i < origGroupPlayers.length; i++) {
    var orig   = origGroupPlayers[i] || {};
    var nw     = newGroupPlayers[i]  || {};
    var oEmail = (orig.email || '').toLowerCase().trim();
    var nEmail = (nw.email   || '').toLowerCase().trim();
    if (oEmail && nEmail && oEmail !== nEmail) {
      replaceSchedulePlayer(ss, matchDate, oEmail, nw.name || '', nw.email || '');
    }
  }

  return { success: true, filled: requestorReplaced };
}

// ──────────────────────────────────────────────────
// GMAIL DRAFT — SCHEDULE EMAIL
// ──────────────────────────────────────────────────

// Creates a Gmail draft in the admin's account with the published schedule
// as an HTML table in the body and a CSV attachment (player × date matrix).
// The admin then opens Gmail Drafts, previews, and sends.
// Reads MatchGroups + Players sheets and returns the data needed to build schedule emails.
// Returns null if no published schedule exists.
function buildScheduleDataFromMatchGroups() {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var mgSheet = ss.getSheetByName(TABS.matchGroups);
  if (!mgSheet || mgSheet.getLastRow() < 2) return null;

  var anitaRe   = /^anita\.sub\d+@xgmail\.com$/i;
  var pSheet    = ss.getSheetByName(TABS.players);
  var playerRows = pSheet.getLastRow() > 1
    ? pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 2).getValues() : [];
  var playerNameMap = {};
  var playerEmails  = [];
  playerRows.forEach(function(r) {
    var email = (r[1] || '').toString().toLowerCase().trim();
    var name  = (r[0] || '').toString().trim();
    if (!email || anitaRe.test(email)) return;
    playerNameMap[email] = name;
    playerEmails.push(email);
  });

  var rows = mgSheet.getRange(2, 1, mgSheet.getLastRow() - 1, 16).getValues();
  var latestMonth = '';
  rows.forEach(function(r) {
    var m = normalizeMonth(r[1]);
    if (m > latestMonth) latestMonth = m;
  });
  if (!latestMonth) return null;

  var dateParts  = latestMonth.split('-');
  var monthDate  = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, 1);
  var monthLabel = Utilities.formatDate(monthDate, Session.getScriptTimeZone(), 'MMMM yyyy');

  var dateMap = {};
  rows.forEach(function(r) {
    if (normalizeMonth(r[1]) !== latestMonth) return;
    var date = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : r[2].toString();
    var letter       = r[3] ? r[3].toString() : '';
    var sitOutName   = r[12] ? r[12].toString().trim() : '';
    var sitOutEmail  = r[13] ? r[13].toString().toLowerCase().trim() : '';
    var sitOut2Name  = r[14] ? r[14].toString().trim() : '';
    var sitOut2Email = r[15] ? r[15].toString().toLowerCase().trim() : '';
    if (!date || !letter) return;
    if (!dateMap[date]) dateMap[date] = {
      groups: {},
      sitOut:  sitOutName  ? { name: sitOutName,  email: sitOutEmail  } : null,
      sitOut2: sitOut2Name ? { name: sitOut2Name, email: sitOut2Email } : null
    };
    var players = [];
    for (var pi = 0; pi < 4; pi++) {
      var nm = r[4 + pi * 2] ? r[4 + pi * 2].toString().trim() : '';
      var em = r[5 + pi * 2] ? r[5 + pi * 2].toString().toLowerCase().trim() : '';
      if (nm && !anitaRe.test(em)) players.push({ name: nm, email: em, isCaptain: pi === 0 });
    }
    dateMap[date].groups[letter] = players;
  });

  var sortedDates = Object.keys(dateMap).sort();
  return { dateMap: dateMap, sortedDates: sortedDates, monthLabel: monthLabel,
           playerNameMap: playerNameMap, playerEmails: playerEmails };
}

function createScheduleDraft(params) {
  var scheduleUrl = (params.scheduleUrl || '').toString().trim();
  var sd = buildScheduleDataFromMatchGroups();
  if (!sd || !sd.sortedDates.length) return { success: false, error: 'No schedule data.' };
  if (!sd.playerEmails.length) return { success: false, error: 'No player emails found.' };

  var htmlBody    = buildScheduleHtml(sd.dateMap, sd.sortedDates, sd.monthLabel, scheduleUrl);
  var csvContent  = buildScheduleCsv(sd.dateMap, sd.sortedDates, sd.monthLabel, sd.playerNameMap);
  var csvFileName = sd.monthLabel.replace(/\s/g, '_') + '_Schedule.csv';
  var subject     = 'MWF Tennis League — ' + sd.monthLabel + ' Schedule';
  var config      = getConfig();

  // ── Send via Brevo if enabled ────────────────────────────────────────
  if (config.brevoScheduleEmail && config.brevoApiKey) {
    try {
      var recipients = sd.playerEmails.map(function(e) {
        return { email: e, name: sd.playerNameMap[e] || '' };
      });
      sendBrevoEmail({
        apiKey:       config.brevoApiKey,
        senderName:   'MWF Tennis League',
        senderEmail:  config.senderEmail,
        recipients:   recipients,
        subject:      subject,
        htmlContent:  htmlBody,
        attachments:  [{ content: Utilities.base64Encode('﻿' + csvContent), name: csvFileName }]
      });
      return { success: true, month: sd.monthLabel, emailsSent: sd.playerEmails.length };
    } catch(e) {
      return { success: false, error: 'Brevo send failed: ' + e.message };
    }
  }

  // ── Send via GmailApp / MailApp ──────────────────────────────────────
  var csvBlob = Utilities.newBlob('﻿' + csvContent, 'text/csv', csvFileName);
  var toList  = sd.playerEmails.join(', ');
  var opts    = { htmlBody: htmlBody, attachments: [csvBlob], name: 'MWF Tennis League' };
  try {
    if (config.senderEmail) {
      try {
        GmailApp.sendEmail(toList, subject, '', Object.assign({}, opts, { from: config.senderEmail, replyTo: config.senderEmail }));
        return { success: true, month: sd.monthLabel, emailsSent: sd.playerEmails.length };
      } catch(ge) {
        Logger.log('GmailApp failed (' + ge.message + '), falling back to MailApp');
      }
    }
    MailApp.sendEmail(Object.assign({}, opts, { to: toList, subject: subject, body: '' }));
    return { success: true, month: sd.monthLabel, emailsSent: sd.playerEmails.length };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function buildScheduleHtml(dateMap, sortedDates, monthLabel, scheduleUrl) {
  var MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  var DAYS   = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  var thStyle = 'padding:8px 12px;text-align:left;color:white;';
  var tdBase  = 'padding:6px 12px;vertical-align:top;';

  var html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#111;max-width:750px;">' +
    '<h2 style="color:#1a5c3a;margin-bottom:12px;">MWF Tennis League — ' + monthLabel + ' Schedule</h2>' +
    '<p style="margin-bottom:20px;">The ' + monthLabel + ' schedule has been published.' +
    (scheduleUrl ? ' <a href="' + scheduleUrl + '">View Schedule</a>.' : '') + '</p>' +
    '<table style="border-collapse:collapse;width:100%;">' +
    '<tr style="background:#1a5c3a;">' +
    '<th style="' + thStyle + 'width:20%">Date</th>' +
    '<th style="' + thStyle + 'width:6%;text-align:center">Grp</th>' +
    '<th style="' + thStyle + '">Players</th>' +
    '</tr>';

  sortedDates.forEach(function(date, di) {
    var entry   = dateMap[date];
    var dp      = date.split('-');
    var d       = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]));
    var dateLabel = DAYS[d.getDay()].slice(0,3) + ', ' + MONTHS[d.getMonth()].slice(0,3) + ' ' + parseInt(dp[2]);
    var letters   = Object.keys(entry.groups).sort();
    var rowBg     = di % 2 === 0 ? '#f5f9f7' : '#ffffff';
    var altCount  = (entry.sitOut && entry.sitOut.name ? 1 : 0) + (entry.sitOut2 && entry.sitOut2.name ? 1 : 0);
    var totalRows = letters.length + altCount;

    letters.forEach(function(letter, gi) {
      var players    = entry.groups[letter];
      var playerStr  = players.map(function(p) {
        return p.name + (p.isCaptain ? ' <strong>(C)</strong>' : '');
      }).join(', ');
      var borderTop  = (di > 0 && gi === 0) ? '2px solid #c96048' : (gi > 0 ? '1px solid #ddd' : 'none');

      html += '<tr style="background:' + rowBg + ';border-top:' + borderTop + ';">';
      if (gi === 0) {
        html += '<td rowspan="' + totalRows + '" style="' + tdBase + 'font-weight:bold;white-space:nowrap;">' + dateLabel + '</td>';
      }
      html += '<td style="' + tdBase + 'text-align:center;font-weight:bold;color:#1a5c3a;">' + letter + '</td>';
      html += '<td style="' + tdBase + '">' + playerStr + '</td>';
      html += '</tr>';
    });

    if (entry.sitOut && entry.sitOut.name) {
      html += '<tr style="background:' + rowBg + ';">' +
        '<td style="' + tdBase + 'text-align:center;color:#888;font-style:italic;font-size:12px;">Alt</td>' +
        '<td style="' + tdBase + 'color:#888;font-style:italic;font-size:12px;">' + entry.sitOut.name + '</td>' +
        '</tr>';
    }
    if (entry.sitOut2 && entry.sitOut2.name) {
      html += '<tr style="background:' + rowBg + ';">' +
        '<td style="' + tdBase + 'text-align:center;color:#888;font-style:italic;font-size:12px;">Alt</td>' +
        '<td style="' + tdBase + 'color:#888;font-style:italic;font-size:12px;">' + entry.sitOut2.name + '</td>' +
        '</tr>';
    }
  });

  html += '</table>' +
    '<p style="margin-top:24px;color:#888;font-size:12px;">MWF Tennis League</p>' +
    '</div>';
  return html;
}


function buildScheduleCsv(dateMap, sortedDates, monthLabel, playerNameMap) {
  var MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  var DAYS   = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  var anitaRe = /^anita\.sub\d+@xgmail\.com$/i;

  // Header row: Player, then each date formatted short
  var headerCols = ['Player'].concat(sortedDates.map(function(date) {
    var dp = date.split('-');
    var d  = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]));
    return DAYS[d.getDay()].slice(0,3) + ', ' + MONTHS[d.getMonth()].slice(0,3) + ' ' + parseInt(dp[2]);
  }));

  // Build cellData[email][date] = value
  var cellData = {};
  sortedDates.forEach(function(date) {
    var entry = dateMap[date];
    Object.keys(entry.groups).forEach(function(letter) {
      entry.groups[letter].forEach(function(p) {
        if (!p.email || anitaRe.test(p.email)) return;
        if (!cellData[p.email]) cellData[p.email] = {};
        cellData[p.email][date] = letter + (p.isCaptain ? ' [C]' : '');
      });
    });
    if (entry.sitOut && entry.sitOut.email && !anitaRe.test(entry.sitOut.email)) {
      if (!cellData[entry.sitOut.email]) cellData[entry.sitOut.email] = {};
      cellData[entry.sitOut.email][date] = 'Avail';
    }
    if (entry.sitOut2 && entry.sitOut2.email && !anitaRe.test(entry.sitOut2.email)) {
      if (!cellData[entry.sitOut2.email]) cellData[entry.sitOut2.email] = {};
      cellData[entry.sitOut2.email][date] = 'Avail';
    }
  });

  // Sort players by Last, First
  var emails = Object.keys(cellData).sort(function(a, b) {
    return csvLastFirst(playerNameMap[a] || a).localeCompare(csvLastFirst(playerNameMap[b] || b));
  });

  function csvQ(v) { return '"' + (v || '').replace(/"/g, '""') + '"'; }

  var lines = [headerCols.map(csvQ).join(',')];
  emails.forEach(function(email) {
    var name = csvLastFirst(playerNameMap[email] || email);
    var row  = [csvQ(name)].concat(sortedDates.map(function(d) {
      return csvQ((cellData[email] || {})[d] || '');
    }));
    lines.push(row.join(','));
  });
  return lines.join('\r\n');
}

function csvLastFirst(name) {
  var parts = name.trim().split(/\s+/);
  if (parts.length < 2) return name;
  return parts[parts.length - 1] + ', ' + parts.slice(0, -1).join(' ');
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
  const hasTBDTime      = !matchTime;
  const effectiveTime   = (matchTime || '08:00').trim();
  const { phase, skillWindow } = getDispatchPhase(req, config);
  const lastMinute      = phase === 'last-minute';
  const requireAllTimes = phase === 'pre-schedule' && !hasTBDTime;

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
    // Look up player record for rating and no8am flag
    const vol = players.find(p => p.email.toLowerCase() === v.email.toLowerCase());
    if (!vol) return false;
    if (Math.abs(vol.rating - reqRating) > skillWindow) return false;
    // No8am volunteers must never be matched to an 8am slot or a TBD request
    // (TBD defaults to effectiveTime '08:00', which could turn out to be 8am).
    if (vol && vol.no8am && effectiveTime === '08:00') return false;
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

  // For tomorrow's requests there is no rating restriction, so large skill gaps are possible.
  // Sort by closest rating to minimize variation; use timestamp only as tiebreaker.
  // For all other dates keep the original behavior: earliest submission first.
  var tomorrowD = new Date();
  tomorrowD.setDate(tomorrowD.getDate() + 1);
  var tomorrowStr  = Utilities.formatDate(tomorrowD, 'America/New_York', 'yyyy-MM-dd');
  var isTomorrow   = matchDate === tomorrowStr;

  candidates.sort((a, b) => {
    if (isTomorrow) {
      if (a.ratingDiff !== b.ratingDiff) return a.ratingDiff - b.ratingDiff;
      return a.timestamp.localeCompare(b.timestamp);
    }
    if (a.timestamp !== b.timestamp) return a.timestamp.localeCompare(b.timestamp);
    return a.ratingDiff - b.ratingDiff;
  });

  return {
    candidates: candidates.slice(0, 5),
    skillWindow: skillWindow,
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

  if (isEmailEnabled()) sendLeagueEmail(emailParams);
}

function saveMatchTimeReminderSettings(params) {
  var sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);
  var enabled = params.enabled === 'true' || params.enabled === true;
  var time    = (params.time || '10:00').trim();

  sheet.getRange('A28').setValue('Match Time Reminder Enabled');
  sheet.getRange('B28').setValue(enabled);
  sheet.getRange('A29').setValue('Match Time Reminder Time (ET)');
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
  var siteUrl  = APP_BASE_URL + '#request';
  var notified = 0;

  requests.forEach(function(req) {
    if (req.status !== 'open') return;
    if (req.matchTime) return; // already has a time

    // Check if match date is within 60 hours (use 8:00 AM for TBD times)
    var matchDT = new Date(req.matchDate + 'T08:00:00');
    var diffHrs = (matchDT - now) / 36e5;
    if (diffHrs <= 0 || diffHrs > 60) return;

    var groupPlayers = req.groupPlayers || [];
    var isAnitaSub = /^anita\.sub\d+@xgmail\.com$/i.test(req.email || '');
    var captain = isAnitaSub ? (groupPlayers[0] || {}) : null;
    var recipientEmail = isAnitaSub ? (captain.email || '') : req.email;
    var greetingName  = isAnitaSub ? (captain.name  || 'Captain') : req.name;
    if (!recipientEmail) return;

    var dateStr = formatDate(req.matchDate);
    var subject = 'MWF Tennis League — Court time needed for your sub request: ' + dateStr;

    var body =
      'Hi ' + greetingName + ',\n\n' +
      'You have an open sub request for ' + dateStr + ' and no court time has been assigned yet.\n\n' +
      'Once Chelsea has scheduled a court, please add the court time to your request on the Request a Sub page:\n' + siteUrl + '\n\n' +
      'If you are on Overflow, do nothing. Rally will still try to find a sub.\n\n' +
      'Note: Non 8am players are ineligible to fill a sub request without a court time assigned.\n\n' +
      'MWF Tennis League';

    var htmlBody =
      'Hi ' + greetingName + ',<br><br>' +
      'You have an open sub request for <strong>' + dateStr + '</strong> and no court time has been assigned yet.<br><br>' +
      'Once Chelsea has scheduled a court, please add the court time to your request on the <a href="' + siteUrl + '">Request a Sub</a> page.<br><br>' +
      '<em>If you are on Overflow, do nothing. Rally will still try to find a sub.</em><br><br>' +
      '<em>Note: Non 8am players are ineligible to fill a sub request without a court time assigned.</em><br><br>' +
      'MWF Tennis League';

    var ccList = groupPlayers.map(function(p) { return p.email; }).filter(function(e) { return e && e !== recipientEmail; });
    var emailParams = {
      to:       recipientEmail,
      subject:  subject,
      body:     body,
      htmlBody: htmlBody,
      name:     'MWF Tennis League'
    };
    if (ccList.length) emailParams.cc = ccList.join(', ');
    if (isEmailEnabled()) sendLeagueEmail(emailParams);
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
  var directoryUrl = APP_BASE_URL + '#directory';
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
  if (isEmailEnabled()) sendLeagueEmail(emailParams);
}

function sendSubNeededTomorrowEmail(req) {
  if (!isEmailEnabled()) return;

  var isAnitaSub   = /^anita\.sub\d+@xgmail\.com$/i.test(req.email || '');
  var groupPlayers = req.groupPlayers || [];

  var toEmail, greetingName, ccPlayers;
  if (isAnitaSub) {
    var captain  = groupPlayers[0] || {};
    toEmail      = captain.email || '';
    greetingName = captain.name  || 'Captain';
    ccPlayers    = groupPlayers.slice(1);
  } else {
    toEmail      = req.email || '';
    greetingName = req.name  || 'A player';
    ccPlayers    = groupPlayers;
  }
  if (!toEmail) return;

  var dateStr = formatDate(req.matchDate);
  var timeStr = req.matchTime ? (TIME_LABELS[req.matchTime] || req.matchTime) : 'TBD';

  var subject = 'MWF Tennis League — Unable to find substitute: ' + dateStr + (req.matchTime ? ' at ' + timeStr : '');
  var directoryUrl = APP_BASE_URL + '#directory';
  var body =
    'Hi ' + greetingName + ',\n\n' +
    'Unfortunately, we were unable to find a volunteer to fill the sub request for your match:\n\n' +
    '  Date: ' + dateStr + '\n' +
    '  Time: ' + timeStr + '\n\n' +
    'Player email addresses and phone numbers can be found on the Directory page: ' + directoryUrl + '\n\n' +
    'MWF Tennis League';
  var htmlBody =
    'Hi ' + greetingName + ',<br><br>' +
    'Unfortunately, we were unable to find a volunteer to fill the sub request for your match:<br><br>' +
    '&nbsp;&nbsp;Date: ' + dateStr + '<br>' +
    '&nbsp;&nbsp;Time: ' + timeStr + '<br><br>' +
    'Player email addresses and phone numbers can be found on the <a href="' + directoryUrl + '">Directory</a> page.<br><br>' +
    'MWF Tennis League';

  var ccList = ccPlayers.map(function(p) { return p.email; }).filter(function(e) {
    return e && !/^anita\.sub\d+@xgmail\.com$/i.test(e);
  });
  var emailParams = { to: toEmail, subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' };
  if (ccList.length) emailParams.cc = ccList.join(', ');
  sendLeagueEmail(emailParams);
}

function cancelRequest(params) {
  var requests = getRequests();
  var req = requests.find(function(r) { return r.id === params.requestId; });
  if (!req) return { success: false, error: 'Request not found' };
  if (req.status === 'filled') return { success: false, error: 'Cannot cancel a filled request.' };
  var reqSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
  reqSheet.getRange(parseInt(req.rowIndex), 7).setValue('cancelled');
  return { success: true };
}

function manuallyAssignSub(params) {
  var requestId = (params.requestId || '').toString().trim();
  var subName   = (params.subName   || '').toString().trim();
  var subEmail  = (params.subEmail  || '').toString().trim();
  if (!requestId || !subName || !subEmail) return { success: false, error: 'Missing params' };

  var requests = getRequests();
  var req = requests.find(function(r) { return r.id === requestId; });
  if (!req) return { success: false, error: 'Request not found' };

  var ss       = SpreadsheetApp.openById(SHEET_ID);
  var reqSheet = ss.getSheetByName(TABS.requests);
  reqSheet.getRange(parseInt(req.rowIndex), 7).setValue('filled');
  reqSheet.getRange(parseInt(req.rowIndex), 8).setValue(subEmail);

  updateScheduleForSub(ss, {
    matchDate:      req.matchDate,
    requestorEmail: req.email,
    subName:        subName,
    subEmail:       subEmail
  });

  markVolunteerMatched(ss, subEmail, req.matchDate);

  sendConfirmationEmails({
    requestorName:  req.name,
    requestorEmail: req.email,
    subName:        subName,
    subEmail:       subEmail,
    matchDate:      req.matchDate,
    matchTime:      req.matchTime
  }, req.groupPlayers || []);

  return { success: true };
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

// Returns the dispatch phase and skill window for a request given the 4-window config.
// Phase  | Hours until match       | Skill window
// -------+-------------------------+---------------------------------
// last-minute  | <= lastMinuteThresholdHrs  | skillWindowLastMinute
// urgent       | <= urgentThresholdHrs      | skillWindowUrgent
// post-schedule| <= preScheduleThresholdHrs | skillWindowMid
// pre-schedule | > preScheduleThresholdHrs  | skillWindowFarOut
function getDispatchPhase(req, config) {
  if (!req.matchDate) return { phase: 'pre-schedule', skillWindow: config.skillWindowFarOut };
  var timeStr = req.matchTime || '08:00';
  var matchDT = new Date(req.matchDate + 'T' + timeStr + ':00');
  var diffHrs = (matchDT - new Date()) / 36e5;
  if (diffHrs <= (config.lastMinuteThresholdHrs  || 24)) return { phase: 'last-minute',  skillWindow: config.skillWindowLastMinute || 2.8 };
  if (diffHrs <= (config.urgentThresholdHrs       || 48)) return { phase: 'urgent',        skillWindow: config.skillWindowUrgent  || 2.0 };
  if (diffHrs <= (config.preScheduleThresholdHrs  || 72)) return { phase: 'post-schedule', skillWindow: config.skillWindowMid     || 1.0 };
  return { phase: 'pre-schedule', skillWindow: config.skillWindowFarOut || 0.5 };
}

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

  var missing = getPlayersWithoutSubmission(config.targetMonth).filter(function(p) {
    return !/^anita\.sub\d+@xgmail\.com$/i.test(p.email || '');
  });
  if (!missing.length) {
    Logger.log('checkAvailabilityWindow: T-' + daysUntilClose + ' reminder — all players already submitted');
    return;
  }

  var closeDateLabel = closeDate.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
  var urgency        = daysUntilClose === 1 ? 'tomorrow' : 'in 2 days';
  var avUrl          = APP_BASE_URL + '#availability';
  var subject        = 'Reminder: Submit your availability for ' + config.targetMonthLabel + ' — closes ' + urgency;
  var body =
    'Just a reminder — the availability window for ' + config.targetMonthLabel + ' closes ' + urgency + ' (' + closeDateLabel + ').\n\n' +
    'Please submit your available dates before the window closes so we can include you in the schedule.\n\n' +
    'Open the My Availability page to submit:\n' +
    avUrl + '\n\n' +
    'See you on the court!\n' +
    'MWF Tennis League';
  var htmlBody =
    'Just a reminder — the availability window for <strong>' + config.targetMonthLabel + '</strong> closes ' + urgency + ' (' + closeDateLabel + ').<br><br>' +
    'Please submit your available dates before the window closes so we can include you in the schedule.<br><br>' +
    'Open the <a href="' + avUrl + '">My Availability</a> page to submit.<br><br>' +
    'See you on the court!<br>' +
    'MWF Tennis League';

  Logger.log('checkAvailabilityWindow: T-' + daysUntilClose + ' reminder → ' + missing.length + ' player(s)');
  if (!isEmailEnabled()) return;

  var toList = missing.map(function(p) { return p.email; }).join(', ');
  sendLeagueEmail({ to: toList, subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
}

function testAvailabilityEmail() {
  var config = getAvailabilityConfig();
  var closeDateLabel = 'Friday, April 25';
  var avUrl = APP_BASE_URL + '#availability';
  var subject = '[TEST] MWF League - Submit your availability for ' + config.targetMonthLabel;
  var body =
    'Hi,\n\n' +
    'It\'s time to submit your availability for ' + config.targetMonthLabel + '.\n\n' +
    'Please submit your available dates by ' + closeDateLabel + '.\n\n' +
    'Open the My Availability page to get started:\n' +
    avUrl + '\n\n' +
    'See you on the court!\n' +
    'MWF Tennis League';
  var htmlBody =
    'Hi,<br><br>' +
    'It\'s time to submit your availability for <strong>' + config.targetMonthLabel + '</strong>.<br><br>' +
    'Please submit your available dates by ' + closeDateLabel + '.<br><br>' +
    'Open the <a href="' + avUrl + '">My Availability</a> page to get started.<br><br>' +
    'See you on the court!<br>' +
    'MWF Tennis League';
  MailApp.sendEmail({ to: 'brianna.biesecker@gmail.com, marobria@gmail.com', subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
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

  // Flush writes before reading config back (prevents stale-cache reads)
  SpreadsheetApp.flush();

  // Send email blast to all players — wrapped so a send failure doesn't undo the window open
  var emailError = null;
  var emailCount = 0;
  try {
    // Exclude fictitious Anita Sub players — they don't need availability emails
    const allPlayers = getPlayers().filter(function(p) {
      return p.email && !/^anita\.sub\d+@xgmail\.com$/i.test(p.email);
    });
    emailCount = allPlayers.length;
    if (allPlayers.length && isEmailEnabled()) {
      const availConfig    = getAvailabilityConfig();
      const mailConfig     = getConfig();
      const closeDateLabel = new Date(closeDate + 'T12:00:00').toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
      const avUrl          = APP_BASE_URL + '#availability';
      const subject        = 'MWF League - Submit your availability for ' + availConfig.targetMonthLabel;
      const body =
        'It\'s time to submit your availability for ' + availConfig.targetMonthLabel + '.\n\n' +
        'Please submit your available dates by ' + closeDateLabel + '.\n\n' +
        'Open the My Availability page to get started:\n' +
        avUrl + '\n\n' +
        'See you on the court!\n' +
        'MWF Tennis League';
      const htmlBody =
        'It\'s time to submit your availability for <strong>' + availConfig.targetMonthLabel + '</strong>.<br><br>' +
        'Please submit your available dates by ' + closeDateLabel + '.<br><br>' +
        'Open the <a href="' + avUrl + '">My Availability</a> page to get started.<br><br>' +
        'See you on the court!<br>' +
        'MWF Tennis League';
      if (mailConfig.brevoAvailNotification && mailConfig.brevoApiKey) {
        const recipients = allPlayers.map(function(p) { return { email: p.email, name: p.name }; });
        sendBrevoEmail({
          apiKey:      mailConfig.brevoApiKey,
          senderName:  'MWF Tennis League',
          senderEmail: mailConfig.senderEmail,
          recipients:  recipients,
          subject:     subject,
          htmlContent: htmlBody,
          textContent: body
        });
      } else {
        const toList = allPlayers.map(function(p) { return p.email; }).join(', ');
        sendLeagueEmail({ to: toList, subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
      }
    }
  } catch(e) {
    emailError = e.message;
    Logger.log('openAvailabilityWindow email error: ' + e.message);
  }

  return { success: true, playerCount: emailCount, emailError: emailError };
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

    const avUrl   = APP_BASE_URL + '#availability';
    const subject = 'MWF League - Your availability for ' + avConfig.targetMonthLabel + ' is confirmed';
    const body =
      'Hi ' + name + ',\n\n' +
      'We received your availability for ' + avConfig.targetMonthLabel + '.\n\n' +
      'Your selected dates:\n' + (dateLines || '  (none selected)') + '\n\n' +
      (notes ? 'Notes: ' + notes + '\n\n' : '') +
      'If you need to make changes before the window closes, visit the My Availability page:\n' +
      avUrl + '\n\n' +
      'See you on the court!\n' +
      'MWF Tennis League';

    const htmlDateRows = dates.map(function(d) {
      return '<div>' + new Date(d + 'T12:00:00').toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' }) + '</div>';
    }).join('');
    const htmlBody =
      'Hi ' + name + ',<br><br>' +
      'We received your availability for <strong>' + avConfig.targetMonthLabel + '</strong>.<br><br>' +
      'Your selected dates:<br>' + (htmlDateRows || '(none selected)') + '<br>' +
      (notes ? 'Notes: ' + notes + '<br><br>' : '') +
      'If you need to make changes before the window closes, visit the <a href="' + avUrl + '">My Availability</a> page.<br><br>' +
      'See you on the court!<br>' +
      'MWF Tennis League';

    if (isEmailEnabled()) sendLeagueEmail({ to: email, subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
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
// Reads scheduler weight rows from Config tab (B20–B25, B31–B32).
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
    var rrRaw  = configSheet.getRange('B31:B32').getValues();
    var rrLimit = parseFloat(rrRaw[0][0]);
    var wMRR    = parseFloat(rrRaw[1][0]);
    var settings = {
      weightTeamVariance:    isNaN(wTV)     ? 1.0 : wTV,
      weightGroupVariance:   isNaN(wGV)     ? 0.5 : wGV,
      weightSocialVariety:   isNaN(wSV)     ? 2.0 : wSV,
      weightRecency:         isNaN(wRec)    ? 1.5 : wRec,
      solverIterations:      isNaN(iters)   ? 800  : iters,
      solverRestarts:        isNaN(rests)   ? 10   : rests,
      ratingRangeLimit:      isNaN(rrLimit) ? 2.0  : rrLimit,
      weightMaxRatingRange:  isNaN(wMRR)   ? 0.0  : wMRR
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
      weightTeamVariance:   1.0,
      weightGroupVariance:  0.5,
      weightSocialVariety:  2.0,
      weightRecency:        1.5,
      solverIterations:     800,
      solverRestarts:       10,
      ratingRangeLimit:     2.0,
      weightMaxRatingRange: 0.0,
      targetMonth:          '',
      targetMonthLabel:     '',
      submissionCount:      0
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

    // Scheduler weights (B20–B25) + max rating range (B31–B32)
    var raw = configSheet.getRange('B20:B25').getValues();
    var wTV   = parseFloat(raw[0][0]);
    var wGV   = parseFloat(raw[1][0]);
    var wSV   = parseFloat(raw[2][0]);
    var wRec  = parseFloat(raw[3][0]);
    var iters = parseInt(raw[4][0]);
    var rests = parseInt(raw[5][0]);
    var rrRaw  = configSheet.getRange('B31:B32').getValues();
    var rrLimit = parseFloat(rrRaw[0][0]);
    var wMRR    = parseFloat(rrRaw[1][0]);

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
      weightTeamVariance:   isNaN(wTV)     ? 1.0 : wTV,
      weightGroupVariance:  isNaN(wGV)     ? 0.5 : wGV,
      weightSocialVariety:  isNaN(wSV)     ? 2.0 : wSV,
      weightRecency:        isNaN(wRec)    ? 1.5 : wRec,
      solverIterations:     isNaN(iters)   ? 800  : iters,
      solverRestarts:       isNaN(rests)   ? 10   : rests,
      ratingRangeLimit:     isNaN(rrLimit) ? 2.0  : rrLimit,
      weightMaxRatingRange: isNaN(wMRR)   ? 0.0  : wMRR
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
      no8am:  playerMap[email] ? playerMap[email].no8am : false,
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

  // Shuffle dates so the Social Variety goal doesn't unfairly benefit end-of-month dates.
  // Results are re-sorted chronologically before returning so the preview stays readable.
  var slotKeys = Object.keys(slotMap);
  for (var si = slotKeys.length - 1; si > 0; si--) {
    var sj = Math.floor(Math.random() * (si + 1));
    var st = slotKeys[si]; slotKeys[si] = slotKeys[sj]; slotKeys[sj] = st;
  }

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
    slotResults.push({ date: date, skipped: false, groups: result.groups, sitOut: result.sitOut, sitOut2: result.sitOut2 || null });

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
    if (result.sitOut2) {
      sitOutCounts[result.sitOut2.email] = (sitOutCounts[result.sitOut2.email] || 0) + 1;
    }
  });

  // Re-sort chronologically so the schedule preview is in date order
  slotResults.sort(function(a, b) { return (a.date || '').localeCompare(b.date || ''); });

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

// Picks (and removes) one alternate from pool: prefer no8am players, and among
// those prefer players who haven't sat out yet this month (max 1 sit-out/month),
// falling back to the full pool if no candidates meet a preference.
function pickAlternate(pool, sitOutCounts) {
  var no8amPool = pool.filter(function(p) { return p.no8am; });
  var basePool  = no8amPool.length > 0 ? no8amPool : pool;
  var notYetSatOut = basePool.filter(function(p) { return (sitOutCounts[p.email] || 0) === 0; });
  var candidates   = notYetSatOut.length > 0 ? notYetSatOut : basePool;
  var chosen       = candidates[Math.floor(Math.random() * candidates.length)];
  return pool.splice(pool.indexOf(chosen), 1)[0];
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
    groupSizes = fillArray((n - 2) / 4, 4);
  } else {
    groupSizes = fillArray(Math.floor(n / 4), 4).concat([3]);
  }

  var sitOutPlayer  = null;
  var sitOutPlayer2 = null;
  var pool = available.slice();

  if (remainder === 1) {
    sitOutPlayer = pickAlternate(pool, sitOutCounts);
  } else if (remainder === 2) {
    sitOutPlayer  = pickAlternate(pool, sitOutCounts);
    sitOutPlayer2 = pickAlternate(pool, sitOutCounts);
  }

  var iters    = settings.solverIterations || 800;
  var restarts = settings.solverRestarts   || 10;
  var wTV    = settings.weightTeamVariance    || 1.0;
  var wGV    = settings.weightGroupVariance   || 0.5;
  var wSV    = settings.weightSocialVariety   || 2.0;
  var wMRR   = settings.weightMaxRatingRange  || 0.0;
  var rrLimit = settings.ratingRangeLimit !== undefined ? settings.ratingRangeLimit : 2.0;

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
    var gv, tv, rMax, rMin;
    if (sz === 4) {
      var r3 = group[3].rating;
      var m4 = (r0 + r1 + r2 + r3) * 0.25;
      gv = ((r0-m4)*(r0-m4) + (r1-m4)*(r1-m4) + (r2-m4)*(r2-m4) + (r3-m4)*(r3-m4)) * 0.25;
      var d01 = r0 - r1, d23 = r2 - r3;
      tv = (d01*d01 + d23*d23) * 0.25;
      rMax = Math.max(r0, r1, r2, r3);
      rMin = Math.min(r0, r1, r2, r3);
    } else {
      var m3 = (r0 + r1 + r2) / 3;
      gv = ((r0-m3)*(r0-m3) + (r1-m3)*(r1-m3) + (r2-m3)*(r2-m3)) / 3;
      tv = gv;
      rMax = Math.max(r0, r1, r2);
      rMin = Math.min(r0, r1, r2);
    }
    var rangePenalty = (rMax - rMin) > rrLimit ? wMRR : 0;
    return tv * wTV + gv * wGV + social + rangePenalty;
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

  return { groups: outputGroups, sitOut: sitOutPlayer, sitOut2: sitOutPlayer2 };
}

// ── Chunked Publish Helpers ─────────────────────────
// Step 1: clear existing rows for the month.
function clearAnitaRecords() {
  var ss           = SpreadsheetApp.openById(SHEET_ID);
  var anitaPattern = /^Anita Sub\d+$/;
  var anitaEmailRe = /^anita\.sub\d+@xgmail\.com$/i;
  var today        = formatSheetDate(new Date());

  // ── Step 1: find which Anita emails still have open FUTURE sub requests ──
  // We must preserve those players and requests so current-month play continues.
  var rSheet           = ss.getSheetByName(TABS.requests);
  var activeAnitaEmails = {};   // email → true
  if (rSheet && rSheet.getLastRow() >= 2) {
    var reqAll = rSheet.getRange(2, 1, rSheet.getLastRow() - 1, 7).getValues();
    reqAll.forEach(function(r) {
      var email     = (r[3] || '').toString().trim().toLowerCase();
      var matchDate = formatSheetDate(r[4]);
      var status    = (r[6] || '').toString();
      if (anitaEmailRe.test(email) && status === 'open' && matchDate > today) {
        activeAnitaEmails[email] = true;
      }
    });
  }

  // ── Step 2: remove only Anita players with no active future requests ──
  var pSheet = ss.getSheetByName(TABS.players);
  if (pSheet && pSheet.getLastRow() >= 2) {
    var numCols = Math.max(pSheet.getLastColumn(), 1);
    var allData = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, numCols).getValues();
    var keep = allData.filter(function(r) {
      if (!anitaPattern.test((r[0] || '').toString().trim())) return true;
      var email = (r[1] || '').toString().trim().toLowerCase();
      return !!activeAnitaEmails[email]; // keep if still needed
    });
    var removed = allData.length - keep.length;
    if (removed > 0) {
      pSheet.getRange(2, 1, allData.length, numCols).clearContent();
      if (keep.length > 0) pSheet.getRange(2, 1, keep.length, numCols).setValues(keep);
      pSheet.deleteRows(keep.length + 2, removed);
    }
  }

  // Dispatch's expireUpToToday() handles cleanup of past records.
  // Publishing must never change the status of any sub request or volunteer record.
}

function publishScheduleStart(params) {
  var month = params.month || '';
  if (!month) return { error: 'Month required.' };

  clearAnitaRecords();

  // MatchGroups: read-filter-rewrite (one batch delete instead of N deleteRow calls)
  var sheet = getOrCreateMatchGroupsSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var numCols  = Math.max(sheet.getLastColumn(), 14);
    var allRows  = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var keep     = allRows.filter(function(r) { return normalizeMonth(r[1]) !== month; });
    var removed  = allRows.length - keep.length;
    if (removed > 0) {
      sheet.getRange(2, 1, allRows.length, numCols).clearContent();
      if (keep.length > 0) sheet.getRange(2, 1, keep.length, numCols).setValues(keep);
      sheet.deleteRows(keep.length + 2, removed); // remove leftover blank rows in one call
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
  var sitOutName   = slot.sitOut  ? slot.sitOut.name   : '';
  var sitOutEmail  = slot.sitOut  ? slot.sitOut.email  : '';
  var sitOut2Name  = slot.sitOut2 ? slot.sitOut2.name  : '';
  var sitOut2Email = slot.sitOut2 ? slot.sitOut2.email : '';

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
      var anitaRating = Math.round(((partnerRating + avgOf3) / 2) * 100) / 100;

      // Add Anita to Players sheet — build row using getColMap so it works for both layouts
      var anitaCol = getColMap(pSheet);
      var anitaRow = [];
      anitaRow[anitaCol.name]   = anitaName;
      anitaRow[anitaCol.email]  = anitaEmail;
      if (anitaCol.phone >= 0) anitaRow[anitaCol.phone] = '';
      anitaRow[anitaCol.rating] = anitaRating;
      anitaRow[anitaCol.no8am]  = false;
      anitaRow[anitaCol.isAdmin]= false;
      // Fill any undefined gaps so appendRow doesn't truncate
      for (var ai = 0; ai < anitaCol.isAdmin + 1; ai++) {
        if (anitaRow[ai] === undefined) anitaRow[ai] = '';
      }
      pSheet.appendRow(anitaRow);
      pSheet.getRange(pSheet.getLastRow(), anitaCol.rating + 1).setNumberFormat('0.0');

      // Create Sub Request for Anita — captain goes first in groupPlayers so the
      // captain can identify and manage this request on the Request A Sub page.
      var groupForRequest = workingGroup.slice().sort(function(a, b) {
        if (a.email === captainEmail) return -1;
        if (b.email === captainEmail) return 1;
        return 0;
      });
      var groupPlayersJSON = JSON.stringify(groupForRequest.map(function(p) {
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
      var captainPlayer = workingGroup.find(function(p) { return p.email.toLowerCase() === captainEmail.toLowerCase(); });
      sendCaptainThreePlayerNotification(captainPlayer ? captainPlayer.name : '', captainEmail, slot.date, anitaName);
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
      sitOutName, sitOutEmail,
      sitOut2Name, sitOut2Email
    ]);
    saved++;
  });

  // Create a Volunteer record for the sit-out player so they can be matched as a sub
  if (sitOutEmail && sitOutName) {
    // Check No8am flag — reuse pSheet if already loaded, otherwise open now
    var sitOutTimes = '08_00,09_30,11_00,12_30';
    var lookupSheet = pSheet || ss.getSheetByName(TABS.players);
    if (lookupSheet && lookupSheet.getLastRow() >= 2) {
      var pLookup = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, 5).getValues();
      for (var pi = 0; pi < pLookup.length; pi++) {
        if ((pLookup[pi][1] || '').toLowerCase().trim() === sitOutEmail.toLowerCase().trim()) {
          var no8am = pLookup[pi][4]; // col E
          if (no8am === true || (no8am && no8am.toString().toUpperCase() === 'TRUE')) {
            sitOutTimes = '09_30,11_00,12_30'; // exclude 8:00 AM
          }
          break;
        }
      }
    }
    var volSheet = ss.getSheetByName(TABS.volunteers);
    var volRange = volSheet.getRange(volSheet.getLastRow() + 1, 1, 1, 7);
    volRange.setNumberFormats([['@','@','@','@','@','@','@']]);
    var thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
    volRange.setValues([[
      uid(), thirtyDaysAgo,
      sitOutName, sitOutEmail.toLowerCase(),
      slot.date, sitOutTimes, 'pending'
    ]]);
    Logger.log('Created volunteer record for sit-out: ' + sitOutName + ' on ' + slot.date + ' times: ' + sitOutTimes + ' (timestamp backdated 30 days)');
    sendSitOutNotification(sitOutName, sitOutEmail, slot.date);
  }

  // Create a Volunteer record for the 2nd alternate (remainder===2 case)
  if (sitOut2Email && sitOut2Name) {
    var sitOut2Times = '08_00,09_30,11_00,12_30';
    var lookupSheet2 = pSheet || ss.getSheetByName(TABS.players);
    if (lookupSheet2 && lookupSheet2.getLastRow() >= 2) {
      var pLookup2 = lookupSheet2.getRange(2, 1, lookupSheet2.getLastRow() - 1, 5).getValues();
      for (var pi2 = 0; pi2 < pLookup2.length; pi2++) {
        if ((pLookup2[pi2][1] || '').toLowerCase().trim() === sitOut2Email.toLowerCase().trim()) {
          var no8am2 = pLookup2[pi2][4];
          if (no8am2 === true || (no8am2 && no8am2.toString().toUpperCase() === 'TRUE')) {
            sitOut2Times = '09_30,11_00,12_30';
          }
          break;
        }
      }
    }
    var volSheet2 = ss.getSheetByName(TABS.volunteers);
    var volRange2 = volSheet2.getRange(volSheet2.getLastRow() + 1, 1, 1, 7);
    volRange2.setNumberFormats([['@','@','@','@','@','@','@']]);
    var thirtyDaysAgo2 = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
    volRange2.setValues([[
      uid(), thirtyDaysAgo2,
      sitOut2Name, sitOut2Email.toLowerCase(),
      slot.date, sitOut2Times, 'pending'
    ]]);
    Logger.log('Created volunteer record for 2nd alternate: ' + sitOut2Name + ' on ' + slot.date + ' times: ' + sitOut2Times + ' (timestamp backdated 30 days)');
    sendSitOutNotification(sitOut2Name, sitOut2Email, slot.date);
  }

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

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();

  // Build dateMap across all months; track latestMonth for the header label only
  var latestMonth = '';
  var dateMap = {};
  rows.forEach(function(r) {
    var m = normalizeMonth(r[1]);
    if (!m) return;
    if (m > latestMonth) latestMonth = m;

    var date = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (r[2] ? r[2].toString() : '');
    var letter = r[3] ? r[3].toString() : '';
    var sitOutName   = r[12] ? r[12].toString() : '';
    var sitOutEmail  = r[13] ? r[13].toString() : '';
    var sitOut2Name  = r[14] ? r[14].toString() : '';
    var sitOut2Email = r[15] ? r[15].toString() : '';

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
      sitOut:  sitOutName  ? { name: sitOutName,  email: sitOutEmail  } : null,
      sitOut2: sitOut2Name ? { name: sitOut2Name, email: sitOut2Email } : null
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
          sitOut:  dateMap[date][letter].sitOut,
          sitOut2: dateMap[date][letter].sitOut2
        };
      })
    };
  });

  return { month: latestMonth, dates: dates, no8amEmails: no8amEmails };
}

// Builds a CSV schedule attachment that opens in Excel.
// Uses only Utilities.newBlob — no new OAuth scopes required.
function buildScheduleAttachments(schedule, monthLabel) {
  var safe     = monthLabel.replace(/\s/g, '_');
  var csvLines = ['"MWF Tennis League — ' + monthLabel + ' Schedule"', ''];

  schedule.dates.forEach(function(dayObj) {
    var dateLabel = new Date(dayObj.date + 'T12:00:00').toLocaleDateString('en-US',
      { weekday: 'long', month: 'long', day: 'numeric' });
    csvLines.push('"' + dateLabel.replace(/"/g, '""') + '"');
    dayObj.groups.forEach(function(grp) {
      var real = grp.players.filter(function(p) {
        return p.name && !/^anita\.sub\d+@xgmail\.com$/i.test(p.email || '');
      });
      var row = ['Group ' + grp.letter];
      real.forEach(function(p) { row.push(p.name); });
      while (row.length < 5) row.push('');
      row.push(grp.sitOut ? '(sub needed)' : '');
      csvLines.push(row.map(function(v) { return '"' + (v || '').replace(/"/g, '""') + '"'; }).join(','));
    });
    csvLines.push('');
  });

  // BOM (﻿) ensures Excel reads UTF-8 correctly on Windows
  return [Utilities.newBlob('﻿' + csvLines.join('\r\n'), 'text/csv', safe + '_Schedule.csv')];
}

// Diagnostic: sends a test email to the admin and returns full diagnostic info.
function sendTestEmail() {
  var diag = { emailEnabled: false, playerCount: 0, senderEmail: '', scheduleFound: false };
  try {
    diag.emailEnabled = isEmailEnabled();
    diag.senderEmail  = getConfig().senderEmail || '';
    var players = getPlayersWithRatings();
    diag.playerCount  = players.length;
    var sched = getPublishedSchedule();
    diag.scheduleFound = !!(sched && sched.month);
    diag.scheduleMonth = sched ? sched.month : '';

    if (!diag.emailEnabled) {
      return { success: false, error: 'Email is disabled — turn on the Email Enabled toggle in Admin settings (Config B27).', diag: diag };
    }
    if (!players.length) {
      return { success: false, error: 'No players found in the Players sheet.', diag: diag };
    }

    var adminEmail = Session.getActiveUser().getEmail() || 'marobria@gmail.com';
    var body = 'Rally test email.\n\nDiagnostics:\n' +
      'Email enabled: ' + diag.emailEnabled + '\n' +
      'Player count: ' + diag.playerCount + '\n' +
      'Sender email: ' + (diag.senderEmail || '(none — will send from script account)') + '\n' +
      'Schedule found: ' + diag.scheduleFound + ' (' + diag.scheduleMonth + ')';

    if (diag.senderEmail) {
      try {
        GmailApp.sendEmail(adminEmail, 'Rally — Test Email', body, {
          from: diag.senderEmail, replyTo: diag.senderEmail, name: 'MWF Tennis League'
        });
        return { success: true, sentTo: adminEmail, sentFrom: diag.senderEmail, diag: diag };
      } catch(ge) {
        diag.gmailError = ge.message;
      }
    }
    MailApp.sendEmail({ to: adminEmail, subject: 'Rally — Test Email', body: body, name: 'MWF Tennis League' });
    return { success: true, sentTo: adminEmail, sentFrom: 'script account (MailApp)', diag: diag };
  } catch(e) {
    return { success: false, error: e.message, diag: diag };
  }
}

// Sends the published schedule to ALL players in one email (all addresses on To line) with CSV attachment.
function buildScheduleEmailParts(schedule) {
  var parts = schedule.month.split('-');
  var monthLabel = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, 1)
    .toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  var scheduleUrl = APP_BASE_URL + '#schedule';
  var textLines = [], htmlRows = [];
  schedule.dates.forEach(function(dayObj) {
    var dateLabel = new Date(dayObj.date + 'T12:00:00').toLocaleDateString('en-US',
      { weekday: 'long', month: 'long', day: 'numeric' });
    textLines.push(dateLabel.toUpperCase());
    htmlRows.push('<tr><td colspan="2" style="padding:10px 0 4px;font-weight:700;' +
      'border-top:1px solid #E8EBF0;">' + dateLabel + '</td></tr>');
    dayObj.groups.forEach(function(grp) {
      var realPlayers = grp.players.filter(function(p) {
        return p.name && !/^anita\.sub\d+@xgmail\.com$/i.test(p.email || '');
      });
      var names = realPlayers.map(function(p) { return p.name; }).join(', ');
      textLines.push('  Group ' + grp.letter + ': ' + names + (grp.sitOut ? ' (sub needed)' : ''));
      htmlRows.push('<tr><td style="padding:2px 16px;color:#374151;font-weight:600;">' +
        'Group ' + grp.letter + '</td><td style="padding:2px 8px;">' + names +
        (grp.sitOut ? ' <em style="color:#8A4F0B;">(sub needed)</em>' : '') + '</td></tr>');
    });
    textLines.push('');
  });
  var body = 'The MWF Tennis League schedule for ' + monthLabel + ' has been published.\n\n' +
    textLines.join('\n') +
    'Court times will be announced separately as each date approaches.\n\n' +
    'View the schedule online: ' + scheduleUrl + '\n\n' +
    'The schedule is also attached as a spreadsheet file (CSV) that opens in Excel.';
  var htmlBody = '<p>The MWF Tennis League schedule for <strong>' + monthLabel +
    '</strong> has been published.</p>' +
    '<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;">' +
    htmlRows.join('') + '</table>' +
    '<p style="margin-top:16px;">Court times will be announced separately as each date approaches.</p>' +
    '<p><a href="' + scheduleUrl + '">View Schedule</a></p>' +
    '<p style="color:#666;font-size:12px;margin-top:12px;">The schedule is also attached as a spreadsheet file (CSV, opens in Excel).</p>';
  return { subject: 'MWF Tennis League — ' + monthLabel + ' Schedule Published', body: body, htmlBody: htmlBody };
}

function sendScheduleEmails(params) {
  if (!isEmailEnabled()) return { success: true, emailsSent: 0, skipped: 'email_disabled' };

  var schedule = getPublishedSchedule();
  if (!schedule.month || !schedule.dates || !schedule.dates.length) {
    return { success: false, error: 'No published schedule found.' };
  }

  var emailParts  = buildScheduleEmailParts(schedule);
  var config      = getConfig();
  var allPlayers  = getPlayersWithRatings()
    .filter(function(p) { return p.email && !/^anita\.sub\d+@xgmail\.com$/i.test(p.email); });
  if (!allPlayers.length) return { success: true, emailsSent: 0 };

  try {
    if (config.brevoScheduleEmail && config.brevoApiKey) {
      var recipients = allPlayers.map(function(p) { return { email: p.email, name: p.name }; });
      sendBrevoEmail({
        apiKey:       config.brevoApiKey,
        senderName:   'MWF Tennis League',
        senderEmail:  config.senderEmail,
        recipients:   recipients,
        subject:      emailParts.subject,
        htmlContent:  emailParts.htmlBody,
        textContent:  emailParts.body
      });
    } else {
      var toList = allPlayers.map(function(p) { return p.email; }).join(', ');
      var opts   = { name: 'MWF Tennis League', htmlBody: emailParts.htmlBody };
      if (config.senderEmail) {
        try {
          GmailApp.sendEmail(toList, emailParts.subject, emailParts.body,
            Object.assign({}, opts, { from: config.senderEmail, replyTo: config.senderEmail }));
          return { success: true, emailsSent: allPlayers.length };
        } catch(ge) {
          Logger.log('GmailApp failed (' + ge.message + '), falling back to MailApp');
        }
      }
      opts.to = toList; opts.subject = emailParts.subject; opts.body = emailParts.body;
      MailApp.sendEmail(opts);
    }
  } catch(e) {
    return { success: false, error: 'Email failed: ' + e.message };
  }

  return { success: true, emailsSent: allPlayers.length };
}

function sendTestScheduleEmail() {
  var config = getConfig();
  if (!config.brevoApiKey) {
    return { success: false, error: 'Brevo API key not set. Enter it in Config sheet B35.' };
  }

  var sd = buildScheduleDataFromMatchGroups();
  if (!sd || !sd.sortedDates.length) {
    return { success: false, error: 'No published schedule found.' };
  }

  // getPlayersWithRatings() auto-inits the Test column header if missing
  var testPlayers = getPlayersWithRatings()
    .filter(function(p) {
      return p.email && !/^anita\.sub\d+@xgmail\.com$/i.test(p.email) && p.isTest;
    })
    .map(function(p) { return { email: p.email, name: p.name }; });
  if (!testPlayers.length) {
    return { success: false, error: 'No test players found — add "Yes" in the Test column of the Players sheet.' };
  }

  var scheduleUrl = APP_BASE_URL + '#schedule';
  var htmlBody    = buildScheduleHtml(sd.dateMap, sd.sortedDates, sd.monthLabel, scheduleUrl);
  var csvContent  = buildScheduleCsv(sd.dateMap, sd.sortedDates, sd.monthLabel, sd.playerNameMap);
  var csvFileName = sd.monthLabel.replace(/\s/g, '_') + '_Schedule.csv';
  var subject     = 'MWF Tennis League — ' + sd.monthLabel + ' Schedule';

  try {
    sendBrevoEmail({
      apiKey:      config.brevoApiKey,
      senderName:  'MWF Tennis League',
      senderEmail: config.senderEmail,
      recipients:  testPlayers,
      subject:     subject,
      htmlContent: htmlBody,
      attachments: [{ content: Utilities.base64Encode('﻿' + csvContent), name: csvFileName }]
    });
  } catch(e) {
    return { success: false, error: 'Brevo send failed: ' + e.message };
  }
  return { success: true, emailsSent: testPlayers.length };
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
