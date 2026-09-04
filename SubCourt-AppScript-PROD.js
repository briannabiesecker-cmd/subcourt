// ══════════════════════════════════════════════════
// SUBCOURT — Apps Script Web App
// MWF Tennis League
// ══════════════════════════════════════════════════

const SHEET_ID = '1hA-ZPhV62pp376qtWRDfQQkFv6y9U5Wkm0nUyKCHC6o';

// Completely side-effect-free: no sheet reads or writes, no email, no Gmail/Drive
// calls. Exists solely as an obviously-safe function to run once from the Apps
// Script editor's Run button, to trigger Google's permission-grant screen after
// new scopes are added to the manifest — without doing anything else.
function authorizeApp() {
  Logger.log('authorizeApp: no-op. If you see this line, authorization succeeded.');
}

// Execution-level cache for getConfig() — resets between trigger/HTTP invocations.
var _configCache = null;

// deploy.sh replaces 'rally-tennis-prod.html' with 'rally-tennis-prod.html' when pushing to prod.
const APP_BASE_URL  = 'https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-prod.html';
const SCRIPT_URL    = 'https://script.google.com/macros/s/AKfycbzb3EnQsxBt5dLTaQpg7VJjtoBtHTyGpB2VgpfJ9TDuvezk0ihjhn5oW48a9oKiIAyYMg/exec';

// Email enabled state is stored in Config B20 and toggled from the Admin UI.
// Do not hardcode this — use isEmailEnabled() instead.
function isEmailEnabled() {
  try {
    var v = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config).getRange('B27').getValue();
    return v === true || v.toString().toUpperCase() === 'TRUE';
  } catch(e) { return false; }
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

function sendBrevoEmail(params) {
  // params: { apiKey, recipients: [{email, name}], cc, bcc, subject, htmlContent, textContent, attachments, replyTo: {email, name} }
  var payload = {
    sender: { name: 'MWF Tennis League', email: 'noreply@mtctennis.com' },
    to: params.recipients,
    subject: params.subject
  };
  if (params.cc)            payload.cc          = params.cc;
  if (params.bcc)           payload.bcc         = params.bcc;
  if (params.htmlContent)  payload.htmlContent = params.htmlContent;
  if (params.textContent)  payload.textContent = params.textContent;
  if (params.attachments)  payload.attachment  = params.attachments;
  if (params.replyTo)      payload.replyTo     = params.replyTo;
  payload.headers = {
    'List-Unsubscribe':      '<mailto:noreply@mtctennis.com?subject=Unsubscribe>',
    'List-Unsubscribe-Post': 'List-Unsubscribe=One-Click'
  };
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

// Called via google.script.run from the confirmation page — no navigation needed.
// Canonical Overflow check, used everywhere a match time needs to be told apart
// from Overflow — a time counts as Overflow only when the MatchGroups cell (or a
// value about to be written to it) literally is that exact string. Blank,
// whitespace-only, "TBD", or anything else is never Overflow — those are treated
// as unknown and must be asked for, never silently assumed to be Overflow.
function _isOverflowTime(v) {
  return (v || '').toString().trim() === 'Overflow';
}

// Resolves what the player's own match time is on matchDate (per MatchGroups),
// mirroring the Volunteer to Sub screen's warnScheduledDateRow logic: not
// scheduled that day → no conflict possible. Scheduled with a real time or
// Overflow → that's authoritative. Scheduled but still blank/TBD → the caller
// must supply ownMatchTime (collected from the player first) before this can
// resolve; if it isn't supplied yet, needsMatchTime comes back true so the
// caller can ask, then re-call with the answer. A supplied ownMatchTime is
// saved back to MatchGroups so it only has to be asked once.
function _resolvePlayerOwnMatchTime(ss, matchDate, playerEmail, ownMatchTime, playerName) {
  var emailLower = (playerEmail || '').toLowerCase();
  var playingElsewhere = _getPlayerMatchTimesForDate(ss, matchDate);
  if (!Object.prototype.hasOwnProperty.call(playingElsewhere, emailLower)) {
    return { scheduled: false, matchTime: '' };
  }
  var theirTime = (playingElsewhere[emailLower] || '').toString().trim();
  if (theirTime) return { scheduled: true, matchTime: theirTime };

  ownMatchTime = (ownMatchTime || '').toString().trim();
  if (!ownMatchTime || (TIMES.indexOf(ownMatchTime) === -1 && !_isOverflowTime(ownMatchTime))) {
    return { scheduled: true, matchTime: '', needsMatchTime: true };
  }
  var groupRow = _findMatchGroupRow(ss, matchDate, [playerEmail]);
  if (groupRow) {
    _setMatchGroupTime(matchDate, groupRow.letter, ownMatchTime,
      'Ask match time (Volunteer to Sub / I CAN Sub)', playerName, playerEmail);
  }
  return { scheduled: true, matchTime: ownMatchTime };
}

// True when the player's own resolved match time (see above) exactly conflicts
// with matchTime — Overflow, blank, or "not scheduled that day" are never a
// conflict, same as the Volunteer to Sub screen's rules.
function _isExactMatchConflict(ownMatch, matchTime) {
  return !!(ownMatch.matchTime && !_isOverflowTime(ownMatch.matchTime) && matchTime && ownMatch.matchTime === matchTime);
}

function processVolunteerFromEmail(requestId, playerEmail, ownMatchTime, playTwiceChoice) {
  requestId       = (requestId       || '').trim();
  playerEmail     = (playerEmail     || '').trim().toLowerCase();
  playTwiceChoice = (playTwiceChoice || '').trim();
  if (!requestId || !playerEmail) return { success: false, error: 'Invalid parameters.' };
  var requests = getRequests();
  var req;
  for (var i = 0; i < requests.length; i++) {
    if (requests[i].id === requestId) { req = requests[i]; break; }
  }
  if (!req) return { success: false, error: 'This sub request could not be found. It may have already been filled.' };

  var players    = getPlayers();
  var playerName = '';
  var found      = false;
  for (var j = 0; j < players.length; j++) {
    if (players[j].email && players[j].email.toLowerCase() === playerEmail) {
      playerName = players[j].name || '';
      found = true;
      break;
    }
  }
  if (!found) return { success: false, error: 'That email address was not found in the league roster. Please check your email and try again.' };

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ownMatch = _resolvePlayerOwnMatchTime(ss, req.matchDate, playerEmail, ownMatchTime, playerName);
  if (ownMatch.needsMatchTime) {
    return { success: false, needsMatchTime: true, dateStr: formatDate(req.matchDate) };
  }
  if (_isExactMatchConflict(ownMatch, req.matchTime)) {
    return { success: false, error: 'You are already scheduled to play at ' + (TIME_LABELS[req.matchTime] || req.matchTime) + ' that day, so you can\'t sub for this request.' };
  }

  // Scheduled to play at a different time (or Overflow) this same day — same as
  // the Volunteer to Sub screen, ask whether they want to play twice or have
  // Rally look for a sub at their own match time instead, before creating the
  // volunteer record. "Change times" checks for (and creates, if missing) an
  // open sub request for their own match, exactly like _resolveScheduledConflict
  // does for the calendar screen.
  if (ownMatch.scheduled && !playTwiceChoice) {
    return {
      success: false,
      needsPlayTwiceChoice: true,
      ownMatchTimeLabel: TIME_LABELS[ownMatch.matchTime] || ownMatch.matchTime,
      dateStr: formatDate(req.matchDate)
    };
  }
  var ownRequestCreated = false;
  if (ownMatch.scheduled && playTwiceChoice === 'changeTimes') {
    try {
      var groupRow = _findMatchGroupRow(ss, req.matchDate, [playerEmail]);
      if (groupRow) {
        var hasOpenRequest = requests.some(function(r) {
          return r.email.toLowerCase() === playerEmail &&
            r.matchDate === req.matchDate &&
            r.matchTime === ownMatch.matchTime &&
            r.status === 'open';
        });
        if (!hasOpenRequest) {
          var partners = (groupRow.players || []).filter(function(p) {
            return p.email && p.email.toLowerCase() !== playerEmail && p.name;
          });
          var ownReqResult = submitRequest({
            name: playerName, email: playerEmail, matchDate: req.matchDate,
            matchTime: ownMatch.matchTime, groupLetter: groupRow.letter, groupPlayers: partners
          });
          ownRequestCreated = !!(ownReqResult && ownReqResult.success);
        }
      }
    } catch (e) {
      Logger.log('processVolunteerFromEmail: change-times sub request failed: ' + e.message);
    }
  }

  var alreadyFilled = req.status !== 'open';
  var filledNote = req.status === 'filled' ? 'has already been filled' : 'is no longer active';
  var timeCode = req.matchTime
    ? req.matchTime.replace(':', '_')
    : TIMES.map(function(t) { return t.replace(':', '_'); }).join(',');
  var volSheet = ss.getSheetByName(TABS.volunteers);
  // A player never gets more than one non-cancelled volunteer record on the same
  // date: if one already exists, this time slot is merged into it rather than a
  // second record being created for the day.
  var upsertResult = upsertVolunteerTimes(volSheet, playerName, playerEmail, req.matchDate, timeCode.split(','));
  if (upsertResult.created) {
    try { _notifyLateVolunteerForTomorrow(playerName, playerEmail, req.matchDate, timeCode.split(',')); }
    catch(e) { Logger.log('_notifyLateVolunteerForTomorrow failed: ' + e.message); }
  }
  Logger.log('Volunteer from email: ' + playerName + ' (' + playerEmail + ') for request ' + requestId);
  return {
    success: true, playerName: playerName, dateStr: formatDate(req.matchDate),
    shortDateStr: formatDateShort(req.matchDate), timeStr: TIME_LABELS[req.matchTime] || req.matchTime || '',
    alreadyFilled: alreadyFilled, filledNote: filledNote,
    ownRequestCreated: ownRequestCreated,
    ownRequestTimeStr: ownRequestCreated ? (TIME_LABELS[ownMatch.matchTime] || ownMatch.matchTime) : ''
  };
}

function handleVolunteerFromEmail(e) {
  var p            = e.parameter || {};
  var requestId    = (p.requestId    || '').trim();
  var playerEmail  = (p.playerEmail  || '').trim().toLowerCase();
  var notAvailable = p.notAvailable === 'true' || p.notAvailable === '1';
  var css = 'body{font-family:Arial,sans-serif;max-width:480px;margin:40px auto;padding:0 20px;color:#111;}' +
            'h2{color:#1a5c3a;}p{line-height:1.6;font-size:15px;}';
  var wrap = function(body) {
    return HtmlService.createHtmlOutput(
      '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<style>' + css + '</style></head><body>' + body + '</body></html>'
    );
  };
  if (!requestId) return wrap('<p>Invalid link.</p>');

  var requests = getRequests();
  var req;
  for (var i = 0; i < requests.length; i++) { if (requests[i].id === requestId) { req = requests[i]; break; } }

  if (!req) return wrap('<p>This sub request could not be found. It may have already been filled.</p>');

  var dateStr      = formatDate(req.matchDate);
  var shortDateStr = formatDateShort(req.matchDate);
  var timeStr      = TIME_LABELS[req.matchTime] || req.matchTime || '';
  var alreadyFilled = req.status !== 'open';
  var filledNote = req.status === 'filled' ? 'has already been filled' : 'is no longer active';

  if (notAvailable) {
    return wrap(
      '<p style="font-size:18px;font-weight:bold;color:#1a5c3a;">No problem!</p>' +
      '<p>Thanks for letting us know. We\'ll keep looking for a sub.</p>' +
      '<p style="color:#6b7280;font-size:13px;margin-top:24px;">MWF Tennis League</p>'
    );
  }

  if (!playerEmail) {
    // BCC path — confirmation page; uses google.script.run so no navigation needed.
    // Mirrors the Volunteer to Sub screen's conflict handling: if the player is
    // scheduled to play that day with no match time recorded yet,
    // processVolunteerFromEmail comes back with needsMatchTime and this page
    // collects it (Overflow included, no TBD option) before resubmitting. Once
    // the time is known, an exact conflict is blocked outright; scheduled at a
    // different time (or Overflow) instead comes back with needsPlayTwiceChoice,
    // prompting the player the same "play twice or change times" question the
    // calendar screen asks on submit.
    var timeLabel = timeStr ? ' at ' + timeStr : '';
    var reqIdJs   = requestId.replace(/'/g, "\\'");
    var timeOptionsHtml = (TIMES.map(function(t) {
      return '<option value="' + t + '">' + TIME_LABELS[t] + '</option>';
    }).join('') + '<option value="Overflow">Overflow</option>').replace(/'/g, "\\'");
    return wrap(
      '<div id="pg">' +
        '<h2 style="color:#1a5c3a;font-size:22px;margin-bottom:8px;">I can sub' + timeLabel + '<br>on ' + dateStr + '</h2>' +
        '<p style="margin-bottom:16px;">Enter your email address to confirm:</p>' +
        '<input type="email" id="em" placeholder="your@email.com" autocomplete="email" ' +
          'style="width:100%;padding:12px;font-size:16px;border:1px solid #ccc;border-radius:4px;margin-bottom:16px;box-sizing:border-box;">' +
        '<div id="btns" style="display:flex;gap:12px;">' +
          '<button id="btnC" style="flex:1;padding:14px;background:#1a5c3a;color:#fff;border:none;border-radius:4px;font-size:16px;font-weight:bold;cursor:pointer;">Confirm</button>' +
          '<button id="btnD" style="flex:1;padding:14px;background:#e5e7eb;color:#374151;border:none;border-radius:4px;font-size:16px;font-weight:bold;cursor:pointer;">Not Available</button>' +
        '</div>' +
        '<p id="errmsg" style="display:none;color:#dc2626;font-size:14px;margin-top:8px;"></p>' +
        '<p id="msg" style="display:none;color:#6b7280;margin-top:12px;">Processing…</p>' +
      '</div>' +
      '<p style="color:#6b7280;font-size:13px;margin-top:24px;">MWF Tennis League</p>' +
      '<script>' +
        'var RID=\'' + reqIdJs + '\';' +
        'var EM=null;' +
        'var TIME_OPTS=\'' + timeOptionsHtml + '\';' +
        'function submitVol(mt,choice){' +
          'var btns=document.getElementById(\'btns\');' +
          'if(btns){btns.style.display=\'none\';document.getElementById(\'msg\').style.display=\'block\';}' +
          'google.script.run' +
            '.withSuccessHandler(function(r){' +
              'if(!r.success){' +
                'if(r.needsMatchTime){showTimePicker(r.dateStr);return;}' +
                'if(r.needsPlayTwiceChoice){showPlayTwicePrompt(r.dateStr,r.ownMatchTimeLabel);return;}' +
                'var b=document.getElementById(\'btns\');' +
                'if(b){' +
                  'b.style.display=\'flex\';' +
                  'document.getElementById(\'msg\').style.display=\'none\';' +
                  'document.getElementById(\'errmsg\').textContent=r.error||\'An error occurred.\';' +
                  'document.getElementById(\'errmsg\').style.display=\'block\';' +
                '}else{' +
                  'document.getElementById(\'pg\').innerHTML=' +
                    '\'<h2 style="color:#c0392b;">Match time conflict</h2>\'+' +
                    '\'<p>\'+(r.error||\'An error occurred.\')+\'</p>\';' +
                '}' +
                'return;' +
              '}' +
              'var n=r.playerName?(", "+r.playerName.split(" ")[0]):"";' +
              'var statusLine=r.alreadyFilled' +
                '?(\'<p style="color:#c0392b;margin:0 0 4px;">Note: The \'+(r.timeStr?r.timeStr+\' \':\'\')+\'sub request on \'+r.shortDateStr+\' \'+(r.filledNote||\'is no longer active\')+\'.</p>\'+\'<p style="margin:0;">However, a volunteer record has been created for you.</p>\')' +
                ':\'<p>You will be notified if you are selected as a substitute.</p>\';' +
              'var ownRequestLine=r.ownRequestCreated' +
                '?\'<p style="margin-top:12px;">Rally has also submitted a sub request for your own match on <strong>\'+r.dateStr+\'</strong> at <strong>\'+r.ownRequestTimeStr+\'</strong>. You will be notified once a sub is found.</p>\'' +
                ':\'\';' +
              'document.getElementById(\'pg\').innerHTML=' +
                '\'<h2>Thank you\'+n+\'!</h2>\'+' +
                '\'<p>You have volunteered to sub on <strong>\'+r.dateStr+\'</strong>\'+' +
                '(r.timeStr?\' at <strong>\'+r.timeStr+\'</strong>\':\'\')+\'.</p>\'+' +
                'statusLine+' +
                'ownRequestLine;' +
            '})' +
            '.withFailureHandler(function(){' +
              'var b=document.getElementById(\'btns\');' +
              'if(b){b.style.display=\'flex\';document.getElementById(\'msg\').style.display=\'none\';}' +
              'alert(\'Something went wrong. Please try again.\');' +
            '})' +
            '.processVolunteerFromEmail(RID,EM,mt||\'\',choice||\'\');' +
        '}' +
        'function showPlayTwicePrompt(ds,timeLabel){' +
          'document.getElementById(\'pg\').innerHTML=' +
            '\'<h2 style="color:#1a5c3a;font-size:22px;margin-bottom:8px;">Do you want to play twice or change times?</h2>\'+' +
            '\'<p style="margin-bottom:16px;">You are scheduled to play at \'+timeLabel+\' on \'+ds+\'.</p>\'+' +
            '\'<div style="display:flex;gap:12px;">\'+' +
              '\'<button id="btnPT" style="flex:1;padding:14px;background:#1a5c3a;color:#fff;border:none;border-radius:4px;font-size:16px;font-weight:bold;cursor:pointer;">Play twice</button>\'+' +
              '\'<button id="btnCT" style="flex:1;padding:14px;background:#e5e7eb;color:#374151;border:none;border-radius:4px;font-size:16px;font-weight:bold;cursor:pointer;">Change times</button>\'+' +
            '\'</div>\';' +
          'document.getElementById(\'btnPT\').onclick=function(){' +
            'document.getElementById(\'pg\').innerHTML=\'<p style="color:#6b7280;">Processing…</p>\';' +
            'submitVol(\'\',\'playTwice\');' +
          '};' +
          'document.getElementById(\'btnCT\').onclick=function(){' +
            'document.getElementById(\'pg\').innerHTML=\'<p style="color:#6b7280;">Processing…</p>\';' +
            'submitVol(\'\',\'changeTimes\');' +
          '};' +
        '}' +
        'function showTimePicker(ds){' +
          'document.getElementById(\'pg\').innerHTML=' +
            '\'<h2 style="color:#1a5c3a;font-size:22px;margin-bottom:8px;">What time are you playing on \'+ds+\'?</h2>\'+' +
            '\'<p style="margin-bottom:16px;">You are scheduled to play tennis that day. Please select your match time so Rally can check for a conflict.</p>\'+' +
            '\'<select id="mt" style="width:100%;padding:12px;font-size:16px;border:1px solid #ccc;border-radius:4px;margin-bottom:16px;box-sizing:border-box;">\'+' +
              '\'<option value="" selected disabled>Select a time…</option>\'+' +
              'TIME_OPTS+' +
            '\'</select>\'+' +
            '\'<button id="btnMT" style="width:100%;padding:14px;background:#1a5c3a;color:#fff;border:none;border-radius:4px;font-size:16px;font-weight:bold;cursor:pointer;">Continue</button>\'+' +
            '\'<p id="errmsg2" style="display:none;color:#dc2626;font-size:14px;margin-top:8px;"></p>\';' +
          'document.getElementById(\'btnMT\').onclick=function(){' +
            'var sel=document.getElementById(\'mt\').value;' +
            'if(!sel){' +
              'var em2=document.getElementById(\'errmsg2\');' +
              'em2.textContent=\'Please select a time.\';' +
              'em2.style.display=\'block\';' +
              'return;' +
            '}' +
            'document.getElementById(\'pg\').innerHTML=\'<p style="color:#6b7280;">Processing…</p>\';' +
            'submitVol(sel);' +
          '};' +
        '}' +
        'document.getElementById(\'btnC\').onclick=function(){' +
          'var val=(document.getElementById(\'em\').value||\'\').trim();' +
          'if(!val){alert(\'Please enter your email address.\');return;}' +
          'EM=val;' +
          'submitVol(\'\');' +
        '};' +
        'document.getElementById(\'btnD\').onclick=function(){' +
          'document.getElementById(\'pg\').innerHTML=' +
            '\'<p style="font-size:18px;font-weight:bold;color:#1a5c3a;">No problem!</p>\'+' +
            '\'<p>Thanks for letting us know. We\\\'ll keep looking for a sub.</p>\';' +
        '};' +
      '<\/script>'
    );
  }

  // Email present — look up name from Players sheet and create volunteer record
  var ss2 = SpreadsheetApp.openById(SHEET_ID);
  var ownMatch2 = _resolvePlayerOwnMatchTime(ss2, req.matchDate, playerEmail, '');
  if (ownMatch2.needsMatchTime) {
    // This static link has no form to collect a match time interactively —
    // send them to the Volunteer to Sub screen, which will ask for it.
    return wrap(
      '<h2 style="color:#c0392b;">Please confirm your match time first</h2>' +
      '<p>You are scheduled to play on <strong>' + dateStr + '</strong> but Rally does not know your match time yet. ' +
      'Open the <a href="' + APP_BASE_URL + '#volunteer">Volunteer to Sub</a> page, which will ask for it before letting you sub.</p>' +
      '<p style="color:#6b7280;font-size:13px;margin-top:24px;">MWF Tennis League</p>'
    );
  }
  if (_isExactMatchConflict(ownMatch2, req.matchTime)) {
    return wrap(
      '<h2 style="color:#c0392b;">Match time conflict</h2>' +
      '<p>You are already scheduled to play at <strong>' + timeStr + '</strong> on <strong>' + dateStr + '</strong>, so you can not volunteer to sub for this request.</p>' +
      '<p style="color:#6b7280;font-size:13px;margin-top:24px;">MWF Tennis League</p>'
    );
  }
  var players    = getPlayers();
  var playerName = '';
  for (var j = 0; j < players.length; j++) {
    if (players[j].email && players[j].email.toLowerCase() === playerEmail) {
      playerName = players[j].name || '';
      break;
    }
  }
  var timeCode = req.matchTime
    ? req.matchTime.replace(':', '_')
    : TIMES.map(function(t) { return t.replace(':', '_'); }).join(',');
  var volSheet2 = ss2.getSheetByName(TABS.volunteers);
  // A player never gets more than one non-cancelled volunteer record on the same
  // date: if one already exists, this time slot is merged into it rather than a
  // second record being created for the day.
  var upsertResult2 = upsertVolunteerTimes(volSheet2, playerName, playerEmail, req.matchDate, timeCode.split(','));
  if (upsertResult2.created) {
    try { _notifyLateVolunteerForTomorrow(playerName, playerEmail, req.matchDate, timeCode.split(',')); }
    catch(e) { Logger.log('_notifyLateVolunteerForTomorrow failed: ' + e.message); }
  }
  Logger.log('Volunteer from email: ' + playerName + ' (' + playerEmail + ') for request ' + requestId);

  var statusLine = alreadyFilled
    ? '<p style="color:#c0392b;margin:0 0 4px;">Note: The ' +
      (timeStr ? timeStr + ' ' : '') + 'sub request on ' + shortDateStr + ' ' + filledNote + '.</p>' +
      '<p style="margin:0;">However, a volunteer record has been created for you.</p>'
    : '<p>You will be notified if you are selected as a substitute.</p>';

  return wrap(
    '<h2>Thank you' + (playerName ? ', ' + playerName.split(' ')[0] : '') + '!</h2>' +
    '<p>You have volunteered to sub on <strong>' + dateStr + '</strong>' +
    (timeStr ? ' at <strong>' + timeStr + '</strong>' : '') + '.</p>' +
    statusLine +
    '<p style="color:#6b7280;font-size:13px;margin-top:24px;">MWF Tennis League</p>'
  );
}

var BCC_CHUNK_SIZE = 45; // MailApp per-message recipient limit is ~50; stay safely under it

// Same-day identical-content sends are skipped — guards against duplicate emails
// from overlapping trigger runs or manual re-triggers within the same day.
function _getEmailThrottleDateKey(date) {
  try {
    return Utilities.formatDate(date || new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return (date || new Date()).toISOString().slice(0, 10);
  }
}

function _buildEmailContentSignature(params) {
  var payload = JSON.stringify({
    to: params.to || '',
    cc: params.cc || '',
    bcc: params.bcc || '',
    subject: params.subject || '',
    body: params.body || '',
    htmlBody: params.htmlBody || '',
    name: params.name || ''
  });
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload);
  return Utilities.base64Encode(digest);
}

function _isTomorrowOrDayAfterTomorrow(matchDate) {
  try {
    var tz = Session.getScriptTimeZone();
    var today = new Date(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd') + 'T00:00:00');
    var target = new Date(matchDate + 'T00:00:00');
    var diffDays = Math.round((target.getTime() - today.getTime()) / 86400000);
    return diffDays === 1 || diffDays === 2;
  } catch (e) {
    return false;
  }
}

// Calendar-day difference between today and matchDate (negative/0 = today or past).
// Used to split the Substitute Confirmed email into Future/Urgent variants.
function _daysUntilMatch(matchDate) {
  try {
    var tz = Session.getScriptTimeZone();
    var today = new Date(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd') + 'T00:00:00');
    var target = new Date(matchDate + 'T00:00:00');
    return Math.round((target.getTime() - today.getTime()) / 86400000);
  } catch (e) {
    return 99; // unparseable date — default to the Future variant
  }
}

function _getVolunteerCcEmailsForMatch(matchDate, matchTime, players) {
  if (!matchDate || !players || !players.length) return [];
  var effectiveTime = (matchTime || '08:00').trim();
  var volunteerMatches = getVolunteers().filter(function(v) {
    if (!v || !v.email) return false;
    if (v.date !== matchDate) return false;
    if (!v.times || !v.times.length) return false;
    if (v.status && ['matched', 'cancelled', 'expired'].indexOf(String(v.status).toLowerCase()) >= 0) return false;
    return v.times.some(function(t) { return String(t).trim() === effectiveTime; });
  });

  var emails = [];
  volunteerMatches.forEach(function(v) {
    var email = _resolveEmail(v.name, v.email, players);
    if (!email) return;
    if (emails.map(function(existing) { return existing.toLowerCase(); }).indexOf(email.toLowerCase()) >= 0) return;
    emails.push(email);
  });
  return emails;
}

// Drops toEmail out of a BCC list — used wherever an admin address is both the To:
// and, since the admin is also a real player, would otherwise also land in a
// roster-wide BCC list, causing the admin to get every broadcast twice.
function _excludeFromBcc(emails, toEmail) {
  var skip = (toEmail || '').toLowerCase();
  return (emails || []).filter(function(e) { return e && e.toLowerCase() !== skip; });
}

// Unified email sender for every real Rally email — routes through Brevo when
// configured (own quota, unaffected by MailApp's), falling back to MailApp otherwise
// or if Brevo itself fails.
//
// comcast.net recipients are routed around Brevo entirely, straight through MailApp.
// Measured via getBrevoBounceSummary: every comcast.net address in the roster soft-
// bounces through Brevo's shared sending IP at a ~16% rate (554 "server not available"
// from Comcast's resimta MTA), while every other domain bounces at 0% — a domain-level
// rejection Brevo's own automatic retries never recover from. MailApp (Google's own
// sending infrastructure) is a separate reputation path that isn't affected by it.
function sendLeagueEmail(params) {
  var props = PropertiesService.getScriptProperties();
  var throttleKey = 'emailThrottle:' + _getEmailThrottleDateKey(new Date()) + ':' + _buildEmailContentSignature(params);
  if (props.getProperty(throttleKey)) {
    Logger.log('Skipping duplicate email for unchanged content: ' + (params.subject || ''));
    return;
  }

  var config = getConfig();
  var split  = _splitOffComcastRecipients(params);

  if (split.comcast) _sendLeagueEmailViaMailApp(split.comcast, config);

  if (split.rest) {
    // Brevo is the primary path for everyone else — it has its own quota, independent of
    // MailApp's daily recipient cap. Falls through to MailApp below if Brevo isn't
    // configured or fails.
    var sentViaBrevo = false;
    if (config.brevoApiKey) {
      try {
        _sendLeagueEmailViaBrevo(split.rest, config);
        _logEmail(split.rest.to, split.rest.subject, 'sent via Brevo');
        sentViaBrevo = true;
      } catch(e) {
        Logger.log('Brevo send failed for "' + split.rest.subject + '", falling back to MailApp: ' + e.message);
        _logEmail(split.rest.to, split.rest.subject, 'Brevo failed (' + e.message + '), trying MailApp');
      }
    }
    if (!sentViaBrevo) _sendLeagueEmailViaMailApp(split.rest, config);
  }

  props.setProperty(throttleKey, 'sent');
}

function _splitAddrList(str) {
  return (str || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean);
}

// Partitions an email's to/cc/bcc into a comcast.net-only params object and an
// everyone-else params object (either may be null if that group has no recipients),
// each carrying the rest of the original params (subject/body/htmlBody/name/etc.)
// unchanged. Borrows a recipient into "to" if a group's own To list is empty — MailApp
// and Brevo both require a non-empty "to" on every send.
function _splitOffComcastRecipients(params) {
  var toList  = _splitAddrList(params.to);
  var ccList  = _splitAddrList(params.cc);
  var bccList = _splitAddrList(params.bcc);
  var isComcast = function(a) { return /@comcast\.net$/i.test(a); };

  var comcastTo = [], comcastCc = [], comcastBcc = [];
  var restTo    = [], restCc    = [], restBcc    = [];
  toList.forEach(function(a)  { (isComcast(a) ? comcastTo  : restTo).push(a); });
  ccList.forEach(function(a)  { (isComcast(a) ? comcastCc  : restCc).push(a); });
  bccList.forEach(function(a) { (isComcast(a) ? comcastBcc : restBcc).push(a); });

  function build(to, cc, bcc) {
    if (!to.length && !cc.length && !bcc.length) return null;
    if (!to.length) to = bcc.length ? [bcc.shift()] : [cc.shift()];
    var p = {};
    for (var k in params) p[k] = params[k];
    p.to  = to.join(', ');
    p.cc  = cc.length  ? cc.join(', ')  : undefined;
    p.bcc = bcc.length ? bcc.join(', ') : undefined;
    return p;
  }

  return {
    comcast: build(comcastTo, comcastCc, comcastBcc),
    rest:    build(restTo, restCc, restBcc)
  };
}

function _sendLeagueEmailViaMailApp(params, config) {
  var senderEmail = config.senderEmail || '';
  var baseOpts = { name: params.name || 'MWF Tennis League' };
  if (params.htmlBody)     baseOpts.htmlBody = params.htmlBody;
  if (params.cc)           baseOpts.cc       = params.cc;
  if (params.attachments)  baseOpts.attachments = params.attachments;
  if (senderEmail)         baseOpts.replyTo  = senderEmail;
  else if (params.replyTo) baseOpts.replyTo  = params.replyTo;

  var bccAddrs = params.bcc
    ? params.bcc.split(',').map(function(s) { return s.trim(); }).filter(Boolean)
    : [];

  if (bccAddrs.length <= BCC_CHUNK_SIZE) {
    var opts = { name: baseOpts.name };
    if (baseOpts.htmlBody)    opts.htmlBody    = baseOpts.htmlBody;
    if (baseOpts.cc)          opts.cc          = baseOpts.cc;
    if (baseOpts.replyTo)     opts.replyTo     = baseOpts.replyTo;
    if (baseOpts.attachments) opts.attachments = baseOpts.attachments;
    if (bccAddrs.length)      opts.bcc         = bccAddrs.join(',');
    try {
      MailApp.sendEmail(params.to, params.subject, params.body, opts);
      _logEmail(params.to, params.subject, 'sent via MailApp (1 to + ' + bccAddrs.length + ' bcc: ' + bccAddrs.join('; ') + ')');
    } catch(e) {
      _logEmail(params.to, params.subject, 'failed: ' + e.message);
      _sendAdminFallbackEmail(params);
      throw e;
    }
  } else {
    // Split BCC into chunks to stay under per-message recipient limit. Only the first
    // chunk uses params.to (the visible admin address) as the "To" — MailApp requires
    // a non-empty "to" on every call, so later chunks borrow one of their own bcc
    // addresses for that slot instead of reusing params.to, which previously sent the
    // admin one full duplicate email per extra chunk.
    for (var i = 0; i < bccAddrs.length; i += BCC_CHUNK_SIZE) {
      var chunk    = bccAddrs.slice(i, i + BCC_CHUNK_SIZE);
      var chunkTo  = i === 0 ? params.to : chunk[0];
      var chunkBcc = i === 0 ? chunk : chunk.slice(1);
      var chunkOpts = { name: baseOpts.name };
      if (chunkBcc.length)      chunkOpts.bcc    = chunkBcc.join(',');
      if (baseOpts.htmlBody)    chunkOpts.htmlBody    = baseOpts.htmlBody;
      if (baseOpts.cc)          chunkOpts.cc          = baseOpts.cc;
      if (baseOpts.replyTo)     chunkOpts.replyTo     = baseOpts.replyTo;
      if (baseOpts.attachments) chunkOpts.attachments = baseOpts.attachments;
      if (i > 0) Utilities.sleep(500);
      try {
        MailApp.sendEmail(chunkTo, params.subject, params.body, chunkOpts);
        _logEmail(chunkTo, params.subject, 'sent via MailApp (bcc chunk ' + (Math.floor(i / BCC_CHUNK_SIZE) + 1) +
          ', 1 to + ' + chunkBcc.length + ' bcc: ' + chunkBcc.join('; ') + ')');
      } catch(e) {
        _logEmail(chunkTo, params.subject, 'failed chunk ' + (Math.floor(i / BCC_CHUNK_SIZE) + 1) + ': ' + e.message);
        _sendAdminFallbackEmail(params);
        throw e;
      }
    }
  }
}

// Recipient-count chunking mirrors the MailApp path above — kept conservative since
// Brevo's exact per-call recipient ceiling isn't documented as precisely as MailApp's.
function _sendLeagueEmailViaBrevo(params, config) {
  var toAddrs  = (params.to  || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  var ccAddrs  = params.cc   ? params.cc.split(',').map(function(s) { return s.trim(); }).filter(Boolean)  : [];
  var bccAddrs = params.bcc  ? params.bcc.split(',').map(function(s) { return s.trim(); }).filter(Boolean) : [];
  if (!toAddrs.length) throw new Error('No "to" address for Brevo send.');

  var replyToEmail = config.senderEmail || params.replyTo || '';

  function toRecipients(list) { return list.map(function(e) { return { email: e }; }); }

  function buildParams(bccChunk, toOverride) {
    var p = {
      apiKey:      config.brevoApiKey,
      recipients:  toRecipients(toOverride || toAddrs),
      subject:     params.subject,
      textContent: params.body
    };
    if (params.htmlBody)    p.htmlContent = params.htmlBody;
    if (ccAddrs.length)     p.cc          = toRecipients(ccAddrs);
    if (bccChunk.length)    p.bcc         = toRecipients(bccChunk);
    if (replyToEmail)       p.replyTo     = { email: replyToEmail };
    if (params.attachments) p.attachments = _blobsToBrevoAttachments(params.attachments);
    return p;
  }

  if (bccAddrs.length <= BCC_CHUNK_SIZE) {
    sendBrevoEmail(buildParams(bccAddrs));
  } else {
    // Only the first chunk uses toAddrs (the visible admin address) as "recipients" —
    // later chunks borrow one of their own bcc addresses instead, so toAddrs doesn't
    // get a duplicate copy per chunk.
    for (var i = 0; i < bccAddrs.length; i += BCC_CHUNK_SIZE) {
      if (i > 0) Utilities.sleep(300);
      var chunk = bccAddrs.slice(i, i + BCC_CHUNK_SIZE);
      if (i === 0) { sendBrevoEmail(buildParams(chunk)); continue; }
      sendBrevoEmail(buildParams(chunk.slice(1), [chunk[0]]));
    }
  }
}

function _blobsToBrevoAttachments(blobs) {
  return blobs.map(function(b) {
    return { content: Utilities.base64Encode(b.getBytes()), name: b.getName() };
  });
}

function _logEmail(to, subject, status) {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(TABS.emailLog);
    if (!sheet) {
      sheet = ss.insertSheet(TABS.emailLog);
      sheet.appendRow(['Timestamp', 'To', 'Subject', 'Status']);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), to, subject, status]);
  } catch(e) {
    Logger.log('_logEmail error: ' + e.message);
  }
}

// When sendLeagueEmail fails, sends the same email to the admin for manual forwarding.
// Uses 1 quota slot. If <20 addressees, lists them at the top; otherwise says "send to all".
function _sendAdminFallbackEmail(params) {
  var ADMIN_EMAIL = 'marobria@gmail.com';
  var ADDR_THRESHOLD = 20;

  var allAddrs = [];
  if (params.to) allAddrs.push(params.to);
  if (params.cc) params.cc.split(',').forEach(function(a) { var t = a.trim(); if (t) allAddrs.push(t); });
  if (params.bcc) params.bcc.split(',').forEach(function(a) { var t = a.trim(); if (t) allAddrs.push(t); });

  var recipientText, recipientHtml;
  if (allAddrs.length < ADDR_THRESHOLD) {
    recipientText = 'Forward to:\n' + allAddrs.join('\n');
    recipientHtml = '<strong>Forward to:</strong><br>' + allAddrs.join('<br>');
  } else {
    recipientText = 'Forward to: send to all (' + allAddrs.length + ' recipients)';
    recipientHtml = '<strong>Forward to:</strong> send to all (' + allAddrs.length + ' recipients)';
  }

  var bannerHtml =
    '<div style="background:#fff3cd;border:2px solid #f0ad4e;border-radius:4px;' +
    'padding:12px 16px;margin:0 0 16px 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#333;">' +
    recipientHtml + '</div>';

  var htmlBody;
  if (params.htmlBody) {
    // Inject banner immediately after the opening <body> tag
    var injected = params.htmlBody.replace(/(<body[^>]*>)/i, '$1' + bannerHtml);
    htmlBody = (injected !== params.htmlBody) ? injected : bannerHtml + params.htmlBody;
  } else {
    htmlBody = bannerHtml + '<pre style="font-family:Arial,sans-serif;font-size:14px;">' + (params.body || '') + '</pre>';
  }

  try {
    MailApp.sendEmail({
      to:       ADMIN_EMAIL,
      subject:  '[Forward this email] ' + (params.subject || ''),
      body:     recipientText + '\n\n---\n\n' + (params.body || ''),
      htmlBody: htmlBody,
      name:     'MWF Tennis League'
    });
    Logger.log('_sendAdminFallbackEmail: sent fallback for "' + params.subject + '" to ' + ADMIN_EMAIL);
  } catch(e2) {
    Logger.log('_sendAdminFallbackEmail: also failed: ' + e2.message);
  }
}

// Sends one email per player; brief pause between sends to stay within Gmail rate limits.
function sendBulkEmails(players, buildParamsFn) {
  var sent = 0, firstError = null;
  var props = PropertiesService.getScriptProperties();
  players.forEach(function(player, i) {
    try {
      if (i > 0) Utilities.sleep(500);
      sendLeagueEmail(buildParamsFn(player));
      sent++;
    } catch(e) {
      Logger.log('sendBulkEmails: failed for ' + (player.email || '?') + ': ' + e.message);
      if (!firstError) {
        firstError = { email: player.email, error: e.message };
        // Write immediately so we capture it even if the script is later interrupted.
        try { props.setProperty('broadcastLog', JSON.stringify({
          time: new Date().toISOString(), total: players.length, sent: sent, firstError: firstError
        })); } catch(pe) {}
      }
    }
  });
  try {
    props.setProperty('broadcastLog', JSON.stringify({
      time: new Date().toISOString(), total: players.length, sent: sent, firstError: firstError
    }));
  } catch(pe) {}
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
    'Do not create a duplicate sub request.\n\n' +
    'MWF Tennis League';
  var htmlBody =
    'Hi ' + captainName + ',<br><br>' +
    'You are the captain of a 3-player group on ' + dateStr + ' and therefore a sub request has automatically been created for ' + anitaSubName + '.<br><br>' +
    'When <a href="https://midlothian.chelseareservations.com/login.aspx">Chelsea</a> assigns a court time, update the sub request on the <a href="' + reqUrl + '">Request a Sub</a> page.<br><br>' +
    'If Rally is unable to fill the request on ' + dayBefore + ', you will be notified by email. At that time, you should use the email/phone process to find a 4th player.<br><br>' +
    '<span style="color:#c0392b;">Do not create a duplicate sub request.</span><br><br>' +
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
  matchGroups:  'MatchGroups',
  emailLog:     'EmailLog'
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
  if (_configCache) return _configCache;
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);
    // Write labels/defaults for B31–B32 on first use (cells empty)
    var b31 = sheet.getRange('B31').getValue();
    var b32 = sheet.getRange('B32').getValue();
    if (b31 === '' || b31 === null) { sheet.getRange('A31').setValue('Rating Range Limit');      sheet.getRange('B31').setValue(2.0); }
    if (b32 === '' || b32 === null) { sheet.getRange('A32').setValue('Weight Maximum Rating Range'); sheet.getRange('B32').setValue(0.0); }
    // Brevo email section — auto-init on first use
    var b37 = sheet.getRange('B37').getValue();
    if (b37 === '' || b37 === null) {
      sheet.getRange('A34').setValue('── Brevo Email ──');
      sheet.getRange('A35').setValue('Brevo API Key');
      sheet.getRange('A37').setValue('Address Schedule Email on To: vs BCC:');
      sheet.getRange('B37').setValue('No');
    }
    // One-time migration: relabel A37 for sheets already past the block above.
    if (sheet.getRange('A37').getValue() === 'Use Brevo: Schedule Email') {
      sheet.getRange('A37').setValue('Address Schedule Email on To: vs BCC:');
    }
    // One-time migration: retired field, no longer used anywhere — clear it rather
    // than delete the row (deleting would shift every hardcoded cell reference below it).
    if (sheet.getRange('A36').getValue() === 'Use Brevo: Availability Notification') {
      sheet.getRange('A36').setValue('');
      sheet.getRange('B36').setValue('');
    }
    // Urgent sub emails section — auto-init on first use
    var b39 = sheet.getRange('B39').getValue();
    if (b39 === '' || b39 === null) {
      sheet.getRange('A38').setValue('── Urgent Sub Emails ──');
      sheet.getRange('A39').setValue('Send Urgent Sub Emails');
      sheet.getRange('B39').setValue('Yes');
    }
    // Pre-match-day dispatch schedule — auto-init on first use (rows 41–47)
    var b43 = sheet.getRange('B43').getValue();
    if (b43 === '' || b43 === null) {
      sheet.getRange('A41').setValue('── Pre-Match Day Dispatch ──');
      sheet.getRange('A42').setValue('Run');
      sheet.getRange('B42').setValue('Time (ET)');
      sheet.getRange('C42').setValue('Dispatch');
      sheet.getRange('D42').setValue('Broadcast');
      sheet.getRange('E42').setValue('Cancel if open');
      sheet.getRange('A43:E47').setValues([
        ['1', '8:00 AM',  'Yes', 'Yes', 'No' ],
        ['2', '11:00 AM', 'Yes', 'Yes', 'No' ],
        ['3', '2:00 PM',  'Yes', 'Yes', 'No' ],
        ['4', '5:00 PM',  'Yes', 'Yes', 'No' ],
        ['5', '8:00 PM',  'Yes', 'Yes', 'Yes']
      ]);
    }
    // Expand Volunteers column — auto-init on first use (column F, rows 42–47).
    // Defaults to Yes only on the last run, mirroring "Cancel if open" — the last
    // run is the one that used to auto-relax candidate selection on its own;
    // now that's an explicit, per-run choice instead.
    var f42 = sheet.getRange('F42').getValue();
    if (f42 === '' || f42 === null) {
      sheet.getRange('F42').setValue('Expand Volunteers');
      sheet.getRange('F43:F47').setValues([['No'], ['No'], ['No'], ['No'], ['Yes']]);
    }
    var schedRows = sheet.getRange('A43:F47').getValues();
    var preMatchSchedule = schedRows.map(function(row) {
      return {
        run:              row[0].toString(),
        time:             formatSheetTime(row[1]) || row[1].toString().trim(),
        dispatch:         row[2] !== 'No' && row[2] !== false,
        broadcast:        row[3] !== 'No' && row[3] !== false,
        cancel:           row[4] === 'Yes' || row[4] === true,
        expandVolunteers: row[5] === 'Yes' || row[5] === true
      };
    });
    // Match Day -2 dispatch schedule — auto-init on first use (rows 49–55)
    var b50 = sheet.getRange('B50').getValue();
    if (b50 === '' || b50 === null) {
      sheet.getRange('A49').setValue('── Match Day -2 Dispatch ──');
      sheet.getRange('A50').setValue('Run');
      sheet.getRange('B50').setValue('Time (ET)');
      sheet.getRange('C50').setValue('Dispatch');
      sheet.getRange('D50').setValue('Broadcast');
      sheet.getRange('A51:D55').setValues([
        ['1', '8:00 AM',  'Yes', 'Yes'],
        ['2', '11:00 AM', 'Yes', 'Yes'],
        ['3', '2:00 PM',  'Yes', 'Yes'],
        ['4', '5:00 PM',  'Yes', 'Yes'],
        ['5', '8:00 PM',  'Yes', 'Yes']
      ]);
    }
    // Match time reminder column (former E) — removed. The reminder email now fires
    // directly from checkChelseaCourtTimes() instead of a manual per-run schedule flag.
    // One-time cleanup for sheets that still have the old header/values.
    if (sheet.getRange('E50').getValue() === 'Time Reminder') {
      sheet.getRange('E50:E55').clearContent();
    }
    // Overflow-detect column — auto-init on first use (column F, rows 50–55).
    var f50 = sheet.getRange('F50').getValue();
    if (f50 === '' || f50 === null) {
      sheet.getRange('F50').setValue('Overflow Detect');
      sheet.getRange('F51:F55').setValues([['No'], ['No'], ['No'], ['No'], ['No']]);
    }
    var md2Rows = sheet.getRange('A51:F55').getValues();
    var matchDayMinus2Schedule = md2Rows.map(function(row) {
      return {
        run:            row[0].toString(),
        time:           formatSheetTime(row[1]) || row[1].toString().trim(),
        dispatch:       row[2] !== 'No' && row[2] !== false,
        broadcast:      row[3] !== 'No' && row[3] !== false,
        overflowDetect: row[5] === 'Yes' || row[5] === true
      };
    });
    // Friday auto dispatch — auto-init on first use (rows 57–59)
    var b58 = sheet.getRange('B58').getValue();
    if (b58 === '' || b58 === null) {
      sheet.getRange('A57').setValue('── Friday Auto Dispatch ──');
      sheet.getRange('A58').setValue('Auto Dispatch Enabled');
      sheet.getRange('B58').setValue(true);
      sheet.getRange('A59').setValue('Time (ET)');
      sheet.getRange('B59').setNumberFormat('@');
      sheet.getRange('B59').setValue('13:00');
    }
    // Allow player name change on delete — auto-init on first use (row 61)
    var b61 = sheet.getRange('B61').getValue();
    if (b61 === '' || b61 === null) {
      sheet.getRange('A61').setValue('Allow Player Name Change on Delete Sub Request');
      sheet.getRange('B61').setValue('Yes');
    }
    // MTC contact — auto-init labels only on first use (rows 62–64). B63/B64 are left
    // blank on purpose: an empty MTC email address is the normal, expected default.
    if (sheet.getRange('A63').getValue() !== 'MTC Email Address 1') {
      sheet.getRange('A62').setValue('── MTC Contact ──');
      sheet.getRange('A63').setValue('MTC Email Address 1');
      sheet.getRange('A64').setValue('MTC Email Address 2');
    }
    // Chelsea court-sheet check window — auto-init on first use (rows 65–69)
    if (sheet.getRange('A65').getValue() !== '── Chelsea Court Sheet Check ──') {
      sheet.getRange('A65').setValue('── Chelsea Court Sheet Check ──');
      sheet.getRange('A66').setValue('Days to Check (e.g. Sat,Mon,Wed)');
      sheet.getRange('B66').setValue('Sat,Mon,Wed');
      sheet.getRange('A67').setValue('Start Time (ET, 24h HH:MM)');
      sheet.getRange('B67').setNumberFormat('@');
      sheet.getRange('B67').setValue('07:45');
      sheet.getRange('A68').setValue('End Time (ET, 24h HH:MM)');
      sheet.getRange('B68').setNumberFormat('@');
      sheet.getRange('B68').setValue('09:30');
      sheet.getRange('A69').setValue('Check Frequency (minutes)');
      sheet.getRange('B69').setValue(15);
    }
    // Chelsea import on/off — auto-init on first use (row 71). Defaults to
    // No/suspended, since this was added specifically to let the import be
    // switched off; an admin re-enables it from the Dispatch screen.
    var b71 = sheet.getRange('B71').getValue();
    if (b71 === '' || b71 === null) {
      sheet.getRange('A71').setValue('Chelsea Import Enabled');
      sheet.getRange('B71').setValue('No');
    }
    var cfg = {
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
      // Dispatch automation (Friday only) — rows 58–59
      autoDispatchEnabled:      (function() { var v = sheet.getRange('B58').getValue(); return v === true || v.toString().toUpperCase() === 'TRUE'; })(),
      autoDispatchTimeET:       formatSheetTime(sheet.getRange('B59').getValue()) || '13:00',
      // Sender email — row 30
      senderEmail: (sheet.getRange('B30').getValue() || '').toString().trim(),
      // Players Email Group — row 33
      playersGroupEmail: (sheet.getRange('B33').getValue() || '').toString().trim(),
      // Brevo — rows 35, 37
      brevoApiKey:            (sheet.getRange('B35').getValue() || '').toString().trim(),
      brevoScheduleEmail:      (function() { var v = sheet.getRange('B37').getValue(); return v === 'Yes' || v === true; })(),
      urgentSubEmailsEnabled:  (function() { var v = sheet.getRange('B39').getValue(); return v !== 'No' && v !== false; })(),
      preMatchSchedule: preMatchSchedule,
      // How many of today's Pre-Match Day dispatch runs haven't happened yet — 0 means
      // today's last scheduled dispatch for tomorrow's matches has already passed.
      remainingPreMatchRunsToday: _remainingPreMatchRunsToday({ preMatchSchedule: preMatchSchedule }),
      matchDayMinus2Schedule: matchDayMinus2Schedule,
      // Availability window — rows 16–18
      availWindowOpenDate:      (function() { var v = sheet.getRange('B16').getValue(); return v instanceof Date ? formatSheetDate(v) : (v ? v.toString() : ''); })(),
      availWindowCloseDate:     (function() { var v = sheet.getRange('B17').getValue(); return v instanceof Date ? formatSheetDate(v) : (v ? v.toString() : ''); })(),
      availWindowActive:        (function() { var v = sheet.getRange('B18').getValue(); return v === true || v.toString().toUpperCase() === 'TRUE'; })(),
      // Delete-request name change — row 61
      allowPlayerNameChangeOnDelete: (function() { var v = sheet.getRange('B61').getValue(); return v !== 'No' && v !== false; })(),
      // MTC contact — rows 63–64
      mtcEmail1: (sheet.getRange('B63').getValue() || '').toString().trim(),
      mtcEmail2: (sheet.getRange('B64').getValue() || '').toString().trim(),
      // Chelsea court-sheet check window — rows 66–69
      chelseaCheckDays:             (sheet.getRange('B66').getValue() || 'Sat,Mon,Wed').toString().trim(),
      chelseaCheckStartTime:        formatSheetTime(sheet.getRange('B67').getValue()) || '07:45',
      chelseaCheckEndTime:          formatSheetTime(sheet.getRange('B68').getValue()) || '09:30',
      chelseaCheckFrequencyMinutes: parseInt(sheet.getRange('B69').getValue()) || 15,
      chelseaCheckSubject:          (sheet.getRange('B70').getValue() || 'Upcoming Court Sheet').toString().trim(),
      chelseaImportEnabled:         (function() { var v = sheet.getRange('B71').getValue(); return v === 'Yes' || v === true; })(),
    };
    _configCache = cfg;
    return cfg;
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
      senderEmail: '',
      playersGroupEmail: '',
      brevoApiKey: '',
      brevoScheduleEmail: false,
      urgentSubEmailsEnabled: true,
      preMatchSchedule: [
        { run:'1', time:'8:00 AM',  dispatch:true, broadcast:true, cancel:false },
        { run:'2', time:'11:00 AM', dispatch:true, broadcast:true, cancel:false },
        { run:'3', time:'2:00 PM',  dispatch:true, broadcast:true, cancel:false },
        { run:'4', time:'5:00 PM',  dispatch:true, broadcast:true, cancel:false },
        { run:'5', time:'8:00 PM',  dispatch:true, broadcast:true, cancel:true  }
      ],
      matchDayMinus2Schedule: [
        { run:'1', time:'8:00 AM',  dispatch:true, broadcast:true },
        { run:'2', time:'11:00 AM', dispatch:true, broadcast:true },
        { run:'3', time:'2:00 PM',  dispatch:true, broadcast:true },
        { run:'4', time:'5:00 PM',  dispatch:true, broadcast:true },
        { run:'5', time:'8:00 PM',  dispatch:true, broadcast:true }
      ],
      availWindowOpenDate:     '',
      availWindowCloseDate:    '',
      availWindowActive:       false,
      allowPlayerNameChangeOnDelete: true,
      mtcEmail1: '',
      mtcEmail2: '',
      chelseaCheckDays: 'Sat,Mon,Wed',
      chelseaCheckStartTime: '07:45',
      chelseaCheckEndTime: '09:30',
      chelseaCheckFrequencyMinutes: 15,
      chelseaCheckSubject: 'Upcoming Court Sheet',
      chelseaImportEnabled: false,
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
// config watcher. Re-runs automatically when B58/B59
// are edited thereafter.
// ──────────────────────────────────────────────────

function setupTriggers() {
  // Remove all managed triggers
  var managed = ['runAutoDispatch','onConfigEdit','cleanupOldRecords','checkAvailabilityWindow',
                 'runPreMatchDayDispatch','runPreMatchDayDispatchFinal',
                 'runFollowupDispatchT1','runFollowupDispatchT2','runMatchTimeReminder',
                 '_runQueuedAvailBlast','_runMatchTimeReminderCheck','runMatchDayMinus2Dispatch',
                 'checkChelseaCourtTimes'];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (managed.indexOf(t.getHandlerFunction()) !== -1) {
      try { ScriptApp.deleteTrigger(t); } catch(e) {}
    }
  });

  // onEdit watcher for Config tab
  ScriptApp.newTrigger('onConfigEdit').forSpreadsheet(SHEET_ID).onEdit().create();

  // Monthly cleanup of old records across all data tabs
  ScriptApp.newTrigger('cleanupOldRecords').timeBased().onMonthDay(1).atHour(4).create();

  // Daily check to auto-close availability window
  ScriptApp.newTrigger('checkAvailabilityWindow').timeBased().atHour(4).everyDays(1).create();

  // Chelsea court-time email check — runs on the configured frequency (Config B69),
  // self-guards to the configured days/window (B66–B68) inside the handler (Apps
  // Script triggers can't be restricted to specific weekdays/times at minute
  // granularity, so this follows the same pattern as the daily dispatch triggers
  // below: install unconditionally, guard in the handler).
  updateChelseaCheckTrigger();

  // Daily dispatch (handles T+2 broadcast; T+1 is handled by pre-match-day triggers below)
  updateDispatchTrigger();

  // Pre-match-day dispatch: 5 runs on Sun/Tue/Thu (day before Mon/Wed/Fri matches).
  updatePreMatchDayTriggers();

  // Match day -2 dispatch: 5 runs on Sat/Mon/Wed (2 days before Mon/Wed/Fri matches).
  updateMatchDayMinus2Triggers();

  var config = getConfig();
  Logger.log('Triggers installed. Dispatch: ' +
    (config.autoDispatchEnabled ? 'Fridays at ' + config.autoDispatchTimeET + ' ET' : 'disabled') +
    '. Match day -2 runs: Sat/Mon/Wed at 8am, 11am, 2pm, 5pm, 8pm ET.' +
    '. Pre-match-day runs: Sun/Tue/Thu at 8am, 11am, 2pm, 5pm, 8pm ET.' +
    ' Chelsea check: ' + config.chelseaCheckDays + ' every ' + config.chelseaCheckFrequencyMinutes +
    'm from ' + config.chelseaCheckStartTime + ' to ' + config.chelseaCheckEndTime + ' ET' +
    ' (fires the court-time reminder email once read or given up).');
}

function onConfigEdit(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== TABS.config) return;
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (col === 2 && (row === 58 || row === 59)) {
    updateDispatchTrigger();
    Logger.log('Config changed — dispatch trigger updated.');
  }
  if (col === 2 && row >= 43 && row <= 47) {
    updatePreMatchDayTriggers();
    Logger.log('Pre-match schedule time changed — triggers updated.');
  }
  if ((col === 2 || col === 3 || col === 4) && row >= 51 && row <= 55) {
    updateMatchDayMinus2Triggers();
    Logger.log('Match day -2 schedule changed — triggers updated.');
  }
  if (col === 2 && row >= 66 && row <= 69) {
    updateChelseaCheckTrigger();
    Logger.log('Chelsea check schedule changed — trigger updated.');
  }
}

function updatePreMatchDayTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runPreMatchDayDispatch') {
      try { ScriptApp.deleteTrigger(t); } catch(e) {}
    }
  });
  var config = getConfig();
  var hours = (config.preMatchSchedule || []).map(function(r) { return _parseConfigHour(r.time); })
                .filter(function(h) { return h >= 0; });
  if (!hours.length) hours = [8, 11, 14, 17, 20];
  // Daily triggers — the function itself guards to Sun/Tue/Thu only.
  // Using daily instead of 3×weekly to stay under the 20-trigger limit.
  hours.forEach(function(hour) {
    ScriptApp.newTrigger('runPreMatchDayDispatch')
      .timeBased().atHour(hour).everyDays(1)
      .inTimezone('America/New_York').create();
  });
}

function updateMatchDayMinus2Triggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runMatchDayMinus2Dispatch') {
      try { ScriptApp.deleteTrigger(t); } catch(e) {}
    }
  });
  var config = getConfig();
  var hours = (config.matchDayMinus2Schedule || []).map(function(r) { return _parseConfigHour(r.time); })
                .filter(function(h) { return h >= 0; });
  if (!hours.length) hours = [8, 11, 14, 17, 20];
  // Daily triggers — the function itself guards to Sat/Mon/Wed only.
  hours.forEach(function(hour) {
    ScriptApp.newTrigger('runMatchDayMinus2Dispatch')
      .timeBased().atHour(hour).everyDays(1)
      .inTimezone('America/New_York').create();
  });
}

// Re-installs the checkChelseaCourtTimes trigger at the configured frequency
// (Config B69). Apps Script's ClockTriggerBuilder.everyMinutes() only accepts
// 1, 5, 10, 15, or 30 — an out-of-range configured value falls back to 15.
function updateChelseaCheckTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkChelseaCourtTimes') {
      try { ScriptApp.deleteTrigger(t); } catch(e) {}
    }
  });
  var config = getConfig();
  var allowed = [1, 5, 10, 15, 30];
  var freq = allowed.indexOf(config.chelseaCheckFrequencyMinutes) !== -1 ? config.chelseaCheckFrequencyMinutes : 15;
  ScriptApp.newTrigger('checkChelseaCourtTimes').timeBased().everyMinutes(freq).create();
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
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(hourET)
    .nearMinute(minET)
    .inTimezone('America/New_York')
    .create();

  Logger.log('Dispatch trigger set for Fridays at ' + timeET + ' ET.');
}

// Records that a scheduled dispatch trigger actually ran (past its own day/enabled
// check, not just that the trigger fired) — backs the "Last dispatch ran" admin display.
function _recordDispatchRun(label) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('lastDispatchRun', nowEasternISO());
    props.setProperty('lastDispatchRunLabel', label);
  } catch(e) { Logger.log('_recordDispatchRun failed: ' + e.message); }
}

function getDispatchStatus() {
  var props = PropertiesService.getScriptProperties();
  var next  = _computeNextDispatchRun();
  return {
    lastRun:      props.getProperty('lastDispatchRun') || '',
    lastRunLabel: props.getProperty('lastDispatchRunLabel') || '',
    nextRun:      next ? next.time : '',
    nextRunLabel: next ? next.label : ''
  };
}

// Scans up to 8 days ahead against all three dispatch schedules (Pre-Match Day,
// Match Day -2, Friday auto-dispatch) and returns the single earliest upcoming run.
// Apps Script's Trigger API doesn't expose a "next fire time" for time-based
// triggers, so this is computed directly from the same Config-driven schedules
// the triggers themselves use.
function _computeNextDispatchRun() {
  var config = getConfig();
  var tz     = Session.getScriptTimeZone();
  var now    = new Date();
  var best   = null;

  function consider(dowSet, hours, label) {
    hours.forEach(function(h) {
      if (isNaN(h) || h < 0) return;
      for (var add = 0; add <= 8; add++) {
        var d = new Date(now.getTime() + add * 86400000);
        var dow = parseInt(Utilities.formatDate(d, tz, 'u'));
        if (dowSet.indexOf(dow) === -1) continue;
        var dateStr   = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        // Utilities.parseDate (not the bare Date constructor) is required here: Apps
        // Script's V8 runtime resolves an offset-less date-time string against
        // America/Los_Angeles regardless of the script's configured time zone, which
        // silently shifted every candidate run time off from real Eastern time.
        var candidate = Utilities.parseDate(dateStr + ' ' + (h < 10 ? '0' + h : h) + ':00:00', tz, 'yyyy-MM-dd HH:mm:ss');
        if (candidate.getTime() <= now.getTime()) continue;
        if (!best || candidate.getTime() < best.time.getTime()) best = { time: candidate, label: label };
        break;
      }
    });
  }

  consider([7, 2, 4], (config.preMatchSchedule || []).map(function(r) { return _parseConfigHour(r.time); }), 'Pre-Match Day Dispatch'); // Sun/Tue/Thu
  consider([6, 1, 3], (config.matchDayMinus2Schedule || []).map(function(r) { return _parseConfigHour(r.time); }), 'Match Day -2 Dispatch'); // Sat/Mon/Wed
  if (config.autoDispatchEnabled) {
    var hourET = parseInt((config.autoDispatchTimeET || '13:00').split(':')[0]);
    consider([5], [hourET], 'Friday Auto-Dispatch'); // Fri
  }

  if (!best) return null;
  return { time: Utilities.formatDate(best.time, tz, "yyyy-MM-dd'T'HH:mm:ssXXX"), label: best.label };
}

// Processing order for open sub requests during Dispatch — not chronological:
// 8:00 first, then 12:30, 11:00, 9:30, and finally TBD/blank (and Overflow,
// which has no confirmed time either) last. Requests for an earlier match date
// are processed before a later one; within the same match date, ties on match
// time break by earliest submission timestamp.
var DISPATCH_TIME_PRIORITY = ['08:00', '12:30', '11:00', '09:30'];
function _sortRequestsForDispatch(requests) {
  return requests.slice().sort(function(a, b) {
    if (a.matchDate !== b.matchDate) return a.matchDate < b.matchDate ? -1 : 1;
    var ai = DISPATCH_TIME_PRIORITY.indexOf(a.matchTime);
    var bi = DISPATCH_TIME_PRIORITY.indexOf(b.matchTime);
    if (ai === -1) ai = DISPATCH_TIME_PRIORITY.length;
    if (bi === -1) bi = DISPATCH_TIME_PRIORITY.length;
    if (ai !== bi) return ai - bi;
    return a.timestamp.localeCompare(b.timestamp);
  });
}

// Maps a zero-candidate runMatch() result to a more specific DispatchLog Result
// string when the reason is known, for easier admin diagnosis. Falls back to
// the generic 'no_candidates' when nobody matched the requested date/time at
// all, or the reason was something else (no8am mismatch, already assigned via
// a different open request, etc).
function _dispatchNoCandidateResult(result) {
  if (result.matchTime === 'Overflow') return 'Overflow';
  if (result.noCandidateReason === 'alreadyScheduled') return 'Candidate already scheduled';
  if (result.noCandidateReason === 'outOfRange') return 'Candidate out of range';
  return 'no_candidates';
}

function runAutoDispatch() {
  var config = getConfig();
  if (!config.autoDispatchEnabled) {
    Logger.log('runAutoDispatch: disabled, exiting.');
    return { skipped: 'disabled' };
  }
  _recordDispatchRun('Friday Auto-Dispatch');

  // Step 1: expire all sub requests and volunteer records on or before today
  expireUpToToday();

  // Step 2: fetch open requests (after expiry, so already-expired ones are excluded)
  var requests  = getRequests();
  var open      = _sortRequestsForDispatch(requests.filter(function(r) { return r.status === 'open'; }));
  var logSheet  = getOrCreateDispatchLog();
  var reqSheet  = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
  var timestamp = nowEasternISO();

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
          groupLetter:       req.groupLetter,
          volunteerRowIndex: best.rowIndex || null,
          groupPlayers:      JSON.stringify(req.groupPlayers || [])
        });
        assignedThisRun[best.email.toLowerCase() + '|' + req.matchDate] = true;
        logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, 'matched', best.name, best.email, '']);
        Logger.log('Auto-dispatched: ' + req.name + ' → ' + best.name);
      } else {
        // No match found
        if (isLastMinute(req, config.lastMinuteThresholdHrs) && !config.urgentSubEmailsEnabled) {
          // Original last-minute behaviour: cancel immediately (only when urgent sub emails are off)
          var emailNote = 'broadcast sent — last-minute, no candidates, cancelled';
          try {
            sendSubNeededTomorrowEmail(req);
          } catch(emailErr) {
            emailNote = 'email failed (' + emailErr.message + ') — last-minute, no candidates, cancelled';
            Logger.log('sendSubNeededTomorrowEmail failed for ' + req.id + ': ' + emailErr.message);
          }
          if (reqSheet) reqSheet.getRange(req.rowIndex, 7).setValue('cancelled');
          logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, _dispatchNoCandidateResult(result), '', '', emailNote]);
          Logger.log('No candidates (last-minute, cancelled): ' + req.name + ' — ' + emailNote);
        } else {
          logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, _dispatchNoCandidateResult(result), '', '', '']);
          Logger.log('No candidates for: ' + req.name + ' (' + req.id + ')');
        }
      }
    } catch(err) {
      logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, 'error', '', '', err.message]);
      Logger.log('Auto-dispatch error for ' + req.id + ': ' + err.message);
    }
  });

  // One-shot broadcast for T+2 open requests.
  // T+1 is handled by runPreMatchDayDispatch / runPreMatchDayDispatchFinal,
  // which run 5 times on Sun/Tue/Thu via fixed weekly triggers.
  if (config.urgentSubEmailsEnabled) {
    var currentOpen = getRequests().filter(function(r) { return r.status === 'open'; });
    var dateT2 = getDateStr(2);
    var openT2 = currentOpen.filter(function(r) { return r.matchDate === dateT2; });
    if (openT2.length) {
      try { sendUrgentSubBroadcast(openT2, dateT2); }
      catch(e) { Logger.log('T2 urgent sub broadcast failed: ' + e.message); }
    }
  }

  return { dispatched: open.length };
}

// ──────────────────────────────────────────────────
// ROUTING
// ──────────────────────────────────────────────────

function doGet(e) {
  const action   = e.parameter.action;
  const callback = e.parameter.callback;

  // Volunteer from email link — returns an HTML confirmation page, not JSONP
  if (action === 'volunteerFromEmail') return handleVolunteerFromEmail(e);

  let result;

  try {
    if (action === 'getRequests')          result = getRequests();
    else if (action === 'getRequestPageData') result = getRequestPageData();
    else if (action === 'getVolunteers')   result = getVolunteers();
    else if (action === 'getPlayers')      result = getPlayers();
    else if (action === 'getHomeData')     result = getHomeData();
    else if (action === 'getConfig') {
      // Public endpoint — never leak the Brevo API key to callers.
      result = Object.assign({}, getConfig());
      delete result.brevoApiKey;
    }
    else if (action === 'submitRequest')   result = submitRequest(e.parameter);
    else if (action === 'submitVolunteer') result = submitVolunteer(e.parameter);
    else if (action === 'confirmSub')      result = confirmSub(e.parameter);
    else if (action === 'runMatch')        result = runMatch(e.parameter);
    else if (action === 'updateVolunteer')  result = updateVolunteer(e.parameter);
    else if (action === 'deleteVolunteer')  result = deleteVolunteer(e.parameter);
    else if (action === 'getDispatchLog')    result = getDispatchLog();
    else if (action === 'retireRequest')          result = retireRequest(e.parameter);
    else if (action === 'cancelRequest')          result = cancelRequest(e.parameter);
    else if (action === 'manuallyAssignSub')      result = manuallyAssignSub(e.parameter);
    else if (action === 'getAdminConfigTables')        result = getAdminConfigTables();
    else if (action === 'saveDispatchConfigTable')      result = saveDispatchConfigTable(e.parameter);
    else if (action === 'saveSettingsConfigTable')      result = saveSettingsConfigTable(e.parameter);
    else if (action === 'updateRequestTime')          result = updateRequestTime(e.parameter);
    else if (action === 'updateMatchGroupTime')       result = updateMatchGroupTime(e.parameter);
    else if (action === 'debugRunChelseaImport')      result = debugRunChelseaImport(e.parameter);
    else if (action === 'getBrevoBounceSummary')      result = getBrevoBounceSummary(e.parameter);
    else if (action === 'debugCheckSentEmail')        result = debugCheckSentEmail(e.parameter);
    else if (action === 'debugSendTestMail')          result = debugSendTestMail(e.parameter);
    else if (action === 'recalculateAnitaRatings')    result = recalculateAnitaRatings();
    else if (action === 'sendAdminCode')          result = sendAdminCode(e.parameter);
    else if (action === 'verifyAdminCode')         result = verifyAdminCode(e.parameter);
    else if (action === 'debugAdmin')              result = debugAdmin(e.parameter);
    else if (action === 'getCoordinatorRatings')   result = getCoordinatorRatings(e.parameter);
    else if (action === 'getPlayersForAdmin')       result = getPlayersForAdmin();
    else if (action === 'addPlayer')               result = addPlayer(e.parameter);
    else if (action === 'updatePlayer')            result = updatePlayer(e.parameter);
    else if (action === 'propagateEmailChange')    result = propagateEmailChange(e.parameter);
    else if (action === 'deletePlayer')            result = deletePlayer(e.parameter);
    else if (action === 'saveCoordinatorRatings')  result = saveCoordinatorRatings(e.parameter);
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
    else if (action === 'sendTestSubAlertEmail')        result = sendTestSubAlertEmail();
    else if (action === 'getDispatchStatus')            result = getDispatchStatus();
    else if (action === 'updateRequest')             result = updateRequest(e.parameter);
    else if (action === 'editRequestPlayers')         result = editRequestPlayers(e.parameter);
    else if (action === 'getMatchSlot')               result = getMatchSlot(e.parameter);
    else if (action === 'createScheduleDraft')         result = createScheduleDraft(e.parameter);
    else if (action === 'getRecentEmailLog')         result = getRecentEmailLog(e.parameter);
    else if (action === 'resendUrgentSubBroadcast')  result = resendUrgentSubBroadcast(e.parameter);
    else if (action === 'checkEmailQuotaNow')        result = { remaining: MailApp.getRemainingDailyQuota() };
    else if (action === 'sendBroadcastFallbackToAdmin') result = sendBroadcastFallbackToAdmin(e.parameter);
    else if (action === 'backfillNo8amFlags')        result = backfillNo8amFlags();
    else if (action === 'backfillGroupLetters')      result = backfillGroupLetters();
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
      } else if (req.matchTime === 'Overflow') {
        result = { req: { id: req.id, matchDate: req.matchDate, matchTime: req.matchTime, email: req.email },
                   requireAllTimes: false, skillWindow: null, trace: [], overflow: true };
      } else {
        const reqPlayer       = players.find(p => p.email === req.email.toLowerCase());
        const reqRating       = reqPlayer ? reqPlayer.rating : null;
        const matchDate       = req.matchDate;
        const matchTime       = req.matchTime;
        const hasTBDTime      = !matchTime;
        const effectiveTime   = (matchTime || '08:00').trim();
        const { phase: _phase, skillWindow } = getDispatchPhase(req, config);
        const lastMinute      = _phase === 'last-minute';
        const requireAllTimes = hasTBDTime || _phase === 'pre-schedule';
        const reqHasNo8am = [req.email].concat((req.groupPlayers || []).map(p => p.email)).filter(Boolean)
          .some(e => {
            const p = players.find(pl => pl.email.toLowerCase() === e.toLowerCase().trim());
            return !!(p && p.no8am);
          });
        const timesNeeded = (hasTBDTime && reqHasNo8am) ? TIMES.filter(t => t !== '08:00') : TIMES;
        const trace = vols.map(v => {
          const volTimes     = v.times.map(t => t.trim());
          const dateMatch    = v.date.trim() === matchDate.trim();
          const notRequestor = v.email.toLowerCase() !== req.email.toLowerCase();
          const timeMatch    = requireAllTimes
                                 ? timesNeeded.every(t => volTimes.includes(t))
                                 : volTimes.includes(effectiveTime);
          const skillOk      = (() => {
            const p = players.find(p => p.email.toLowerCase() === v.email.toLowerCase());
            return p ? Math.abs(p.rating - reqRating) <= skillWindow : false;
          })();
          const no8amOk      = (() => {
            const p = players.find(p => p.email.toLowerCase() === v.email.toLowerCase());
            return !(p && p.no8am && effectiveTime === '08:00' && !(hasTBDTime && reqHasNo8am));
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
            dateMatch, notRequestor, timeMatch, skillOk, no8amOk, notAssigned, notPlaying,
            passes: dateMatch && notRequestor && timeMatch && skillOk && no8amOk && notAssigned && notPlaying,
            playingTriggers
          };
        });
        result = {
          req: { id: req.id, matchDate, matchTime, email: req.email },
          lastMinute, requireAllTimes, reqHasNo8am,
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

// Combined read for the Request a Sub tab — one round trip instead of the page
// firing getRequests + getPublishedSchedule as separate parallel doGet calls,
// which was piling up concurrent executions on tab load/refresh.
function getRequestPageData() {
  return { requests: getRequests(), publishedSchedule: getPublishedSchedule() };
}

function getRequests() {
  const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // Column J (10) is the No8am flag (_flagNo8amOnRequestRow) — groupLetter lives in
  // column K (11) instead, alongside it rather than reshuffling existing columns.
  const rows = sheet.getRange(1, 1, lastRow, 11).getValues();
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
    groupPlayers: (function() { try { return JSON.parse(r[8] || '[]'); } catch(e) { return []; } })(),
    groupLetter:  r[10] ? r[10].toString().trim() : ''
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

// Looks up a player's current email by name from a pre-loaded players array.
// Falls back to storedEmail if the name isn't found — never returns empty when
// storedEmail has a value.
function _resolveEmail(name, storedEmail, players) {
  if (!name || !players || !players.length) return storedEmail || '';
  var lower = name.toLowerCase();
  var match = players.find(function(p) { return p.name && p.name.toLowerCase() === lower; });
  return (match && match.email) ? match.email : (storedEmail || '');
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

// ISO 8601 timestamp with the New York offset (handles EST/EDT automatically),
// so SubRequests/Volunteers rows show Eastern time instead of UTC.
function nowEasternISO() {
  return Utilities.formatDate(new Date(), 'America/New_York', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

// Flags column J ('No8am') on a SubRequests row Yes/No based on whether any of the
// given emails belongs to a No8am player. Purely a spreadsheet-visible marker for
// coordinators — matching logic recomputes this itself rather than trusting the cell,
// so it can't go stale.
function _flagNo8amOnRequestRow(reqSheet, rowNum, emails, allPlayers) {
  try {
    var players = allPlayers || getPlayersWithRatings();
    var hasNo8am = emails.filter(Boolean).some(function(e) {
      var p = players.find(function(pl) { return (pl.email || '').toLowerCase() === e.toLowerCase().trim(); });
      return !!(p && p.no8am);
    });
    var headerCell = reqSheet.getRange(1, 10);
    if (!headerCell.getValue()) headerCell.setValue('No8am');
    reqSheet.getRange(rowNum, 10).setValue(hasNo8am ? 'Yes' : 'No');
  } catch(e) {
    Logger.log('_flagNo8amOnRequestRow failed: ' + e.message);
  }
}

// Records which MatchGroups group (letter) a SubRequests row belongs to, in column
// K — so a later sub-confirmation can find the exact row to update in MatchGroups
// instead of matching on email alone, which is ambiguous if the requester happens
// to appear in more than one group on the same day.
function _setGroupLetterOnRequestRow(reqSheet, rowNum, groupLetter) {
  try {
    var headerCell = reqSheet.getRange(1, 11);
    if (!headerCell.getValue()) headerCell.setValue('Group Letter');
    reqSheet.getRange(rowNum, 11).setValue(groupLetter || '');
  } catch(e) {
    Logger.log('_setGroupLetterOnRequestRow failed: ' + e.message);
  }
}

// One-off backfill: applies the No8am flag (column J) to every currently open
// SubRequests row. New requests get it automatically going forward via
// _flagNo8amOnRequestRow (submitRequest, publishScheduleSlot's Anita Sub creation).
function backfillNo8amFlags() {
  var reqSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
  if (!reqSheet || reqSheet.getLastRow() < 2) return { success: true, updated: 0 };
  var allPlayers = getPlayersWithRatings();
  var rows = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, 9).getValues();
  var updated = 0;
  rows.forEach(function(r, i) {
    if ((r[6] || '').toString().toLowerCase().trim() !== 'open') return;
    var groupPlayers = [];
    try { groupPlayers = JSON.parse(r[8] || '[]'); } catch(e) {}
    var emails = [r[3]].concat(groupPlayers.map(function(p) { return p.email; }));
    _flagNo8amOnRequestRow(reqSheet, i + 2, emails, allPlayers);
    updated++;
  });
  return { success: true, updated: updated };
}

// One-off backfill: looks up and records the MatchGroups group letter (column K)
// for every currently open SubRequests row that doesn't have one yet — requests
// created before that field existed. New requests get it automatically going
// forward via submitRequest / publishScheduleSlot's Anita Sub creation.
function backfillGroupLetters() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var reqSheet = ss.getSheetByName(TABS.requests);
  if (!reqSheet || reqSheet.getLastRow() < 2) return { success: true, updated: 0, notFound: 0 };
  var rows = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, 11).getValues();
  var updated  = 0;
  var notFound = 0;
  rows.forEach(function(r, i) {
    if ((r[6] || '').toString().toLowerCase().trim() !== 'open') return;
    if (r[10] && r[10].toString().trim()) return; // already has a letter

    var email = (r[3] || '').toString().toLowerCase().trim();
    var matchDate = formatSheetDate(r[4]);
    var groupPlayers = [];
    try { groupPlayers = JSON.parse(r[8] || '[]'); } catch(e) {}
    var partnerEmails = groupPlayers.map(function(p) { return (p.email || '').toLowerCase(); });
    var matchGroupRow = _findMatchGroupRow(ss, matchDate, [email].concat(partnerEmails));
    if (matchGroupRow && matchGroupRow.letter) {
      _setGroupLetterOnRequestRow(reqSheet, i + 2, matchGroupRow.letter);
      updated++;
    } else {
      notFound++;
      Logger.log('backfillGroupLetters: no MatchGroups row found for row ' + (i + 2) + ' (' + email + ', ' + matchDate + ')');
    }
  });
  return { success: true, updated: updated, notFound: notFound };
}

function submitRequest(params) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
  var groupPlayersArr = [];
  if (params.groupPlayers) {
    try { groupPlayersArr = typeof params.groupPlayers === 'string' ? JSON.parse(params.groupPlayers) : params.groupPlayers; }
    catch(e) { groupPlayersArr = []; }
  }

  // Guard against duplicate requests (double-click, slow-network retry, two tabs/devices) —
  // the frontend does a similar pre-submit check, but only on some entry points, so it's
  // enforced here too where it can't be bypassed.
  var reqEmail = (params.email || '').toString().toLowerCase();
  var reqDate  = params.matchDate ? params.matchDate.toString() : '';
  var partnerEmails = groupPlayersArr.map(function(p) { return (p.email || '').toLowerCase(); });
  var isDuplicate = getRequests().some(function(r) {
    if (r.email.toLowerCase() !== reqEmail) return false;
    if (r.matchDate !== reqDate || r.status === 'cancelled') return false;
    if (!r.groupPlayers || !r.groupPlayers.length) return true;
    return r.groupPlayers.some(function(p) { return partnerEmails.indexOf((p.email || '').toLowerCase()) !== -1; });
  });
  if (isDuplicate) {
    return { success: false, error: 'A sub request for this date already exists.' };
  }

  // Look up the MatchGroups row this request belongs to — captured on the request
  // itself so a later sub-confirmation can find the exact group/row to update
  // instead of matching on email alone, which breaks if the requester happens to
  // appear in more than one group on the same day. When the caller already knows
  // which group (e.g. the My Matches schedule table, which is keyed by date+letter),
  // it passes groupLetter explicitly and that's looked up directly; otherwise fall
  // back to the best-effort email scan.
  var explicitGroupLetter = (params.groupLetter || '').toString().trim();
  var ssForGroupLookup = SpreadsheetApp.openById(SHEET_ID);
  var matchGroupRow = explicitGroupLetter
    ? _findMatchGroupRowByLetter(ssForGroupLookup, reqDate, explicitGroupLetter)
    : (reqDate ? _findMatchGroupRow(ssForGroupLookup, reqDate, [reqEmail].concat(partnerEmails)) : null);
  var groupLetter = matchGroupRow ? matchGroupRow.letter : explicitGroupLetter;

  // Court times only ever exist within 2 days of the match, and MatchGroups is
  // authoritative in that window. A group still marked Overflow needs the player
  // to supply a real time before a request can go through — same as a blank/TBD
  // time — rather than being blocked outright; server-side, not just a hidden
  // frontend control.
  var submittedTime = params.matchTime ? params.matchTime.toString().trim() : '';
  var nearTermGroup = _isTomorrowOrDayAfterTomorrow(reqDate) ? matchGroupRow : null;
  if (nearTermGroup && nearTermGroup.time === 'Overflow' && !submittedTime) {
    return { success: false, error: 'This match is in Overflow status — please select a match time, or contact your coordinator.' };
  }

  const groupPlayers = JSON.stringify(groupPlayersArr);
  const row = [
    uid(),
    nowEasternISO(),
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
  _setGroupLetterOnRequestRow(sheet, lastRow, groupLetter);
  _flagNo8amOnRequestRow(sheet, lastRow, [params.email].concat(groupPlayersArr.map(function(p) { return p.email; })));

  // If the player supplied or changed a time within the 2-day window, reconcile it
  // back into MatchGroups (and any other open request for the group) — only when
  // it's actually new/different, not when they just accepted what was already shown.
  try {
    if (submittedTime && nearTermGroup && nearTermGroup.time !== submittedTime) {
      var setResult = _setMatchGroupTime(reqDate, nearTermGroup.letter, submittedTime,
        'Request a Sub form', params.name, params.email);
      if (setResult.success) _syncGroupTimeToOpenRequests(reqDate, setResult.emails, submittedTime);
    }
  } catch(e) {
    Logger.log('submitRequest: MatchGroups time reconcile failed: ' + e.message);
  }

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

  // If this request is for tomorrow (the Pre-Match Day target) and fewer than 3 of the
  // scheduled Pre-Match Day dispatch runs are still left today, don't wait for the next
  // one — queue an immediate broadcast resend so it doesn't sit unseen for hours.
  // Reuses _runQueuedBroadcast's one-shot-trigger pattern (same as the admin's manual
  // Run Now button) so this HTTP response isn't held up by a slow bulk email send.
  try {
    if (params.matchDate === getDateStr(1) && _remainingPreMatchRunsToday(getConfig()) < 3) {
      ScriptApp.getProjectTriggers().forEach(function(t) {
        if (t.getHandlerFunction() === '_runQueuedBroadcast') ScriptApp.deleteTrigger(t);
      });
      ScriptApp.newTrigger('_runQueuedBroadcast').timeBased().after(60000).create();
    }
  } catch(e) {
    Logger.log('Immediate pre-match broadcast check failed: ' + e.message);
  }

  return { success: true };
}

// Ensures a player never ends up with more than one non-cancelled volunteer
// record on the same date: if one already exists (and is still 'pending'),
// the given times are merged into it instead of a second row being created.
// An already-matched/expired record represents a finalized outcome, so those
// are left alone rather than reopened. Times are passed/stored as arrays of
// encoded codes (e.g. "08_00") to prevent Sheets auto-converting them.
function upsertVolunteerTimes(sheet, name, email, date, times, timestampISO) {
  const emailLower = (email || '').toLowerCase();
  const lastRow    = sheet.getLastRow();
  const existing   = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 7).getValues() : [];
  const existingIdx = existing.findIndex(r =>
    (r[3] || '').toLowerCase() === emailLower &&
    formatSheetDate(r[4]) === date &&
    (r[6] || '').toLowerCase() !== 'cancelled'
  );
  if (existingIdx !== -1) {
    const existingRow = existing[existingIdx];
    const status = (existingRow[6] || '').toLowerCase();
    if (status !== 'pending') {
      Logger.log('upsertVolunteerTimes: skipped merge for ' + emailLower + ' on ' + date + ' — existing record is ' + status);
      return { created: false, merged: false };
    }
    const existingTimes = (existingRow[5] || '').toString().split(',').map(t => t.trim()).filter(Boolean);
    const mergedTimes   = existingTimes.slice();
    times.forEach(t => { if (mergedTimes.indexOf(t) === -1) mergedTimes.push(t); });
    if (mergedTimes.length === existingTimes.length) return { created: false, merged: false }; // nothing new to add
    const rowNum = existingIdx + 2; // +1 for header row, +1 for 1-based sheet rows
    sheet.getRange(rowNum, 6).setNumberFormat('@').setValue(mergedTimes.join(','));
    Logger.log('upsertVolunteerTimes: merged times for ' + emailLower + ' on ' + date + ' -> ' + mergedTimes.join(','));
    return { created: false, merged: true };
  }
  const nextRow = sheet.getLastRow() + 1;
  const range   = sheet.getRange(nextRow, 1, 1, 7);
  // Set number format first to prevent auto-conversion
  range.setNumberFormats([['@','@','@','@','@','@','@']]);
  range.setValues([[
    uid(),
    timestampISO || nowEasternISO(),
    name,
    email,
    date,
    times.join(','),  // stored as 08_00,09_30 etc to prevent Sheets time conversion
    'pending'
  ]]);
  return { created: true, merged: false };
}

function submitVolunteer(params) {
  const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.volunteers);
  const entries = JSON.parse(params.entries);
  entries.forEach(entry => {
    var upsertResult = upsertVolunteerTimes(sheet, params.name, params.email, entry.date, entry.times);
    if (upsertResult.created) {
      try { _notifyLateVolunteerForTomorrow(params.name, params.email, entry.date, entry.times); }
      catch(e) { Logger.log('_notifyLateVolunteerForTomorrow failed: ' + e.message); }
    }
  });

  // Confirmation email to volunteer
  if (params.email && isEmailEnabled() && entries.length > 0) {
    var volUrl  = APP_BASE_URL + '#volunteer';
    var subject = 'MWF Tennis League — Volunteer to sub confirmed';
    var dateLines = entries.map(function(entry) {
      var times = entry.times.map(function(t) { var tc = t.replace('_', ':'); return TIME_LABELS[tc] || tc; }).join(', ');
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
      var times = entry.times.map(function(t) { var tc = t.replace('_', ':'); return TIME_LABELS[tc] || tc; }).join(', ');
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
    try { propagateEmailChange({ oldEmail: oldEmail, newEmail: email }); }
    catch(e) { Logger.log('propagateEmailChange failed: ' + e.message); }
    notifyGroupRosterChange({
      remove: [{ name: oldName, email: oldEmail }],
      add:    [{ name: name, email: email }]
    });
  }
  return { success: true };
}

// Keeps open SubRequests and pending Volunteers records pointing at a player's current
// email after it changes — as the primary requestor/volunteer, and as a groupPlayers
// reference on someone else's open request. Without this, a stale email on an open
// request silently breaks matching (runMatch can't resolve the requestor to a Player
// row), with no visible symptom beyond "no candidates."
function propagateEmailChange(params) {
  var oldEmail = (params.oldEmail || '').toString().toLowerCase().trim();
  var newEmail = (params.newEmail || '').toString().trim();
  if (!oldEmail || !newEmail) return { success: false, error: 'oldEmail and newEmail are required.' };
  if (oldEmail === newEmail.toLowerCase()) return { success: true, requestsUpdated: 0, groupRefsUpdated: 0, volunteersUpdated: 0 };

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var requestsUpdated = 0, groupRefsUpdated = 0, volunteersUpdated = 0;

  var reqSheet = ss.getSheetByName(TABS.requests);
  if (reqSheet && reqSheet.getLastRow() >= 2) {
    var reqRows = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, 9).getValues();
    reqRows.forEach(function(r, i) {
      var rowNum = i + 2;
      if ((r[6] || '').toString().toLowerCase().trim() !== 'open') return;

      if ((r[3] || '').toString().toLowerCase().trim() === oldEmail) {
        reqSheet.getRange(rowNum, 4).setValue(newEmail);
        requestsUpdated++;
      }

      var groupPlayers;
      try { groupPlayers = JSON.parse(r[8] || '[]'); } catch(e) { return; }
      if (!Array.isArray(groupPlayers) || !groupPlayers.length) return;
      var changed = false;
      var updated = groupPlayers.map(function(p) {
        if ((p.email || '').toString().toLowerCase().trim() !== oldEmail) return p;
        changed = true;
        var copy = {};
        for (var k in p) { copy[k] = p[k]; }
        copy.email = newEmail;
        return copy;
      });
      if (changed) {
        var cell = reqSheet.getRange(rowNum, 9);
        cell.setNumberFormat('@');
        cell.setValue(JSON.stringify(updated));
        groupRefsUpdated++;
      }
    });
  }

  var volSheet = ss.getSheetByName(TABS.volunteers);
  if (volSheet && volSheet.getLastRow() >= 2) {
    var volRows = volSheet.getRange(2, 1, volSheet.getLastRow() - 1, 7).getValues();
    volRows.forEach(function(r, i) {
      var rowNum = i + 2;
      var status = (r[6] || 'pending').toString().toLowerCase().trim();
      if (status !== 'pending') return;
      if ((r[3] || '').toString().toLowerCase().trim() === oldEmail) {
        volSheet.getRange(rowNum, 4).setValue(newEmail);
        volunteersUpdated++;
      }
    });
  }

  Logger.log('propagateEmailChange: ' + oldEmail + ' -> ' + newEmail +
    ' | requests=' + requestsUpdated + ' groupRefs=' + groupRefsUpdated + ' volunteers=' + volunteersUpdated);
  return { success: true, requestsUpdated: requestsUpdated, groupRefsUpdated: groupRefsUpdated, volunteersUpdated: volunteersUpdated };
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

// View Schedule's editable time cell — same 2-day rule as everywhere else: court
// times never exist more than 2 days before a match, so this rejects the write
// outright rather than just relying on the frontend not offering the control.
function updateMatchGroupTime(params) {
  var matchDate    = (params.matchDate || '').toString().trim();
  var groupLetter  = (params.groupLetter || '').toString().trim();
  var matchTime    = (params.matchTime || '').toString().trim();
  var source       = (params.source || 'View Schedule manual edit').toString().trim();
  var playerName   = (params.playerName || '').toString().trim();
  var playerEmail  = (params.playerEmail || '').toString().trim();
  if (!matchDate || !groupLetter) return { success: false, error: 'Missing matchDate or groupLetter.' };
  if (!_isTomorrowOrDayAfterTomorrow(matchDate)) {
    return { success: false, error: 'Court times can only be set within 2 days of the match.' };
  }
  var setResult = _setMatchGroupTime(matchDate, groupLetter, matchTime, source, playerName, playerEmail);
  if (!setResult.success) return { success: false, error: 'No matching MatchGroups row found.' };
  var updatedRequests = _syncGroupTimeToOpenRequests(matchDate, setResult.emails, matchTime);
  return { success: true, updatedRequests: updatedRequests };
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

  // 4. Replace requestor in groupPlayers of any other open sub requests on the same day
  updateRelatedOpenRequests(ss, params);

  // 5. Parse group players
  var groupPlayers = [];
  try { groupPlayers = JSON.parse(params.groupPlayers || '[]'); } catch(e) {}

  // 6. Send email
  sendConfirmationEmails(params, groupPlayers);

  return { success: true };
}

// Replaces the requestor's player slot in MatchGroups with the confirmed sub.
function updateScheduleForSub(ss, params) {
  var matchDate      = (params.matchDate      || '').toString().trim();
  var groupLetter    = (params.groupLetter    || '').toString().trim();
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
    // The request records which group it came from — when present, only that
    // exact row is a candidate. This is what stops the wrong group's slot from
    // getting overwritten when the requester happens to appear in more than one
    // group on the same day (matching on email+date alone picked whichever row
    // came first). Requests created before this field existed have no letter, so
    // they fall back to the old email-only match.
    if (groupLetter && (r[3] || '').toString().trim() !== groupLetter) continue;

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

// When a sub is confirmed, update the groupPlayers field of any other open sub
// requests on the same day that list the requester, swapping in the sub's identity.
function updateRelatedOpenRequests(ss, params) {
  var requestorEmail = (params.requestorEmail || '').toLowerCase().trim();
  var matchDate      = (params.matchDate      || '').toString().trim();
  var subName        = (params.subName        || '').toString().trim();
  var subEmail       = (params.subEmail       || '').toString().trim();
  if (!requestorEmail || !matchDate || !subName || !subEmail) return;

  var sheet   = ss.getSheetByName(TABS.requests);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if ((r[6] || '').toString().toLowerCase() !== 'open') continue;
    if (formatSheetDate(r[4]) !== matchDate) continue;

    var groupPlayers;
    try { groupPlayers = JSON.parse(r[8] || '[]'); } catch(e) { continue; }
    if (!Array.isArray(groupPlayers) || !groupPlayers.length) continue;

    var changed = false;
    var updated = groupPlayers.map(function(p) {
      if ((p.email || '').toLowerCase().trim() === requestorEmail) {
        changed = true;
        return { name: subName, email: subEmail };
      }
      return p;
    });

    if (changed) {
      var cell = sheet.getRange(i + 2, 9);
      cell.setNumberFormat('@');
      cell.setValue(JSON.stringify(updated));
      Logger.log('updateRelatedOpenRequests: swapped ' + requestorEmail + ' → ' + subEmail + ' in row ' + (i + 2));
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
// ──────────────────────────────────────────────────
// MATCHGROUPS TIME — shared helpers
// ──────────────────────────────────────────────────
// Court times only ever exist within 2 days of the match (Chelsea assigns them
// 2 days out; nothing writes a time any earlier than that). Column 17 (Q) holds
// it — added alongside the original 16 MatchGroups columns rather than reshuffling
// existing ones.

// Scans MatchGroups for the row on matchDate containing any of the given emails.
// Returns { letter, time, emails, rowIndex } (emails = all 4 slots in that group)
// or null if no row matches.
function _findMatchGroupRow(ss, matchDate, emails) {
  var sheet = ss.getSheetByName(TABS.matchGroups);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var lookFor = {};
  (emails || []).forEach(function(e) { if (e) lookFor[e.toLowerCase().trim()] = true; });
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var rowDate = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (r[2] ? r[2].toString() : '');
    if (rowDate !== matchDate) continue;
    var groupEmails = [];
    var groupPlayers = [];
    var isMatch = false;
    for (var pi = 0; pi < 4; pi++) {
      var nm = (r[4 + pi * 2] || '').toString().trim();
      var em = (r[5 + pi * 2] || '').toString().trim();
      if (em) {
        groupEmails.push(em);
        groupPlayers.push({ name: nm, email: em });
        if (lookFor[em.toLowerCase()]) isMatch = true;
      }
    }
    if (isMatch) {
      return { letter: r[3] ? r[3].toString() : '', time: (r[16] || '').toString().trim(), emails: groupEmails, players: groupPlayers, rowIndex: i + 2 };
    }
  }
  return null;
}

// Looks up a MatchGroups row by date + group letter directly, rather than by
// scanning for an email — used when the caller already knows exactly which group
// a request belongs to, since an email scan is ambiguous when the same player
// appears in more than one group on the same day. Returns the same shape as
// _findMatchGroupRow (letter, time, emails, rowIndex) or null.
function _findMatchGroupRowByLetter(ss, matchDate, groupLetter) {
  var sheet = ss.getSheetByName(TABS.matchGroups);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var rowDate = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (r[2] ? r[2].toString() : '');
    if (rowDate !== matchDate) continue;
    var letter = r[3] ? r[3].toString().trim() : '';
    if (letter !== groupLetter) continue;
    var emails = [];
    for (var pi = 0; pi < 4; pi++) {
      var em = (r[5 + pi * 2] || '').toString().trim();
      if (em) emails.push(em);
    }
    return { letter: letter, time: (r[16] || '').toString().trim(), emails: emails, rowIndex: i + 2 };
  }
  return null;
}

// Writes a MatchGroups row's time by date + group letter (not by email — used by
// the Chelsea import and the View Schedule dropdown, which both already know the
// letter). Returns { success, emails } so callers can cascade to SubRequests
// without a second scan.
// source/playerName/playerEmail are optional — pass them so the MatchTimeLog
// audit row (appended below, only when the value actually changes) records
// who/what made the change and why, since this cell otherwise leaves no trace
// of who touched it.
function _setMatchGroupTime(matchDate, groupLetter, timeValue, source, playerName, playerEmail) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.matchGroups);
  if (!sheet || sheet.getLastRow() < 2) return { success: false, emails: [] };
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var rowDate = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (r[2] ? r[2].toString() : '');
    var letter = r[3] ? r[3].toString() : '';
    if (rowDate === matchDate && letter === groupLetter) {
      var oldTime = (r[16] || '').toString().trim();
      sheet.getRange(i + 2, 17).setNumberFormat('@').setValue(timeValue);
      var emails = [];
      for (var pi = 0; pi < 4; pi++) {
        var em = (r[5 + pi * 2] || '').toString().trim();
        if (em) emails.push(em);
      }
      if (oldTime !== timeValue) {
        try {
          getOrCreateMatchTimeLog().appendRow([
            nowEasternISO(), matchDate, groupLetter, oldTime, timeValue,
            source || 'unknown', playerName || '', playerEmail || ''
          ]);
        } catch(e) {
          Logger.log('_setMatchGroupTime: MatchTimeLog append failed: ' + e.message);
        }
      }
      return { success: true, emails: emails };
    }
  }
  return { success: false, emails: [] };
}

function getOrCreateMatchTimeLog() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('MatchTimeLog');
  if (!sheet) {
    sheet = ss.insertSheet('MatchTimeLog');
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Timestamp', 'MatchDate', 'GroupLetter', 'OldTime', 'NewTime', 'Source', 'PlayerName', 'PlayerEmail'
    ]]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  return sheet;
}

// Pushes a group's time onto every still-open SubRequests row for that date whose
// requester is one of the group's players.
function _syncGroupTimeToOpenRequests(matchDate, groupEmails, timeValue) {
  var reqSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.requests);
  if (!reqSheet || reqSheet.getLastRow() < 2) return 0;
  var emailSet = {};
  (groupEmails || []).forEach(function(e) { if (e) emailSet[e.toLowerCase().trim()] = true; });
  var rows = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, 9).getValues();
  var updated = 0;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var status  = (r[6] || '').toString();
    var reqDate = formatSheetDate(r[4]);
    var reqEmail = (r[3] || '').toString().toLowerCase().trim();
    if (status === 'open' && reqDate === matchDate && emailSet[reqEmail]) {
      var cell = reqSheet.getRange(i + 2, 6);
      cell.setNumberFormat('@');
      cell.setValue(timeValue);
      updated++;
    }
  }
  return updated;
}

function getMatchSlot(params) {
  var playerEmail = (params.playerEmail || '').toLowerCase().trim();
  var matchDate   = (params.matchDate   || '').toString().trim();
  if (!playerEmail || !matchDate) return { found: false };

  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.matchGroups);
  if (!sheet || sheet.getLastRow() < 2) return { found: false };

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
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

      var matchTime = '';
      var overflow  = false;

      // Court times only ever exist within 2 days of the match. Within that
      // window, MatchGroups (populated by the Chelsea import or a manual View
      // Schedule edit) is authoritative.
      if (_isTomorrowOrDayAfterTomorrow(matchDate)) {
        var mgTime = (r[16] || '').toString().trim();
        if (mgTime === 'Overflow') {
          overflow = true;
        } else if (mgTime) {
          matchTime = mgTime;
        }

        // Fall back to the old cross-request lookup only if MatchGroups doesn't
        // have a time yet (e.g. Chelsea hasn't sent it yet this cycle).
        if (!matchTime && !overflow) {
          var groupEmails = players.map(function(p) { return p.email; });
          var reqSheet = ss.getSheetByName(TABS.requests);
          if (reqSheet && reqSheet.getLastRow() >= 2) {
            var reqRows = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, 6).getValues();
            for (var j = 0; j < reqRows.length; j++) {
              var rr = reqRows[j];
              var reqDate  = formatSheetDate(rr[4]);
              var reqTime  = (rr[5] ? rr[5].toString().trim() : '');
              var reqEmail = (rr[3] || '').toString().toLowerCase().trim();
              if (reqDate === matchDate && reqTime && reqTime !== 'Overflow' && groupEmails.indexOf(reqEmail) !== -1) {
                matchTime = reqTime;
                break;
              }
            }
          }
        }
      }

      return { found: true, matchTime: matchTime, overflow: overflow, players: players, letter: r[3] ? r[3].toString().trim() : '' };
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
    var csvB64   = Utilities.base64Encode('﻿' + csvContent);
    var sent = 0, sendErrors = [];
    sd.playerEmails.forEach(function(email) {
      var name      = sd.playerNameMap[email] || '';
      var recipient = { email: email, name: name };
      try {
        sendBrevoEmail({
          apiKey:      config.brevoApiKey,
          recipients:  [recipient],
          subject:     subject,
          htmlContent: buildScheduleHtml(sd.dateMap, sd.sortedDates, sd.monthLabel, scheduleUrl, name),
          textContent: buildScheduleTextBody(sd.dateMap, sd.sortedDates, sd.monthLabel, scheduleUrl, name),
          attachments: [{ content: csvB64, name: csvFileName }]
        });
        sent++;
      } catch(e) {
        Logger.log('Brevo send failed for ' + email + ': ' + e.message);
        sendErrors.push(email);
      }
    });
    return { success: sent > 0, month: sd.monthLabel, emailsSent: sent,
             errors: sendErrors.length ? sendErrors : undefined };
  }

  // ── Send via MailApp (BCC-chunked through sendLeagueEmail) ───────────
  var csvBlob    = Utilities.newBlob('﻿' + csvContent, 'text/csv', csvFileName);
  var adminEmail = 'marobria@gmail.com';
  try {
    sendLeagueEmail({
      to:          adminEmail,
      bcc:         _excludeFromBcc(sd.playerEmails, adminEmail).join(','),
      subject:     subject,
      body:        '',
      htmlBody:    htmlBody,
      attachments: [csvBlob],
      name:        'MWF Tennis League'
    });
    return { success: true, month: sd.monthLabel, emailsSent: sd.playerEmails.length };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function buildScheduleHtml(dateMap, sortedDates, monthLabel, scheduleUrl, recipientName) {
  var DAYS   = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  var MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  var firstName   = recipientName ? recipientName.split(' ')[0] : '';
  var greetingRow = firstName
    ? '<tr><td style="padding-bottom:12px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">Hi ' + firstName + ',</td></tr>'
    : '';

  var dateRows = '';
  sortedDates.forEach(function(date) {
    var entry  = dateMap[date];
    var dp     = date.split('-');
    var d      = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]));
    var label  = DAYS[d.getDay()] + ', ' + MONTHS[d.getMonth()] + ' ' + parseInt(dp[2]);
    var groupLines = '';
    Object.keys(entry.groups).sort().forEach(function(letter) {
      var players = entry.groups[letter].map(function(p) {
        return p.name + (p.isCaptain ? ' <strong>(C)</strong>' : '');
      }).join(', ');
      groupLines += '<tr><td style="padding:2px 0 2px 14px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">Group ' + letter + ': ' + players + '</td></tr>';
    });
    if (entry.sitOut  && entry.sitOut.name)  groupLines += '<tr><td style="padding:2px 0 2px 14px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#8A4F0B;">Alternate: ' + entry.sitOut.name  + '</td></tr>';
    if (entry.sitOut2 && entry.sitOut2.name) groupLines += '<tr><td style="padding:2px 0 2px 14px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#8A4F0B;">Alternate: ' + entry.sitOut2.name + '</td></tr>';
    dateRows +=
      '<tr><td style="padding:10px 0 4px 0;border-top:1px solid #e5e7eb;">' +
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">' +
          '<tr><td style="font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:700;color:#111111;padding-bottom:4px;">' + label + '</td></tr>' +
          groupLines +
        '</table>' +
      '</td></tr>';
  });

  var viewLinkRow = scheduleUrl
    ? '<tr><td style="padding-bottom:16px;font-family:Arial,Helvetica,sans-serif;font-size:14px;"><a href="' + scheduleUrl + '" style="color:#1a5c3a;text-decoration:underline;">View Schedule Online</a></td></tr>'
    : '';

  return '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">' +
    '<html xmlns="http://www.w3.org/1999/xhtml"><head>' +
    '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0" />' +
    '<title>MWF Tennis League Schedule</title></head>' +
    '<body style="margin:0;padding:0;background-color:#f9fafb;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f9fafb;">' +
    '<tr><td style="padding:20px 12px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" style="max-width:600px;width:100%;background-color:#ffffff;border:1px solid #e5e7eb;border-radius:6px;">' +
    '<tr><td style="padding:24px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">' +
    greetingRow +
    '<tr><td style="padding-bottom:12px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">The MWF Tennis League schedule for <strong>' + monthLabel + '</strong> has been published.</td></tr>' +
    viewLinkRow +
    dateRows +
    '<tr><td style="padding-top:16px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;">Court times will be announced separately as each date approaches.</td></tr>' +
    '</table></td></tr>' +
    '<tr><td style="padding:12px 24px;font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#9ca3af;background-color:#f9fafb;border-top:1px solid #e5e7eb;border-radius:0 0 6px 6px;">' +
    'MWF Tennis League &bull; You are receiving this email as a registered player in the league.</td></tr>' +
    '</table></td></tr></table>' +
    '</body></html>';
}

function buildScheduleTextBody(dateMap, sortedDates, monthLabel, scheduleUrl, recipientName) {
  var MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  var DAYS   = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  var textLines = [];
  var firstName = recipientName ? recipientName.split(' ')[0] : '';

  textLines.push('MWF Tennis League — ' + monthLabel + ' Schedule');
  textLines.push('');
  if (firstName) textLines.push('Hi ' + firstName + ',');
  textLines.push('');
  textLines.push('The schedule for ' + monthLabel + ' has been published.');
  textLines.push('');

  sortedDates.forEach(function(date) {
    var entry = dateMap[date];
    var dp = date.split('-');
    var d = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]));
    var dateLabel = DAYS[d.getDay()] + ', ' + MONTHS[d.getMonth()] + ' ' + parseInt(dp[2]);
    textLines.push(dateLabel.toUpperCase());
    Object.keys(entry.groups).sort().forEach(function(letter) {
      var players = entry.groups[letter].map(function(p) {
        return p.name + (p.isCaptain ? ' (C)' : '');
      }).join(', ');
      textLines.push('  Group ' + letter + ': ' + players);
    });
    if (entry.sitOut && entry.sitOut.name) {
      textLines.push('  Alternate: ' + entry.sitOut.name);
    }
    if (entry.sitOut2 && entry.sitOut2.name) {
      textLines.push('  Alternate: ' + entry.sitOut2.name);
    }
    textLines.push('');
  });

  if (scheduleUrl) {
    textLines.push('View the schedule online: ' + scheduleUrl);
    textLines.push('');
  }
  textLines.push('Court times will be announced separately as each date approaches.');
  return textLines.join('\n');
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

// Maps every player scheduled to play (per MatchGroups) on matchDate to their own
// match time ('' if unknown/TBD, or 'Overflow'). One full-sheet scan shared across
// every volunteer being checked in runMatch, instead of a per-volunteer scan.
// First match wins if a player somehow appears on more than one row for the same
// date (a sheet data-entry slip) — consistent with _findMatchGroupRow, which also
// returns the first row it finds, so both stay in agreement about which group's
// time is authoritative for that player.
function _getPlayerMatchTimesForDate(ss, matchDate) {
  var sheet = ss.getSheetByName(TABS.matchGroups);
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  rows.forEach(function(r) {
    var rowDate = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (r[2] ? r[2].toString() : '');
    if (rowDate !== matchDate) return;
    var time = (r[16] || '').toString().trim();
    for (var pi = 0; pi < 4; pi++) {
      var email = (r[5 + pi * 2] || '').toString().toLowerCase().trim();
      if (email && !Object.prototype.hasOwnProperty.call(map, email)) map[email] = time;
    }
  });
  return map;
}

// A volunteer turns out to already be scheduled to play at the exact time they
// volunteered for — removes just that time from their record, or cancels the
// whole record if it was their only time. Runs as a side effect of runMatch
// discovering the conflict, so the record stops looking (wrongly) available.
function _removeConflictingVolunteerTime(ss, v, conflictingTime) {
  var volSheet = ss.getSheetByName(TABS.volunteers);
  if (!volSheet) return;
  var remaining = v.times.filter(function(t) { return t !== conflictingTime; });
  if (remaining.length) {
    var encoded = remaining.map(function(t) { return t.replace(':', '_'); }).join(',');
    volSheet.getRange(v.rowIndex, 6).setNumberFormat('@').setValue(encoded);
    Logger.log('runMatch: removed ' + conflictingTime + ' from ' + v.email + '\'s volunteer record for ' +
      v.date + ' — already scheduled to play then');
  } else {
    volSheet.getRange(v.rowIndex, 7).setValue('cancelled');
    Logger.log('runMatch: cancelled ' + v.email + '\'s volunteer record for ' + v.date +
      ' — its only time (' + conflictingTime + ') conflicted with a match they\'re already scheduled to play');
  }
}

function runMatch(params) {
  const config     = getConfig();
  const requests   = getRequests();
  const volunteers = getVolunteers();
  const players    = getPlayersWithRatings();
  const ss         = SpreadsheetApp.openById(SHEET_ID);
  // Only ever true on a Pre-Match Day Dispatch run whose configured row has
  // "Expand Volunteers" checked — see the own-match-conflict rules below.
  const expandVolunteers = params.expandVolunteers === true || params.expandVolunteers === 'true';

  const req = requests.find(r => r.id === params.requestId);
  if (!req) return { error: 'Request not found' };
  // Overflow requests have no real court time (group hasn't heard from Chelsea yet) —
  // never assign a volunteer to one.
  if (req.matchTime === 'Overflow') {
    return { candidates: [], skillWindow: null, requireAllTimes: false, phase: null, matchTime: 'Overflow' };
  }

  const reqPlayer = players.find(p => p.email === req.email.toLowerCase());
  if (!reqPlayer) return { error: 'Requestor not found in Players sheet' };

  const reqRating    = reqPlayer.rating;
  const matchDate    = req.matchDate;
  const matchTime    = req.matchTime;
  const hasTBDTime      = !matchTime;
  const effectiveTime   = (matchTime || '08:00').trim();
  const { phase, skillWindow } = getDispatchPhase(req, config);
  const lastMinute      = phase === 'last-minute';
  const requireAllTimes = hasTBDTime || phase === 'pre-schedule';
  // If the request's own group already includes a No8am player, this match was never
  // going to be scheduled at 8am in the first place.
  const reqHasNo8am = [req.email].concat((req.groupPlayers || []).map(p => p.email)).filter(Boolean)
    .some(e => {
      const p = players.find(pl => pl.email.toLowerCase() === e.toLowerCase().trim());
      return !!(p && p.no8am);
    });
  // ...so for a TBD request from that group, 8am was never a realistic outcome — nobody
  // needs to cover it to be a valid candidate, No8am or not.
  const timesNeeded = (hasTBDTime && reqHasNo8am) ? TIMES.filter(t => t !== '08:00') : TIMES;

  // Every player already scheduled to play (MatchGroups) on matchDate, mapped to
  // their own match time — computed once and reused for every volunteer below.
  const playingElsewhere = _getPlayerMatchTimesForDate(ss, matchDate);

  // Set when a volunteer otherwise available for this exact date+time gets
  // excluded specifically for that reason — lets the caller log a more useful
  // DispatchLog Result than a generic "no candidates" when that's the whole
  // story (checked in this priority order: already-scheduled, then out-of-range).
  let anyOutOfRange = false;
  let anyAlreadyScheduled = false;

  let candidates = volunteers.filter(v => {
    if (v.date.trim() !== matchDate.trim()) return false;
    if (v.email.toLowerCase() === req.email.toLowerCase()) return false;
    if (v.status === 'matched' || v.status === 'cancelled' || v.status === 'expired') return false;
    const volTimes = v.times.map(t => t.trim());
    if (requireAllTimes) {
      if (!timesNeeded.every(t => volTimes.includes(t))) return false;
    } else {
      if (!volTimes.includes(effectiveTime)) return false;
    }
    // Look up player record for rating and no8am flag
    const vol = players.find(p => p.email.toLowerCase() === v.email.toLowerCase());
    if (!vol) return false;
    if (Math.abs(vol.rating - reqRating) > skillWindow) { anyOutOfRange = true; return false; }
    // No8am volunteers must never be matched to a confirmed 8am slot. A TBD request
    // defaults to effectiveTime '08:00' too (it could turn out to be 8am) — unless the
    // request's own group already includes a No8am player, which rules that out.
    if (vol && vol.no8am && effectiveTime === '08:00' && !(hasTBDTime && reqHasNo8am)) return false;
    const alreadyAssigned = requests.some(r =>
      r.assignedSub && r.assignedSub.toLowerCase() === v.email.toLowerCase() &&
      r.matchDate === matchDate && r.status === 'filled' &&
      (!matchTime || !r.matchTime || r.matchTime === matchTime)
    );
    if (alreadyAssigned) return false;
    const alreadyPlayingRequest = requests.some(r =>
      r.email.toLowerCase() === v.email.toLowerCase() &&
      r.matchDate === matchDate && r.status !== 'open' &&
      (!matchTime || !r.matchTime || r.matchTime === matchTime)
    );
    if (alreadyPlayingRequest) return false;

    // The volunteer may already be scheduled to play a real match this same day
    // per MatchGroups — a separate thing from anything in SubRequests above.
    const emailLower = v.email.toLowerCase();
    if (Object.prototype.hasOwnProperty.call(playingElsewhere, emailLower)) {
      const theirTime = playingElsewhere[emailLower];

      // Exact same day+time as this request is a literal double-booking — never
      // assign, and correct the volunteer record since it's wrong about them
      // being free then.
      if (matchTime && theirTime && theirTime === matchTime) {
        _removeConflictingVolunteerTime(ss, v, matchTime);
        anyAlreadyScheduled = true;
        return false;
      }

      // Own match is Overflow — no confirmed court time yet, so it isn't a real
      // conflict and nothing further applies.
      if (theirTime !== 'Overflow') {
        // Scheduled at a different (or still-unknown) time this same day.
        // Whether they can be assigned depends on their own open sub request
        // (if any) and this dispatch run's Expand Volunteers setting.
        const ownOpenRequest = requests.find(r =>
          r.email.toLowerCase() === emailLower &&
          r.matchDate === matchDate &&
          r.status === 'open'
        );
        const ownOpenNonEightAm = ownOpenRequest && TIMES.includes(ownOpenRequest.matchTime) && ownOpenRequest.matchTime !== '08:00';
        if (ownOpenNonEightAm) {
          // Own open request at 9:30, 11:00, or 12:30 — only an 8:00 assignment
          // is allowed, regardless of Expand Volunteers.
          if (matchTime !== '08:00') { anyAlreadyScheduled = true; return false; }
        } else if (!expandVolunteers) {
          // Own open request at 8am, or no open request at all — only allow
          // when this dispatch run has Expand Volunteers turned on.
          anyAlreadyScheduled = true;
          return false;
        }
      }
    }
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

  // Within 48 hours: sort by closest rating (no skill restriction, so minimize variation).
  // Beyond 48 hours: FIFO — earliest submission first, rating as tiebreaker.
  var useRatingSort = phase === 'last-minute' || phase === 'urgent';

  candidates.sort((a, b) => {
    if (useRatingSort) {
      if (a.ratingDiff !== b.ratingDiff) return a.ratingDiff - b.ratingDiff;
      return a.timestamp.localeCompare(b.timestamp);
    }
    if (a.timestamp !== b.timestamp) return a.timestamp.localeCompare(b.timestamp);
    return a.ratingDiff - b.ratingDiff;
  });

  // Only meaningful when candidates is empty — tells the caller why, for a more
  // specific DispatchLog Result than a generic "no candidates" when the reason
  // is known. Priority: already-scheduled beats out-of-range if both occurred.
  var noCandidateReason = 'none';
  if (!candidates.length) {
    if (anyAlreadyScheduled) noCandidateReason = 'alreadyScheduled';
    else if (anyOutOfRange) noCandidateReason = 'outOfRange';
  }

  return {
    candidates: candidates.slice(0, 5),
    skillWindow: skillWindow,
    requireAllTimes,
    phase,
    noCandidateReason,
    matchTime: matchTime ? TIME_LABELS[matchTime] : null
  };
}

// ──────────────────────────────────────────────────
// EMAIL
// ──────────────────────────────────────────────────

// isReminder=true is the Sub Reminder email (runSubReminder) — its content is untouched
// below. Otherwise this splits into two variants based on how far out the match is:
//   Future Substitute Confirm (>2 days out): adds the Chelsea "Confirm #" instruction line.
//   Urgent Substitute Confirm (<=2 days out): CCs MTC contacts if any are set, and swaps
//   the Chelsea instruction line for a manual-update prompt.
function sendConfirmationEmails(data, groupPlayers, subjectPrefix, isReminder) {
  groupPlayers = groupPlayers || [];
  const players    = getPlayers();
  const dateStr    = formatDate(data.matchDate);
  const timeStr    = data.matchTime ? TIME_LABELS[data.matchTime] : 'TBD';
  const senderName = 'MWF Tennis League';

  // To: requestor + sub   CC: group partners — always resolve against current Players sheet
  const resolvedRequestorEmail = _resolveEmail(data.requestorName, data.requestorEmail, players);
  const resolvedSubEmail       = _resolveEmail(data.subName,       data.subEmail,       players);
  const toAddresses = [resolvedRequestorEmail, resolvedSubEmail].filter(Boolean).join(', ');
  const groupCcList = groupPlayers.map(function(p) { return _resolveEmail(p.name, p.email, players); }).filter(Boolean);
  // CC anyone who volunteered for this date/time slot when the match is tomorrow or the day after,
  // so near-term volunteers see it's already filled.
  var volunteerCcList = [];
  if (_isTomorrowOrDayAfterTomorrow(data.matchDate)) {
    volunteerCcList = _getVolunteerCcEmailsForMatch(data.matchDate, data.matchTime, players);
  }

  var chelseaLine     = 'Make updates in Chelsea as required.';
  var chelseaLineHtml = 'Make updates in <a href="https://midlothian.chelseareservations.com/login.aspx">Chelsea</a> as required.';
  var mtcCcList     = [];
  var extraLine     = null;
  var extraLineHtml = null;

  if (!isReminder) {
    if (_daysUntilMatch(data.matchDate) <= 2) {
      var config = getConfig();
      mtcCcList = [config.mtcEmail1, config.mtcEmail2].filter(Boolean);
      if (mtcCcList.length) {
        chelseaLine     = 'MTC Admin: please update Chelsea per the information above';
        chelseaLineHtml = chelseaLine;
      } else {
        chelseaLine     = 'Call MTC to change the player name in Chelsea';
        chelseaLineHtml = chelseaLine;
      }
    } else {
      extraLine     = "Chelsea > Request > Edit A Request > Click on 'Confirm #'";
      extraLineHtml = "Chelsea &gt; Request &gt; Edit A Request &gt; Click on 'Confirm #'";
    }
  }

  const ccList = groupCcList.concat(volunteerCcList).concat(mtcCcList).filter(function(email, index, arr) {
    return email && arr.map(function(item) { return String(item).toLowerCase(); }).indexOf(String(email).toLowerCase()) === index;
  });
  const ccAddresses = ccList.join(', ');

  const subject =
    (subjectPrefix || '') + 'MWF Tennis League — Substitute confirmed: ' + data.subName + ' for ' + data.requestorName;

  const greetingText = isReminder ? 'Reminder' : 'Hi team,';
  const greetingHtml = isReminder ? '<span style="color:#c0392b;">Reminder</span>' : 'Hi team,';

  const body =
    greetingText + '\n\n' +
    data.subName + ' will be substituting for ' + data.requestorName +
    ' on ' + dateStr + ' at ' + timeStr + '.\n\n' +
    chelseaLine + (extraLine ? '\n' + extraLine : '') + '\n\n' +
    'See you on the court!\n\n' +
    'MWF Tennis League';

  const htmlBody =
    greetingHtml + '<br><br>' +
    data.subName + ' will be substituting for ' + data.requestorName +
    ' on ' + dateStr + ' at ' + timeStr + '.<br><br>' +
    chelseaLineHtml + (extraLineHtml ? '<br>' + extraLineHtml : '') + '<br><br>' +
    'See you on the court!<br><br>' +
    'MWF Tennis League';

  var emailParams = {
    to:       toAddresses,
    subject:  subject,
    body:     body,
    htmlBody: htmlBody,
    name:     senderName
  };
  if (ccAddresses) emailParams.cc = ccAddresses;

  if (isEmailEnabled()) sendLeagueEmail(emailParams);
}

// Runs daily at 3:00 AM ET (see setupSubReminderTrigger). Only does anything on
// Sun/Tue/Thu — the night before a Mon/Wed/Fri match day — when it re-sends the
// dispatch confirmation email for every filled request on the next match date.
function runSubReminder() {
  var tz  = Session.getScriptTimeZone();
  var now = new Date();
  var dow = parseInt(Utilities.formatDate(now, tz, 'u')); // 1=Mon … 6=Sat, 7=Sun
  if (dow !== 7 && dow !== 2 && dow !== 4) return { skipped: 'not a reminder day' };

  var tomorrowStr = Utilities.formatDate(new Date(now.getTime() + 24 * 60 * 60 * 1000), tz, 'yyyy-MM-dd');
  var requests = getRequests();
  var players  = getPlayers();
  var sent = 0;

  requests.forEach(function(req) {
    if (req.status !== 'filled') return;
    if (req.matchDate !== tomorrowStr) return;
    if (!req.assignedSub) return;

    var subPlayer = players.find(function(p) { return p.email && p.email.toLowerCase() === req.assignedSub.toLowerCase(); });
    var data = {
      requestorName:  req.name,
      requestorEmail: req.email,
      subName:        subPlayer ? subPlayer.name : req.assignedSub,
      subEmail:       req.assignedSub,
      matchDate:      req.matchDate,
      matchTime:      req.matchTime
    };
    sendConfirmationEmails(data, req.groupPlayers, 'Sub Reminder: ', true);
    sent++;
  });

  Logger.log('runSubReminder: sent ' + sent + ' reminder(s) for ' + tomorrowStr);
  return { success: true, sent: sent, matchDate: tomorrowStr };
}

// One-time setup — run manually from the Apps Script editor to install the daily
// 3:00 AM ET trigger. Safe to re-run; clears any existing runSubReminder trigger first.
function setupSubReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runSubReminder') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runSubReminder')
    .timeBased().atHour(3).nearMinute(0).everyDays(1)
    .inTimezone('America/New_York').create();
}

// Fires from checkChelseaCourtTimes() — once a Chelsea court sheet is read for the
// day, or once the check window closes without one arriving.
function runMatchTimeReminder() {
  var now       = new Date();
  var FOUR_HOURS = 4 * 60 * 60 * 1000;

  // Track last-sent time per request ID so we resend at most once per 4 hours,
  // in case column E is Yes on more than one run in the same day.
  var props = PropertiesService.getScriptProperties();
  var log = {};
  try { log = JSON.parse(props.getProperty('matchTimeReminderLog') || '{}'); } catch(e) {}

  var requests  = getRequests();
  var players   = getPlayers();
  var siteUrl   = APP_BASE_URL + '#request';
  var notified  = 0;
  var activeIds = {};

  requests.forEach(function(req) {
    if (req.status !== 'open') return;
    if (req.matchTime) return;

    // Only remind for matches 2 or fewer days out (use 8:00 AM for TBD times)
    var matchDT = new Date(req.matchDate + 'T08:00:00');
    var diffHrs = (matchDT - now) / 36e5;
    if (diffHrs <= 0 || diffHrs > 48) return;

    activeIds[req.id] = true;

    // Skip if a reminder was already sent within the last 4 hours
    if (log[req.id] && (now.getTime() - log[req.id]) < FOUR_HOURS) return;

    var groupPlayers = req.groupPlayers || [];
    var isAnitaSub = /^anita\.sub\d+@xgmail\.com$/i.test(req.email || '');
    var requesterName = isAnitaSub ? 'Your group' : req.name;

    // Build recipient list: requester + all group members — always resolve against current Players sheet
    var allEmails = [];
    if (!isAnitaSub && req.name) allEmails.push(_resolveEmail(req.name, req.email, players));
    groupPlayers.forEach(function(p) { if (p.name || p.email) allEmails.push(_resolveEmail(p.name, p.email, players)); });
    var seen = {};
    allEmails = allEmails.filter(function(e) {
      var k = e.toLowerCase(); if (seen[k]) return false; seen[k] = true; return true;
    });
    if (!allEmails.length) return;

    var dateStr = formatDate(req.matchDate);
    var subject = 'MWF Tennis League — Court time needed for your sub request: ' + dateStr;

    var body =
      'Hi team,\n\n' +
      requesterName + ' has an open sub request for ' + dateStr + ' and no court time has been assigned yet.\n\n' +
      'UPDATE THE COURT TIME NOW, on the Request a Sub page:\n' + siteUrl + '\n\n' +
      'If you are on Overflow, do nothing. Rally will still try to find a sub.\n\n' +
      'Note: Non 8am players are ineligible to fill a sub request without a court time assigned.\n\n' +
      'MWF Tennis League';

    var htmlBody =
      'Hi team,<br><br>' +
      requesterName + ' has an open sub request for <strong>' + dateStr + '</strong> and no court time has been assigned yet.<br><br>' +
      '<strong>UPDATE THE COURT TIME NOW</strong>, on the <a href="' + siteUrl + '">Request a Sub</a> page.<br><br>' +
      '<em>If you are on Overflow, do nothing. Rally will still try to find a sub.</em><br><br>' +
      '<em>Note: Non 8am players are ineligible to fill a sub request without a court time assigned.</em><br><br>' +
      'MWF Tennis League';

    if (isEmailEnabled()) sendLeagueEmail({ to: allEmails.join(', '), subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
    log[req.id] = now.getTime();
    notified++;
  });

  // Prune log entries for requests that are no longer open/TBD
  Object.keys(log).forEach(function(id) { if (!activeIds[id]) delete log[id]; });
  props.setProperty('matchTimeReminderLog', JSON.stringify(log));

  Logger.log('runMatchTimeReminder: notified ' + notified + ' requestor(s).');
  return { success: true, notified: notified };
}

// ──────────────────────────────────────────────────
// ADMIN CONFIG TABLES — Dispatch screen + Settings screen
// Lets admins edit Config-sheet values from the web app instead of the Sheet directly.
// ──────────────────────────────────────────────────

function getAdminConfigTables() {
  var config = getConfig();
  var sched  = getSchedulerSettings();
  return {
    autoDispatchEnabled:      config.autoDispatchEnabled,
    autoDispatchTimeET:       config.autoDispatchTimeET,
    calendarLookaheadDays:    config.calendarLookaheadDays,
    allowPlayerNameChangeOnDelete: config.allowPlayerNameChangeOnDelete,
    preScheduleThresholdHrs:  config.preScheduleThresholdHrs,
    skillWindowFarOut:        config.skillWindowFarOut,
    urgentThresholdHrs:       config.urgentThresholdHrs,
    skillWindowMid:           config.skillWindowMid,
    lastMinuteThresholdHrs:   config.lastMinuteThresholdHrs,
    skillWindowUrgent:        config.skillWindowUrgent,
    skillWindowLastMinute:    config.skillWindowLastMinute,
    preMatchSchedule:         config.preMatchSchedule,
    matchDayMinus2Schedule:   config.matchDayMinus2Schedule,
    weightTeamVariance:       sched.weightTeamVariance,
    weightGroupVariance:      sched.weightGroupVariance,
    weightSocialVariety:      sched.weightSocialVariety,
    weightRecency:            sched.weightRecency,
    solverIterations:         sched.solverIterations,
    solverRestarts:           sched.solverRestarts,
    ratingRangeLimit:         sched.ratingRangeLimit,
    weightMaxRatingRange:     sched.weightMaxRatingRange,
    emailEnabled:             isEmailEnabled(),
    senderEmail:              config.senderEmail,
    brevoApiKey:              config.brevoApiKey,
    brevoScheduleEmail:       config.brevoScheduleEmail,
    urgentSubEmailsEnabled:   config.urgentSubEmailsEnabled,
    mtcEmail1:                config.mtcEmail1,
    mtcEmail2:                config.mtcEmail2,
    chelseaCheckDays:             config.chelseaCheckDays,
    chelseaCheckStartTime:        config.chelseaCheckStartTime,
    chelseaCheckEndTime:          config.chelseaCheckEndTime,
    chelseaCheckFrequencyMinutes: config.chelseaCheckFrequencyMinutes,
    chelseaCheckSubject:          config.chelseaCheckSubject,
    chelseaImportEnabled:         config.chelseaImportEnabled
  };
}

// Saves every value on the Dispatch screen's config table, then re-installs the
// triggers that depend on the schedule/time values so changes take effect immediately.
function saveDispatchConfigTable(params) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);

  var enabled = params.autoDispatchEnabled === 'true' || params.autoDispatchEnabled === true;
  var time    = (params.autoDispatchTimeET || '13:00').trim();
  sheet.getRange('B58').setValue(enabled);
  var timeCell = sheet.getRange('B59');
  timeCell.setNumberFormat('@');
  timeCell.setValue(time);

  sheet.getRange('B10').setValue(parseInt(params.calendarLookaheadDays) || 30);

  var allowNameChange = params.allowPlayerNameChangeOnDelete === 'true' || params.allowPlayerNameChangeOnDelete === true;
  sheet.getRange('B61').setValue(allowNameChange ? 'Yes' : 'No');

  sheet.getRange('B63').setValue((params.mtcEmail1 || '').toString().trim());
  sheet.getRange('B64').setValue((params.mtcEmail2 || '').toString().trim());

  sheet.getRange('B4').setValue(parseInt(params.preScheduleThresholdHrs)   || 72);
  sheet.getRange('C4').setValue(parseFloat(params.skillWindowFarOut)       || 0.5);
  sheet.getRange('B5').setValue(parseInt(params.urgentThresholdHrs)       || 48);
  sheet.getRange('C5').setValue(parseFloat(params.skillWindowMid)         || 1.0);
  sheet.getRange('B6').setValue(parseInt(params.lastMinuteThresholdHrs)   || 24);
  sheet.getRange('C6').setValue(parseFloat(params.skillWindowUrgent)      || 2.0);
  sheet.getRange('C7').setValue(parseFloat(params.skillWindowLastMinute)  || 2.8);

  var preMatch = JSON.parse(params.preMatchSchedule || '[]');
  if (preMatch.length === 5) {
    sheet.getRange('A43:F47').setValues(preMatch.map(function(r, i) {
      return [String(i + 1), r.time || '', r.dispatch ? 'Yes' : 'No', r.broadcast ? 'Yes' : 'No',
              r.cancel ? 'Yes' : 'No', r.expandVolunteers ? 'Yes' : 'No'];
    }));
  }

  var md2 = JSON.parse(params.matchDayMinus2Schedule || '[]');
  if (md2.length === 5) {
    sheet.getRange('A51:F55').setValues(md2.map(function(r, i) {
      return [String(i + 1), r.time || '', r.dispatch ? 'Yes' : 'No', r.broadcast ? 'Yes' : 'No',
              '', r.overflowDetect ? 'Yes' : 'No'];
    }));
  }

  SpreadsheetApp.flush();
  _configCache = null;

  try { updateDispatchTrigger(enabled, time); } catch(e) { Logger.log('updateDispatchTrigger error: ' + e.message); }
  try { updatePreMatchDayTriggers(); } catch(e) { Logger.log('updatePreMatchDayTriggers error: ' + e.message); }
  try { updateMatchDayMinus2Triggers(); } catch(e) { Logger.log('updateMatchDayMinus2Triggers error: ' + e.message); }

  return { success: true };
}

// Saves every value on the Settings screen's config table.
// Availability window dates/active flag are managed on the Scheduler screen, not here.
function saveSettingsConfigTable(params) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.config);

  sheet.getRange('B20').setValue(parseFloat(params.weightTeamVariance)  || 0);
  sheet.getRange('B21').setValue(parseFloat(params.weightGroupVariance) || 0);
  sheet.getRange('B22').setValue(parseFloat(params.weightSocialVariety) || 0);
  sheet.getRange('B23').setValue(parseFloat(params.weightRecency)       || 0);
  sheet.getRange('B24').setValue(parseInt(params.solverIterations)      || 800);
  sheet.getRange('B25').setValue(parseInt(params.solverRestarts)        || 10);

  sheet.getRange('B31').setValue(parseFloat(params.ratingRangeLimit)     || 2.0);
  sheet.getRange('B32').setValue(parseFloat(params.weightMaxRatingRange) || 0.0);

  sheet.getRange('B27').setValue(params.emailEnabled === 'true' || params.emailEnabled === true);
  sheet.getRange('B30').setValue((params.senderEmail || '').toString().trim());

  sheet.getRange('B35').setValue((params.brevoApiKey || '').toString().trim());
  sheet.getRange('B37').setValue((params.brevoScheduleEmail    === 'true' || params.brevoScheduleEmail    === true) ? 'Yes' : 'No');
  sheet.getRange('B39').setValue((params.urgentSubEmailsEnabled === 'true' || params.urgentSubEmailsEnabled === true) ? 'Yes' : 'No');

  sheet.getRange('B66').setValue((params.chelseaCheckDays || 'Sat,Mon,Wed').toString().trim());
  var chelseaStartCell = sheet.getRange('B67');
  chelseaStartCell.setNumberFormat('@');
  chelseaStartCell.setValue((params.chelseaCheckStartTime || '07:45').toString().trim());
  var chelseaEndCell = sheet.getRange('B68');
  chelseaEndCell.setNumberFormat('@');
  chelseaEndCell.setValue((params.chelseaCheckEndTime || '09:30').toString().trim());
  sheet.getRange('B69').setValue(parseInt(params.chelseaCheckFrequencyMinutes) || 15);
  sheet.getRange('B70').setValue((params.chelseaCheckSubject || 'Upcoming Court Sheet').toString().trim());
  sheet.getRange('B71').setValue((params.chelseaImportEnabled === 'true' || params.chelseaImportEnabled === true) ? 'Yes' : 'No');

  SpreadsheetApp.flush();
  _configCache = null;

  try { updateChelseaCheckTrigger(); } catch(e) { Logger.log('updateChelseaCheckTrigger error: ' + e.message); }

  return { success: true };
}

function sendRetirementEmail(req) {
  var players      = getPlayers();
  var toEmail      = _resolveEmail(req.name, req.email, players);
  var dateStr      = formatDate(req.matchDate);
  var timeStr      = req.matchTime ? TIME_LABELS[req.matchTime] : 'TBD';
  var subject      = 'MWF Tennis League — Unable to find substitute: ' + dateStr + (req.matchTime ? ' at ' + timeStr : '');
  var directoryUrl = APP_BASE_URL + '#directory-emailall';
  var body =
    'Hi ' + req.name + ',\n\n' +
    'Unfortunately, we were unable to find a volunteer to fill the sub request for your match:\n\n' +
    '  Date: ' + dateStr + '\n' +
    '  Time: ' + timeStr + '\n\n' +
    'If you\'d like to launch an email to the entire group, visit the Directory page: ' + directoryUrl + '\n\n' +
    'MWF Tennis League';
  var htmlBody =
    'Hi ' + req.name + ',<br><br>' +
    'Unfortunately, we were unable to find a volunteer to fill the sub request for your match:<br><br>' +
    '&nbsp;&nbsp;Date: ' + dateStr + '<br>' +
    '&nbsp;&nbsp;Time: ' + timeStr + '<br><br>' +
    'Click on <a href="' + directoryUrl + '">Directory</a>, if you\'d like to launch an email to the entire group.<br><br>' +
    'MWF Tennis League';
  var groupPlayers = req.groupPlayers || [];
  var ccList = groupPlayers.map(function(p) { return _resolveEmail(p.name, p.email, players); }).filter(Boolean);
  var emailParams = { to: toEmail, subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' };
  if (ccList.length) emailParams.cc = ccList.join(', ');
  if (isEmailEnabled()) sendLeagueEmail(emailParams);
}

function sendSubNeededTomorrowEmail(req) {
  if (!isEmailEnabled()) return;

  var players      = getPlayers();
  var isAnitaSub   = /^anita\.sub\d+@xgmail\.com$/i.test(req.email || '');
  var groupPlayers = req.groupPlayers || [];

  var toEmail, greetingName, ccPlayers;
  if (isAnitaSub) {
    var captain  = groupPlayers[0] || {};
    toEmail      = _resolveEmail(captain.name, captain.email, players);
    greetingName = captain.name  || 'Captain';
    ccPlayers    = groupPlayers.slice(1);
  } else {
    toEmail      = _resolveEmail(req.name, req.email, players);
    greetingName = req.name  || 'A player';
    ccPlayers    = groupPlayers;
  }
  if (!toEmail) return;

  var dateStr = formatDate(req.matchDate);
  var timeStr = req.matchTime ? (TIME_LABELS[req.matchTime] || req.matchTime) : 'TBD';

  var subject = 'MWF Tennis League — Unable to find substitute: ' + dateStr + (req.matchTime ? ' at ' + timeStr : '');
  var directoryUrl = APP_BASE_URL + '#directory-emailall';
  var body =
    'Hi ' + greetingName + ',\n\n' +
    'Unfortunately, we were unable to find a volunteer to fill the sub request for your match:\n\n' +
    '  Date: ' + dateStr + '\n' +
    '  Time: ' + timeStr + '\n\n' +
    'If you\'d like to launch an email to the entire group, visit the Directory page: ' + directoryUrl + '\n\n' +
    'MWF Tennis League';
  var htmlBody =
    'Hi ' + greetingName + ',<br><br>' +
    'Unfortunately, we were unable to find a volunteer to fill the sub request for your match:<br><br>' +
    '&nbsp;&nbsp;Date: ' + dateStr + '<br>' +
    '&nbsp;&nbsp;Time: ' + timeStr + '<br><br>' +
    'Click on <a href="' + directoryUrl + '">Directory</a>, if you\'d like to launch an email to the entire group.<br><br>' +
    'MWF Tennis League';

  var ccList = ccPlayers.map(function(p) { return _resolveEmail(p.name, p.email, players); }).filter(function(e) {
    return e && !/^anita\.sub\d+@xgmail\.com$/i.test(e);
  });
  var emailParams = { to: toEmail, subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' };
  if (ccList.length) emailParams.cc = ccList.join(', ');
  sendLeagueEmail(emailParams);
}

// ──────────────────────────────────────────────────
// CHELSEA COURT-TIME IMPORT
// ──────────────────────────────────────────────────
// Chelsea (via MTC) emails a "Court Sheet" to mwfmtctennis@gmail.com with real
// court times, always exactly 2 days before the match day it covers — as a PDF
// (OCR'd) or a native .xlsx spreadsheet (read directly), whichever MTC sends that
// day. This reads that email, extracts the table, and writes matching group times
// into MatchGroups.

var CHELSEA_TIMES = ['08:00', '09:30', '11:00', '12:30']; // must match frontend TIMES

// Trigger handler — installed at the configured frequency (Config B69, see
// updateChelseaCheckTrigger), self-guards to the configured days/window (Config
// B66–B68) so it only actually does anything in that range.
//
// Also fires the "no court time" reminder email (runMatchTimeReminder) — right
// after a Chelsea PDF is successfully read for the day, or, if none ever arrives,
// once after the check window closes ("Rally stops trying"). This replaces the old
// manually-scheduled Match Day -2 "Time Reminder" flag.
function checkChelseaCourtTimes() {
  var config = getConfig();
  if (!config.chelseaImportEnabled) return;
  var tz  = Session.getScriptTimeZone();
  var now = new Date();
  var dow = parseInt(Utilities.formatDate(now, tz, 'u')); // 1=Mon...7=Sun
  var allowedDows = _parseConfigDows(config.chelseaCheckDays);
  if (allowedDows.length && allowedDows.indexOf(dow) === -1) return;

  var props    = PropertiesService.getScriptProperties();
  var todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  if (props.getProperty('chelseaProcessedDate') === todayStr) return; // already read + reminded today

  var minutesOfDay = parseInt(Utilities.formatDate(now, tz, 'H')) * 60 + parseInt(Utilities.formatDate(now, tz, 'm'));
  var startMin = _parseConfigMinutesOfDay(config.chelseaCheckStartTime);
  var endMin   = _parseConfigMinutesOfDay(config.chelseaCheckEndTime);
  if (startMin < 0) startMin = 7 * 60 + 45;
  if (endMin < 0)   endMin   = 9 * 60 + 30;
  if (minutesOfDay < startMin) return;

  var reminderSentToday = props.getProperty('chelseaReminderSentDate') === todayStr;

  if (minutesOfDay > endMin) {
    // Window already closed today — Rally gave up finding the email. Send the
    // reminder once, if it hasn't gone out yet, then stop; nothing left to check.
    if (!reminderSentToday) {
      try { runMatchTimeReminder(); } catch(e) { Logger.log('runMatchTimeReminder failed: ' + e.message); }
      props.setProperty('chelseaReminderSentDate', todayStr);
    }
    return;
  }

  var isLastRun = (minutesOfDay + config.chelseaCheckFrequencyMinutes) > endMin;

  _runChelseaImport();

  var nowProcessed = props.getProperty('chelseaProcessedDate') === todayStr;
  if ((nowProcessed || isLastRun) && !reminderSentToday) {
    try { runMatchTimeReminder(); } catch(e) { Logger.log('runMatchTimeReminder failed: ' + e.message); }
    props.setProperty('chelseaReminderSentDate', todayStr);
  }
}

// Converts "Sat,Mon,Wed" style config text into ISO weekday numbers (1=Mon...7=Sun).
function _parseConfigDows(str) {
  var map = { SUN: 7, MON: 1, TUE: 2, WED: 3, THU: 4, FRI: 5, SAT: 6 };
  return (str || '').split(',')
    .map(function(s) { return map[s.trim().slice(0, 3).toUpperCase()]; })
    .filter(function(n) { return !!n; });
}

// Parses a 24h "HH:MM" config string into minutes since midnight, or -1 if invalid.
function _parseConfigMinutesOfDay(timeStr) {
  var m = (timeStr || '').toString().trim().match(/^(\d{1,2}):(\d{2})$/);
  return m ? (parseInt(m[1]) * 60 + parseInt(m[2])) : -1;
}

// Core import — split from the trigger guard above so it can also be run on demand
// (manual test / admin retry via debugRunChelseaImport) without waiting for the
// schedule window.
function _runChelseaImport(opts) {
  opts = opts || {};
  var tz = Session.getScriptTimeZone();
  var config = getConfig();
  var targetDate = opts.overrideTargetDate || getDateStr(2); // Chelsea always assigns times exactly 2 days out
  var result = { targetDate: targetDate, applied: [], skipped: [], dateMismatch: false, error: '' };
  var props = PropertiesService.getScriptProperties();

  try {
    var subject = (config.chelseaCheckSubject || 'Upcoming Court Sheet').replace(/"/g, '\\"');
    var todayForQuery = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd'); // restrict to today's mail only — old emails left in the inbox shouldn't match
    var threads = GmailApp.search('to:mwfmtctennis@gmail.com subject:"' + subject + '" has:attachment after:' + todayForQuery);
    if (!threads.length) { result.error = 'no matching email found'; Logger.log('_runChelseaImport: ' + result.error); return result; }

    // Most recent matching message with a PDF or Excel court-sheet attachment.
    // MTC has sent this report both as a PDF (OCR'd below) and as a native .xlsx
    // spreadsheet (read directly — far more reliable, no OCR needed). Whichever
    // is attached is used; the PDF wins if a message somehow has both.
    var pdfBlob = null, xlsxBlob = null;
    for (var ti = threads.length - 1; ti >= 0 && !pdfBlob; ti--) {
      var msgs = threads[ti].getMessages();
      for (var mi = msgs.length - 1; mi >= 0 && !pdfBlob; mi--) {
        var atts = msgs[mi].getAttachments();
        for (var ai = 0; ai < atts.length; ai++) {
          if (atts[ai].getContentType() === 'application/pdf') { pdfBlob = atts[ai]; break; }
          if (!xlsxBlob && _isChelseaXlsxAttachment(atts[ai])) xlsxBlob = atts[ai];
        }
      }
    }
    if (!pdfBlob && !xlsxBlob) { result.error = 'no PDF or Excel attachment found'; Logger.log('_runChelseaImport: ' + result.error); return result; }

    var text, grid = null;
    if (pdfBlob) {
      text = _extractPdfText(pdfBlob);
    } else {
      grid = _extractXlsxGrid(xlsxBlob);
      text = grid.map(function(row) { return row.join(' '); }).join('\n');
    }
    result.extractedChars = text.length;
    if (opts.dumpText) { result.text = text; return result; } // debug only — inspect raw extraction, nothing else

    // Found and read a real attachment — done for today regardless of what's inside it.
    // A stuck/wrong file won't resolve itself by retrying every 15 minutes.
    // (Dry runs don't count — they shouldn't suppress the real scheduled check.)
    if (!opts.dryRun) props.setProperty('chelseaProcessedDate', Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'));

    if (!opts.skipDateCheck && text.indexOf(_chelseaLongDate(targetDate)) === -1) {
      result.dateMismatch = true;
      Logger.log('_runChelseaImport: attachment date text did not match expected target date ' + targetDate + ' — aborting.');
      return result;
    }

    var parsed = grid ? _parseChelseaRowsFromGrid(grid) : _parseChelseaRows(text);
    if (parsed.error) {
      result.error = parsed.error;
      Logger.log('_runChelseaImport: ' + parsed.error);
      return result;
    }
    var rows    = parsed.rows;
    var players = getPlayers();
    var ss      = SpreadsheetApp.openById(SHEET_ID);

    rows.forEach(function(row) {
      var matched = _matchChelseaRowPlayers(row.players, players);
      if (matched.length === 0) return; // no real players in this row — ignore silently
      if (matched.length === 1) {
        result.skipped.push({ time: row.time, reason: '1 match', names: matched[0].name });
        Logger.log('_runChelseaImport: 1-match skip — ' + matched[0].name + ' at ' + row.time);
        return;
      }
      var matchedEmails = matched.map(function(p) { return p.email; });
      var group = _findMatchGroupRow(ss, targetDate, matchedEmails);
      if (!group) {
        result.skipped.push({ time: row.time, reason: 'no matching group', names: matched.map(function(p){return p.name;}).join(', ') });
        Logger.log('_runChelseaImport: 2+ matches but no MatchGroups row found for ' + targetDate + ' — ' + matched.map(function(p){return p.name;}).join(', '));
        return;
      }
      if (opts.dryRun) {
        result.applied.push({ group: group.letter, time: row.time, names: matched.map(function(p){return p.name;}).join(', '), dryRun: true });
        return;
      }
      var setResult = _setMatchGroupTime(targetDate, group.letter, row.time, 'Chelsea import');
      if (setResult.success) {
        _syncGroupTimeToOpenRequests(targetDate, setResult.emails, row.time);
        result.applied.push({ group: group.letter, time: row.time, names: matched.map(function(p){return p.name;}).join(', ') });
      }
    });

    Logger.log('_runChelseaImport: applied ' + result.applied.length + ', skipped ' + result.skipped.length + ' for ' + targetDate);
  } catch (e) {
    result.error = e.message;
    Logger.log('_runChelseaImport error: ' + e.message);
  }
  return result;
}

// Converts a PDF blob to text via the Advanced Drive Service (Apps Script has no
// built-in PDF text extraction) — uploads as a temp Google Doc with OCR conversion,
// reads it, then deletes the temp file.
function _extractPdfText(pdfBlob) {
  // resource.mimeType must describe the SOURCE file being uploaded (the PDF) —
  // not the OCR conversion target. Declaring it as GOOGLE_DOCS here (an earlier
  // mistake) told Drive the file was already a Doc, so it refused to OCR it.
  var resource = { title: 'chelsea-court-sheet-temp', mimeType: pdfBlob.getContentType() };
  var file = Drive.Files.insert(resource, pdfBlob, { ocr: true, ocrLanguage: 'en' });
  try {
    return DocumentApp.openById(file.id).getBody().getText();
  } finally {
    try { Drive.Files.remove(file.id); } catch (e) { Logger.log('_extractPdfText cleanup failed: ' + e.message); }
  }
}

// True if a Gmail attachment looks like the Excel "Booked Short Court Listing"
// export MTC sometimes sends instead of the PDF (e.g. "ReportTNShortBookL.xlsx").
// Checked by filename as well as content type since Gmail doesn't always report
// the spreadsheet MIME type consistently.
function _isChelseaXlsxAttachment(att) {
  var name = (att.getName() || '').toLowerCase();
  var type = (att.getContentType() || '').toLowerCase();
  return /\.xlsx?$/.test(name) || type.indexOf('spreadsheetml') !== -1 || type === 'application/vnd.ms-excel';
}

// Converts an .xlsx blob to a 2D array of cell values via the Advanced Drive
// Service — uploads as a temp Google Sheet, reads its grid, then deletes the
// temp file. Far more reliable than the PDF's OCR path since this report's
// Excel export is already a clean structured table, not scanned text.
function _extractXlsxGrid(xlsxBlob) {
  var resource = { title: 'chelsea-court-sheet-temp', mimeType: MimeType.GOOGLE_SHEETS };
  var file = Drive.Files.insert(resource, xlsxBlob, { convert: true });
  try {
    var sheet = SpreadsheetApp.openById(file.id).getSheets()[0];
    return sheet.getDataRange().getValues().map(function(row) {
      return row.map(function(cell) { return (cell === null || cell === undefined) ? '' : String(cell); });
    });
  } finally {
    try { Drive.Files.remove(file.id); } catch (e) { Logger.log('_extractXlsxGrid cleanup failed: ' + e.message); }
  }
}

// Builds the "Wednesday - August 12, 2026" style string Chelsea's PDF states its
// date as, to sanity-check the email actually covers the expected target date.
function _chelseaLongDate(dateStr) {
  var d = new Date(dateStr + 'T12:00:00'); // noon anchor avoids timezone edge cases
  var weekday  = d.toLocaleDateString('en-US', { weekday: 'long' });
  var monthDay = d.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
  return weekday + ' - ' + monthDay + ', ' + d.getFullYear();
}

function _escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

// Placeholder slot text ("999999 * Unavailable" / "999999 * Instructional") is never
// a real player — filtered out wherever a name is extracted, regardless of which
// report layout produced the surrounding text.
var CHELSEA_PLACEHOLDER_NAMES = { UNAVAILABLE: true, INSTRUCTIONAL: true };

// Extracts player names from one row's chunk of text. MTC has sent this report in two
// different cell layouts so far — "<ID> <NAME>" (e.g. "19150 MARC THORESON", the
// regular "Short Court Listing") and "<NAME>- #<ID>" (e.g. "ROGER YAUCHZY- #16810",
// the "Booked Short Court Listing" variant) — so both patterns are tried on every
// chunk rather than assuming one layout. They don't collide: real content only ever
// satisfies one of the two, though the ID-then-name pattern can pick up a harmless
// trailing-hyphen fragment of the *next* cell when parsing name-then-ID text (e.g.
// matching "JON BRANNON-" out of "...#16810 JON BRANNON- #14400...") — left in
// rather than suppressed, since it never equals a real player's exact name and so
// never causes a false match.
function _extractChelseaNames(chunk) {
  var names = [];
  var m;

  var idThenName = /\d{4,6}\s+([A-Za-z][A-Za-z .'\-]+)/g;
  while ((m = idThenName.exec(chunk)) !== null) {
    var n1 = m[1].trim();
    if (n1 && !CHELSEA_PLACEHOLDER_NAMES[n1.toUpperCase()]) names.push(n1);
  }

  var nameThenId = /([A-Za-z][A-Za-z .'\-]*?)-\s*#\d{4,6}/g;
  while ((m = nameThenId.exec(chunk)) !== null) {
    var n2 = m[1].trim();
    if (n2 && !CHELSEA_PLACEHOLDER_NAMES[n2.toUpperCase()]) names.push(n2);
  }

  return names;
}

// Extracts { time, players: [name,...] } rows from the extracted PDF text, keeping
// only rows at 8:00/9:30/11:00/12:30. Report format (as of 2026-08): "Epic-Padel"
// court listing, one row per "<Facility>\n<Court#>\n<slot1>\n<slot2>\n<slot3>\n<slot4>",
// where each slot is blank, a placeholder (see above), or a real player name+ID pair.
//
// Google's OCR conversion does NOT reliably keep a row's Time value adjacent to it,
// so rows and times are extracted independently and zipped by position — but only
// within each page (split on the repeating "Time Facility Court Player 1-4" header),
// not across the whole document. A page's counts not lining up (seen in practice: the
// last page, once down to a single remaining time slot, sometimes runs
// "<time> <Facility> <Court>" inline instead of one-per-line) only costs that page
// instead of aborting the whole import — better to skip one page's rows than
// misassign a time to the wrong group, but no reason a single page's OCR quirk should
// block every other page's valid data too.
function _parseChelseaRows(text) {
  // Strip the "M/D/YYYY H:MM:SS AM/PM" print-timestamp — its H:MM:SS otherwise
  // produces a spurious time-fragment match (e.g. "31:10 PM" out of "12:31:10 PM").
  var body = text.replace(/\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*(AM|PM)/gi, '');

  var facilityMatch = body.match(/Facility\s+([A-Za-z][A-Za-z ]*?)\s*(?:\n|Time\s+Facility)/);
  var facilityName = facilityMatch ? facilityMatch[1].trim() : '';
  if (!facilityName) return { rows: [], error: 'could not determine facility name from PDF header' };
  var esc = _escapeRegex(facilityName);

  var pages = body.split(/Time\s+Facility\s+Court\s+Player\s*1\s+Player\s*2\s+Player\s*3\s+Player\s*4/i).slice(1);
  if (!pages.length) return { rows: [], error: 'could not find any "Time Facility Court..." page headers in PDF text' };

  var timeRe = /(\d{1,2}):(\d{2})\s*(AM|PM)/gi;
  var rows = [];
  var mismatchedPages = 0;

  pages.forEach(function(page, pageIndex) {
    var rowRe = new RegExp(esc + '\\s*\\n\\s*(\\d{1,2})\\s*\\n([\\s\\S]*?)(?=' + esc + '\\s*\\n\\s*\\d{1,2}\\s*\\n|$)', 'g');
    var rowBlocks = [];
    var m;
    while ((m = rowRe.exec(page)) !== null) {
      rowBlocks.push(_extractChelseaNames(m[2])); // keep even empty/unavailable rows — needed to stay aligned with times below
    }

    var times = [];
    var tm;
    timeRe.lastIndex = 0;
    while ((tm = timeRe.exec(page)) !== null) {
      var h = parseInt(tm[1]);
      var ap = tm[3].toUpperCase();
      if (ap === 'PM' && h !== 12) h += 12;
      if (ap === 'AM' && h === 12) h = 0;
      times.push((h < 10 ? '0' + h : h) + ':' + tm[2]);
    }

    if (times.length !== rowBlocks.length) {
      mismatchedPages++;
      Logger.log('_parseChelseaRows: page ' + (pageIndex + 1) + ' row/time count mismatch (' +
        rowBlocks.length + ' rows, ' + times.length + ' times) — skipping this page.');
      return;
    }
    for (var i = 0; i < rowBlocks.length; i++) {
      if (CHELSEA_TIMES.indexOf(times[i]) === -1) continue; // not one of the MWF slots
      if (!rowBlocks[i].length) continue; // no real players in this slot
      rows.push({ time: times[i], players: rowBlocks[i] });
    }
  });

  if (mismatchedPages === pages.length) {
    return { rows: [], error: 'row/time count mismatch on all ' + pages.length + ' page(s) — aborting' };
  }
  return { rows: rows, error: '' };
}

// Parses { time, players: [name,...] } rows directly from the Excel "Booked Short
// Court Listing" grid — one clean row per court slot (Time, Facility, Court,
// Player 1-4 columns), unlike the PDF's OCR text. The header row is located by
// its "Time" / "Player 1" cell text rather than fixed column letters, in case
// MTC's column layout shifts.
function _parseChelseaRowsFromGrid(grid) {
  function findCol(row, label) {
    for (var c = 0; c < row.length; c++) {
      if ((row[c] || '').toString().trim().toLowerCase() === label) return c;
    }
    return -1;
  }

  var headerRowIdx = -1, timeCol = -1, playerCols = [];
  for (var r = 0; r < grid.length; r++) {
    var row = grid[r];
    var tCol = findCol(row, 'time');
    var p1Col = findCol(row, 'player 1');
    if (tCol !== -1 && p1Col !== -1) {
      headerRowIdx = r;
      timeCol = tCol;
      ['player 1', 'player 2', 'player 3', 'player 4'].forEach(function(label) {
        var c = findCol(row, label);
        if (c !== -1) playerCols.push(c);
      });
      break;
    }
  }
  if (headerRowIdx === -1) return { rows: [], error: 'could not find "Time / Player 1-4" header row in Excel sheet' };

  var rows = [];
  for (var i = headerRowIdx + 1; i < grid.length; i++) {
    var dataRow = grid[i];
    var timeVal = (dataRow[timeCol] || '').toString().trim();
    if (!timeVal) continue; // blank spacer row — report keeps going after this in practice, so skip rather than stop

    var time24 = _parseChelseaTimeCell(timeVal);
    if (!time24 || CHELSEA_TIMES.indexOf(time24) === -1) continue; // not one of the MWF slots

    var names = [];
    playerCols.forEach(function(c) {
      var n = _normalizeChelseaCellName(dataRow[c]);
      if (n) names.push(n);
    });
    if (!names.length) continue; // no real players in this slot (e.g. all "* Instructional")

    rows.push({ time: time24, players: names });
  }
  return { rows: rows, error: '' };
}

// Parses an Excel time cell like "08:00  AM" into 24h "HH:MM", or null if unparseable.
function _parseChelseaTimeCell(str) {
  var m = /^(\d{1,2}):(\d{2})\s*(AM|PM)/i.exec((str || '').toString().trim());
  if (!m) return null;
  var h  = parseInt(m[1]);
  var ap = m[3].toUpperCase();
  if (ap === 'PM' && h !== 12) h += 12;
  if (ap === 'AM' && h === 12) h = 0;
  return (h < 10 ? '0' + h : h) + ':' + m[2];
}

// Strips the "* " placeholder marker (e.g. "* Instructional") and filters out
// non-player placeholder text, using the same placeholder set the PDF parser
// checks — the Excel export marks these slots the same way, just without an ID.
function _normalizeChelseaCellName(raw) {
  var v = (raw || '').toString().trim().replace(/^\*\s*/, '');
  if (!v || CHELSEA_PLACEHOLDER_NAMES[v.toUpperCase()]) return null;
  return v;
}

// Matches PDF player-name fragments against the Players sheet, case-insensitively.
// A name matching more than one player is treated as no match for that slot —
// safer than guessing which one was meant.
function _matchChelseaRowPlayers(names, players) {
  var matched = [];
  names.forEach(function(rawName) {
    var norm = rawName.replace(/\s+/g, ' ').trim().toLowerCase();
    if (!norm) return;
    var hits = players.filter(function(p) { return (p.name || '').replace(/\s+/g, ' ').trim().toLowerCase() === norm; });
    if (hits.length === 1) matched.push(hits[0]);
  });
  return matched;
}

// Manual/on-demand run for testing — same import the schedule would run, without
// waiting for Sat/Mon/Wed 7:45-9:30am or the "already ran today" guard.
function debugRunChelseaImport(params) {
  params = params || {};
  return _runChelseaImport({
    dumpText:          params.dumpText === '1',
    skipDateCheck:     params.skipDateCheck === '1',
    overrideTargetDate: params.overrideTargetDate || '',
    dryRun:            params.dryRun === '1'
  });
}

// Debug/admin tool — pulls bounce events from Brevo's Events API and summarizes
// counts per recipient address, so repeat-offender addresses (vs. one-off/broad
// provider throttling) are easy to spot without exporting Brevo's Logs page by hand.
function getBrevoBounceSummary(params) {
  params = params || {};
  var config = getConfig();
  if (!config.brevoApiKey) return { success: false, error: 'Brevo API key not set (Config B35).' };

  var days = parseInt(params.days) || 90;
  if (days > 90) days = 90;

  function fetchEvents(eventType) {
    var url = 'https://api.brevo.com/v3/smtp/statistics/events?event=' + eventType + '&days=' + days + '&limit=2500&sort=desc';
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'api-key': config.brevoApiKey },
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code < 200 || code >= 300) {
      throw new Error('Brevo error ' + code + ': ' + response.getContentText().substring(0, 300));
    }
    return JSON.parse(response.getContentText()).events || [];
  }

  var bounceEvents, requestEvents;
  try {
    bounceEvents  = fetchEvents('bounces');
    requestEvents = fetchEvents('requests'); // one per email actually sent — denominator for failure rate
  } catch(e) {
    return { success: false, error: e.message };
  }

  var byEmail = {};
  function getRec(email) {
    if (!byEmail[email]) byEmail[email] = { email: email, hardCount: 0, softCount: 0, sentCount: 0, lastDate: '', lastEvent: '', lastReason: '' };
    return byEmail[email];
  }
  bounceEvents.forEach(function(ev) {
    var email = (ev.email || '').toLowerCase();
    if (!email) return;
    var rec = getRec(email);
    if (ev.event === 'hardBounces') rec.hardCount++;
    else if (ev.event === 'softBounces') rec.softCount++;
    if (!rec.lastDate || ev.date > rec.lastDate) {
      rec.lastDate   = ev.date   || '';
      rec.lastEvent  = ev.event  || '';
      rec.lastReason = ev.reason || '';
    }
  });
  requestEvents.forEach(function(ev) {
    var email = (ev.email || '').toLowerCase();
    if (!email) return;
    getRec(email).sentCount++;
  });

  var summary = Object.keys(byEmail).map(function(k) {
    var r = byEmail[k];
    r.totalCount   = r.hardCount + r.softCount;
    r.domain       = r.email.split('@')[1] || '';
    r.failPercent  = r.sentCount ? Math.round((r.totalCount / r.sentCount) * 1000) / 10 : null;
    return r;
  }).sort(function(a, b) { return b.totalCount - a.totalCount; });

  var domainCounts = {};
  summary.forEach(function(r) { domainCounts[r.domain] = (domainCounts[r.domain] || 0) + r.totalCount; });

  return {
    success: true,
    days: days,
    totalEvents: bounceEvents.length,
    uniqueAddresses: summary.length,
    byDomain: domainCounts,
    summary: summary
  };
}

// Debug tool — dumps To/Cc/Bcc headers for recent Sent-folder messages matching a
// Gmail search query, to directly confirm what actually went out on a MailApp send
// (Gmail's Sent-folder UI/search only surfaces the visible "To" recipient — Bcc'd
// recipients on the sender's own copy aren't independently searchable there, which
// can look like "only one recipient got it" when everyone in Bcc actually did).
function debugCheckSentEmail(params) {
  params = params || {};
  var query = params.query || 'in:sent subject:"subs needed"';
  var threads = GmailApp.search(query, 0, parseInt(params.limit) || 5);
  var messages = [];
  threads.forEach(function(t) {
    t.getMessages().forEach(function(m) {
      messages.push({
        date:    m.getDate().toISOString(),
        from:    m.getFrom(),
        to:      m.getTo(),
        cc:      m.getCc(),
        bcc:     m.getBcc(),
        subject: m.getSubject()
      });
    });
  });
  return { query: query, count: messages.length, messages: messages };
}

// One-off diagnostic — sends a known bcc list via MailApp; check delivery afterward
// (with debugCheckSentEmail, giving Gmail time to index) via query "in:anywhere
// subject:<marker>" against the bcc address, and the sent-folder Bcc header via
// "in:sent subject:<marker>".
function debugSendTestMail(params) {
  params = params || {};
  var marker = 'debugTestBcc-' + Date.now();
  var toAddr  = params.to  || 'marobria@gmail.com';
  var bccAddr = params.bcc || 'mwfmtctennis@gmail.com';
  MailApp.sendEmail(toAddr, marker, 'Diagnostic test — checking Bcc delivery.', { bcc: bccAddr, name: 'Rally Diagnostic' });
  return { marker: marker, sentTo: toAddr, sentBcc: bccAddr };
}

// Marks still-TBD sub requests for targetDate as 'Overflow' (Match Day -2 Dispatch,
// Config column F) before dispatch runs. A TBD request this close to match day likely
// means the group has no Chelsea court time yet — runMatch() skips 'Overflow' requests
// so dispatch never assigns a volunteer to one.
function markOverflowRequests(targetDate) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var reqSheet = ss.getSheetByName(TABS.requests);
  var requests = getRequests().filter(function(r) {
    return r.status === 'open' && r.matchDate === targetDate && !r.matchTime;
  });
  requests.forEach(function(req) {
    var cell = reqSheet.getRange(req.rowIndex, 6);
    cell.setNumberFormat('@');
    cell.setValue('Overflow');

    // Keep MatchGroups in sync — the request's group should show Overflow too.
    try {
      var emails = [req.email].concat((req.groupPlayers || []).map(function(p) { return p.email; }));
      var group = _findMatchGroupRow(ss, targetDate, emails);
      if (group) _setMatchGroupTime(targetDate, group.letter, 'Overflow', 'Match Day -2 auto-Overflow');
    } catch(e) {
      Logger.log('markOverflowRequests: MatchGroups sync failed for ' + req.id + ': ' + e.message);
    }

    try { sendOverflowNotification(req); } catch(e) {
      Logger.log('sendOverflowNotification failed for ' + req.id + ': ' + e.message);
    }
  });
  Logger.log('markOverflowRequests: ' + requests.length + ' request(s) marked Overflow for ' + targetDate);
  return requests.length;
}

function sendOverflowNotification(req) {
  if (!isEmailEnabled()) return;

  var players      = getPlayers();
  var isAnitaSub    = /^anita\.sub\d+@xgmail\.com$/i.test(req.email || '');
  var groupPlayers  = req.groupPlayers || [];

  var allEmails = [];
  if (!isAnitaSub && req.name) allEmails.push(_resolveEmail(req.name, req.email, players));
  groupPlayers.forEach(function(p) { if (p.name || p.email) allEmails.push(_resolveEmail(p.name, p.email, players)); });
  var seen = {};
  allEmails = allEmails.filter(function(e) {
    var k = e.toLowerCase(); if (seen[k]) return false; seen[k] = true; return true;
  });
  if (!allEmails.length) return;

  var reqUrl  = APP_BASE_URL + '#request';
  var dateStr = formatDate(req.matchDate);
  var subject = 'MWF Tennis League — Sub request assumed Overflow: ' + dateStr;
  var body =
    'This sub request cannot be filled because there is no match time, so it is assumed to be in \'Overflow\'.\n\n' +
    'Update that match time when you get a court assignment from Chelsea, on the Request a Sub page:\n' +
    reqUrl + '\n\n' +
    'MWF Tennis League';
  var htmlBody =
    'This sub request cannot be filled because there is no match time, so it is assumed to be in <strong>Overflow</strong>.<br><br>' +
    'Update that match time when you get a court assignment from Chelsea, on the <a href="' + reqUrl + '">Request a Sub</a> page.<br><br>' +
    'MWF Tennis League';

  sendLeagueEmail({ to: allEmails.join(', '), subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
}

// ──────────────────────────────────────────────────
// URGENT SUB EMAIL BROADCAST & FOLLOWUP DISPATCH
// ──────────────────────────────────────────────────

function getDateStr(daysFromToday) {
  var d = new Date();
  d.setDate(d.getDate() + daysFromToday);
  return formatSheetDate(d);
}

function getOpenRequestsForDate(targetDate) {
  return getRequests().filter(function(r) {
    return r.status === 'open' && r.matchDate === targetDate;
  });
}

// Reviews every open request, regardless of match date — not just the one date a
// given trigger run happens to target. A request 3 weeks out gets the same shot at
// being filled as one due tomorrow; getDispatchPhase()/runMatch() already size the
// skill window per-request based on hours-until-match, so this works at any distance.
// expandVolunteers is only ever true when called from a Pre-Match Day Dispatch
// run whose configured row has "Expand Volunteers" checked — every other caller
// (Match Day -2, Friday Auto-Dispatch, manual admin dispatch) omits it, which
// keeps the stricter default behavior for runMatch's own-match-conflict rules.
function runDispatchAllOpen(expandVolunteers) {
  var requests  = _sortRequestsForDispatch(getRequests().filter(function(r) { return r.status === 'open'; }));
  if (!requests.length) return 0;
  var logSheet  = getOrCreateDispatchLog();
  var timestamp = nowEasternISO();
  var assigned  = {};
  var matched   = 0;
  requests.forEach(function(req) {
    try {
      var result = runMatch({ requestId: req.id, expandVolunteers: !!expandVolunteers });
      if (result.candidates && result.candidates.length) {
        var eligible = result.candidates.filter(function(c) {
          return !assigned[c.email.toLowerCase() + '|' + req.matchDate];
        });
        if (!eligible.length) return;
        var best = eligible[0];
        confirmSub({
          requestId: req.id, requestRowIndex: req.rowIndex,
          subEmail: best.email, subName: best.name,
          requestorName: req.name, requestorEmail: req.email,
          matchDate: req.matchDate, matchTime: req.matchTime,
          groupLetter: req.groupLetter,
          volunteerRowIndex: best.rowIndex || null,
          groupPlayers: JSON.stringify(req.groupPlayers || [])
        });
        assigned[best.email.toLowerCase() + '|' + req.matchDate] = true;
        logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, 'matched', best.name, best.email, 'followup dispatch']);
        matched++;
      } else {
        logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, _dispatchNoCandidateResult(result), '', '', 'followup dispatch']);
      }
    } catch(e) {
      logSheet.appendRow([timestamp, req.id, req.name, req.matchDate, req.matchTime, 'error', '', '', 'followup: ' + e.message]);
    }
  });
  return matched;
}

// Friendly "h:mm AM/PM" formatting for the next-dispatch-run note appended to the
// subs-needed broadcast — separate from TIME_LABELS since this covers whatever
// hour is actually configured, not just the 4 canonical match time slots.
function _formatClockTime(date, tz) {
  return Utilities.formatDate(date, tz, 'h:mm a');
}

// One-sentence heads-up appended to the bottom of the subs-needed broadcast, telling
// players when Dispatch will next attempt to fill any remaining open requests.
// Returns '' if there's no scheduled run to report (e.g. every dispatch schedule
// is disabled/empty) so callers can skip the note entirely.
function _nextDispatchRunWindow() {
  var next = _computeNextDispatchRun();
  if (!next) return null;
  var tz    = Session.getScriptTimeZone();
  var start = new Date(next.time);
  var end   = new Date(start.getTime() + 60 * 60 * 1000);
  return { start: _formatClockTime(start, tz), end: _formatClockTime(end, tz) };
}

function buildSubNeededEmailHtml(requests, scriptUrl) {
  var dateStr    = formatDate(requests[0].matchDate);
  var headerRow =
    '<tr style="border-bottom:2px solid #e5e7eb;">' +
    '<th style="text-align:left;padding:6px 12px 6px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;font-weight:600;"></th>' +
    '<th style="text-align:left;padding:6px 12px 6px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;font-weight:600;">Needs Sub</th>' +
    '<th style="text-align:left;padding:6px 12px 6px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;font-weight:600;white-space:nowrap;">Time</th>' +
    '<th style="text-align:left;padding:6px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;font-weight:600;">Other Players</th>' +
    '</tr>';
  var dataRows = requests.map(function(req) {
    var timeLabel  = TIME_LABELS[req.matchTime] || req.matchTime || 'TBD';
    var otherNames = (req.groupPlayers || [])
      .filter(function(p) { return (p.email || '').toLowerCase() !== (req.email || '').toLowerCase(); })
      .map(function(p) { return p.name; }).join(', ');
    var linkUrl = scriptUrl + '?action=volunteerFromEmail&requestId=' + encodeURIComponent(req.id);
    var buttonHtml = req.matchTime === 'Overflow'
      ? '<span style="display:inline-block;padding:7px 14px;background-color:#9ca3af;color:#f3f4f6;border-radius:4px;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;white-space:nowrap;">I CAN Sub</span>'
      : '<a href="' + linkUrl + '" style="display:inline-block;padding:7px 14px;background-color:#1a5c3a;color:#ffffff;text-decoration:none;border-radius:4px;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:700;white-space:nowrap;">I CAN Sub</a>';
    return '<tr style="border-bottom:1px solid #f0f0f0;">' +
      '<td style="padding:10px 12px 10px 0;vertical-align:middle;">' +
        buttonHtml +
      '</td>' +
      '<td style="padding:10px 12px 10px 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;vertical-align:middle;">' + (req.name || '') + '</td>' +
      '<td style="padding:10px 12px 10px 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;vertical-align:middle;white-space:nowrap;">' + timeLabel + '</td>' +
      '<td style="padding:10px 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;vertical-align:middle;">' + otherNames + '</td>' +
      '</tr>';
  }).join('');

  var nextRun    = _nextDispatchRunWindow();
  var nextRunRow = nextRun
    ? '<tr><td colspan="4" style="padding-top:12px;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#111111;">' +
        'The Dispatch process will run again between <strong>' + nextRun.start + '</strong> and <strong>' + nextRun.end + '</strong>.' +
      '</td></tr>'
    : '';

  return '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">' +
    '<html xmlns="http://www.w3.org/1999/xhtml"><head>' +
    '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0" />' +
    '<title>Subs Needed</title></head>' +
    '<body style="margin:0;padding:0;background-color:#f9fafb;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f9fafb;">' +
    '<tr><td style="padding:20px 12px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" style="max-width:650px;width:100%;background-color:#ffffff;border:1px solid #e5e7eb;border-radius:6px;">' +
    '<tr><td style="padding:24px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">' +
    '<tr><td colspan="4" style="padding-bottom:16px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">' +
      'Substitutes are needed for <strong>' + dateStr + '</strong>. Click <strong>I CAN Sub</strong> if you are available to substitute for any of the players listed below.' +
    '</td></tr>' +
    '<tr><td colspan="4"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">' +
    headerRow + dataRows +
    '</table></td></tr>' +
    nextRunRow +
    '<tr><td colspan="4" style="padding-top:16px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;">Do not reply to this email.</td></tr>' +
    '</table></td></tr>' +
    '<tr><td style="padding:12px 24px;font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#9ca3af;background-color:#f9fafb;border-top:1px solid #e5e7eb;border-radius:0 0 6px 6px;">' +
    'MWF Tennis League &bull; You are receiving this email as a registered player in the league.</td></tr>' +
    '</table></td></tr></table></body></html>';
}

function buildSubNeededEmailText(requests, targetDate) {
  var dateStr = formatDate(targetDate);
  var lines   = [];
  lines.push('Substitutes are needed for ' + dateStr + '.');
  lines.push('If you can sub for any of the players below, click the "I CAN Sub" link in the HTML version of this email.');
  lines.push('');
  requests.forEach(function(req) {
    var timeLabel  = TIME_LABELS[req.matchTime] || req.matchTime || 'TBD';
    var otherNames = (req.groupPlayers || [])
      .filter(function(p) { return (p.email || '').toLowerCase() !== (req.email || '').toLowerCase(); })
      .map(function(p) { return p.name; }).join(', ');
    lines.push('Needs Sub: ' + (req.name || '') + '  |  Time: ' + timeLabel + '  |  Other Players: ' + otherNames);
  });
  var nextRun = _nextDispatchRunWindow();
  if (nextRun) {
    lines.push('');
    lines.push('The Dispatch process will run again between ' + nextRun.start + ' and ' + nextRun.end + '.');
  }
  lines.push('');
  lines.push('Do not reply to this email.');
  lines.push('');
  lines.push('MWF Tennis League');
  return lines.join('\n');
}

// Guards against the same broadcast going out twice in quick succession from two
// independent trigger paths (the regularly scheduled Pre-Match Day / Match Day -2
// dispatch runs, and the one-shot _runQueuedBroadcast trigger submitRequest() queues
// for 60s later on a late-arriving request). sendLeagueEmail()'s own throttle only
// catches byte-identical resends, so it misses this case whenever a request's field
// (e.g. matchTime getting confirmed) changes between the two sends. Keyed on target
// date + the open request IDs, not full content, so a genuinely new/changed situation
// later in the day still sends.
function _alreadyBroadcastRecently(targetDate, openRequests) {
  var props = PropertiesService.getScriptProperties();
  var sig   = targetDate + '|' + openRequests.map(function(r) { return r.id; }).sort().join(',');
  var key   = 'urgentBroadcastSig';
  var last  = props.getProperty(key);
  var lastTime = parseInt(props.getProperty(key + 'Time') || '0', 10);
  var recent   = (Date.now() - lastTime) < 60 * 60 * 1000; // 1 hour cooldown
  props.setProperty(key, sig);
  props.setProperty(key + 'Time', String(Date.now()));
  return last === sig && recent;
}

function sendUrgentSubBroadcast(openRequests, targetDate) {
  if (!openRequests.length || !isEmailEnabled()) return;
  if (_alreadyBroadcastRecently(targetDate, openRequests)) {
    Logger.log('sendUrgentSubBroadcast: skipped — same open requests for ' + targetDate + ' already broadcast within the last hour');
    return;
  }
  var config    = getConfig();
  var d         = new Date(targetDate + 'T12:00:00');
  var monthDay  = d.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
  var subject   = 'MWF Tennis, subs needed ' + monthDay;
  var scriptUrl = SCRIPT_URL;
  var players   = getPlayersWithRatings().filter(function(p) {
    return p.email && !/^anita\.sub\d+@xgmail\.com$/i.test(p.email);
  });
  if (!players.length) return;
  var adminEmail = 'marobria@gmail.com';
  var bccList   = _excludeFromBcc(players.map(function(p) { return p.email; }), adminEmail).join(',');
  sendLeagueEmail({
    to:       adminEmail,
    bcc:      bccList,
    subject:  subject,
    body:     buildSubNeededEmailText(openRequests, targetDate),
    htmlBody: buildSubNeededEmailHtml(openRequests, scriptUrl),
    name:     'MWF Tennis League'
  });
  Logger.log('Urgent sub broadcast sent via BCC to ' + players.length + ' player(s) for ' + targetDate);
}

// Volunteers for targetDate who were never used as a sub — status is either still
// 'pending' (open) or already flagged 'expired'. Excludes 'matched'/'cancelled'.
function getLeftoverVolunteersForDate(targetDate) {
  return getVolunteers().filter(function(v) {
    return v.date === targetDate && (v.status === 'pending' || v.status === 'expired');
  });
}

// Unique {name,email} of every player in a MatchGroups slot for targetDate.
// Sit-out players are excluded — they aren't playing that day.
function getPlayersScheduledForDate(targetDate) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.matchGroups);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
  var seen = {};
  var players = [];
  rows.forEach(function(r) {
    var rowDate = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (r[2] ? r[2].toString() : '');
    if (rowDate !== targetDate) return;
    for (var pi = 0; pi < 4; pi++) {
      var name  = r[4 + pi * 2]     ? r[4 + pi * 2].toString().trim()     : '';
      var email = r[4 + pi * 2 + 1] ? r[4 + pi * 2 + 1].toString().trim() : '';
      if (!name || !email) continue;
      var key = email.toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;
      players.push({ name: name, email: email });
    }
  });
  return players;
}

// Match groups (court assignments) for targetDate, sorted by group letter — for the
// "groups playing tomorrow" table on the leftover-volunteers email.
function getMatchGroupsForDate(targetDate) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.matchGroups);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  var groups = [];
  rows.forEach(function(r) {
    var rowDate = r[2] instanceof Date
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (r[2] ? r[2].toString() : '');
    if (rowDate !== targetDate) return;
    var letter = r[3] ? r[3].toString().trim() : '';
    if (!letter) return;
    var players = [];
    for (var pi = 0; pi < 4; pi++) {
      var name  = r[4 + pi * 2]     ? r[4 + pi * 2].toString().trim()     : '';
      var email = r[4 + pi * 2 + 1] ? r[4 + pi * 2 + 1].toString().trim() : '';
      if (name) players.push({ name: name, email: email, isCaptain: pi === 0 });
    }
    if (players.length) groups.push({ letter: letter, time: r[16] ? r[16].toString().trim() : '', players: players });
  });
  groups.sort(function(a, b) { return a.letter.localeCompare(b.letter); });
  return groups;
}

// Match-time label for the "groups playing" table — the leftover-volunteers email
// always shows the actual court time in the Group column when one is set, falling
// back to the plain Group letter only if it isn't known yet.
function _groupTimeLabel(g) {
  return g.time ? (TIME_LABELS[g.time] || g.time) : g.letter;
}

function buildLeftoverVolunteersEmailHtml(volunteers, groups) {
  var introText = 'No more sub requests can be filled for tomorrow. The following players are available if needed.';
  var dataRows = volunteers.length
    ? volunteers.map(function(v) {
        var times = (v.times || []).map(function(t) { return TIME_LABELS[t] || t; }).join(', ') || 'TBD';
        return '<tr style="border-bottom:1px solid #f0f0f0;">' +
          '<td style="padding:8px 12px 8px 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">' + v.name + '</td>' +
          '<td style="padding:8px 12px 8px 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">' + v.email + '</td>' +
          '<td style="padding:8px 12px 8px 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">' + (v.phone || '—') + '</td>' +
          '<td style="padding:8px 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">' + times + '</td>' +
          '</tr>';
      }).join('')
    : '<tr style="border-bottom:1px solid #f0f0f0;">' +
        '<td colspan="4" style="padding:8px 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">None</td>' +
        '</tr>';

  var groupRows = (groups || []).map(function(g) {
    var names = g.players.map(function(p) { return p.isCaptain ? '<strong>' + p.name + '</strong>' : p.name; }).join(', ');
    return '<tr style="border-bottom:1px solid #f0f0f0;">' +
      '<td style="padding:8px 12px 8px 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;font-weight:600;">' + _groupTimeLabel(g) + '</td>' +
      '<td style="padding:8px 0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">' + names + '</td>' +
      '</tr>';
  }).join('');

  var groupsSection = (groups && groups.length)
    ? '<tr><td colspan="4" style="padding-top:20px;padding-bottom:8px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;font-weight:600;">Groups playing tomorrow</td></tr>' +
      '<tr><td colspan="4"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">' +
      '<tr style="border-bottom:2px solid #e5e7eb;">' +
      '<th style="text-align:left;padding:6px 12px 6px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;font-weight:600;">Group</th>' +
      '<th style="text-align:left;padding:6px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;font-weight:600;">Players</th>' +
      '</tr>' + groupRows +
      '</table></td></tr>'
    : '';

  return '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">' +
    '<html xmlns="http://www.w3.org/1999/xhtml"><head>' +
    '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0" />' +
    '<title>Extra Subs Available</title></head>' +
    '<body style="margin:0;padding:0;background-color:#f9fafb;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f9fafb;">' +
    '<tr><td style="padding:20px 12px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" style="max-width:650px;width:100%;background-color:#ffffff;border:1px solid #e5e7eb;border-radius:6px;">' +
    '<tr><td style="padding:24px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">' +
    '<tr><td colspan="4" style="padding-bottom:16px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">' +
      introText +
    '</td></tr>' +
    '<tr><td colspan="4"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">' +
    '<tr style="border-bottom:2px solid #e5e7eb;">' +
    '<th style="text-align:left;padding:6px 12px 6px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;font-weight:600;">Name</th>' +
    '<th style="text-align:left;padding:6px 12px 6px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;font-weight:600;">Email</th>' +
    '<th style="text-align:left;padding:6px 12px 6px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;font-weight:600;">Phone</th>' +
    '<th style="text-align:left;padding:6px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;font-weight:600;">Available Times</th>' +
    '</tr>' + dataRows +
    '</table></td></tr>' +
    groupsSection +
    '<tr><td colspan="4" style="padding-top:16px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6b7280;">Do not reply to this email.</td></tr>' +
    '</table></td></tr>' +
    '<tr><td style="padding:12px 24px;font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#9ca3af;background-color:#f9fafb;border-top:1px solid #e5e7eb;border-radius:0 0 6px 6px;">' +
    'MWF Tennis League &bull; You are receiving this email as a registered player in the league.</td></tr>' +
    '</table></td></tr></table></body></html>';
}

function buildLeftoverVolunteersEmailText(volunteers, groups) {
  var lines = [];
  lines.push('No more sub requests can be filled for tomorrow. The following players are available if needed.');
  lines.push('');
  if (volunteers.length) {
    volunteers.forEach(function(v) {
      var times = (v.times || []).map(function(t) { return TIME_LABELS[t] || t; }).join(', ') || 'TBD';
      lines.push(v.name + '  |  ' + v.email + '  |  ' + (v.phone || '—') + '  |  ' + times);
    });
  } else {
    lines.push('None');
  }
  if (groups && groups.length) {
    lines.push('');
    lines.push('Groups playing tomorrow:');
    groups.forEach(function(g) {
      var names = g.players.map(function(p) { return p.name; }).join(', ');
      lines.push('  ' + _groupTimeLabel(g) + ': ' + names);
    });
  }
  lines.push('');
  lines.push('Do not reply to this email.');
  lines.push('');
  lines.push('MWF Tennis League');
  return lines.join('\n');
}

// Fires on the final Pre-Match Day run (row.cancel), win or lose — a request can
// stay unfilled even with a leftover volunteer if their rating misses the skill
// window. Always sends (even with zero leftover volunteers, shown as "None") so
// everyone scheduled to play knows who's still around if a spot opens up.
// Called right after a brand-new (not merged) volunteer record is created for
// tomorrow, once today's last chance at an automated Pre-Match Day Dispatch
// run has passed (see _hasRemainingPreMatchDayDispatchToday) — Dispatch won't
// see this volunteer today. Notifies each still-open sub request on that date
// whose match time the volunteer just offered, so the requester knows to
// reach out directly instead of waiting on a run that isn't coming today.
function _notifyLateVolunteerForTomorrow(volunteerName, volunteerEmail, date, times) {
  try {
    if (date !== getDateStr(1)) return; // only "tomorrow" — the whole point is today's dispatch window
    if (_hasRemainingPreMatchDayDispatchToday(getConfig())) return; // dispatch could still catch it today

    var volunteerEmailLower = (volunteerEmail || '').toLowerCase();
    var offeredTimes = (times || []).map(function(t) { return t.toString().replace('_', ':'); });
    var openReqs = getRequests().filter(function(r) {
      return r.status === 'open' && r.matchDate === date &&
        r.email.toLowerCase() !== volunteerEmailLower &&
        offeredTimes.indexOf(r.matchTime) !== -1;
    });
    openReqs.forEach(function(req) {
      try { _sendLateVolunteerNotification(req, volunteerName, volunteerEmail); }
      catch(e) { Logger.log('_sendLateVolunteerNotification failed for ' + req.id + ': ' + e.message); }
    });
  } catch(e) {
    Logger.log('_notifyLateVolunteerForTomorrow failed: ' + e.message);
  }
}

function _sendLateVolunteerNotification(req, volunteerName, volunteerEmail) {
  if (!isEmailEnabled()) return;
  var players = getPlayers();
  var reqEmail = _resolveEmail(req.name, req.email, players) || req.email;
  var volEmail = _resolveEmail(volunteerName, volunteerEmail, players) || volunteerEmail;
  if (!reqEmail || !volEmail) return;

  var phoneByEmail = {};
  players.forEach(function(p) { if (p.email) phoneByEmail[p.email.toLowerCase()] = p.phone || ''; });
  var reqPhone = phoneByEmail[reqEmail.toLowerCase()] || '';
  var volPhone = phoneByEmail[volEmail.toLowerCase()] || '';

  var dateStr = formatDate(req.matchDate);
  var timeStr = TIME_LABELS[req.matchTime] || req.matchTime;

  var ccList = (req.groupPlayers || [])
    .map(function(p) { return _resolveEmail(p.name, p.email, players); })
    .concat([volEmail])
    .filter(Boolean)
    .filter(function(e) { return e.toLowerCase() !== reqEmail.toLowerCase(); });
  var seenCc = {};
  ccList = ccList.filter(function(e) {
    var lc = e.toLowerCase();
    if (seenCc[lc]) return false;
    seenCc[lc] = true;
    return true;
  });

  var subject = 'MWF Tennis League — ' + volunteerName + ' just volunteered for ' + dateStr;

  var body =
    'Hi ' + (req.name || 'there') + ',\n\n' +
    volunteerName + ' has just volunteered to sub on ' + dateStr + ' at ' + timeStr +
    ' — the same day and time as your open sub request — but it is too late for today\'s Dispatch to assign them automatically.\n\n' +
    'If you are still looking for a sub, contact each other directly:\n\n' +
    '  ' + volunteerName + '  |  ' + volEmail + (volPhone ? '  |  ' + volPhone : '') + '\n' +
    '  ' + (req.name || '') + '  |  ' + reqEmail + (reqPhone ? '  |  ' + reqPhone : '') + '\n\n' +
    'MWF Tennis League';

  var htmlBody =
    'Hi ' + (req.name || 'there') + ',<br><br>' +
    '<strong>' + volunteerName + '</strong> has just volunteered to sub on <strong>' + dateStr + '</strong> at <strong>' + timeStr + '</strong>' +
    ' — the same day and time as your open sub request — but it is too late for today\'s Dispatch to assign them automatically.<br><br>' +
    'If you are still looking for a sub, contact each other directly:<br><br>' +
    '<table style="font-family:Arial,sans-serif;font-size:14px;border-collapse:collapse;">' +
      '<tr><td style="padding:3px 12px 3px 0;font-weight:600;">' + volunteerName + '</td><td style="padding:3px 12px 3px 0;">' + volEmail + '</td><td style="padding:3px 0;">' + (volPhone || '—') + '</td></tr>' +
      '<tr><td style="padding:3px 12px 3px 0;font-weight:600;">' + (req.name || '') + '</td><td style="padding:3px 12px 3px 0;">' + reqEmail + '</td><td style="padding:3px 0;">' + (reqPhone || '—') + '</td></tr>' +
    '</table><br>' +
    'MWF Tennis League';

  sendLeagueEmail({
    to:       reqEmail,
    cc:       ccList.join(','),
    subject:  subject,
    body:     body,
    htmlBody: htmlBody,
    name:     'MWF Tennis League'
  });
  Logger.log('Late-volunteer notification sent: ' + reqEmail + ' <- ' + volunteerName + ' (' + req.id + ')');
}

function sendLeftoverVolunteersEmail(targetDate) {
  if (!isEmailEnabled()) return;
  var volunteers = getLeftoverVolunteersForDate(targetDate);

  var scheduledPlayers = getPlayersScheduledForDate(targetDate);
  // Once a request is filled, updateScheduleForSub swaps the substitute into the
  // schedule slot — the original requestor drops out of scheduledPlayers entirely and
  // isn't picked up anywhere else. Add them back explicitly.
  var filledRequestors = getRequests().filter(function(r) {
    return r.matchDate === targetDate && r.status === 'filled' &&
      r.email && !/^anita\.sub\d+@xgmail\.com$/i.test(r.email);
  });
  var recipients = {};
  volunteers.forEach(function(v) { if (v.email) recipients[v.email.toLowerCase()] = v.email; });
  scheduledPlayers.forEach(function(p) { if (p.email) recipients[p.email.toLowerCase()] = p.email; });
  filledRequestors.forEach(function(r) { if (r.email) recipients[r.email.toLowerCase()] = r.email; });
  var toList = Object.keys(recipients).map(function(k) { return recipients[k]; });
  if (!toList.length) return;

  var allPlayers = getPlayers();
  var phoneByEmail = {};
  allPlayers.forEach(function(p) { if (p.email) phoneByEmail[p.email.toLowerCase()] = p.phone || ''; });
  var volunteersWithPhone = volunteers.map(function(v) {
    return { name: v.name, email: v.email, times: v.times, phone: phoneByEmail[(v.email || '').toLowerCase()] || '' };
  });

  var groups = getMatchGroupsForDate(targetDate);

  var config  = getConfig();
  var dateStr = formatDate(targetDate);

  var emailParams = {
    to:       toList.join(','),
    subject:  'MWF Tennis League — Players available for ' + dateStr + ' if needed',
    body:     buildLeftoverVolunteersEmailText(volunteersWithPhone, groups),
    htmlBody: buildLeftoverVolunteersEmailHtml(volunteersWithPhone, groups),
    name:     'MWF Tennis League'
  };
  if (config.senderEmail) emailParams.cc = config.senderEmail;

  sendLeagueEmail(emailParams);
  Logger.log('Leftover volunteers email sent for ' + targetDate + ' to ' + toList.length + ' recipient(s).');
}

// ── Pre-match-day dispatch (runs 5× on Sun/Tue/Thu via fixed weekly triggers) ──
// Config rows 43–47 (Time / Dispatch / Broadcast / Cancel) control each run.

function _preMatchDayTargetDate() {
  var tz  = Session.getScriptTimeZone();
  var dow = parseInt(Utilities.formatDate(new Date(), tz, 'u')); // 1=Mon … 7=Sun
  if (dow !== 7 && dow !== 2 && dow !== 4) return null; // Sun/Tue/Thu only
  return getDateStr(1);
}

// Count of configured Pre-Match Day dispatch runs (Config A43:E47) whose hour hasn't
// passed yet today.
function _remainingPreMatchRunsToday(config) {
  var tz = Session.getScriptTimeZone();
  var currentHour = parseInt(Utilities.formatDate(new Date(), tz, 'H'));
  var remaining = 0;
  (config.preMatchSchedule || []).forEach(function(r) {
    var h = _parseConfigHour(r.time);
    if (h > currentHour) remaining++;
  });
  return remaining;
}

// True once today no longer has any chance of an automated Pre-Match Day
// Dispatch run — either today isn't a Pre-Match Day at all (not Sun/Tue/Thu,
// so none was ever coming), or it is but every enabled run's hour has already
// passed. Used to decide whether a volunteer who just offered to sub tomorrow
// missed Dispatch's window entirely for today.
function _hasRemainingPreMatchDayDispatchToday(config) {
  if (!_preMatchDayTargetDate()) return false;
  var tz = Session.getScriptTimeZone();
  var currentHour = parseInt(Utilities.formatDate(new Date(), tz, 'H'));
  return (config.preMatchSchedule || []).some(function(r) {
    return r.dispatch && _parseConfigHour(r.time) > currentHour;
  });
}

function _parseConfigHour(timeStr) {
  var s = (timeStr || '').toString().trim().toUpperCase();
  var m = s.match(/^(\d+)(?::\d+)?\s*(AM|PM)$/);
  if (m) {
    var h = parseInt(m[1]);
    if (m[2] === 'PM' && h !== 12) h += 12;
    if (m[2] === 'AM' && h === 12) h = 0;
    return h;
  }
  var m2 = s.match(/^(\d+):(\d+)$/);
  if (m2) return parseInt(m2[1]);
  return -1;
}

// Cancels the given player's own still-open (pending) Volunteers record for
// date, if any — used when their own sub request for that same day just got
// auto-cancelled with no sub found, so they're no longer available to sub for
// someone else's match on a day their own situation is still unresolved.
// Leaves an already-matched volunteer record alone (that's a real confirmed
// assignment for a different match, not just an open offer).
function _cancelOwnOpenVolunteerRecord(ss, email, date) {
  var emailLower = (email || '').toLowerCase();
  var match = getVolunteers().find(function(v) {
    return v.email.toLowerCase() === emailLower && v.date === date && v.status === 'pending';
  });
  if (!match) return false;
  ss.getSheetByName(TABS.volunteers).getRange(match.rowIndex, 7).setValue('cancelled');
  Logger.log('runPreMatchDayDispatch: cancelled ' + email + '\'s own volunteer record for ' + date +
    ' — their sub request for that day was just auto-cancelled');
  return true;
}

function runPreMatchDayDispatch() {
  var targetDate = _preMatchDayTargetDate();
  if (!targetDate) { Logger.log('runPreMatchDayDispatch: not a pre-match day'); return; }
  _recordDispatchRun('Pre-Match Day Dispatch');
  try { expireUpToToday(); } catch(e) { Logger.log('expireUpToToday failed: ' + e.message); }

  var config      = getConfig();
  var tz          = Session.getScriptTimeZone();
  var currentHour = parseInt(Utilities.formatDate(new Date(), tz, 'H'));

  // Find the config row whose scheduled hour is within 1 h of now
  var row = null;
  (config.preMatchSchedule || []).forEach(function(r) {
    var h = _parseConfigHour(r.time);
    if (h >= 0 && Math.abs(h - currentHour) <= 1) row = r;
  });
  if (!row) row = { dispatch: true, broadcast: true, cancel: false, expandVolunteers: false }; // safe fallback

  Logger.log('runPreMatchDayDispatch: ' + targetDate + ' hour=' + currentHour +
    ' dispatch=' + row.dispatch + ' broadcast=' + row.broadcast + ' cancel=' + row.cancel +
    ' expandVolunteers=' + !!row.expandVolunteers);

  if (row.dispatch) runDispatchAllOpen(!!row.expandVolunteers);

  var openReqs = getOpenRequestsForDate(targetDate);
  if (openReqs.length && row.broadcast && isEmailEnabled() && config.urgentSubEmailsEnabled) {
    try { sendUrgentSubBroadcast(openReqs, targetDate); }
    catch(e) { Logger.log('sendUrgentSubBroadcast failed: ' + e.message); }
  }
  if (row.cancel) {
    if (openReqs.length) {
      var ss       = SpreadsheetApp.openById(SHEET_ID);
      var reqSheet = ss.getSheetByName(TABS.requests);
      openReqs.forEach(function(req) {
        reqSheet.getRange(req.rowIndex, 7).setValue('cancelled');
        try { sendSubNeededTomorrowEmail(req); } catch(e) {
          Logger.log('Cancel notify failed for ' + req.id + ': ' + e.message);
        }
        // The requester may have also volunteered to sub elsewhere this same day
        // (e.g. trying to switch groups) — with their own request now cancelled
        // and no sub found, that offer should go away too rather than risk Rally
        // assigning them as a sub for someone else's match.
        try { _cancelOwnOpenVolunteerRecord(ss, req.email, targetDate); } catch(e) {
          Logger.log('Own-volunteer-record cancel failed for ' + req.id + ': ' + e.message);
        }
      });
    }
    // Independent of whether every request got filled — a volunteer can go unused
    // even with an open request if their rating falls outside the match's skill window.
    try { sendLeftoverVolunteersEmail(targetDate); } catch(e) {
      Logger.log('Leftover volunteers notify failed for ' + targetDate + ': ' + e.message);
    }
  }
}

// Alias kept so the 8 PM trigger name still resolves after any old trigger references
function runPreMatchDayDispatchFinal() { runPreMatchDayDispatch(); }

function _matchDayMinus2TargetDate() {
  var tz  = Session.getScriptTimeZone();
  var dow = parseInt(Utilities.formatDate(new Date(), tz, 'u')); // 1=Mon…7=Sun
  if (dow !== 6 && dow !== 1 && dow !== 3) return null; // Sat/Mon/Wed only
  return getDateStr(2); // target match is 2 days out
}

function runMatchDayMinus2Dispatch() {
  var targetDate = _matchDayMinus2TargetDate();
  if (!targetDate) { Logger.log('runMatchDayMinus2Dispatch: not a match day -2'); return; }
  _recordDispatchRun('Match Day -2 Dispatch');
  try { expireUpToToday(); } catch(e) { Logger.log('expireUpToToday failed: ' + e.message); }

  var config      = getConfig();
  var tz          = Session.getScriptTimeZone();
  var currentHour = parseInt(Utilities.formatDate(new Date(), tz, 'H'));

  var row = null;
  (config.matchDayMinus2Schedule || []).forEach(function(r) {
    var h = _parseConfigHour(r.time);
    if (h >= 0 && Math.abs(h - currentHour) <= 1) row = r;
  });
  if (!row) row = { dispatch: true, broadcast: true };

  Logger.log('runMatchDayMinus2Dispatch: ' + targetDate + ' hour=' + currentHour +
    ' dispatch=' + row.dispatch + ' broadcast=' + row.broadcast +
    ' overflowDetect=' + !!row.overflowDetect);

  // Before filling anything: TBD requests this close to match day likely mean the group
  // is on Overflow with no Chelsea court time yet. Mark them Overflow so dispatch skips
  // them instead of assigning a volunteer to a request that may never get a real time.
  if (row.overflowDetect) {
    try { markOverflowRequests(targetDate); } catch(e) { Logger.log('markOverflowRequests failed: ' + e.message); }
  }

  if (row.dispatch) runDispatchAllOpen();

  var openReqs = getOpenRequestsForDate(targetDate);
  if (openReqs.length && row.broadcast && isEmailEnabled() && config.urgentSubEmailsEnabled) {
    try { sendUrgentSubBroadcast(openReqs, targetDate); }
    catch(e) { Logger.log('sendUrgentSubBroadcast failed: ' + e.message); }
  }
}

// One-shot trigger handler — fires ~1 min after being queued by submitRequest()'s
// immediate-broadcast-resend check.
function _runQueuedBroadcast() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === '_runQueuedBroadcast') ScriptApp.deleteTrigger(t);
  });
  var targetDate = getDateStr(1);
  var openReqs = getOpenRequestsForDate(targetDate);
  if (openReqs.length && isEmailEnabled() && getConfig().urgentSubEmailsEnabled) {
    try { sendUrgentSubBroadcast(openReqs, targetDate); }
    catch(e) { Logger.log('sendUrgentSubBroadcast failed: ' + e.message); }
  }
}

function sendTestSubAlertEmail() {
  // Find open requests — use the earliest upcoming date
  var allOpen = getRequests().filter(function(r) { return r.status === 'open'; });
  if (!allOpen.length) {
    return { success: false, error: 'No open sub requests found to preview.' };
  }
  var targetDate = Object.keys(allOpen.reduce(function(acc, r) { acc[r.matchDate] = true; return acc; }, {})).sort()[0];
  var openReqs   = allOpen.filter(function(r) { return r.matchDate === targetDate; });

  var testPlayers = getPlayersWithRatings().filter(function(p) {
    return p.email && !/^anita\.sub\d+@xgmail\.com$/i.test(p.email) && p.isTest;
  });
  if (!testPlayers.length) {
    return { success: false, error: 'No test players found — add "Yes" in the Test column of the Players sheet.' };
  }

  var d        = new Date(targetDate + 'T12:00:00');
  var monthDay = d.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
  var subject  = 'MWF Tennis, subs needed ' + monthDay;
  var scriptUrl = SCRIPT_URL;

  var sent = 0, errors = [];
  testPlayers.forEach(function(player) {
    try {
      sendLeagueEmail({
        to:       player.email,
        subject:  subject,
        body:     buildSubNeededEmailText(openReqs, targetDate),
        htmlBody: buildSubNeededEmailHtml(openReqs, scriptUrl),
        name:     'MWF Tennis League'
      });
      sent++;
    } catch(e) {
      Logger.log('Test sub alert failed for ' + player.email + ': ' + e.message);
      errors.push(player.email + ': ' + e.message);
    }
  });

  if (sent === 0) return { success: false, error: 'All sends failed. ' + (errors[0] || '') };
  return { success: true, emailsSent: sent, date: targetDate };
}

// Sends the broadcast sub-needed email to marobria@gmail.com only (1 quota slot),
// for manual forwarding to the league when the scheduled broadcast fails.
function sendBroadcastEmailToAdmin() {
  var targetDate = getDateStr(1);
  var openReqs   = getOpenRequestsForDate(targetDate);
  if (!openReqs.length) {
    return { success: false, error: 'No open sub requests for ' + targetDate };
  }
  var d        = new Date(targetDate + 'T12:00:00');
  var monthDay = d.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
  var subject  = '[FORWARD TO LEAGUE] MWF Tennis, subs needed ' + monthDay;
  try {
    MailApp.sendEmail({
      to:       'marobria@gmail.com',
      subject:  subject,
      body:     buildSubNeededEmailText(openReqs, targetDate),
      htmlBody: buildSubNeededEmailHtml(openReqs, SCRIPT_URL),
      name:     'MWF Tennis League'
    });
    Logger.log('sendBroadcastEmailToAdmin: sent for ' + targetDate + ' (' + openReqs.length + ' open requests)');
    return { success: true, targetDate: targetDate, openRequests: openReqs.length };
  } catch(e) {
    Logger.log('sendBroadcastEmailToAdmin failed: ' + e.message);
    return { success: false, error: e.message };
  }
}

// Removes duplicate volunteer rows, keeping the best record per email+date:
// matched > earliest pending. Run once from the Apps Script editor to clean up existing data.
function deduplicateVolunteers() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.volunteers);
  if (!sheet || sheet.getLastRow() < 2) return { removed: 0 };
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();

  var seen     = {}; // email|date → { rowIndex, status, timestamp }
  var toDelete = [];

  rows.forEach(function(r, i) {
    var email     = (r[3] || '').toLowerCase().trim();
    var date      = formatSheetDate(r[4]);
    var status    = (r[6] || '').toLowerCase();
    var timestamp = r[1] ? new Date(r[1]).getTime() : 0;
    var rowIndex  = i + 2;
    if (!email || !date) return;
    var key = email + '|' + date;
    if (!seen[key]) {
      seen[key] = { rowIndex: rowIndex, status: status, timestamp: timestamp };
    } else {
      var kept = seen[key];
      var keepCurrent = (status === 'matched' && kept.status !== 'matched') ||
                        (status === kept.status && timestamp < kept.timestamp);
      if (keepCurrent) {
        toDelete.push(kept.rowIndex);
        seen[key] = { rowIndex: rowIndex, status: status, timestamp: timestamp };
      } else {
        toDelete.push(rowIndex);
      }
    }
  });

  toDelete.sort(function(a, b) { return b - a; }); // delete from bottom up
  toDelete.forEach(function(rowIndex) { sheet.deleteRow(rowIndex); });
  Logger.log('deduplicateVolunteers: removed ' + toDelete.length + ' duplicate(s)');
  return { removed: toDelete.length };
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
    groupLetter:    req.groupLetter,
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

function formatDateShort(str) {
  if (!str) return '—';
  const d = new Date(str + 'T12:00:00');
  return d.toLocaleDateString('en-US', { month: 'short' }) + ', ' + d.getDate();
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

  if (daysUntilClose !== 1 && daysUntilClose !== 0) return;

  var currentHour = parseInt(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'H'));
  if (currentHour >= 12) { Logger.log('checkAvailabilityWindow: skipping — afternoon/evening run'); return; }

  var missing = getPlayersWithoutSubmission(config.targetMonth).filter(function(p) {
    return !/^anita\.sub\d+@xgmail\.com$/i.test(p.email || '');
  });
  if (!missing.length) {
    Logger.log('checkAvailabilityWindow: T-' + daysUntilClose + ' reminder — all players already submitted');
    return;
  }

  var closeDateLabel = closeDate.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
  var urgency        = daysUntilClose === 0 ? 'today' : 'tomorrow';
  var avUrl          = APP_BASE_URL + '#availability';
  var subject        = 'Reminder: Submit your availability for ' + config.targetMonthLabel + ' — closes ' + urgency;
  var body =
    'Your monthly availability has not been received.\n\n' +
    'Just a reminder — the availability window for ' + config.targetMonthLabel + ' closes ' + urgency + ' (' + closeDateLabel + ').\n\n' +
    'Please submit your available dates before the window closes so we can include you in the schedule.\n\n' +
    'Open the My Availability page to submit:\n' +
    avUrl + '\n\n' +
    'See you on the court!\n' +
    'MWF Tennis League';
  var htmlBody =
    'Your monthly availability has <u>not</u> been received.<br><br>' +
    'Just a reminder — the availability window for <strong>' + config.targetMonthLabel + '</strong> closes ' + urgency + ' (' + closeDateLabel + ').<br><br>' +
    'Please submit your available dates before the window closes so we can include you in the schedule.<br><br>' +
    'Open the <a href="' + avUrl + '">My Availability</a> page to submit.<br><br>' +
    'See you on the court!<br>' +
    'MWF Tennis League';

  Logger.log('checkAvailabilityWindow: T-' + daysUntilClose + ' reminder → ' + missing.length + ' player(s)');
  if (!isEmailEnabled()) return;

  var adminEmail = 'marobria@gmail.com';
  sendLeagueEmail({
    to:       adminEmail,
    bcc:      _excludeFromBcc(missing.map(function(p) { return p.email; }), adminEmail).join(','),
    subject:  subject,
    body:     body,
    htmlBody: htmlBody,
    name:     'MWF Tennis League'
  });
}

function testCheckAvailabilityWindowEmail() {
  var config      = getAvailabilityConfig();
  var closeDate   = new Date(config.closeDate + 'T00:00:00');
  var closeDateLabel = closeDate.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
  var avUrl       = APP_BASE_URL + '#availability';
  var subject     = 'Reminder: Submit your availability for ' + config.targetMonthLabel + ' — closes today';
  var body =
    'Your monthly availability has not been received.\n\n' +
    'Just a reminder — the availability window for ' + config.targetMonthLabel + ' closes today (' + closeDateLabel + ').\n\n' +
    'Please submit your available dates before the window closes so we can include you in the schedule.\n\n' +
    'Open the My Availability page to submit:\n' +
    avUrl + '\n\n' +
    'See you on the court!\n' +
    'MWF Tennis League';
  var htmlBody =
    'Your monthly availability has <u>not</u> been received.<br><br>' +
    'Just a reminder — the availability window for <strong>' + config.targetMonthLabel + '</strong> closes today (' + closeDateLabel + ').<br><br>' +
    'Please submit your available dates before the window closes so we can include you in the schedule.<br><br>' +
    'Open the <a href="' + avUrl + '">My Availability</a> page to submit.<br><br>' +
    'See you on the court!<br>' +
    'MWF Tennis League';
  MailApp.sendEmail({ to: 'marobria@gmail.com', subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
  return { success: true };
}

// Diagnostic: returns EmailLog rows newer than params.hours ago (default 24), so a
// failed dispatch run can be audited without opening the Sheet.
// Manual recovery: resends the urgent-sub broadcast for an explicit date, using
// whichever requests are open right now. For when a scheduled broadcast failed
// outright (e.g. MailApp quota) and needs a manual retry for that date.
// Manual recovery, quota-safe: sends the subs-needed content for an explicit date to
// marobria@gmail.com only (1 recipient) with the full player address list appended,
// for manual forward/BCC when MailApp quota is too low to BCC the whole roster directly.
function sendBroadcastFallbackToAdmin(params) {
  var targetDate = (params.matchDate || '').toString().trim();
  if (!targetDate) return { success: false, error: 'matchDate is required (yyyy-MM-dd).' };
  var openReqs = getOpenRequestsForDate(targetDate);
  if (!openReqs.length) return { success: true, sent: false, reason: 'No open requests for ' + targetDate };

  var players = getPlayersWithRatings().filter(function(p) {
    return p.email && !/^anita\.sub\d+@xgmail\.com$/i.test(p.email);
  });
  var addressList = players.map(function(p) { return p.name + ' <' + p.email + '>'; });

  var d        = new Date(targetDate + 'T12:00:00');
  var monthDay = d.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
  var subject  = '[FORWARD TO LEAGUE] MWF Tennis, subs needed ' + monthDay;

  var body = buildSubNeededEmailText(openReqs, targetDate) +
    '\n\n---\nForward to (or BCC):\n' + addressList.join('\n');
  var htmlBody = buildSubNeededEmailHtml(openReqs, SCRIPT_URL) +
    '<div style="margin-top:20px;padding:16px;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#111;">' +
    '<strong>Forward to (or BCC):</strong><br>' + addressList.map(function(a) { return a.replace(/</g, '&lt;').replace(/>/g, '&gt;'); }).join('<br>') +
    '</div>';

  MailApp.sendEmail({ to: 'marobria@gmail.com', subject: subject, body: body, htmlBody: htmlBody, name: 'MWF Tennis League' });
  return { success: true, sent: true, recipientCount: addressList.length };
}

function resendUrgentSubBroadcast(params) {
  var targetDate = (params.matchDate || '').toString().trim();
  if (!targetDate) return { success: false, error: 'matchDate is required (yyyy-MM-dd).' };
  if (!isEmailEnabled()) return { success: false, error: 'Email is disabled (Config B27).' };
  var config = getConfig();
  if (!config.urgentSubEmailsEnabled) return { success: false, error: 'Urgent Sub Emails disabled (Config B39).' };
  var openReqs = getOpenRequestsForDate(targetDate);
  if (!openReqs.length) return { success: true, sent: false, reason: 'No open requests for ' + targetDate };
  sendUrgentSubBroadcast(openReqs, targetDate);
  return { success: true, sent: true, openRequests: openReqs.length };
}

function getRecentEmailLog(params) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TABS.emailLog);
  if (!sheet || sheet.getLastRow() < 2) return { rows: [] };
  var hours  = parseFloat(params.hours) || 24;
  var cutoff = new Date(Date.now() - hours * 60 * 60 * 1000);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  var rows = data
    .map(function(r) {
      var ts = r[0] instanceof Date ? r[0] : new Date(r[0]);
      return { timestamp: ts, to: r[1], subject: r[2], status: r[3] };
    })
    .filter(function(r) { return r.timestamp >= cutoff; })
    .map(function(r) {
      return {
        timestamp: Utilities.formatDate(r.timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
        to: r.to, subject: r.subject, status: r.status
      };
    });
  return { rows: rows };
}

function checkEmailQuota() {
  var remaining = MailApp.getRemainingDailyQuota();
  Logger.log('Remaining daily email quota: ' + remaining);
  // Self-delete so a scheduled run is one-shot
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkEmailQuota') {
      try { ScriptApp.deleteTrigger(t); } catch(e) {}
    }
  });
  try {
    MailApp.sendEmail({
      to: 'marobria@gmail.com',
      subject: 'Rally: daily email quota at 3:45 AM = ' + remaining,
      body: 'Remaining MailApp recipients for today: ' + remaining + '\n\nThis is an automated diagnostic check.'
    });
  } catch(e) {
    Logger.log('Could not send quota email: ' + e.message);
  }
  return { remaining: remaining };
}

function scheduleCheckEmailQuotaTomorrow() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkEmailQuota') {
      try { ScriptApp.deleteTrigger(t); } catch(e) {}
    }
  });
  ScriptApp.newTrigger('checkEmailQuota')
    .timeBased().atHour(3).nearMinute(45).everyDays(1)
    .inTimezone('America/New_York').create();
  Logger.log('checkEmailQuota scheduled for 3:45 AM ET (will self-delete after first run)');
  return { scheduled: true };
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
  SpreadsheetApp.flush();

  // Queue the email blast via a one-shot trigger so the HTTP response returns immediately.
  // Sending inline times out the JSONP request before all emails complete.
  if (isEmailEnabled()) {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === '_runQueuedAvailBlast') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('_runQueuedAvailBlast').timeBased().after(30 * 1000).create();
  }

  return { success: true };
}

// One-shot trigger handler — fires ~30s after openAvailabilityWindow queues it.
function _runQueuedAvailBlast() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === '_runQueuedAvailBlast') ScriptApp.deleteTrigger(t);
  });
  const availConfig = getAvailabilityConfig();
  if (!availConfig.isOpen || !isEmailEnabled()) return;

  const allPlayers = getPlayers().filter(function(p) {
    return p.email && !/^anita\.sub\d+@xgmail\.com$/i.test(p.email);
  });
  if (!allPlayers.length) return;

  const closeDateLabel = new Date(availConfig.closeDate + 'T12:00:00').toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
  const avUrl    = APP_BASE_URL + '#availability';
  const subject  = 'MWF League - Submit your availability for ' + availConfig.targetMonthLabel;
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

  const adminEmail = 'marobria@gmail.com';
  sendLeagueEmail({
    to:       adminEmail,
    bcc:      _excludeFromBcc(allPlayers.map(function(p) { return p.email; }), adminEmail).join(','),
    subject:  subject,
    body:     body,
    htmlBody: htmlBody,
    name:     'MWF Tennis League'
  });
  Logger.log('Availability blast sent via BCC to ' + allPlayers.length + ' players.');
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
  const emailParam     = (params.email         || '').toLowerCase();
  const month          = params.month          || '';
  const availableDates = params.availableDates || '[]';
  const notes          = params.notes          || '';

  // Always use the current email from the Players sheet
  const players = getPlayers();
  const email   = _resolveEmail(name, emailParam, players);

  Logger.log('submitAvailability called: name=%s email=%s month=%s dates=%s', name, email, month, availableDates);

  if (!name || !email || !month) return { success: false, error: 'Missing required fields.' };

  // Validate window is still open
  const avConfig = getAvailabilityConfig();
  if (!avConfig.isOpen) return { success: false, error: 'The availability window is currently closed.' };

  const sheet   = getOrCreateAvailabilitySheet();
  const lastRow = sheet.getLastRow();

  // Upsert: match existing row by email OR by name (handles email address changes)
  let targetRow = -1;
  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    rows.forEach(function(r, i) {
      var rowEmail = (r[2] || '').toLowerCase();
      var rowName  = (r[1] || '').toLowerCase();
      if (normalizeMonth(r[3]) === month &&
          (rowEmail === email || rowEmail === emailParam || rowName === name.toLowerCase())) {
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

// Deletes records older than the previous month across every data tab, so EmailLog,
// MatchGroups, DispatchLog, SubRequests, Volunteers, and Availability don't grow
// unbounded. Runs monthly via the trigger created in setupTriggers().
function cleanupOldRecords() {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var tz  = Session.getScriptTimeZone();
  var now = new Date();
  var cutoffDate  = new Date(now.getFullYear(), now.getMonth() - 1, 1); // first of previous month
  var cutoffStr   = Utilities.formatDate(cutoffDate, tz, 'yyyy-MM-dd');
  var cutoffMonth = Utilities.formatDate(cutoffDate, tz, 'yyyy-MM');

  var results = {
    emailLog:     _deleteRowsBeforeDate(ss, TABS.emailLog,      1, cutoffStr),
    matchGroups:  _deleteRowsBeforeDate(ss, TABS.matchGroups,   3, cutoffStr),
    dispatchLog:  _deleteRowsBeforeDate(ss, 'DispatchLog',      4, cutoffStr),
    subRequests:  _deleteRowsBeforeDate(ss, TABS.requests,      5, cutoffStr),
    volunteers:   _deleteRowsBeforeDate(ss, TABS.volunteers,    5, cutoffStr),
    availability: _deleteRowsBeforeMonth(ss, TABS.availability, 4, cutoffMonth)
  };
  Logger.log('cleanupOldRecords: ' + JSON.stringify(results));
  return results;
}

// Deletes rows (bottom-up, to avoid index shifting as rows are removed) whose date in
// `colIndex1Based` falls before `cutoffStr` (yyyy-MM-dd). Handles both Date-object and
// string cells, since Sheets silently converts date-like strings to Date on write.
function _deleteRowsBeforeDate(ss, tabName, colIndex1Based, cutoffStr) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var values  = sheet.getRange(2, colIndex1Based, lastRow - 1, 1).getValues();
  var deleted = 0;
  for (var i = values.length - 1; i >= 0; i--) {
    var raw = values[i][0];
    if (!raw) continue;
    var dateStr = raw instanceof Date
      ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : raw.toString().trim().slice(0, 10);
    if (dateStr && dateStr < cutoffStr) {
      sheet.deleteRow(i + 2);
      deleted++;
    }
  }
  return deleted;
}

// Same as _deleteRowsBeforeDate but for a "YYYY-MM" Month column (Availability).
function _deleteRowsBeforeMonth(ss, tabName, colIndex1Based, cutoffMonth) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var values  = sheet.getRange(2, colIndex1Based, lastRow - 1, 1).getValues();
  var deleted = 0;
  for (var i = values.length - 1; i >= 0; i--) {
    var monthStr = normalizeMonth(values[i][0]);
    if (monthStr && monthStr < cutoffMonth) {
      sheet.deleteRow(i + 2);
      deleted++;
    }
  }
  return deleted;
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
// Per-player targets can be overridden below (e.g. 0.10 = ~10% captaincy goal).
// Adds slot.captains = [emailForGroupA, emailForGroupB, ...] to every active slot.
var CAPTAIN_TARGETS = {
  'marobria@gmail.com': 0
};

// Hard cap — no player is made captain more than this many times in one month's
// schedule, regardless of how far below their target ratio they are.
var CAPTAIN_MAX_PER_MONTH = 3;

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

  // Greedy assignment: pick the player whose captaincy ratio is furthest below
  // their individual target (ratio / target — lower score = more overdue), skipping
  // anyone who has already hit the monthly cap.
  var captainCounts = {};
  slotResults.forEach(function(slot) {
    if (slot.skipped) return;
    var captains = [];
    slot.groups.forEach(function(group) {
      var best = null;
      var bestScore = Infinity;
      group.forEach(function(p) {
        if (!p.email) return;
        var emailKey = p.email.toLowerCase();
        var target = CAPTAIN_TARGETS.hasOwnProperty(emailKey) ? CAPTAIN_TARGETS[emailKey] : 0.25;
        if (target === 0) return;
        if ((captainCounts[p.email] || 0) >= CAPTAIN_MAX_PER_MONTH) return;
        var ratio  = (captainCounts[p.email] || 0) / (appearanceCounts[p.email] || 1);
        var score  = ratio / target;
        if (score < bestScore) { bestScore = score; best = p; }
      });
      // Every candidate in this group already hit the cap (only possible with a
      // very small, repeatedly-paired roster) — fall back to ignoring the cap
      // rather than leaving the group without a captain ("captain is always P1").
      if (!best) {
        group.forEach(function(p) {
          if (!p.email) return;
          var emailKey = p.email.toLowerCase();
          var target = CAPTAIN_TARGETS.hasOwnProperty(emailKey) ? CAPTAIN_TARGETS[emailKey] : 0.25;
          if (target === 0) return;
          var ratio = (captainCounts[p.email] || 0) / (appearanceCounts[p.email] || 1);
          if (ratio < bestScore) { bestScore = ratio; best = p; }
        });
      }
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
        uid(), nowEasternISO(),
        anitaName, anitaEmail,
        slot.date, '', 'open', '', groupPlayersJSON
      ]);
      var lastReqRow = rSheet.getLastRow();
      rSheet.getRange(lastReqRow, 5).setNumberFormat('@');
      rSheet.getRange(lastReqRow, 6).setNumberFormat('@');
      rSheet.getRange(lastReqRow, 9).setNumberFormat('@');
      _setGroupLetterOnRequestRow(rSheet, lastReqRow, String.fromCharCode(65 + gi));
      _flagNo8amOnRequestRow(rSheet, lastReqRow, groupForRequest.map(function(p) { return p.email; }));

      Logger.log('Created ' + anitaName + ' (rating ' + anitaRating + ') for ' + slot.date + ' group ' + String.fromCharCode(65 + gi));
      var captainPlayer = workingGroup.find(function(p) { return p.email.toLowerCase() === captainEmail.toLowerCase(); });
      try { sendCaptainThreePlayerNotification(captainPlayer ? captainPlayer.name : '', captainEmail, slot.date, anitaName); }
      catch(emailErr) { Logger.log('Captain notify failed (email): ' + emailErr.message); }
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
      sitOut2Name, sitOut2Email,
      '' // Time — populated later by the Chelsea import or a manual View Schedule edit
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
    var thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
    upsertVolunteerTimes(volSheet, sitOutName, sitOutEmail.toLowerCase(), slot.date, sitOutTimes.split(','), thirtyDaysAgo);
    Logger.log('Recorded volunteer availability for sit-out: ' + sitOutName + ' on ' + slot.date + ' times: ' + sitOutTimes);
    try { sendSitOutNotification(sitOutName, sitOutEmail, slot.date); }
    catch(emailErr) { Logger.log('Sit-out notify failed (email): ' + emailErr.message); }
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
    var thirtyDaysAgo2 = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
    upsertVolunteerTimes(volSheet2, sitOut2Name, sitOut2Email.toLowerCase(), slot.date, sitOut2Times.split(','), thirtyDaysAgo2);
    Logger.log('Recorded volunteer availability for 2nd alternate: ' + sitOut2Name + ' on ' + slot.date + ' times: ' + sitOut2Times);
    try { sendSitOutNotification(sitOut2Name, sitOut2Email, slot.date); }
    catch(emailErr) { Logger.log('Sit-out2 notify failed (email): ' + emailErr.message); }
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

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();

  // Callers only ever use today-forward dates (View Schedule, Request a Sub,
  // Availability lock-checks) — skip past rows so response size doesn't grow
  // unbounded as MatchGroups accumulates month over month.
  var todayStr = getDateStr(0);

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
    if (date && date < todayStr) return;
    var letter = r[3] ? r[3].toString() : '';
    var sitOutName   = r[12] ? r[12].toString() : '';
    var sitOutEmail  = r[13] ? r[13].toString() : '';
    var sitOut2Name  = r[14] ? r[14].toString() : '';
    var sitOut2Email = r[15] ? r[15].toString() : '';
    var time         = r[16] ? r[16].toString().trim() : '';

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
      sitOut2: sitOut2Name ? { name: sitOut2Name, email: sitOut2Email } : null,
      time:    time
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
          sitOut2: dateMap[date][letter].sitOut2,
          time:    dateMap[date][letter].time
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

    var dayHtml = '<div style="padding:10px 0;border-top:1px solid #e5e7eb;">' +
      '<div style="font-weight:700;margin-bottom:6px;">' + dateLabel + '</div>';
    dayObj.groups.forEach(function(grp) {
      var realPlayers = grp.players.filter(function(p) {
        return p.name && !/^anita\.sub\d+@xgmail\.com$/i.test(p.email || '');
      });
      var names = realPlayers.map(function(p) { return p.name; }).join(', ');
      textLines.push('  Group ' + grp.letter + ': ' + names + (grp.sitOut ? ' (sub needed)' : ''));
      dayHtml += '<div style="margin:2px 0 0 12px;">Group ' + grp.letter + ': ' + names +
        (grp.sitOut ? ' <span style="color:#8A4F0B;">(sub needed)</span>' : '') + '</div>';
    });
    dayHtml += '</div>';
    htmlRows.push(dayHtml);
    textLines.push('');
  });
  var body = 'The MWF Tennis League schedule for ' + monthLabel + ' has been published.\n\n' +
    textLines.join('\n') + '\n\n' +
    'Court times will be announced separately as each date approaches.\n\n' +
    'View the schedule online: ' + scheduleUrl + '\n\n' +
    'The schedule is also attached as a spreadsheet file (CSV) that opens in Excel.';
  var htmlBody = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#111;max-width:650px;">' +
    '<p style="margin:0 0 12px 0;">The MWF Tennis League schedule for <strong>' + monthLabel +
    '</strong> has been published.</p>' +
    '<p style="margin:0 0 16px 0;"><a href="' + scheduleUrl + '" style="color:#1a5c3a;">View Schedule</a></p>' +
    htmlRows.join('') +
    '<p style="margin-top:16px;color:#666;font-size:12px;">Court times will be announced separately as each date approaches.</p>' +
    '<p style="margin-top:8px;color:#666;font-size:12px;">The schedule is also attached as a spreadsheet file (CSV, opens in Excel).</p>' +
    '</div>';
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
      var adminEmail2 = 'marobria@gmail.com';
      sendLeagueEmail({
        to:       adminEmail2,
        bcc:      _excludeFromBcc(allPlayers.map(function(p) { return p.email; }), adminEmail2).join(','),
        subject:  emailParts.subject,
        body:     emailParts.body,
        htmlBody: emailParts.htmlBody,
        name:     'MWF Tennis League'
      });
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
  var subject     = 'MWF Tennis League — ' + sd.monthLabel + ' Schedule';
  var sent = 0, sendErrors = [];
  testPlayers.forEach(function(recipient) {
    try {
      sendBrevoEmail({
        apiKey:      config.brevoApiKey,
        recipients:  [recipient],
        subject:     subject,
        htmlContent: buildScheduleHtml(sd.dateMap, sd.sortedDates, sd.monthLabel, scheduleUrl, recipient.name),
        textContent: buildScheduleTextBody(sd.dateMap, sd.sortedDates, sd.monthLabel, scheduleUrl, recipient.name)
      });
      sent++;
    } catch(e) {
      Logger.log('Brevo send failed for ' + recipient.email + ': ' + e.message);
      sendErrors.push(recipient.email + ': ' + e.message);
    }
  });
  if (sent === 0) {
    return { success: false, error: 'All sends failed. ' + (sendErrors[0] || '') };
  }
  return { success: true, emailsSent: sent, errors: sendErrors.length ? sendErrors : undefined };
}

// ── Sheet helper ────────────────────────────────────
function getOrCreateMatchGroupsSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.matchGroups);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.matchGroups);
    sheet.getRange(1, 1, 1, 17).setValues([[
      'Timestamp','Month','Date','Group',
      'P1 Name','P1 Email','P2 Name','P2 Email',
      'P3 Name','P3 Email','P4 Name','P4 Email',
      'SitOut Name','SitOut Email','SitOut2 Name','SitOut2 Email',
      'Time'
    ]]);
    sheet.setFrozenRows(1);
  } else if (!sheet.getRange(1, 17).getValue()) {
    // Existing sheets predate the Time column (col 17) — backfill just the header.
    sheet.getRange(1, 17).setValue('Time');
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
