/*******************************************************
 * Project Timeline → Calendar + Google Tasks Sync
 * - Calendar events with color + emoji by deadline proximity
 * - Guest invites + sendUpdates on change (Advanced Calendar API)
 * - Per-assignee Google Tasks via Domain-Wide Delegation (DWD)
 * - Sync Log sheet + idempotent Task IDs (JSON) column
 * - Conditional formatting for End Date (green/yellow/red)
 *******************************************************/
function test_ServiceAccountKey() {
  try {
    const props = PropertiesService.getScriptProperties();
    const clientEmail = props.getProperty('SA_CLIENT_EMAIL');
    const privateKey = props.getProperty('SA_PRIVATE_KEY');
    if (!clientEmail || !privateKey) {
      throw new Error('Missing SA_CLIENT_EMAIL or SA_PRIVATE_KEY in Script Properties');
    }

    Logger.log('Client Email: ' + clientEmail);

    // Normalize newlines in case key stored as single line
    const normalizedKey = privateKey.replace(/\\n/g, '\n');
    const hasEscaped = privateKey.indexOf('\\n') !== -1;
    const hasReal = normalizedKey.indexOf('\n') !== -1;

    Logger.log('Escaped \\n detected: ' + hasEscaped);
    Logger.log('Contains real newlines: ' + hasReal);
    Logger.log('Key length: ' + normalizedKey.length);

    // Try an actual token request (Tasks scope)
    const token = getSATokenForUser_(Session.getActiveUser().getEmail(), 'https://www.googleapis.com/auth/tasks');
    Logger.log('Access token retrieved successfully (first 100 chars): ' + token.substring(0, 100));

    Logger.log('✅ Service account key works and token exchange succeeded.');
  } catch (e) {
    Logger.log('❌ Service account key test failed: ' + e.message);
  }
}

// Choose color/emoji from daysBefore deadline
function decideColor_(daysBefore) {
  // No deadline? default green.
  if (daysBefore === null || daysBefore === undefined) {
    return { name: 'green', emoji: '🟢', calColor: CalendarApp.EventColor.GREEN };
  }

  // Deadline logic
  if (CONFIG.onDeadlineIsRed) {
    if (daysBefore <= 0) { // on or past the deadline
      return { name: 'red', emoji: '🔴', calColor: CalendarApp.EventColor.RED };
    }
  } else {
    if (daysBefore < 0) {  // strictly past the deadline
      return { name: 'red', emoji: '🔴', calColor: CalendarApp.EventColor.RED };
    }
  }

  if (daysBefore <= CONFIG.yellowWindowDays) {
    return { name: 'yellow', emoji: '🟡', calColor: CalendarApp.EventColor.YELLOW };
  }

  return { name: 'green', emoji: '🟢', calColor: CalendarApp.EventColor.GREEN };
}

/***** CONFIG *****/
const CONFIG = {
  calendarId: 'c_b697a80ae3fb9a9918f40ffdb570a633f9d2286db89a65123374589206aa3ea6@group.calendar.google.com',
  sheetName: null,                // null = active sheet
  headerRow: 1,
  deadlineTaskLabel: 'Project Deadline',
  eventIdHeader: 'Event ID',
  lastSyncedHeader: 'Last Synced',
  deadlineHeader: 'Deadline',
  taskIdsHeader: 'Task IDs (JSON)',      // per-row: { "user@dom": "taskId", ... }
  yellowWindowDays: 2,
  onDeadlineIsRed: true,

  // Calendar behavior
  sendInvites: true,                      // email guests on create
  sendUpdatesOnChange: true,              // email on updates (needs Advanced Calendar API)
  defaultRemindersMins: [1440, 120],      // 24h & 2h popup reminders
  addEmojiPrefix: true,                   // 🟢/🟡/🔴 prefix in event title

  // Google Tasks (DWD)
  createGoogleTasks: true,                // create personal tasks per assignee
  tasksListNameTemplate: 'Project: {{SHEET_NAME}}' // per-user tasks list name

  // SA creds are read from Script Properties: SA_CLIENT_EMAIL, SA_PRIVATE_KEY
};

/***** MENU *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Project')
    .addItem('Sync to Calendar', 'syncTasksToCalendar')
    .addItem('Refresh Deadline & Formatting', 'refreshDeadlineAndFormatting')
    .addItem('Open Sync Log', 'openSyncLog_')
    .addToUi();
}

/***** PUBLIC: refresh named range + conditional formatting only *****/
function refreshDeadlineAndFormatting() {
  const { sheet, headers } = getSheetAndHeaders_();
  const locating = locateDeadline_(sheet, headers);
  if (!locating.deadlineDate) {
    SpreadsheetApp.getActive().toast('No Project Deadline row found (or no date).');
    return;
  }
  ensureDeadlineNamedRange_(locating.deadlineRange);
  ensureEndDateConditionalFormatting_(sheet, headers, locating.firstTaskRow);
  SpreadsheetApp.getActive().toast('Deadline & conditional formatting updated.');
}

/***** MAIN SYNC *****/
function syncTasksToCalendar() {
  const { sheet, headers, values, richValues } = getSheetAndHeaders_({ includeValues: true });
  const locating = locateDeadline_(sheet, headers, values);

  // Ensure named range & conditional formatting
  if (locating.deadlineRange) ensureDeadlineNamedRange_(locating.deadlineRange);
  ensureEndDateConditionalFormatting_(sheet, headers, locating.firstTaskRow);

  // Ensure helper columns
  const eventIdColIndex  = ensureColumn_(sheet, headers, CONFIG.eventIdHeader);
  const lastSyncColIndex = ensureColumn_(sheet, headers, CONFIG.lastSyncedHeader);
  const taskIdsColIndex  = ensureColumn_(sheet, headers, CONFIG.taskIdsHeader);

  const cal = CalendarApp.getCalendarById(CONFIG.calendarId);
  if (!cal) throw new Error('Calendar not found. Check CONFIG.calendarId.');

  const tz = Session.getScriptTimeZone();
  const startR = Math.max(locating.firstTaskRow, CONFIG.headerRow + 1);

  for (let r = startR - 1; r < values.length; r++) {
    const row = values[r];
    const rawTitle = str_(row[headers['Task']]);
    if (!rawTitle || equalsCI_(rawTitle, CONFIG.deadlineTaskLabel)) continue;

    try {
      const start = toDate_(row[headers['Start Date']], tz);
      if (!start) { log_('skip', r + 1, rawTitle, '', 'No start date'); continue; }

      const endInclusive = isDate_(row[headers['End Date']])
        ? toDate_(row[headers['End Date']], tz)
        : addDays_(start, Math.max(1, Number(row[headers['Duration (days)']] || 1)) - 1);

      const endExclusive = addDays_(endInclusive, 1);

      // Row-level or project deadline
      const rowDeadlineCell = headers[CONFIG.deadlineHeader] != null ? row[headers[CONFIG.deadlineHeader]] : '';
      const projectDeadline = locating.deadlineDate;
      const rowDeadline = isDate_(rowDeadlineCell) ? stripTime_(new Date(rowDeadlineCell)) : projectDeadline;

      const daysBefore = rowDeadline ? daysBefore_(endInclusive, rowDeadline) : null;
      const colorInfo = decideColor_(daysBefore);
      const titleForEvent = CONFIG.addEmojiPrefix ? `${colorInfo.emoji} ${rawTitle}` : rawTitle;

      const dependsOn = str_(row[headers['Depends On']]);
      const status    = str_(row[headers['Status']]);
      const notes     = str_(row[headers['Notes']]);

      let description = buildDescription_({ dependsOn, status, notes });
      if (rowDeadline) {
        description += (description ? '\n' : '') +
          `Deadline (${isDate_(rowDeadlineCell) ? 'Row' : 'Project'}): ${Utilities.formatDate(rowDeadline, tz, 'yyyy-MM-dd')}\n` +
          `Days before deadline: ${daysBefore}`;
      }

      // ✅ People-chips + plain emails
      const assigneeRaw  = row[headers['Assigned To (email)']];
      const assigneeRich = richValues ? richValues[r][headers['Assigned To (email)']] : null;
      const guestEmails  = dedupEmails_(extractEmailsFromCell_(str_(assigneeRaw), assigneeRich));

      const existingEventId = row[eventIdColIndex] ? String(row[eventIdColIndex]).trim() : '';

      // Create/Update calendar event
      const event = upsertAllDayEvent_({
        cal,
        existingEventId,
        title: titleForEvent,
        start,
        endExclusive,
        description,
        guestEmails,
        colorInfo
      });

      // Save event id & stamp
      sheet.getRange(r + 1, eventIdColIndex + 1).setValue(event.getId());
      sheet.getRange(r + 1, lastSyncColIndex + 1).setValue(new Date());

      // Per-assignee Google Tasks (DWD)
      const existingMap = parseJsonSafe_(row[taskIdsColIndex]) || {};
      if (CONFIG.createGoogleTasks && guestEmails.length) {
        const updatedMap = upsertTasksForAssignees_({
          assignees: guestEmails,
          baseTitle: rawTitle, // No emoji in personal task
          dueDate: endInclusive,
          sheetName: sheet.getName(),
          eventId: event.getId(),
          existingMap
        });
        sheet.getRange(r + 1, taskIdsColIndex + 1).setValue(JSON.stringify(updatedMap));
      }

      log_('upsert', r + 1, rawTitle, event.getId(), `color=${colorInfo.name}; guests=${guestEmails.join(',')}`);
    } catch (e) {
      log_('error', r + 1, str_(values[r][headers['Task']]), '', e && e.message ? e.message : String(e));
    }
  }

  SpreadsheetApp.getActive().toast('Calendar sync complete.');
}

/***** EVENT UPSERT (invites, reminders, color, sendUpdates) *****/
function upsertAllDayEvent_({ cal, existingEventId, title, start, endExclusive, description, guestEmails, colorInfo }) {
  let event = existingEventId ? cal.getEventById(existingEventId) : null;

  if (!event) {
    const options = {
      description,
      guests: guestEmails.join(','),
      sendInvites: CONFIG.sendInvites
    };
    event = cal.createAllDayEvent(title, start, endExclusive, options);
  } else {
    event.setAllDayDates(start, endExclusive);
    event.setTitle(title);
    event.setDescription(description);
    const current = new Set(event.getGuestList().map(g => g.getEmail().toLowerCase()));
    guestEmails.forEach(e => { if (!current.has(e.toLowerCase())) event.addGuest(e); });
  }

  event.removeAllReminders();
  (CONFIG.defaultRemindersMins || []).forEach(m => event.addPopupReminder(m));
  event.setColor(colorInfo.calColor);

  // send update emails (Advanced Calendar API)
  if (CONFIG.sendUpdatesOnChange) {
    try {
      const body = { summary: event.getTitle(), description: event.getDescription() };
      Calendar.Events.patch(body, CONFIG.calendarId, event.getId().split('@')[0], { sendUpdates: 'all' });
    } catch (_) { /* Advanced service not enabled or not permitted; safe to ignore */ }
  }

  return event;
}

/***** TASKS (DWD) *****/
function upsertTasksForAssignees_({ assignees, baseTitle, dueDate, sheetName, eventId, existingMap }) {
  const updated = { ...(existingMap || {}) };
  const listName = CONFIG.tasksListNameTemplate.replace('{{SHEET_NAME}}', sheetName || 'Project');

  assignees.forEach(email => {
    try {
      const token = getSATokenForUser_(email, 'https://www.googleapis.com/auth/tasks');
      const listId = ensureTasksListForUser_(token, listName);
      const notes = `Linked calendar event:\nhttps://calendar.google.com/calendar/u/0/r/eventedit/${encodeURIComponent(eventId)}`;
      const dueISO = toDueISO_(dueDate);

      const existingTaskId = updated[email];
      if (existingTaskId) {
        const url = `https://tasks.googleapis.com/tasks/v1/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(existingTaskId)}`;
        const res = UrlFetchApp.fetch(url, {
          method: 'patch',
          headers: { Authorization: `Bearer ${token}` },
          contentType: 'application/json',
          muteHttpExceptions: true,
          payload: JSON.stringify({ title: baseTitle, notes, due: dueISO, status: 'needsAction' })
        });
        if (res.getResponseCode() === 404) {
          updated[email] = insertTask_(token, listId, baseTitle, notes, dueDate);
        } else if (res.getResponseCode() !== 200) {
          log_('task_error', 0, baseTitle, '', `PATCH ${email}: ${res.getResponseCode()} ${res.getContentText()}`);
        }
      } else {
        updated[email] = insertTask_(token, listId, baseTitle, notes, dueDate);
      }
    } catch (e) {
      log_('task_error', 0, baseTitle, '', `${email}: ${e && e.message ? e.message : e}`);
    }
  });

  return updated;
}

function insertTask_(token, listId, title, notes, dueDate) {
  const dueISO = toDueISO_(dueDate);
  const res = UrlFetchApp.fetch(
    `https://tasks.googleapis.com/tasks/v1/lists/${encodeURIComponent(listId)}/tasks`,
    {
      method: 'post',
      headers: { Authorization: `Bearer ${token}` },
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({ title, notes, due: dueISO, status: 'needsAction' })
    }
  );
  if (res.getResponseCode() !== 200) {
    throw new Error(`INSERT task failed: ${res.getResponseCode()} ${res.getContentText()}`);
  }
  return JSON.parse(res.getContentText()).id;
}

function ensureTasksListForUser_(token, listName) {
  const listsRes = UrlFetchApp.fetch('https://tasks.googleapis.com/tasks/v1/users/@me/lists?maxResults=100', {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });
  if (listsRes.getResponseCode() !== 200) {
    throw new Error(`LIST lists failed: ${listsRes.getResponseCode()} ${listsRes.getContentText()}`);
  }
  const items = (JSON.parse(listsRes.getContentText()).items || []);
  const found = items.find(l => l.title === listName);
  if (found) return found.id;

  const createRes = UrlFetchApp.fetch('https://tasks.googleapis.com/tasks/v1/users/@me/lists', {
    method: 'post',
    headers: { Authorization: `Bearer ${token}` },
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({ title: listName })
  });
  if (createRes.getResponseCode() !== 200) {
    throw new Error(`CREATE list failed: ${createRes.getResponseCode()} ${createRes.getContentText()}`);
  }
  return JSON.parse(createRes.getContentText()).id;
}

/***** SERVICE ACCOUNT JWT (DWD impersonation) *****/
function getSATokenForUser_(userEmail, scope) {
  const props = PropertiesService.getScriptProperties();
  const clientEmail = props.getProperty('SA_CLIENT_EMAIL');
  let privateKey = props.getProperty('SA_PRIVATE_KEY'); // can be real PEM or one-line with \n
  if (!clientEmail || !privateKey) {
    throw new Error('Missing SA_CLIENT_EMAIL or SA_PRIVATE_KEY in Script Properties');
  }

  // Normalize escaped newlines if the key was stored as a single line
  // e.g., "-----BEGIN...-----\nMIIE...==\n-----END PRIVATE KEY-----\n"
  privateKey = privateKey.replace(/\\n/g, '\n');

  const now = Math.floor(Date.now() / 1000);
  const header = { alg: 'RS256', typ: 'JWT' };
  const claim = {
    iss: clientEmail,
    sub: userEmail,
    scope: scope,
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600
  };

  const enc = o => Utilities.base64EncodeWebSafe(JSON.stringify(o)).replace(/=+$/,'');
  const unsigned = `${enc(header)}.${enc(claim)}`;
  const signature = Utilities.base64EncodeWebSafe(
    Utilities.computeRsaSha256Signature(unsigned, privateKey)
  ).replace(/=+$/,'');
  const assertion = `${unsigned}.${signature}`;

  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: assertion
    },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error(`Token exchange failed: ${res.getResponseCode()} ${res.getContentText()}`);
  }

  return JSON.parse(res.getContentText()).access_token;
}

/***** ADMIN DIRECTORY LOOKUP (DWD via SA) *****/
function getAdminDirectoryToken_() {
  const admin = PropertiesService.getScriptProperties().getProperty('ADMIN_IMPERSONATE_EMAIL');
  if (!admin) throw new Error('Missing ADMIN_IMPERSONATE_EMAIL property');
  return getSATokenForUser_(admin, 'https://www.googleapis.com/auth/admin.directory.user.readonly');
}

/**
 * Resolve one or more display names (comma/semicolon separated) to primary emails
 * via Admin Directory search. Prefers exact full-name match; falls back to first result.
 */
function findDirectoryEmailsByName_(displayText) {
  const names = (displayText || '')
    .split(/[,;]+/).map(s => s.trim()).filter(Boolean);
  if (!names.length) return [];

  const token = getAdminDirectoryToken_();
  const out = new Set();

  names.forEach(name => {
    const url = 'https://admin.googleapis.com/admin/directory/v1/users' +
                '?maxResults=5&viewType=domain_public' +
                '&fields=users(primaryEmail,name/fullName),nextPageToken' +
                '&q=' + encodeURIComponent(`name:${name}`);

    try {
      const res = UrlFetchApp.fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
        muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) {
        log_('dir_warn', 0, name, '', `HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
        return;
      }
      const users = (JSON.parse(res.getContentText()).users || []);
      let match = users.find(u => ((u.name && u.name.fullName) || '').toLowerCase() === name.toLowerCase());
      if (!match && users.length) match = users[0];
      if (match && match.primaryEmail) out.add(match.primaryEmail.toLowerCase());
      else log_('dir_warn', 0, name, '', 'No user match');
    } catch (e) {
      log_('dir_error', 0, name, '', String(e));
    }
  });

  return Array.from(out);
}

/***** DEADLINE LOCATOR *****/
function locateDeadline_(sheet, headers, valuesOpt) {
  const values = valuesOpt || sheet.getDataRange().getValues();
  const taskCol = headers['Task'];
  const startCol = headers['Start Date'];
  const endCol = headers['End Date'];

  let deadlineRow = null;
  for (let i = CONFIG.headerRow; i < values.length; i++) {
    const cell = str_(values[i][taskCol]);
    if (equalsCI_(cell, CONFIG.deadlineTaskLabel)) { deadlineRow = i + 1; break; } // 1-based
  }

  let firstTaskRow = CONFIG.headerRow + 1;
  let deadlineDate = null;
  let deadlineRange = null;

  if (deadlineRow) {
    firstTaskRow = Math.max(firstTaskRow, deadlineRow + 1);
    const endVal = values[deadlineRow - 1][endCol];
    const startVal = values[deadlineRow - 1][startCol];
    deadlineDate = isDate_(endVal) ? stripTime_(new Date(endVal)) :
                   isDate_(startVal) ? stripTime_(new Date(startVal)) : null;

    const colIndexForRange = isDate_(endVal) ? endCol : startCol;
    if (deadlineDate != null) deadlineRange = sheet.getRange(deadlineRow, colIndexForRange + 1, 1, 1);
  }

  return { deadlineRow, firstTaskRow, deadlineDate, deadlineRange };
}

/***** NAMED RANGE *****/
function ensureDeadlineNamedRange_(range) {
  if (!range) return;
  const ss = SpreadsheetApp.getActive();
  const name = 'DEADLINE';
  const existing = ss.getNamedRanges().find(nr => nr.getName() === name);
  if (existing) existing.setRange(range);
  else ss.setNamedRange(name, range);
}

/***** CONDITIONAL FORMATTING (End Date with per-row Deadline) *****/
function ensureEndDateConditionalFormatting_(sheet, headers, firstTaskRow) {
  ensureColumn_(sheet, headers, CONFIG.deadlineHeader);

  const endColIdx1 = headers['End Date'] + 1;           // 1-based
  const dlColIdx1  = headers[CONFIG.deadlineHeader] + 1;
  const endColLtr  = colLetter_(endColIdx1);
  const dlColLtr   = colLetter_(dlColIdx1);

  const maxRow = sheet.getMaxRows();
  const range = sheet.getRange(firstTaskRow, endColIdx1, maxRow - firstTaskRow + 1, 1);

  const r = firstTaskRow; // anchor row
  const rowDl = `IF($${dlColLtr}${r}="", DEADLINE, $${dlColLtr}${r})`;
  const y = CONFIG.yellowWindowDays;
  const yellowUpperOp = CONFIG.onDeadlineIsRed ? '<' : '<=';
  const redOp = CONFIG.onDeadlineIsRed ? '>=' : '>';

  const greenFormula  = `=AND($${endColLtr}${r}<>"", $${endColLtr}${r} < (${rowDl} - ${y}))`;
  const yellowFormula = `=AND($${endColLtr}${r}<>"", $${endColLtr}${r} >= (${rowDl} - ${y}), $${endColLtr}${r} ${yellowUpperOp} ${rowDl})`;
  const redFormula    = `=AND($${endColLtr}${r}<>"", $${endColLtr}${r} ${redOp} ${rowDl})`;

  const existing = sheet.getConditionalFormatRules() || [];
  const filtered = existing.filter(rule => {
    const rgn = rule.getRanges ? rule.getRanges() : [];
    const same = rgn.some(g =>
      g.getA1Notation() === range.getA1Notation() && g.getSheet().getSheetId() === sheet.getSheetId()
    );
    return !same;
  });

  const green = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(greenFormula).setBackground('#34a853').setRanges([range]).build();
  const yellow = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(yellowFormula).setBackground('#fbbc04').setRanges([range]).build();
  const red = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(redFormula).setBackground('#ea4335').setRanges([range]).build();

  sheet.setConditionalFormatRules(filtered.concat([green, yellow, red]));
}

/***** UTIL & LOG *****/
function getSheetAndHeaders_(opts = {}) {
  const ss = SpreadsheetApp.getActive();
  const sheet = CONFIG.sheetName ? ss.getSheetByName(CONFIG.sheetName) : ss.getActiveSheet();
  if (!sheet) throw new Error('Sheet not found.');

  const headerRange = sheet.getRange(CONFIG.headerRow, 1, 1, sheet.getLastColumn());
  const headerVals = headerRange.getValues()[0];

  const headers = {};
  headerVals.forEach((h, i) => { const key = str_(h); if (key) headers[key] = i; });

  ['Task','Depends On','Start Date','Duration (days)','End Date','Assigned To (email)','Status','Notes',
   CONFIG.deadlineHeader, CONFIG.taskIdsHeader, CONFIG.eventIdHeader, CONFIG.lastSyncedHeader]
    .forEach(n => { if (headers[n] == null) headers[n] = ensureColumn_(sheet, headers, n); });

  const dataRange = sheet.getDataRange();
  const values = opts.includeValues ? dataRange.getValues() : null;
  const richValues = opts.includeValues ? dataRange.getRichTextValues() : null;   //  👈 add this

  return { sheet, headers, values, richValues };                                    //  👈 and return it
}

function ensureColumn_(sheet, headers, name) {
  if (headers[name] != null) return headers[name];
  const lastCol = sheet.getLastColumn();
  sheet.getRange(CONFIG.headerRow, lastCol + 1).setValue(name);
  headers[name] = lastCol; // zero-based
  return headers[name];
}

// Returns display text for plain values AND Smart Chips (people chips)
function str_(v) {
  if (v == null) return '';
  if (typeof v === 'object') {
    try {
      if (typeof v.getText === 'function') return String(v.getText()).trim(); // RichTextValue
    } catch (_) {}
  }
  return String(v).trim();
}

function extractEmailsFromCell_(rawText, rich) {
  const out = new Set();

  // Try RichText runs (chips/links)
  if (rich && typeof rich.getText === 'function') {
    try {
      if (typeof rich.getRuns === 'function') {
        const runs = rich.getRuns() || [];
        runs.forEach(run => {
          try {
            const url = run.getLinkUrl && run.getLinkUrl();
            if (url && url.indexOf('mailto:') === 0) out.add(url.replace(/^mailto:/, '').toLowerCase());
            const t = run.getText && run.getText();
            (parseEmails_(t) || []).forEach(e => out.add(e.toLowerCase()));
          } catch (_) {}
        });
      } else {
        const url = rich.getLinkUrl && rich.getLinkUrl();
        if (url && url.indexOf('mailto:') === 0) out.add(url.replace(/^mailto:/, '').toLowerCase());
        (parseEmails_(rich.getText()) || []).forEach(e => out.add(e.toLowerCase()));
      }
    } catch (_) {}
  }

  // Parse plain text fallback (comma/semicolon/space separated)
  (parseEmails_(rawText) || []).forEach(e => out.add(e.toLowerCase()));

  // If still empty, attempt resolving display names via Admin Directory
  if (out.size === 0 && rich && typeof rich.getText === 'function') {
    const display = (rich.getText && rich.getText()) || String(rawText || '');
    findDirectoryEmailsByName_(display).forEach(e => out.add(e));
  }

  return Array.from(out);
}

function equalsCI_(a,b){ return a.toLowerCase() === b.toLowerCase(); }
function isDate_(v){ return v && Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v); }
function toDate_(v, tz){ if (isDate_(v)) return stripTime_(new Date(v)); const d=new Date(v); return isNaN(d)?null:stripTime_(d); }
function stripTime_(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays_(d, n){ const x=new Date(d); x.setDate(x.getDate()+n); return stripTime_(x); }
function daysBefore_(endInclusive, deadline){ const e=stripTime_(endInclusive); const dl=stripTime_(deadline); return Math.floor((dl - e)/(1000*60*60*24)); }
function buildDescription_({ dependsOn, status, notes }) { const p=[]; if (dependsOn) p.push(`Depends On: ${dependsOn}`); if (status) p.push(`Status: ${status}`); if (notes) p.push(`Notes: ${notes}`); return p.join('\n'); }
function parseEmails_(s){ if (!s) return []; return s.split(/[,; ]+/).map(x=>x.trim()).filter(x=>/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x)); }
function colLetter_(n){ let s=''; while(n>0){const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26);} return s; }

function parseJsonSafe_(v){ try{ return v ? JSON.parse(v) : null; } catch { return null; } }

function toDueISO_(d) {
  // d is a date with local time cleared by your helpers; make an exact UTC midnight
  return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())).toISOString();
}

function dedupEmails_(arr) {
  const s = new Set();
  (arr || []).forEach(e => { if (e) s.add(e.toLowerCase()); });
  return Array.from(s);
}

function log_(action, row, task, eventId, note) {
  const ss = SpreadsheetApp.getActive();
  const name = 'Sync Log';
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,6).setValues([['When','Action','Row','Task','Event ID','Notes']]);
  sh.appendRow([new Date(), action, row, task, eventId || '', note || '']);
}
function openSyncLog_(){ const sh = SpreadsheetApp.getActive().getSheetByName('Sync Log'); if (sh) SpreadsheetApp.setActiveSheet(sh); }

/************ QUICK TESTS ************/

// 0) Test people chip parsing for a specific row
function test_AssigneeChipParsing(rowNumber1Based) {
  const { headers, values, richValues } = getSheetAndHeaders_({ includeValues: true });
  const r = (rowNumber1Based || (CONFIG.headerRow + 2)) - 1; // default: first task row
  const raw  = values[r][headers['Assigned To (email)']];
  const rich = richValues[r][headers['Assigned To (email)']];
  const emails = extractEmailsFromCell_(str_(raw), rich);
  Logger.log(JSON.stringify({ row: rowNumber1Based || (r + 1), emails }));
}

// 1) DWD test: list Tasks lists as the active user (or hardcode a user)
function test_TasksImpersonation() {
  const user = Session.getActiveUser().getEmail();
  const token = getSATokenForUser_(user, 'https://www.googleapis.com/auth/tasks');
  const r = UrlFetchApp.fetch('https://tasks.googleapis.com/tasks/v1/users/@me/lists', {
    headers: { Authorization: `Bearer ${token}` }
  });
  Logger.log(r.getResponseCode());
  Logger.log(r.getContentText());
}

// 2) DWD test: create a task in "Project: Test" due tomorrow
function test_CreateTaskForUser() {
  const user = Session.getActiveUser().getEmail();
  const token = getSATokenForUser_(user, 'https://www.googleapis.com/auth/tasks');

  const listName = 'Project: Test';
  const lists = JSON.parse(UrlFetchApp.fetch(
    'https://tasks.googleapis.com/tasks/v1/users/@me/lists?maxResults=100',
    { headers: { Authorization: `Bearer ${token}` } }
  ).getContentText()).items || [];
  let list = lists.find(l => l.title === listName);
  if (!list) {
    list = JSON.parse(UrlFetchApp.fetch(
      'https://tasks.googleapis.com/tasks/v1/users/@me/lists',
      {
        method: 'post',
        headers: { Authorization: `Bearer ${token}` },
        contentType: 'application/json',
        payload: JSON.stringify({ title: listName })
      }
    ).getContentText());
  }

  const due = new Date(); due.setDate(due.getDate() + 1);
  const task = JSON.parse(UrlFetchApp.fetch(
    `https://tasks.googleapis.com/tasks/v1/lists/${encodeURIComponent(list.id)}/tasks`,
    {
      method: 'post',
      headers: { Authorization: `Bearer ${token}` },
      contentType: 'application/json',
      payload: JSON.stringify({
        title: 'DWD test task',
        notes: 'Created via service account impersonation.',
        due: new Date(due.getFullYear(), due.getMonth(), due.getDate()).toISOString()
      })
    }
  ).getContentText());
  Logger.log(`Created task id: ${task.id}`);
}

// 3) Calendar test (as Apps Script user, not SA)
function test_CreateCalendarEvent() {
  const cal = CalendarApp.getCalendarById(CONFIG.calendarId);
  const start = new Date(); const end = new Date(start); end.setDate(end.getDate() + 1);
  const e = cal.createAllDayEvent('Calendar test event', start, end, { description: 'Hello' });
  Logger.log(e.getId());
}
