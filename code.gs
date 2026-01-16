/*******************************************************
 * Project Tasks → Calendar + Google Tasks Sync
 * - Creates calendar events with task details
 * - Invites assignees using @Person mentions
 * - Creates personal Google Tasks via Domain-Wide Delegation
 * - Sync Log for tracking all operations
 *******************************************************/

/***** CONFIG *****/
const CONFIG = {
  calendarId: 'c_b697a80ae3fb9a9918f40ffdb570a633f9d2286db89a65123374589206aa3ea6@group.calendar.google.com',
  sheetName: null,                // null = active sheet
  headerRow: 1,
  eventIdHeader: 'Event ID',
  lastSyncedHeader: 'Last Synced',
  taskIdsHeader: 'Task IDs (JSON)',      // per-row: { "user@dom": "taskId", ... }

  // Calendar behavior
  sendInvites: true,                      // email guests on create
  sendUpdatesOnChange: true,              // email on updates (needs Advanced Calendar API)
  defaultRemindersMins: [1440, 120],      // 24h & 2h popup reminders

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
    .addItem('Open Sync Log', 'openSyncLog_')
    .addToUi();
}

/***** MAIN SYNC *****/
function syncTasksToCalendar() {
  const { sheet, headers, values, richValues } = getSheetAndHeaders_({ includeValues: true });

  // Ensure helper columns
  const eventIdColIndex  = ensureColumn_(sheet, headers, CONFIG.eventIdHeader);
  const lastSyncColIndex = ensureColumn_(sheet, headers, CONFIG.lastSyncedHeader);
  const taskIdsColIndex  = ensureColumn_(sheet, headers, CONFIG.taskIdsHeader);

  const cal = CalendarApp.getCalendarById(CONFIG.calendarId);
  if (!cal) throw new Error('Calendar not found. Check CONFIG.calendarId.');

  const tz = Session.getScriptTimeZone();
  const startR = CONFIG.headerRow + 1;

  for (let r = startR - 1; r < values.length; r++) {
    const row = values[r];
    const rawTitle = str_(row[headers['Task']]);
    if (!rawTitle) continue;

    try {
      const start = toDate_(row[headers['Start Date']], tz);
      if (!start) { log_('skip', r + 1, rawTitle, '', 'No start date'); continue; }

      const endInclusive = isDate_(row[headers['End Date']])
        ? toDate_(row[headers['End Date']], tz)
        : addDays_(start, Math.max(1, Number(row[headers['Duration (days)']] || 1)) - 1);

      const endExclusive = addDays_(endInclusive, 1);

      const dependsOn = str_(row[headers['Depends On']]);
      const status    = str_(row[headers['Status']]);
      const notes     = str_(row[headers['Notes']]);

      const description = buildDescription_({ dependsOn, status, notes });

      // Parse @Person mentions + plain emails
      const assigneeRaw  = row[headers['Assigned To (email)']];
      const assigneeRich = richValues ? richValues[r][headers['Assigned To (email)']] : null;
      const guestEmails  = dedupEmails_(extractEmailsFromCell_(str_(assigneeRaw), assigneeRich));

      const existingEventId = row[eventIdColIndex] ? String(row[eventIdColIndex]).trim() : '';

      // Create/Update calendar event
      const event = upsertAllDayEvent_({
        cal,
        existingEventId,
        title: rawTitle,
        start,
        endExclusive,
        description,
        guestEmails
      });

      // Save event id & stamp
      sheet.getRange(r + 1, eventIdColIndex + 1).setValue(event.getId());
      sheet.getRange(r + 1, lastSyncColIndex + 1).setValue(new Date());

      // Per-assignee Google Tasks (DWD)
      const existingMap = parseJsonSafe_(row[taskIdsColIndex]) || {};
      if (CONFIG.createGoogleTasks && guestEmails.length) {
        const updatedMap = upsertTasksForAssignees_({
          assignees: guestEmails,
          baseTitle: rawTitle,
          dueDate: endInclusive,
          sheetName: sheet.getName(),
          eventId: event.getId(),
          notes: notes,
          existingMap
        });
        sheet.getRange(r + 1, taskIdsColIndex + 1).setValue(JSON.stringify(updatedMap));
      }

      log_('upsert', r + 1, rawTitle, event.getId(), `guests=${guestEmails.join(',')}`);
    } catch (e) {
      log_('error', r + 1, str_(values[r][headers['Task']]), '', e && e.message ? e.message : String(e));
    }
  }

  SpreadsheetApp.getActive().toast('Calendar sync complete.');
}

/***** EVENT UPSERT (invites, reminders) *****/
function upsertAllDayEvent_({ cal, existingEventId, title, start, endExclusive, description, guestEmails }) {
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
function upsertTasksForAssignees_({ assignees, baseTitle, dueDate, sheetName, eventId, notes, existingMap }) {
  const updated = { ...(existingMap || {}) };
  const listName = CONFIG.tasksListNameTemplate.replace('{{SHEET_NAME}}', sheetName || 'Project');

  assignees.forEach(email => {
    try {
      const token = getSATokenForUser_(email, 'https://www.googleapis.com/auth/tasks');
      const listId = ensureTasksListForUser_(token, listName);
      const taskNotes = `${notes ? notes + '\n\n' : ''}Linked calendar event:\nhttps://calendar.google.com/calendar/u/0/r/eventedit/${encodeURIComponent(eventId)}`;
      const dueISO = toDueISO_(dueDate);

      const existingTaskId = updated[email];
      if (existingTaskId) {
        const url = `https://tasks.googleapis.com/tasks/v1/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(existingTaskId)}`;
        const res = UrlFetchApp.fetch(url, {
          method: 'patch',
          headers: { Authorization: `Bearer ${token}` },
          contentType: 'application/json',
          muteHttpExceptions: true,
          payload: JSON.stringify({ title: baseTitle, notes: taskNotes, due: dueISO, status: 'needsAction' })
        });
        if (res.getResponseCode() === 404) {
          updated[email] = insertTask_(token, listId, baseTitle, taskNotes, dueDate);
        } else if (res.getResponseCode() !== 200) {
          log_('task_error', 0, baseTitle, '', `PATCH ${email}: ${res.getResponseCode()} ${res.getContentText()}`);
        }
      } else {
        updated[email] = insertTask_(token, listId, baseTitle, taskNotes, dueDate);
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
  let privateKey = props.getProperty('SA_PRIVATE_KEY');
  if (!clientEmail || !privateKey) {
    throw new Error('Missing SA_CLIENT_EMAIL or SA_PRIVATE_KEY in Script Properties');
  }

  // Normalize escaped newlines if the key was stored as a single line
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
 * Normalize display text from a chip cell to a set of candidate names.
 */
function _normalizeDisplayNames_(displayText) {
  if (!displayText) return [];
  return displayText
    .replace(/@/g, ' ')
    .split(/[\n,;]+/)
    .map(s => s.trim().replace(/\s+/g, ' '))
    .filter(Boolean);
}

/**
 * Resolve one or more display names to emails via Admin Directory.
 */
function findDirectoryEmailsByName_(displayText) {
  const names = _normalizeDisplayNames_(displayText);
  if (!names.length) return [];

  const token = getAdminDirectoryToken_();
  const base = 'https://admin.googleapis.com/admin/directory/v1/users';
  const out = new Set();

  names.forEach(name => {
    const queries = [
      `name:"${name}"`,
      `name:${name}`
    ];

    let found = null;
    for (let i = 0; i < queries.length && !found; i++) {
      const url = `${base}?customer=my_customer&maxResults=5&viewType=admin_view`
                + `&fields=users(primaryEmail,name/fullName)`
                + `&query=${encodeURIComponent(queries[i])}`;
      try {
        const res = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
        if (res.getResponseCode() !== 200) {
          log_('dir_warn', 0, name, '', `HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
          continue;
        }
        const users = (JSON.parse(res.getContentText()).users || []);
        found = users.find(u => ((u.name && u.name.fullName) || '').toLowerCase() === name.toLowerCase()) || users[0];
      } catch (e) {
        log_('dir_error', 0, name, '', String(e));
      }
    }

    if (found && found.primaryEmail) out.add(found.primaryEmail.toLowerCase());
    else log_('dir_warn', 0, name, '', 'No user match');
  });

  return Array.from(out);
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
   CONFIG.taskIdsHeader, CONFIG.eventIdHeader, CONFIG.lastSyncedHeader]
    .forEach(n => { if (headers[n] == null) headers[n] = ensureColumn_(sheet, headers, n); });

  const dataRange = sheet.getDataRange();
  const values = opts.includeValues ? dataRange.getValues() : null;
  const richValues = opts.includeValues ? dataRange.getRichTextValues() : null;

  return { sheet, headers, values, richValues };
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

  // 1) Try RichText runs (chips/links) - this handles @Person mentions
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

  // 2) Parse plain text fallback (comma/semicolon/space separated)
  (parseEmails_(rawText) || []).forEach(e => out.add(e.toLowerCase()));

  // 3) If still empty, attempt resolving display names via Admin Directory
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
function buildDescription_({ dependsOn, status, notes }) { const p=[]; if (dependsOn) p.push(`Depends On: ${dependsOn}`); if (status) p.push(`Status: ${status}`); if (notes) p.push(`Notes: ${notes}`); return p.join('\n'); }
function parseEmails_(s){ if (!s) return []; return s.split(/[,; ]+/).map(x=>x.trim()).filter(x=>/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x)); }

function parseJsonSafe_(v){ try{ return v ? JSON.parse(v) : null; } catch { return null; } }

function toDueISO_(d) {
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
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,5).setValues([['When','Action','Row','Task','Event ID','Notes']]);
  sh.appendRow([new Date(), action, row, task, eventId || '', note || '']);
}
function openSyncLog_(){ const sh = SpreadsheetApp.getActive().getSheetByName('Sync Log'); if (sh) SpreadsheetApp.setActiveSheet(sh); }

/************ TEST FUNCTIONS ************/

// Test service account key validity
function test_ServiceAccountKey() {
  try {
    const props = PropertiesService.getScriptProperties();
    const clientEmail = props.getProperty('SA_CLIENT_EMAIL');
    const privateKey = props.getProperty('SA_PRIVATE_KEY');
    if (!clientEmail || !privateKey) {
      throw new Error('Missing SA_CLIENT_EMAIL or SA_PRIVATE_KEY in Script Properties');
    }

    Logger.log('Client Email: ' + clientEmail);

    const normalizedKey = privateKey.replace(/\\n/g, '\n');
    const hasEscaped = privateKey.indexOf('\\n') !== -1;
    const hasReal = normalizedKey.indexOf('\n') !== -1;

    Logger.log('Escaped \\n detected: ' + hasEscaped);
    Logger.log('Contains real newlines: ' + hasReal);
    Logger.log('Key length: ' + normalizedKey.length);

    const token = getSATokenForUser_(Session.getActiveUser().getEmail(), 'https://www.googleapis.com/auth/tasks');
    Logger.log('Access token retrieved successfully (first 100 chars): ' + token.substring(0, 100));

    Logger.log('✅ Service account key works and token exchange succeeded.');
  } catch (e) {
    Logger.log('❌ Service account key test failed: ' + e.message);
  }
}

// Test Admin Directory API access
function test_DirectoryPing() {
  const token = getAdminDirectoryToken_();
  const url = 'https://admin.googleapis.com/admin/directory/v1/users'
            + '?customer=my_customer&maxResults=5&viewType=admin_view'
            + '&fields=users(primaryEmail,name/fullName)';
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
  Logger.log(res.getResponseCode());
  Logger.log(res.getContentText());
}

// Test @Person chip parsing for a specific row
function test_AssigneeChipParsing(rowNumber1Based) {
  const { headers, values, richValues } = getSheetAndHeaders_({ includeValues: true });
  const r = (rowNumber1Based || (CONFIG.headerRow + 2)) - 1;
  const raw  = values[r][headers['Assigned To (email)']];
  const rich = richValues[r][headers['Assigned To (email)']];
  const emails = extractEmailsFromCell_(str_(raw), rich);
  Logger.log(JSON.stringify({ row: rowNumber1Based || (r + 1), emails }));
}

// Test directory lookup by name
function test_DirectoryLookupByName() {
  const emails = findDirectoryEmailsByName_('Chad Stolle; Spencer Lott');
  Logger.log(JSON.stringify(emails));
}

// Test Tasks impersonation
function test_TasksImpersonation() {
  const user = Session.getActiveUser().getEmail();
  const token = getSATokenForUser_(user, 'https://www.googleapis.com/auth/tasks');
  const r = UrlFetchApp.fetch('https://tasks.googleapis.com/tasks/v1/users/@me/lists', {
    headers: { Authorization: `Bearer ${token}` }
  });
  Logger.log(r.getResponseCode());
  Logger.log(r.getContentText());
}

// Test creating a task
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
        title: 'Test task',
        notes: 'Created via service account impersonation.',
        due: new Date(due.getFullYear(), due.getMonth(), due.getDate()).toISOString()
      })
    }
  ).getContentText());
  Logger.log(`Created task id: ${task.id}`);
}

// Test calendar event creation
function test_CreateCalendarEvent() {
  const cal = CalendarApp.getCalendarById(CONFIG.calendarId);
  const start = new Date(); const end = new Date(start); end.setDate(end.getDate() + 1);
  const e = cal.createAllDayEvent('Calendar test event', start, end, { description: 'Hello' });
  Logger.log(e.getId());
}
