# 📋 Google Sheets Task Assignment with Calendar Sync

> Assign tasks using @Person mentions in Google Sheets and automatically sync to a shared calendar

## 🎯 What This Does

This Google Apps Script syncs tasks from Google Sheets to Google Calendar using the built-in `@Person` mentions (Smart Chips) for task assignment:

- **Creates calendar events** for each task on a shared calendar
- **Sends calendar invites** to people assigned via `@Person` mentions
- **Creates personal Google Tasks** for each assignee with task details
- **Tracks all changes** in a sync log for transparency

### 💡 Perfect For:

- Project managers coordinating team tasks
- Teams collaborating on projects with a shared calendar
- Anyone who wants spreadsheet tasks to sync to calendars automatically

---

## ✨ Key Features

### 📅 Calendar Event Creation
- Creates all-day events with task name, dates, and details
- Automatically invites all assigned people
- Customizable reminders (default: 24 hours and 2 hours before)

### 👥 Easy Assignment with @Person
Assign tasks by typing `@` and selecting a person:
- **People chips**: `@John Doe` (Google Sheets smart chips)
- **Email addresses**: `john@company.com`
- **Multiple assignees**: `@John Doe, @Jane Smith` or `john@company.com; jane@company.com`

### ✅ Personal Google Tasks
- Creates individual Google Tasks for each assignee
- Tasks include: task name, due date, and notes
- Links back to calendar events for context
- Automatically updates when sheet changes

### 📊 Full Transparency
- **Sync Log sheet**: Records every sync operation with timestamps
- Event IDs and Task IDs stored in the sheet for tracking
- Error logging for failed operations

---

## 🚀 Getting Started

### Step 1: Set Up Your Google Sheet

Create a sheet with these columns:

| Column Name | Description | Required | Example |
|------------|-------------|----------|---------|
| **Task** | Task name/description | ✅ Yes | "Design homepage mockup" |
| **Start Date** | When task begins | ✅ Yes | 2025-01-15 |
| **End Date** | When task ends | Recommended | 2025-01-20 |
| **Duration (days)** | Used if no End Date | Optional | 5 |
| **Assigned To (email)** | Who's responsible (use @Person) | Recommended | @John Doe |
| **Depends On** | Prerequisites | Optional | "Task 1, Task 2" |
| **Status** | Current state | Optional | "In Progress" |
| **Notes** | Additional details | Optional | "Needs approval" |

### Step 2: Install the Apps Script

1. Open your Google Sheet
2. Go to **Extensions** > **Apps Script**
3. Delete any existing code
4. Copy the entire contents of `code.gs` from this repository
5. Paste into the Apps Script editor
6. Click **Save** (💾 icon)

### Step 3: Configure Settings

#### A. Update Calendar ID

In the `CONFIG` section at the top of `code.gs`, update:

```javascript
calendarId: 'your-calendar-id@group.calendar.google.com',
```

**To find your Calendar ID:**
1. Open Google Calendar
2. Click the three dots next to the calendar you want to use
3. Select "Settings and sharing"
4. Scroll to "Integrate calendar" → Copy the Calendar ID

#### B. Set Up Service Account (For Tasks & Directory)

This script uses Google Workspace Domain-Wide Delegation to create personal tasks:

1. **Create a Service Account**:
   - Go to [Google Cloud Console](https://console.cloud.google.com)
   - Create a new project (or select existing)
   - Enable these APIs:
     - Google Calendar API
     - Google Tasks API  
     - Admin SDK API
   - Create a Service Account with Domain-Wide Delegation
   - Download the JSON key file

2. **Configure Domain-Wide Delegation**:
   - Go to [Google Workspace Admin Console](https://admin.google.com)
   - Security > API Controls > Domain-wide Delegation
   - Add your service account's Client ID with these scopes:
     ```
     https://www.googleapis.com/auth/calendar
     https://www.googleapis.com/auth/tasks
     https://www.googleapis.com/auth/admin.directory.user.readonly
     ```

3. **Add Script Properties**:
   - In Apps Script: **Project Settings** (⚙️ icon) > **Script Properties** > **Add script property**
   - Add these three properties:

   | Property | Value |
   |----------|-------|
   | `SA_CLIENT_EMAIL` | From your service account JSON: `client_email` |
   | `SA_PRIVATE_KEY` | From your service account JSON: `private_key` (entire key including `-----BEGIN PRIVATE KEY-----`) |
   | `ADMIN_IMPERSONATE_EMAIL` | A Google Workspace admin email (e.g., `admin@company.com`) |

#### C. Enable Advanced Services

In Apps Script editor:
1. Click **Services** (⊕ icon on left sidebar)
2. Find and add: **Admin Directory API**
3. Click **Add**

### Step 4: Test the Setup

Run these test functions from Apps Script to verify everything works:

1. **Test Service Account**: Run `test_ServiceAccountKey()`
2. **Test Directory API Access**: Run `test_DirectoryPing()`
3. **Test Tasks Access**: Run `test_TasksImpersonation()`

### Step 5: First Sync

1. Go back to your Google Sheet
2. You should now see a **"Project"** menu in the toolbar
3. Click **Project** > **Sync to Calendar**
4. Grant permissions when prompted
5. Wait for the "Calendar sync complete" notification

🎉 **Done!** Check your shared calendar and Google Tasks!

---

## 📖 How to Use

### Daily Workflow

1. **Add tasks** to your sheet with task name, dates, and notes
2. **Assign people** using `@Person` mentions in the "Assigned To (email)" column
3. **Run sync**: Click **Project** > **Sync to Calendar**
4. **Assignees automatically get**:
   - Calendar event invitations on the shared calendar
   - Personal tasks in their Google Tasks list
5. **Check the Sync Log** sheet for sync history

### Task Information Synced

Each task includes:
- **Task name** - The title of the event/task
- **Assigned person** - Invited to the calendar event
- **Due date** - End date of the task
- **Notes** - Included in event description and task notes
- **Status & Dependencies** - Added to event description

---

## ⚙️ Configuration Options

Edit the `CONFIG` object in `code.gs`:

```javascript
const CONFIG = {
  // Calendar settings
  calendarId: 'your-calendar@group.calendar.google.com',
  sendInvites: true,                    // Email guests on create
  sendUpdatesOnChange: true,            // Email guests on updates
  defaultRemindersMins: [1440, 120],    // 24h and 2h reminders
  
  // Google Tasks
  createGoogleTasks: true,              // Create personal tasks
  tasksListNameTemplate: 'Project: {{SHEET_NAME}}',
  
  // Sheet structure
  sheetName: null,                      // null = active sheet
  headerRow: 1                          // Row number with column headers
};
```

---

## 🧪 Testing & Troubleshooting

### Test Functions

| Function | Purpose |
|----------|---------|
| `test_ServiceAccountKey()` | Verify service account credentials |
| `test_DirectoryPing()` | Check Admin Directory API access |
| `test_TasksImpersonation()` | Test Google Tasks access |
| `test_AssigneeChipParsing(row)` | Test @Person parsing for a row |
| `test_CreateCalendarEvent()` | Create a test calendar event |

### Common Issues

**❌ "Missing SA_CLIENT_EMAIL or SA_PRIVATE_KEY"**
- Go to Apps Script > Project Settings > Script Properties
- Make sure all three properties are set correctly

**❌ "Calendar not found"**
- Verify `CONFIG.calendarId` matches your target calendar
- Make sure users are invited to the shared calendar

**❌ "Token exchange failed"**
- Verify Domain-Wide Delegation is set up
- Check that all OAuth scopes are authorized

**❌ "@Person not recognized"**
- The script will try to resolve names via Admin Directory
- If that fails, use email addresses directly

---

## 🔒 Privacy & Security

- **Service Account**: Has delegated access only to specified scopes
- **User Data**: Email addresses used only for calendar/task creation
- **Audit Trail**: Sync Log records all operations

---

## 📝 License

This project is provided as-is for use within Google Workspace organizations.

---

**Made with ❤️ for teams who love simple task assignment**
