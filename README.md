# 📊 Google Sheets Project Management with Calendar & Tasks Sync

> Transform Google Sheets into a powerful project management tool with automatic calendar events and personal task assignments

## 🎯 What This Does

This Google Apps Script turns any Google Sheet into a complete project management system by automatically:

- **Creating calendar events** for each task with color-coded urgency indicators (🟢 on track, 🟡 approaching deadline, 🔴 overdue)
- **Sending calendar invites** to all team members assigned to tasks
- **Creating personal Google Tasks** for each team member with their assigned work
- **Managing project deadlines** with automatic conditional formatting
- **Tracking all changes** in a sync log for full transparency

### 💡 Perfect For:

- Project managers coordinating team timelines
- Department heads managing cross-functional initiatives  
- Teams collaborating on multi-phase projects
- Anyone who wants their spreadsheet tasks to automatically sync to calendars and task lists

---

## ✨ Key Features

### 📅 Smart Calendar Integration
- Automatic event creation with task details, dependencies, and status
- **Color-coded by urgency**: 
  - 🟢 **Green**: More than 2 days before deadline
  - 🟡 **Yellow**: Within 2 days of deadline
  - 🔴 **Red**: At or past deadline
- Email notifications to assignees when events are created or updated
- Customizable reminders (default: 24 hours and 2 hours before)

### ✅ Personal Task Management
- Creates individual Google Tasks for each team member
- Tasks appear in each person's Google Tasks list
- Links back to calendar events for full context
- Automatically updates when sheet changes

### 👥 Flexible Assignee Input
You can assign tasks using any of these methods:
- **Email addresses**: `john@company.com`
- **People chips**: `@John Doe` (Google Sheets smart chips)
- **Display names**: `John Doe` (automatically looks up email via Google Directory)
- **Multiple assignees**: `john@company.com; jane@company.com` or `@John Doe, @Jane Smith`

### ⏰ Deadline Management
- Set a project-wide deadline OR per-task deadlines
- Automatic conditional formatting on End Date column
- Visual deadline proximity indicators
- Named range support for easy formula referencing

### 📊 Full Transparency
- **Sync Log sheet**: Records every sync operation with timestamps
- Event IDs and Task IDs stored in the sheet for tracking
- Error logging for failed operations
- Warnings for unresolvable names or emails

---

## 🚀 Getting Started

### Step 1: Set Up Your Google Sheet

Create a sheet with these columns (or use your existing project sheet):

| Column Name | Description | Required | Example |
|------------|-------------|----------|---------|
| **Task** | Task name/description | ✅ Yes | "Design homepage mockup" |
| **Start Date** | When task begins | ✅ Yes | 2025-01-15 |
| **End Date** | When task ends | Recommended | 2025-01-20 |
| **Duration (days)** | Used if no End Date | Optional | 5 |
| **Assigned To (email)** | Who's responsible | Recommended | john@company.com or @John Doe |
| **Depends On** | Prerequisites | Optional | "Task 1, Task 2" |
| **Status** | Current state | Optional | "In Progress" |
| **Notes** | Additional details | Optional | "Needs approval from design team" |
| **Deadline** | Task-specific deadline | Optional | 2025-01-25 |

**Optional Project Deadline**: Add a row with Task = "Project Deadline" and set the End Date. This becomes the default deadline for all tasks.

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

This script uses Google Workspace Domain-Wide Delegation to create personal tasks. You'll need:

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
   - Check **Execution log** for ✅ success message
2. **Test Directory API Access**: Run `test_DirectoryPing()`
   - Should return HTTP 200 and list of users
   - If 403, enable Admin SDK API in GCP and Advanced Service in Apps Script
3. **Test Tasks Access**: Run `test_TasksImpersonation()`
   - Should list your Google Tasks lists
4. **Test Directory Lookup**: Run `test_DirectoryLookupByName()`
   - Update with actual names from your organization first

### Step 5: First Sync

1. Go back to your Google Sheet
2. You should now see a **"Project"** menu in the toolbar
3. Click **Project** > **Sync to Calendar**
4. Grant permissions when prompted
5. Wait for the "Calendar sync complete" notification

🎉 **Done!** Check your calendar and Google Tasks to see your project tasks!

---

## 📖 How to Use

### Daily Workflow

1. **Update your sheet** with new tasks, dates, or assignees
2. **Run sync**: Click **Project** > **Sync to Calendar**
3. **Team members automatically get**:
   - Calendar event invitations
   - Personal tasks in their Google Tasks list
4. **Check the Sync Log** sheet for detailed sync history

### Understanding the Color Codes

The script automatically assigns colors based on how close tasks are to their deadline:

- 🟢 **Green (On Track)**: Task ends more than 2 days before deadline
- 🟡 **Yellow (Approaching)**: Task ends within 2 days of deadline  
- 🔴 **Red (Urgent)**: Task is at or past deadline

**Customize**: Change `yellowWindowDays` and `onDeadlineIsRed` in the `CONFIG` section.

### Conditional Formatting

The End Date column automatically highlights based on proximity to deadline:
- Green background = safe
- Yellow background = approaching
- Red background = overdue

Refresh formatting: **Project** > **Refresh Deadline & Formatting**

### Managing Assignees

**Multiple ways to assign tasks:**

```
Simple email:           john@company.com
Multiple emails:        john@company.com; jane@company.com
People chips:           @John Doe
Mix of both:            @John Doe, jane@company.com
Display names only:     John Doe; Jane Smith
```

The script intelligently resolves all of these to email addresses.

### Per-Task Deadlines

If you need different deadlines for specific tasks:
1. The **Deadline** column is automatically created on first sync
2. Enter a date in this column for any task
3. That task will use its own deadline instead of the project deadline

---

## ⚙️ Configuration Options

Edit the `CONFIG` object in `code.gs` to customize behavior:

```javascript
const CONFIG = {
  // Calendar settings
  calendarId: 'your-calendar@group.calendar.google.com',
  sendInvites: true,                    // Email guests on create
  sendUpdatesOnChange: true,            // Email guests on updates
  defaultRemindersMins: [1440, 120],    // 24h and 2h reminders
  addEmojiPrefix: true,                 // Add 🟢🟡🔴 to event titles
  
  // Color thresholds
  yellowWindowDays: 2,                  // Yellow within X days of deadline
  onDeadlineIsRed: true,                // Red on deadline day (vs. after)
  
  // Google Tasks
  createGoogleTasks: true,              // Create personal tasks
  tasksListNameTemplate: 'Project: {{SHEET_NAME}}',
  
  // Sheet structure
  sheetName: null,                      // null = active sheet
  headerRow: 1,                         // Row number with column headers
  deadlineTaskLabel: 'Project Deadline' // Task name for project deadline row
};
```

---

## 🧪 Testing & Troubleshooting

### Test Functions

Run these from the Apps Script editor to diagnose issues:

| Function | Purpose |
|----------|---------|
| `test_ServiceAccountKey()` | Verify service account credentials are valid |
| `test_TasksImpersonation()` | Test access to Google Tasks for current user |
| `test_CreateTaskForUser()` | Create a test task due tomorrow |
| `test_CreateCalendarEvent()` | Create a test calendar event |
| `test_AssigneeChipParsing(rowNumber)` | Test email extraction for a specific row |
| `test_DirectoryLookupByName()` | Test name-to-email resolution |

### Common Issues

**❌ "Missing SA_CLIENT_EMAIL or SA_PRIVATE_KEY"**
- Go to Apps Script > Project Settings > Script Properties
- Make sure all three properties are set correctly
- Check that private key includes the full `-----BEGIN PRIVATE KEY-----` header

**❌ "Calendar not found"**
- Verify `CONFIG.calendarId` matches your target calendar
- Make sure the service account has been invited to the calendar with "Make changes to events" permission

**❌ "Token exchange failed"**
- Verify Domain-Wide Delegation is set up in Google Workspace Admin
- Check that all three OAuth scopes are authorized
- Confirm `ADMIN_IMPERSONATE_EMAIL` is a valid admin user

**❌ "No user match" in Sync Log**
- The display name couldn't be found in Google Directory
- Try using email addresses or People chips instead
- Verify the person exists in your Google Workspace organization

**❌ Tasks not creating**
- External users (outside your domain) cannot receive Google Tasks
- They will still get calendar invites
- Check Sync Log for "task_error" entries

**⚠️ Sync seems slow**
- Normal for large sheets (100+ rows)
- The script makes API calls for each assignee and task
- Consider breaking very large projects into multiple sheets

---

## 🔒 Privacy & Security

- **Service Account**: Has delegated access only to specified scopes
- **User Data**: Email addresses and names are only used for calendar/task creation
- **Audit Trail**: Sync Log records all operations for transparency
- **External Users**: Cannot be impersonated (by design) - only receive calendar invites

---

## 🛠️ Advanced Customization

### Filtering External Users

To skip Google Tasks creation for external users, add this helper:

```javascript
function isManagedDomainEmail_(email) {
  return /@yourdomain\.com$/i.test(email); // Change to your domain
}
```

Then in `syncTasksToCalendar()` after extracting emails:

```javascript
const managedAssignees = guestEmails.filter(isManagedDomainEmail_);
if (CONFIG.createGoogleTasks && managedAssignees.length) {
  // Use managedAssignees instead of guestEmails
}
```

### Batch Performance Optimization

For very large sheets, replace individual cell writes with batch operations at the end of the loop. Collect values in arrays and call `setValues()` once.

### Custom Color Schemes

Modify the `decideColor_()` function to use different colors or thresholds based on your preferences.

---

## 📝 License

This project is provided as-is for use within Google Workspace organizations. Modify and adapt as needed for your team.

---

## 🤝 Contributing

Found a bug or have a feature request? Contributions and feedback are welcome!

---

## 📞 Support

For issues specific to:
- **Google Workspace setup**: Contact your organization's Google Workspace admin
- **Apps Script errors**: Check the Execution log in Apps Script editor
- **Script functionality**: Review the Sync Log sheet in your Google Sheet

---

**Made with ❤️ for project managers who love spreadsheets but need more automation**
