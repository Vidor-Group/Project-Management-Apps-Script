# Project-Management-Apps-Script
Use this apps script to create project management sheets for timelines of large projects

## Features

- 📅 **Calendar Sync**: Creates color-coded calendar events (🔴/🟡/🟢) based on deadline proximity
- ✅ **Google Tasks**: Per-assignee personal tasks via Domain-Wide Delegation (DWD)
- 👥 **People Chips Support**: Automatically resolves People chips, names, and emails to primary email addresses
- ⏰ **Deadline Management**: Project-wide and per-row deadlines with conditional formatting
- 📊 **Sync Logging**: Maintains audit trail of all sync operations

## People Chips / Names in "Assigned To (email)"

You can type emails, People chips (@First Last), or just names (First Last).

The script will resolve to emails by:

1. **Extracting mailto: links** from rich text runs (when available)
2. **Parsing plain emails** in the text
3. **If none found**, querying Admin Directory (domain-wide delegation required)

### Admin Directory Setup

To enable name-to-email resolution:

1. **Domain-Wide Delegation Scope**: Add `https://www.googleapis.com/auth/admin.directory.user.readonly`
2. **Script Property**: Set `ADMIN_IMPERSONATE_EMAIL` to a super-admin (or user with Directory read access)
3. **Enable Admin Directory API**: 
   - In Apps Script: Resources > Advanced Google Services > Admin Directory API (ON)
   - In GCP Console: Enable Admin SDK API

### Behavior Notes

- ✅ **Internal domain users**: Receive both calendar invites AND personal Google Tasks
- ⚠️ **External users** (non-domain): Receive calendar invites only (cannot be impersonated for Tasks)
- 📝 **Sync Log**: Shows directory warnings only when a name can't be found

## Setup

### Required Script Properties

Set these in Project Settings > Script Properties:

```
SA_CLIENT_EMAIL = your-service-account@project.iam.gserviceaccount.com
SA_PRIVATE_KEY = -----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----
ADMIN_IMPERSONATE_EMAIL = admin@yourdomain.com
```

### Required Scopes

Your service account needs Domain-Wide Delegation with these scopes:

- `https://www.googleapis.com/auth/calendar` (Calendar access)
- `https://www.googleapis.com/auth/tasks` (Tasks impersonation)
- `https://www.googleapis.com/auth/admin.directory.user.readonly` (Directory lookup)

### Calendar Configuration

Update `CONFIG.calendarId` in `code.gs` with your target calendar ID.

## Usage

1. **Prepare your sheet** with columns: Task, Start Date, End Date, Duration (days), Assigned To (email), Status, Notes
2. **Add Project Deadline row** (optional): Set Task = "Project Deadline" and End Date
3. **Run sync**: Project menu > Sync to Calendar
4. **Refresh formatting**: Project menu > Refresh Deadline & Formatting

## Testing

Run these functions from the Apps Script editor:

- `test_AssigneeChipParsing(rowNumber)` - Test email extraction for a specific row
- `test_DirectoryLookupByName()` - Test Admin Directory name resolution
- `test_ServiceAccountKey()` - Verify service account credentials
- `test_TasksImpersonation()` - Test DWD for Tasks
- `test_CreateTaskForUser()` - Test task creation
- `test_CreateCalendarEvent()` - Test calendar event creation
