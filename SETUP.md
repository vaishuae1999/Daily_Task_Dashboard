# Google setup

## 1. Create the master and backup location

In the data owner's Google Drive, create a **new blank Google Sheet** named Daily Support Master and a private folder named Daily Support Backups. Keep the original Excel workbook unchanged. This master is a new dataset; the preparation/import step copies only reviewed records into it.

The Sheet ID is the part between `/d/` and `/edit` in its browser URL. The folder ID is the part following `/folders/`. These IDs are configuration values, not authentication tokens. Do not publish them in GitHub.

## 2. Add the Apps Script project

From the new Google Sheet, choose Extensions > Apps Script. Replace Code.gs with apps-script/Code.gs. Add an HTML file named Index and paste apps-script/Index.html into it. In Project Settings, enable the appsscript.json manifest view and paste apps-script/appsscript.json into that file.

## 3. Configure Script Properties

In Project Settings > Script Properties, add:

| Property | Value |
| --- | --- |
| MASTER_SHEET_ID | ID of the new Google Sheet |
| ADMIN_EMAILS | Comma-separated Google emails for the owner and approved admins |
| BACKUP_FOLDER_ID | ID of the private backup folder |

Include the project owner's Google email and the approved Vaishnavi admin email supplied in the conversation. Do not paste these actual values into repository files. Akshay and other team members are deferred. Only people with script editor rights should be able to change these properties.

This standalone allowlist is separate from the existing ChatGPT Site's access settings. Give approved admins edit access to the master Sheet, and appropriate access to the backup folder if they need manual backups. Do not make the Sheet public. The email must correspond to a Google account; domain policy may require your Workspace administrator's approval.

## 4. Initialize and deploy

Run initialize from the Apps Script editor while signed in as an allowlisted administrator. Review and authorize the requested Google permissions. It creates a Tasks tab without replacing other tabs.

Deploy > New deployment > Web app:

- Execute as: **User accessing the web app**.
- Access: signed-in Google users, or your organization if all admins belong to it. Never anonymous.
- The backend's ADMIN_EMAILS allowlist still applies.

Copy the deployed `/exec` URL. Sign in with an allowlisted account. Each user authorizes the required Google access. If Google or the organization blocks authorization, resolve that through the account administrator; do not bypass verification warnings. Updates require editing the deployment to use a new version.

## 5. Test before importing customer records

1. Confirm an allowlisted admin can open the app and an unlisted account is denied.
2. Create a test task with status Replied; confirm it stays Replied in the Sheet.
3. Edit one task from two windows; the stale save must be rejected after the other window saves.
4. Move the test task to Trash, then restore it.
5. Click Back up now and verify a dated copy exists in the private folder.
6. Import the same small reviewed test batch twice; the second import must skip existing source keys.

Then prepare the real Excel file locally and review the generated files before using Import reviewed JSON. The tool does not update the original workbook or upload any data.

## 6. Enable weekly backups

After the manual backup succeeds, the chosen backup owner runs installWeeklyBackup once from the script editor. That account owns the trigger. It runs Sundays at approximately 11 PM Asia/Kolkata, not at an exact minute. Only one account should install it; different accounts can create separate triggers. Monitor Apps Script Executions and Google's trigger-failure notifications.

The backup owner's account must retain access to the master Sheet and backup folder. The app displays the timestamp of the last successful backup; check it weekly.

## 7. Recovery test

Open a dated backup, verify expected records, and copy it to a separate test Sheet. In a separate test Apps Script project with the same code, set MASTER_SHEET_ID to that test copy and confirm records load. Do not point the production app at a backup without a planned cutover and checking for newer data. Never delete the current master as part of a restore test.
