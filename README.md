# Daily Task Dashboard

Standalone customer-support dashboard foundation for Google Sheets and Google Apps Script.

## Current state

- Apps Script application with Google identity checks and a server-side admin allowlist.
- Shared task register with create, edit, search, agent filter, soft deletion and restore.
- Original source statuses retained, including Replied and Called. These are not treated as Done.
- Phone, AWB, follow-up/calling dates, address and source provenance fields.
- Version checks prevent stale dashboard saves. Script locking serializes writes.
- Reviewed JSON imports, maximum 500 records per batch, deduplicated by source key.
- Manual native Google Sheet backup and optional weekly backup trigger.
- Excel preparation tool that separates usable records from records needing review.

This code is NOT yet deployed to Google. The existing ChatGPT-hosted dashboard is a separate application and is not automatically synchronized with this version. No customer records or account credentials are committed here.

## Start

Follow [SETUP.md](SETUP.md). Google account authorization and deployment must be performed in the data owner's Google account. Do not put credentials, actual administrator emails, spreadsheet IDs or customer data in this public repository.

## Prepare an Excel import locally

Requires Node 22+:

```
npm install
npm run prepare-import -- /absolute/path/to/source.xlsx data/new-import
npm test
```

Read `summary.json` and inspect the ready files privately. Import `ready-001.json`, then the next batch through the deployed app. Inspect `review.json` separately; it includes the source row and reasons. Never upload data/ to GitHub. A fresh output directory is required each run.

Source keys identify an exact source sheet, row number and row contents. Reimporting that same snapshot is safe; reordered or edited workbooks need reconciliation, not a blind second import. Similar records across tabs are flagged for review rather than deleted. Blank source dates and agents are not replaced with invented values. Headerless and incomplete rows remain in review.

## Security and limits

This first standalone version is for trusted administrators only. Every callable data function checks the signed-in Google email against Script Properties. Deploy as **User accessing the web app**, not as owner or anonymous. Admins need edit access to the master Sheet, which also means they can bypass the app by editing that Sheet directly. Do not use this architecture to promise agent-only row restrictions or tamper-proof audit logs. Restrict spreadsheet sharing; add a separate service backend before introducing restricted agent roles.

Only edits through this application increment the record version. Direct Sheet edits are visible on refresh but are not covered by the optimistic concurrency check. Make edits through the app during normal operations. The lock does not prevent manual spreadsheet edits.

Backups are native Google Sheet copies in the selected folder, not automatically refreshed copies of the original .xlsx. Excel backup export, immutable audit history, email automation, WhatsApp, IVR and restricted agent management remain future milestones. Weekly backups run only after installation of the trigger. Google quotas apply. Backup permissions follow Drive behavior; use a private folder and review sharing. Backups are not automatically pruned.

## Verification

`npm test` uses a mocked Google runtime to check denied identities, formula escaping, stale-write conflicts, Trash/restore, duplicate imports and batch validation. Passing local tests does not establish live Google authorization, trigger execution or browser behavior. See the live acceptance checklist in SETUP.md.

## Reference

- https://developers.google.com/apps-script/guides/web
- https://developers.google.com/apps-script/reference/base/session
- https://developers.google.com/apps-script/guides/triggers/installable
