# Setup guide

How to deploy your own copy of the Visa Reconciliation app.

## 1. Before you start

- A Google Workspace account. Cardholders must be in the same Workspace, so the app can tell who is signed in.
- A Shared Drive where the deploying account is a **Manager**.
- An [Anthropic API key](https://console.anthropic.com/).
- Node.js, and clasp: `npm install -g @google/clasp`, then `clasp login`.

## 2. Create the Apps Script project and push the code

1. Create a standalone project at [script.google.com](https://script.google.com), or run `clasp create --type standalone --rootDir src`.
2. Copy `.clasp.json.example` to `.clasp.json` and put the project's script ID in it.
3. Run `clasp push`. The manifest (`src/appsscript.json`) turns on the Drive API v3 advanced service.

## 3. Set Script Properties

In the Apps Script editor: **Project Settings → Script Properties**.

| Property | Value |
| --- | --- |
| `CLAUDE_API_KEY` | Your Anthropic API key |
| `FINANCE_EMAILS` | Comma-separated Finance reviewer emails, e.g. `finance@example.org,vpfinance@example.org` |
| `ED_EMAIL` | The Executive Director's email, e.g. `ed@example.org` |
| `ROOT_FOLDER_ID` | ID of the Shared Drive (or a folder in it). The app creates a `VISA_REC_APP` folder there. |
| `DEPARTMENT_MAP` | JSON map of cardholder email to receipts folder name, e.g. `{"president@example.org":"President","events@example.org":"Events"}` |
| `ALL_FOLDER_NAME` | Optional. Parent folder for cardholders not in `DEPARTMENT_MAP`. Default `Admin (ALL)`. |
| `ALLOWED_DOMAINS` | Optional. Extra email domains that may sign in, added to `DEFAULT_ALLOWED_DOMAINS` in `src/Auth.gs`. Change that list for your organisation. |
| `STATEMENTS_FOLDER_ID` | Optional. Use an existing statements folder instead of the auto-created one. |
| `HISTORICAL_SHEETS_FOLDER_ID` | Optional. Folder of past reconciliation spreadsheets, used to improve code suggestions. |

These are set automatically: `SPREADSHEET_ID` by `initializeDatabase()`, and `FORMAT_TEMPLATE_URL` by `createFormatTemplate()`.

You can also edit `setupFolderConfig()` in `src/Folders.gs` and run it once to set `ROOT_FOLDER_ID`, `ALLOWED_DOMAINS` and `DEPARTMENT_MAP` together.

## 4. Initialise the database

Run `initializeDatabase()` from the editor. It creates the database spreadsheet with its sheets (Users, BudgetCodes, Reconciliations, Transactions, Receipts, AuditLog) and loads the budget codes. Payroll codes (the 56xxx range) start inactive, because payroll never goes on a card.

Optional: run `importHistoricalSheets()` to learn from past reconciliations.

## 5. Drive folders

The app builds and shares this tree itself:

```
VISA_REC_APP
├── Visa Statements                 Finance saves statements here
├── Receipts
│   ├── <Department>/<Month YYYY>    cardholders listed in DEPARTMENT_MAP
│   └── Admin (ALL)/<email>          everyone else, one folder each
├── Approved Visa Recs/<Month YYYY>  approved reconciliations, filed automatically
└── Archives/FY YYYY-YYYY            created by the fiscal year close
```

Name statements `[Name] VISA Statement [Month] [Year].pdf`, for example `Jordan VISA Statement July 2026.pdf`. The name must match the cardholder's name in the Users tab.

## 6. Add the first Finance user

Run this once from the editor with your own details. The deploying account should be a Finance or ED user, because the setup functions below check the role of whoever runs them.

```javascript
addUser('Your Name', 'you@example.org', 'finance', 'ALL');
```

After that, add everyone else from the app's **Users** tab. Adding a user creates and shares their receipts folder and sends a welcome email.

| Role | Can do |
| --- | --- |
| `staff` | Reconcile their own card with their department's budget codes |
| `exec` | Same as staff, and see reports |
| `finance` | Review everyone's reconciliations, correct codes, approve to the ED, manage users, close the fiscal year |
| `ed` | Final approval, manage users, close the fiscal year |

`department_suffix` sets which budget codes a user sees: one suffix (`1200`), several (`1000,1500`), or `ALL`.

## 7. Turn on the scheduled jobs

Run each once from the editor:

| Function | What it schedules |
| --- | --- |
| `setupStatementNotifyTrigger()` | Hourly: email cardholders when a new statement is saved |
| `setupReceiptScanTrigger()` | Hourly: read new receipts and match them to charges |
| `setupReminderTrigger()` | Mondays at 9am: remind cardholders about missing receipts |

## 8. Deploy the web app

1. **Deploy → New deployment → Web app**, or `clasp deploy`.
2. **Execute as:** Me. The app reads the database and Drive as the deploying account, and still identifies each visitor.
3. **Who has access:** Anyone with a Google account. The Users sheet controls who actually gets in.
4. Send the web app URL to your users.

To ship changes without changing the URL, update the same deployment: `clasp push`, then `clasp deploy -i <deploymentId>`.

## Maintenance functions

Run from the Apps Script editor when needed. All require a Finance or ED account.

| Function | Use it to |
| --- | --- |
| `backfillUserFolders()` | Create and share receipts folders for users added without one |
| `repairMappedUserFolders()` | Move users in `DEPARTMENT_MAP` to their department folder |
| `splitSharedReceiptFolders()` | Give each user sharing the Admin (ALL) folder their own folder |
| `deactivatePayrollCodes()` | Switch off payroll budget codes in an existing database |
| `createFormatTemplate()` | Create the sample export linked from the app |

The fiscal year close is in the app: **Reports → Close fiscal year** (Finance and ED only).
