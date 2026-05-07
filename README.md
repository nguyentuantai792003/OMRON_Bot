# OMRON QnA Daily Report Bot

Google Apps Script automation for the OMRON project. Sends daily Q&A status reports to Google Chat and creates daily meeting issues on Backlog.

---

## Features

- **Daily Q&A Report** — reads the Q&A Excel file from Google Drive, groups questions by status, and posts a summary to Google Chat
- **Daily Dev Report** — reads the daily report Excel file, checks who has submitted their report, and posts a per-member summary with mentions for those who haven't
- **Backlog Issue Creation** — automatically creates a `[DAILY MEETING]` issue on Backlog at 17:00 ICT with each member's done tasks, next-day plan, and in-progress TODO items

---

## Project Structure

```
scripts.js   — single Apps Script file containing all logic and config
```

---

## Configuration

All settings are at the top of `scripts.js`.

### Timezone
```js
const TIMEZONE = "Asia/Ho_Chi_Minh";
```

### Google Chat Webhooks
| Constant | Description |
|---|---|
| `WEBHOOK_URL_DEV` | Dev team space |
| `WEBHOOK_URL_TEST` | Test space |
| `WEBHOOK_URL_ALL` | All-members space |
| `WEBHOOK_URL_DEV_APP` | Dev app space |

### Google Drive Files
| Constant | Description |
|---|---|
| `QNA_FILE_ID` | Google Drive ID of the Q&A `.xlsx` file |
| `DAILY_REPORT_FILE_ID` | Google Drive ID of the daily report `.xlsx` file |

### Sheet Names
| Constant | Value |
|---|---|
| `QNA_SHEET_NAME` | `質問票` |
| `TODO_SHEET_NAME` | `Todo list` |

The daily report sheet name is auto-detected from the current month, e.g. `T52026` for May 2026.

### Backlog
| Constant | Description |
|---|---|
| `BACKLOG_API_KEY` | Backlog API key |
| `BACKLOG_SPACE` | Backlog space name (`vti-corp`) |
| `BACKLOG_PROJECT_ID` | Numeric project ID (`763680`) |
| `BACKLOG_ISSUE_TYPE_ID` | Numeric issue type ID (`4108322`) |
| `BACKLOG_PRIORITY_ID` | Priority ID (`3` = Normal) |

### Backlog Trigger Window
```js
const BACKLOG_TRIGGER_HOUR = 17;
const BACKLOG_TRIGGER_START_MINUTE = 0;
const BACKLOG_TRIGGER_END_MINUTE = 20;
```
`sendBacklogReport` will only execute when called between 17:00 and 17:20 ICT. A script property (`lastRun`) prevents duplicate issues on the same day.

### User Mapping
`USER_MAP` maps display names to Google Chat user IDs for `@mention` support. Add or uncomment entries as team members join.

### Excluded Users
`EXCLUDED_USERS` lists members whose rows are skipped when reading the daily report sheet.

---

## Functions

### Entry Points (set these as time-based triggers)

| Function | Recommended Trigger | Description |
|---|---|---|
| `dailyQnAReport()` | Daily, any time (weekdays) | Posts Q&A status summary to Google Chat |
| `sendDailyChatReport()` | Daily, e.g. 17:00 (weekdays) | Posts dev daily report to Google Chat |
| `sendBacklogReport()` | Every 5–10 min, 17:00–17:20 | Creates daily meeting issue on Backlog |

All three functions skip Saturday and Sunday automatically.

### Internal Functions

| Function | Description |
|---|---|
| `getDailyReportData()` | Opens the daily report file and returns today's records grouped by user |
| `createBacklogIssue(reportByUser, today)` | Builds the issue description and calls the Backlog API |
| `parseNote(note)` | Parses a report note into `today` (done) and `next` (plan) sections |
| `getTodoInProgress()` | Reads the Todo list sheet and returns items with status containing "progress" |
| `createTempSpreadsheetFromExcel(fileId, tempName)` | Converts a Drive `.xlsx` file to a temporary Google Spreadsheet for reading |
| `deleteTempFile(fileId)` | Moves a temporary file to trash |
| `sendGoogleChat(webhookUrl, message)` | Posts a plain-text message to a Google Chat webhook |
| `isSameDay(d1, d2)` | Compares two dates by year/month/day |
| `parseDate(value)` | Parses a `MM/DD/YYYY` string into a `Date` object |

---

## Daily Report Note Format

Each member's report note in the spreadsheet should follow this format:

```
Today:
- Task A
- Task B

Next plan:
- Task C

Issue:
- Any blockers (this section is stripped before posting to Backlog)
```

---

## Setup

1. Open [Google Apps Script](https://script.google.com) and create a new project.
2. Paste the contents of `scripts.js` into the editor.
3. Enable the **Google Drive API** (Advanced Services → Drive API v3).
4. Set the correct values for `BACKLOG_PROJECT_ID`, `BACKLOG_ISSUE_TYPE_ID`, and all webhook URLs.
5. Add time-based triggers for `dailyQnAReport`, `sendDailyChatReport`, and `sendBacklogReport`.

---

## Notes

- The script converts `.xlsx` files to temporary Google Spreadsheets on each run and deletes them afterward. Make sure the service account / script owner has edit access to Google Drive.
- `sendBacklogReport` uses `PropertiesService` to store a `lastRun` key and prevent duplicate Backlog issues within the same trigger window.
- To look up your Backlog project and issue type IDs:
  ```
  GET https://<space>.backlog.com/api/v2/projects/<PROJECT_KEY>?apiKey=<API_KEY>
  GET https://<space>.backlog.com/api/v2/projects/<PROJECT_KEY>/issueTypes?apiKey=<API_KEY>
  ```
