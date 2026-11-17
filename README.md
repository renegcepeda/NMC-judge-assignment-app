# Google Sheets Blind-Judging & Scoring Automation

A Google Apps Script workflow that turns a Google Sheet into an end-to-end blind-judging management system for showcases, conferences, hackathons, and grant reviews.

## Features

- **Smart Chip Parsing & Balanced Assignment:** Extracts judge names and email addresses directly from Google Sheets Person Smart Chips (`@email`) and randomly assigns **2 distinct judges per entry** with evenly balanced workloads.
- **Automated Blind-Judging Sheets:** Generates a private Google Sheet for every judge inside a designated Google Drive folder containing *only* their assigned entries—automatically stripping applicant personal details (names, emails, demographics) while preserving project materials.
- **Built-in 20-Point Rubric:** Appends two 10-point scoring criteria, an auto-calculating `Total Score (Out of 20)` formula column, and an optional `Juror Notes` column to every judge's sheet.
- **One-Click Sharing & Email Dispatch:** Grants each judge `Editor` permissions on their private spreadsheet and sends them a customized email with their direct link and scoring instructions.
- **Score Consolidation & Averaging:** Pulls all submitted scores and juror comments back into the master sheet and computes average scores per entry.
- **Event Manager Delegation:** Optionally prompts for an Event Manager's email and grants them automatic `Editor` access across the master workbook and all generated judge spreadsheets.

---

## Workbook Structure

Your Google Spreadsheet must contain two primary tabs:

### 1. `Judging_master`
Stores all submissions, judge assignments, and final aggregated results:
- **Column A:** Timestamp (included in blind sheets)
- **Columns B–H:** Applicant Personal Information *(Redacted/hidden from judges)*
- **Columns I–Q:** Project Details, Links, Descriptions, and Statements *(Shown to judges)*
- **Columns R–S:** `Judge 1` and `Judge 2` *(Populated automatically in Step 1)*
- **Columns T–X:** `Judge 1 Score`, `Judge 2 Score`, `Average Score`, `Judge 1 Notes`, and `Judge 2 Notes` *(Populated automatically in Step 4)*

### 2. `Judges`
Tracks the judging roster and distribution status across 6 columns:
- **Column A (`Judge Chip`):** Insert Google Person Smart Chips (or plain emails) starting at `A2`.
- **Column B (`Judge Name`):** Auto-extracted human name from Column A.
- **Column C (`Judge Email`):** Auto-extracted email address from Column A (or entered manually).
- **Column D (`Assigned Count`):** Number of entries assigned to the judge.
- **Column E (`Private Judge Sheet Link`):** Direct URL to the judge's private spreadsheet.
- **Column F (`Email Status`):** Tracks whether the invite email has been sent.

---

## Setup & Installation

1. **Create a Destination Folder in Google Drive**
   - Create a folder where all generated judge spreadsheets will live.
   - Copy its Folder ID from the URL (`https://drive.google.com/drive/folders/<YOUR_FOLDER_ID>`).

2. **Add the Script to Your Google Sheet**
   - Open your Google Sheet and go to **Extensions > Apps Script**.
   - Paste the contents of `Code.gs` into the editor.
   - Update the configuration constants at the top of `Code.gs`:
     ```javascript
     const TARGET_FOLDER_ID = "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE";
     const DEFAULT_EVENT_MANAGER_EMAIL = ""; // Optional
     const EVENT_NAME = "Showcase";
     ```
   - Save the project (`Ctrl+S` / `Cmd+S`) and refresh your Google Sheet.

3. **Authorize Permissions**
   - When you run your first command from the **Judging Tools** menu, Google will ask you to authorize the script to manage Google Sheets, Google Drive files, and send emails on your behalf.

---

## Usage Workflow

Once refreshed, use the **Judging Tools** menu in your Google Sheet's top toolbar:

1. **1. Assign Judges Evenly (2 per Entry):** Prompts for an Event Manager email, extracts judge names/emails from Column A of `Judges`, and assigns 2 judges per entry in `Judging_master`.
2. **2. Auto-Create Individual Judge Sheets:** Creates a subfolder in your target Drive directory and generates one blind-judging spreadsheet per judge.
3. **3. Email Links & Share With Judges:** Shares each private sheet with its assigned judge and emails them their link. *(Skips rows already marked `Sent` so it is safe to re-run if you add late judges.)*
4. **4. Collect & Average All Scores:** Reads every judge's spreadsheet and writes their scores, averages, and notes into Columns `T–X` of `Judging_master`.

### Utility Commands
- **Grant Event Manager Access Now:** Immediately adds or updates an Event Manager as an Editor on the master workbook and all existing judge spreadsheets.
- **Build / Refresh Instructions Tab:** Creates a formatted reference tab inside the workbook for organizers.
- **Clear Judge Assignments:** Wipes Columns `R` and `S` on `Judging_master` if you need to reshuffle assignments before creating judge sheets.

---

## Deploying with `clasp` (Optional)

If you manage this repository locally using Google's [`clasp`](https://github.com/google/clasp) CLI:

```bash
npm install -g @google/clasp
clasp login
clasp clone <YOUR_SCRIPT_ID>
# Replace Code.gs with this repository's Code.gs
clasp push
```

## License

MIT License — feel free to adapt and customize for your organization or event.
