/**
 * ============================================================================
 * GOOGLE SHEETS BLIND-JUDGING & SCORING AUTOMATION
 * ============================================================================
 * Description:
 *   Automates the end-to-end judging workflow for showcases, competitions,
 *   hackathons, or grant reviews directly inside Google Sheets:
 *     1. Extracts judge names/emails from Google Person Smart Chips and evenly
 *        assigns 2 distinct judges per entry.
 *     2. Generates individual, blind-judging Google Sheets inside a specified
 *        Google Drive folder (hiding applicant personal info).
 *     3. Shares each private spreadsheet with its assigned judge and emails
 *        them a direct link with rubric instructions.
 *     4. Collects submitted scores and notes back into the master sheet and
 *        calculates average scores.
 * ============================================================================
 */

// ============================================================================
// CONFIGURATION (Update these values before running)
// ============================================================================

// Required: The Google Drive Folder ID where individual judge sheets will be created.
// Find this in the URL when viewing your Drive folder: https://drive.google.com/drive/folders/YOUR_FOLDER_ID
const TARGET_FOLDER_ID = "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE";

// Optional: Fallback Event Manager email (leave blank "" to configure via the UI popup prompt).
const DEFAULT_EVENT_MANAGER_EMAIL = "";

// Organization / Event Name used in email templates and instructions.
const EVENT_NAME = "Showcase";

// ============================================================================
// MENU & INITIALIZATION
// ============================================================================

/**
 * Runs automatically when the spreadsheet is opened.
 * Adds the custom "Judging Tools" menu to the top Google Sheets toolbar.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Judging Tools")
    .addItem("1. Assign Judges Evenly (2 per Entry)", "assignJudges")
    .addItem("2. Auto-Create Individual Judge Sheets", "createAllJudgeSheets")
    .addItem("3. Email Links & Share With Judges", "emailAllJudges")
    .addItem("4. Collect & Average All Scores", "collectAndAverageScores")
    .addSeparator()
    .addItem("Grant Event Manager Access Now", "grantEventManagerAccess")
    .addItem("Build / Refresh Instructions Tab", "buildInstructionsTab")
    .addItem("Clear Judge Assignments", "clearAssignments")
    .addToUi();
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Locates the master judging tab case-insensitively ("Judging_master" or "Judging_Master").
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The active spreadsheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null} The master sheet, or null if not found.
 */
function getMasterSheet(ss) {
  return (
    ss.getSheetByName("Judging_master") ||
    ss.getSheetByName("Judging_Master") ||
    ss.getSheets().find(function(s) {
      return s.getName().toLowerCase() === "judging_master";
    }) ||
    null
  );
}

/**
 * Retrieves the saved Event Manager email from Script Properties,
 * falling back to DEFAULT_EVENT_MANAGER_EMAIL if configured.
 *
 * @return {string} Valid email address or an empty string.
 */
function getEventManagerEmail() {
  const saved = PropertiesService.getScriptProperties().getProperty("EVENT_MANAGER_EMAIL");
  if (saved && saved.includes("@")) {
    return saved.trim();
  }
  if (DEFAULT_EVENT_MANAGER_EMAIL && DEFAULT_EVENT_MANAGER_EMAIL.includes("@")) {
    return DEFAULT_EVENT_MANAGER_EMAIL.trim();
  }
  return "";
}

/**
 * Extracts a clean display name and email address from a Google Sheets Smart Chip cell.
 * Uses display values and rich-text links so Person Chips resolve to full names
 * (e.g., "Jane Doe") rather than raw email addresses.
 *
 * @param {*} rawVal - The cell's raw value from getValues().
 * @param {string} displayVal - The cell's visible text from getDisplayValues().
 * @param {GoogleAppsScript.Spreadsheet.RichTextValue} richTextVal - The cell's RichTextValue.
 * @return {{name: string, email: string}} Extracted name and email.
 */
function parseSmartChip(rawVal, displayVal, richTextVal) {
  const raw = String(rawVal || "").trim();
  const disp = String(displayVal || "").trim();
  let email = "";
  let name = "";

  // 1. Inspect RichTextValue link URLs (mailto:) or text runs
  let richText = "";
  if (richTextVal) {
    richText = String(richTextVal.getText() || "").trim();
    const linkUrl = richTextVal.getLinkUrl();
    if (linkUrl && linkUrl.toLowerCase().startsWith("mailto:")) {
      email = linkUrl.replace(/^mailto:/i, "").trim();
    }
    if (!email && richTextVal.getRuns) {
      const runs = richTextVal.getRuns();
      for (let i = 0; i < runs.length; i++) {
        const runUrl = runs[i].getLinkUrl();
        if (runUrl && runUrl.toLowerCase().startsWith("mailto:")) {
          email = runUrl.replace(/^mailto:/i, "").trim();
          break;
        }
      }
    }
  }

  // 2. Extract email from raw or display text if not found in rich text
  const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
  if (!email) {
    const rawMatch = raw.match(emailRegex);
    if (rawMatch) email = rawMatch[0];
  }
  if (!email) {
    const dispMatch = disp.match(emailRegex);
    if (dispMatch) email = dispMatch[0];
  }

  // 3. Determine human-readable name from the Smart Chip
  const candidateName = disp || richText || raw;
  if (candidateName && !candidateName.includes("@")) {
    name = candidateName;
  } else if (candidateName && candidateName.includes("<") && candidateName.includes(">")) {
    // Handle "First Last <email@domain.com>" format
    name = candidateName.split("<")[0].trim().replace(/^["']|["']$/g, "");
  } else if (email) {
    // Fallback: format the email prefix into Title Case
    const prefix = email.split("@")[0];
    name = prefix
      .replace(/[._-]+/g, " ")
      .replace(/\b\w/g, function(char) { return char.toUpperCase(); });
  } else {
    name = candidateName;
  }

  return { name: name, email: email };
}

// ============================================================================
// STEP 1: ASSIGN JUDGES EVENLY
// ============================================================================

/**
 * Prompts for an Event Manager email, grants them Editor access to the master workbook,
 * parses Judge Names & Emails from Smart Chips in the 'Judges' tab, and randomly
 * assigns 2 distinct judges to every submission in 'Judging_master' (Columns R & S).
 */
function assignJudges() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = getMasterSheet(ss);
  const judgesSheet = ss.getSheetByName("Judges");

  if (!masterSheet || !judgesSheet) {
    ui.alert("Missing tabs! Ensure your workbook has tabs named 'Judging_master' and 'Judges'.");
    return;
  }

  // Prompt for Event Manager email
  const currentManager = getEventManagerEmail();
  const promptText = currentManager
    ? "Current Event Manager: " + currentManager + "\n\nEnter a new email to update, leave blank to keep current, or type CLEAR to remove:"
    : "Enter the Event Manager's email address to grant them automatic Editor access to all sheets (or leave blank to skip):";

  const managerPrompt = ui.prompt("Event Manager Setup", promptText, ui.ButtonSet.OK_CANCEL);
  if (managerPrompt.getSelectedButton() === ui.Button.CANCEL) {
    return;
  }

  const enteredText = managerPrompt.getResponseText().trim();
  const scriptProps = PropertiesService.getScriptProperties();

  if (enteredText.toUpperCase() === "CLEAR") {
    scriptProps.deleteProperty("EVENT_MANAGER_EMAIL");
  } else if (enteredText !== "") {
    if (!enteredText.includes("@")) {
      ui.alert("Invalid Email", "Please enter a valid email address containing '@', or leave the box blank to skip.", ui.ButtonSet.OK);
      return;
    }
    scriptProps.setProperty("EVENT_MANAGER_EMAIL", enteredText);
  }

  // Grant master workbook Editor access to the Event Manager
  const activeManager = getEventManagerEmail();
  if (activeManager) {
    try {
      ss.addEditor(activeManager);
    } catch (e) {
      Logger.log("Could not add event manager to master sheet: " + e.message);
    }
  }

  // Standardize the 6-column Judges roster header
  judgesSheet.getRange("A1:F1").setValues([[
    "Judge Chip",
    "Judge Name",
    "Judge Email",
    "Assigned Count",
    "Private Judge Sheet Link",
    "Email Status"
  ]]).setFontWeight("bold");

  if (judgesSheet.getMaxColumns() > 6) {
    judgesSheet.getRange(1, 7, judgesSheet.getMaxRows(), judgesSheet.getMaxColumns() - 6).clearContent();
  }

  const lastJudgeRow = judgesSheet.getLastRow();
  if (lastJudgeRow < 2) {
    ui.alert("Please list at least 2 judges in Column A ('Judge Chip') starting at cell A2.");
    return;
  }

  const numJudgeRows = lastJudgeRow - 1;
  const chipRange = judgesSheet.getRange(2, 1, numJudgeRows, 1);
  const chipRawValues = chipRange.getValues();
  const chipDisplayValues = chipRange.getDisplayValues();
  const chipRichTextValues = chipRange.getRichTextValues();
  const existingBC = judgesSheet.getRange(2, 2, numJudgeRows, 2).getValues();

  const judges = [];
  const updatedBC = [];

  for (let i = 0; i < numJudgeRows; i++) {
    const rawVal = chipRawValues[i][0];
    const dispVal = chipDisplayValues[i][0];
    const richVal = chipRichTextValues[i][0];

    let existingName = String(existingBC[i][0] || "").trim();
    let existingEmail = String(existingBC[i][1] || "").trim();

    if (String(rawVal || dispVal || "").trim() !== "") {
      const parsed = parseSmartChip(rawVal, dispVal, richVal);

      // Prefer clean display name if Column B is empty or contains an email
      let finalName = existingName;
      if (!finalName || finalName.includes("@")) {
        finalName = parsed.name;
      }

      let finalEmail = existingEmail || parsed.email;

      judges.push(finalName);
      updatedBC.push([finalName, finalEmail]);
    } else {
      updatedBC.push([existingName, existingEmail]);
    }
  }

  // Write extracted Judge Name (Col B) and Judge Email (Col C)
  judgesSheet.getRange(2, 2, updatedBC.length, 2).setValues(updatedBC);

  if (judges.length < 2) {
    ui.alert("Please list at least 2 judges in Column A (A2:A) of the 'Judges' tab.");
    return;
  }

  const lastRow = masterSheet.getLastRow();
  const numEntries = lastRow - 1;
  if (numEntries < 1) {
    ui.alert("No entries found in '" + masterSheet.getName() + "'.");
    return;
  }

  // Check if assignments already exist in Columns R & S (18 & 19)
  const existing = masterSheet.getRange(2, 18, numEntries, 2).getValues().flat().filter(Boolean);
  if (existing.length > 0) {
    const response = ui.alert(
      "Overwrite Assignments?",
      "Judge assignments already exist in '" + masterSheet.getName() + "'. Do you want to reshuffle and overwrite them?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  // Fisher-Yates shuffle on judges
  const shuffledJudges = judges.slice();
  for (let i = shuffledJudges.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [shuffledJudges[i], shuffledJudges[j]] = [shuffledJudges[j], shuffledJudges[i]];
  }

  // Fisher-Yates shuffle on entry row indices
  const rowIndices = Array.from({ length: numEntries }, function(_, i) { return i; });
  for (let i = rowIndices.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [rowIndices[i], rowIndices[j]] = [rowIndices[j], rowIndices[i]];
  }

  // Round-robin pair assignment ensuring 2 distinct judges per entry and balanced workloads
  const assignments = new Array(numEntries);
  for (let k = 0; k < numEntries; k++) {
    const targetRow = rowIndices[k];
    const j1 = shuffledJudges[k % shuffledJudges.length];
    const j2 = shuffledJudges[(k + 1) % shuffledJudges.length];
    assignments[targetRow] = [j1, j2];
  }

  masterSheet.getRange("R1:S1").setValues([["Judge 1", "Judge 2"]]).setFontWeight("bold");
  masterSheet.getRange(2, 18, masterSheet.getMaxRows() - 1, 2).clearContent();
  masterSheet.getRange(2, 18, numEntries, 2).setValues(assignments);

  const minLoad = Math.floor((numEntries * 2) / judges.length);
  const maxLoad = Math.ceil((numEntries * 2) / judges.length);
  const managerNote = activeManager ? "\nEvent Manager (" + activeManager + ") granted Editor access." : "";

  ui.alert(
    "Assignments Complete!",
    "Extracted judge names to Column B and assigned " + numEntries + " entries across " + judges.length + " judges (" + minLoad + "–" + maxLoad + " entries per judge)." + managerNote,
    ui.ButtonSet.OK
  );
}

// ============================================================================
// STEP 2: CREATE INDIVIDUAL BLIND-JUDGING SHEETS
// ============================================================================

/**
 * Creates a dedicated subfolder in Google Drive and generates one private Google Sheet
 * per judge containing only their assigned entries from 'Judging_master'.
 * Redacts applicant personal details (Columns B–H) and appends a 20-point scoring rubric.
 */
function createAllJudgeSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = getMasterSheet(ss);
  const judgesSheet = ss.getSheetByName("Judges");

  if (!masterSheet || !judgesSheet) {
    ui.alert("Missing tabs! Ensure your workbook has tabs named 'Judging_master' and 'Judges'.");
    return;
  }

  if (!TARGET_FOLDER_ID || TARGET_FOLDER_ID === "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE") {
    ui.alert("Configuration Required", "Please set your TARGET_FOLDER_ID at the top of the script before creating judge sheets.", ui.ButtonSet.OK);
    return;
  }

  const lastJudgeRow = judgesSheet.getLastRow();
  if (lastJudgeRow < 2) {
    ui.alert("No judges listed in 'Judges' tab starting at Row 2.");
    return;
  }

  const numJudgeRows = lastJudgeRow - 1;
  const judgeRowsData = judgesSheet.getRange(2, 1, numJudgeRows, 2).getValues();
  const activeJudges = judgeRowsData
    .map(function(r) { return String(r[1] || "").trim(); })
    .filter(Boolean);

  if (activeJudges.length === 0) {
    ui.alert("No extracted Judge Names found in Column B. Run Step 1 first.");
    return;
  }

  let parentFolder;
  try {
    parentFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  } catch (e) {
    ui.alert("Could not access the target Google Drive folder. Check permissions for ID: " + TARGET_FOLDER_ID);
    return;
  }

  const allData = masterSheet.getDataRange().getValues();
  if (allData.length < 2) {
    ui.alert("No entries found in '" + masterSheet.getName() + "'.");
    return;
  }
  const headers = allData[0];

  // Rubric column headers (customize criteria descriptions as needed)
  const q1Header = "Criterion 1: Quality of Application (Out of 10) — Focused on a single project, feasible within the presentation format, and aligned with event guidelines.";
  const q2Header = "Criterion 2: Quality of Proposed Project (Out of 10) — Compelling concept, strong execution, and offers a unique perspective.";
  const totalHeader = "Total Score (Out of 20)";
  const notesHeader = "Juror Notes (Optional) — Comments on why this project would or would not be a compelling inclusion.";

  // Blind review columns: Entry ID, Timestamp (Col A), and Project Details (Cols I–Q).
  // Applicant personal info in Columns B–H (indices 1–7) is intentionally excluded.
  const blindHeaders = [
    "Entry ID",
    headers[0],  // Col A: Timestamp
    headers[8],  // Col I: Project Field 1
    headers[9],  // Col J: Project Field 2
    headers[10], // Col K: Project Field 3
    headers[11], // Col L: Project Field 4
    headers[12], // Col M: Project Field 5
    headers[13], // Col N: Project Field 6
    headers[14], // Col O: Project Field 7
    headers[15], // Col P: Project Field 8
    headers[16], // Col Q: Project Field 9
    q1Header,    // Col L (12) in Judge Sheet
    q2Header,    // Col M (13) in Judge Sheet
    totalHeader, // Col N (14) in Judge Sheet (Auto-sum)
    notesHeader  // Col O (15) in Judge Sheet
  ];

  const folder = parentFolder.createFolder(ss.getName() + " - Judge Sheets");
  const activeManager = getEventManagerEmail();
  const statusRows = [];
  let createdCount = 0;

  for (let i = 0; i < judgeRowsData.length; i++) {
    const judgeName = String(judgeRowsData[i][1] || "").trim();

    if (!judgeName) {
      statusRows.push(["", "", ""]);
      continue;
    }

    const judgeEntries = [blindHeaders];

    for (let r = 1; r < allData.length; r++) {
      const row = allData[r];
      const assignedJ1 = String(row[17] || "").trim(); // Col R in master
      const assignedJ2 = String(row[18] || "").trim(); // Col S in master

      if (assignedJ1 === judgeName || assignedJ2 === judgeName) {
        const currentSheetRow = judgeEntries.length + 1;
        judgeEntries.push([
          "ROW-" + (r + 1),
          row[0], row[8], row[9], row[10], row[11],
          row[12], row[13], row[14], row[15], row[16],
          "", // Score 1 (Out of 10)
          "", // Score 2 (Out of 10)
          "=IF(AND(ISBLANK(L" + currentSheetRow + "), ISBLANK(M" + currentSheetRow + ")), \"\", SUM(L" + currentSheetRow + ":M" + currentSheetRow + "))",
          ""  // Juror Notes
        ]);
      }
    }

    const judgeSS = SpreadsheetApp.create("Judging - " + judgeName);
    const sheet = judgeSS.getActiveSheet();
    sheet.setName("My Assigned Entries");
    sheet.getRange(1, 1, judgeEntries.length, blindHeaders.length).setValues(judgeEntries);
    sheet.setFrozenRows(1);

    // Format header row and rubric columns
    sheet.getRange(1, 1, 1, 11).setFontWeight("bold");
    sheet.getRange(1, 12, 1, 4).setFontWeight("bold").setWrap(true).setVerticalAlignment("top");
    sheet.setColumnWidth(12, 220);
    sheet.setColumnWidth(13, 220);
    sheet.setColumnWidth(14, 140);
    sheet.setColumnWidth(15, 320);

    // Shade the Total Score formula column so judges recognize it as auto-calculated
    if (judgeEntries.length > 1) {
      sheet.getRange(2, 14, judgeEntries.length - 1, 1).setBackground("#f1f3f4").setFontWeight("bold");
    }

    // Grant Event Manager Editor access
    if (activeManager) {
      try {
        judgeSS.addEditor(activeManager);
      } catch (e) {
        Logger.log("Could not add event manager to " + judgeName + " sheet: " + e.message);
      }
    }

    const file = DriveApp.getFileById(judgeSS.getId());
    file.moveTo(folder);

    statusRows.push([judgeEntries.length - 1, judgeSS.getUrl(), "Not Sent"]);
    createdCount++;
  }

  // Write Assigned Count (Col D), Private Link (Col E), and Email Status (Col F)
  judgesSheet.getRange(2, 4, judgesSheet.getMaxRows() - 1, 3).clearContent();
  if (statusRows.length > 0) {
    judgesSheet.getRange(2, 4, statusRows.length, 3).setValues(statusRows);
  }

  ui.alert(
    "Judge Sheets Created!",
    "Created " + createdCount + " private spreadsheets inside '" + parentFolder.getName() + " > " + folder.getName() + "'. Links are listed in Columns D–F of the 'Judges' tab.",
    ui.ButtonSet.OK
  );
}

// ============================================================================
// STEP 3: SHARE & EMAIL LINKS TO JUDGES
// ============================================================================

/**
 * Grants each judge Editor access to their private spreadsheet and sends them
 * an instructional email with their direct link. Updates Column F ('Email Status').
 */
function emailAllJudges() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const judgesSheet = ss.getSheetByName("Judges");

  const lastRow = judgesSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("No judges found in the 'Judges' tab.");
    return;
  }

  const rows = judgesSheet.getRange(2, 1, lastRow - 1, 6).getValues();

  const confirm = ui.alert(
    "Send Emails to Judges?",
    "This will grant Editor access to each judge's private spreadsheet and email them their link. Continue?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const activeManager = getEventManagerEmail();
  let sentCount = 0;
  let skippedCount = 0;

  for (let i = 0; i < rows.length; i++) {
    const judgeChip = String(rows[i][0] || "").trim();
    const judgeName = String(rows[i][1] || "").trim();
    let judgeEmail = String(rows[i][2] || "").trim();
    const assignedCount = rows[i][3];
    const sheetUrl = String(rows[i][4] || "").trim();
    const currentStatus = String(rows[i][5] || "").trim();

    if (!judgeChip && !judgeName && !judgeEmail) {
      continue;
    }

    if (!judgeEmail && judgeChip.includes("@")) {
      judgeEmail = judgeChip;
    }

    if (!judgeEmail || !judgeEmail.includes("@") || !sheetUrl) {
      judgesSheet.getRange(i + 2, 6).setValue("Skipped (Missing Email or Link)");
      skippedCount++;
      continue;
    }

    // Avoid duplicate sends if already marked Sent
    if (currentStatus.startsWith("Sent")) {
      continue;
    }

    try {
      const judgeSS = SpreadsheetApp.openByUrl(sheetUrl);
      judgeSS.addEditor(judgeEmail);

      if (activeManager) {
        try {
          judgeSS.addEditor(activeManager);
        } catch (err) {}
      }

      const subject = EVENT_NAME + " Judging Assignment (" + assignedCount + " Entries)";
      const body =
        "Hello " + judgeName + ",\n\n" +
        "Thank you for serving as a judge for the " + EVENT_NAME + "! You have been assigned " + assignedCount + " entries to review.\n\n" +
        "Please open your private judging spreadsheet using the link below:\n" +
        sheetUrl + "\n\n" +
        "Instructions:\n" +
        "1. Review each entry's description and project links.\n" +
        "2. Enter your two scores (out of 10 points each) in Columns L and M.\n" +
        "3. Column N automatically calculates the Total Score out of 20.\n" +
        "4. Leave any optional comments in Column O (Juror Notes).\n" +
        "5. Your changes save automatically.\n\n" +
        "Please let us know if you have any trouble accessing your sheet.\n\n" +
        "Best regards,\n" +
        EVENT_NAME + " Organizing Committee";

      MailApp.sendEmail(judgeEmail, subject, body);
      judgesSheet.getRange(i + 2, 6).setValue("Sent (" + new Date().toLocaleDateString() + ")");
      sentCount++;
    } catch (e) {
      judgesSheet.getRange(i + 2, 6).setValue("Error: " + e.message);
      skippedCount++;
    }
  }

  ui.alert(
    "Emails Processed!",
    "Successfully shared & emailed: " + sentCount + "\nSkipped/Errors: " + skippedCount + "\n\nCheck Column F of the 'Judges' tab for status details.",
    ui.ButtonSet.OK
  );
}

// ============================================================================
// STEP 4: COLLECT & AVERAGE SCORES
// ============================================================================

/**
 * Reads all individual judge spreadsheets, matches scores by Entry ID ("ROW-X")
 * and Judge Name, writes Judge 1 & Judge 2 scores/notes into Columns T–X of
 * 'Judging_master', and calculates each entry's average score.
 */
function collectAndAverageScores() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = getMasterSheet(ss);
  const judgesSheet = ss.getSheetByName("Judges");

  if (!masterSheet || !judgesSheet) {
    ui.alert("Missing tabs! Ensure your workbook has tabs named 'Judging_master' and 'Judges'.");
    return;
  }

  const lastJudgeRow = judgesSheet.getLastRow();
  if (lastJudgeRow < 2) {
    ui.alert("No judge sheet links found in the 'Judges' tab.");
    return;
  }

  // Read Col B (Judge Name) and Col E (Private Judge Sheet Link)
  const linkRows = judgesSheet.getRange(2, 2, lastJudgeRow - 1, 4).getValues().filter(function(r) {
    return r[0] && r[3];
  });
  if (linkRows.length === 0) {
    ui.alert("No judge sheet links found in Column E of the 'Judges' tab. Run Step 2 first.");
    return;
  }

  const scoreMap = {};

  linkRows.forEach(function(linkRow) {
    const judgeName = String(linkRow[0]).trim();
    const sheetUrl = String(linkRow[3]).trim();
    try {
      const judgeSS = SpreadsheetApp.openByUrl(sheetUrl);
      const sheet = judgeSS.getSheetByName("My Assigned Entries") || judgeSS.getSheets()[0];
      const data = sheet.getDataRange().getValues();

      // Col A (0) = Entry ID, Col L (11) = Score 1, Col M (12) = Score 2, Col N (13) = Total, Col O (14) = Notes
      for (let i = 1; i < data.length; i++) {
        const entryId = String(data[i][0] || "").trim();
        const s1 = data[i][11];
        const s2 = data[i][12];
        const rawTotal = data[i][13];
        const notes = String(data[i][14] || "").trim();

        let finalScore = "";
        if (rawTotal !== "" && rawTotal !== null && !isNaN(Number(rawTotal))) {
          finalScore = Number(rawTotal);
        } else if ((s1 !== "" && s1 !== null) || (s2 !== "" && s2 !== null)) {
          finalScore = (Number(s1) || 0) + (Number(s2) || 0);
        }

        if (entryId) {
          scoreMap[entryId + "|" + judgeName] = {
            score: finalScore,
            notes: notes
          };
        }
      }
    } catch (e) {
      Logger.log("Could not open sheet for " + judgeName + ": " + e.message);
    }
  });

  const masterData = masterSheet.getDataRange().getValues();
  const numEntries = masterData.length - 1;
  if (numEntries < 1) return;

  masterSheet.getRange("T1:X1").setValues([[
    "Judge 1 Score (Out of 20)",
    "Judge 2 Score (Out of 20)",
    "Average Score (Out of 20)",
    "Judge 1 Notes",
    "Judge 2 Notes"
  ]]).setFontWeight("bold");

  const results = [];

  for (let r = 1; r < masterData.length; r++) {
    const entryId = "ROW-" + (r + 1);
    const j1 = String(masterData[r][17] || "").trim();
    const j2 = String(masterData[r][18] || "").trim();

    const res1 = scoreMap[entryId + "|" + j1] || { score: "", notes: "" };
    const res2 = scoreMap[entryId + "|" + j2] || { score: "", notes: "" };

    const validScores = [res1.score, res2.score].filter(function(s) {
      return typeof s === "number" && !isNaN(s);
    });
    const avgScore = validScores.length > 0
      ? Math.round((validScores.reduce(function(a, b) { return a + b; }, 0) / validScores.length) * 100) / 100
      : "";

    results.push([
      res1.score,
      res2.score,
      avgScore,
      res1.notes,
      res2.notes
    ]);
  }

  masterSheet.getRange(2, 20, numEntries, 5).setValues(results);

  ui.alert(
    "Scores Collected!",
    "All submitted judge scores (out of 20), averages, and notes have been updated in Columns T–X of '" + masterSheet.getName() + "'.",
    ui.ButtonSet.OK
  );
}

// ============================================================================
// UTILITY TOOLS
// ============================================================================

/**
 * Grants an Event Manager Editor access to the master workbook and all
 * existing individual judge spreadsheets immediately.
 */
function grantEventManagerAccess() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const judgesSheet = ss.getSheetByName("Judges");

  let activeManager = getEventManagerEmail();
  const promptText = activeManager
    ? "Current Event Manager: " + activeManager + "\n\nPress OK to grant access using this email, or type a different email address:"
    : "Enter the Event Manager's email address to grant them Editor access to the master sheet and all judge sheets:";

  const response = ui.prompt("Grant Event Manager Access", promptText, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const input = response.getResponseText().trim();
  if (input !== "") {
    if (!input.includes("@")) {
      ui.alert("Invalid Email", "Please enter a valid email address.", ui.ButtonSet.OK);
      return;
    }
    PropertiesService.getScriptProperties().setProperty("EVENT_MANAGER_EMAIL", input);
    activeManager = input;
  }

  if (!activeManager) {
    ui.alert("No Email Provided", "No Event Manager email was entered.", ui.ButtonSet.OK);
    return;
  }

  let sharedSheets = 0;

  try {
    ss.addEditor(activeManager);
    sharedSheets++;
  } catch (e) {
    Logger.log("Master sheet share error: " + e.message);
  }

  if (judgesSheet) {
    const lastRow = judgesSheet.getLastRow();
    if (lastRow >= 2) {
      const urls = judgesSheet.getRange(2, 5, lastRow - 1, 1).getValues().flat().filter(Boolean);
      urls.forEach(function(url) {
        try {
          const judgeSS = SpreadsheetApp.openByUrl(String(url).trim());
          judgeSS.addEditor(activeManager);
          sharedSheets++;
        } catch (e) {
          Logger.log("Judge sheet share error: " + e.message);
        }
      });
    }
  }

  ui.alert(
    "Event Manager Access Granted!",
    "Added " + activeManager + " as Editor on " + sharedSheets + " spreadsheet(s) (master workbook + existing judge sheets).",
    ui.ButtonSet.OK
  );
}

/**
 * Generates or resets a formatted 'Instructions' tab inside the workbook
 * for event organizers.
 */
function buildInstructionsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Instructions");

  if (!sheet) {
    sheet = ss.insertSheet("Instructions");
  } else {
    sheet.clear();
  }

  const rows = [
    ["STEP / TOPIC", "ACTION / INSTRUCTION", "DETAILS & NOTES"],
    ["OVERVIEW", "How This Workbook Works", "Automates event manager access, balanced 2-judge assignment from 'Judging_master', blind-judging sheet creation with a 20-point rubric, email distribution, and score aggregation."],
    ["TAB STRUCTURE", "Workbook Tabs", "Requires 'Judging_master' (submissions, assignments in Cols R–S, and aggregated scores in Cols T–X) and 'Judges' (6-column roster: Judge Chip, Judge Name, Judge Email, Assigned Count, Private Judge Sheet Link, Email Status)."],
    ["ANNUAL SETUP", "New Event Workflow", "Copy the master workbook, paste or link new submissions into 'Judging_master' (keeping personal info in Cols B–H and project details in Cols I–Q), and add judges to Column A of the 'Judges' tab."],
    ["WORKFLOW STEPS", "Running the Scripts", "Step 1: Prompts for Event Manager email, extracts judge names/emails, and assigns 2 judges per entry. Step 2: Creates private judge sheets. Step 3: Shares sheets and emails links. Step 4: Collects and averages scores into Cols T–X."],
    ["TROUBLESHOOTING", "Common Fixes", "Refresh browser if 'Judging Tools' menu is missing; authorize permissions on first run; use 'Grant Event Manager Access Now' to add/update an event manager; use 'Clear Judge Assignments' before reshuffling."]
  ];

  const range = sheet.getRange(1, 1, rows.length, 3);
  range.setValues(rows);
  range.setWrap(true);
  range.setVerticalAlignment("top");
  range.setFontSize(10);

  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold").setFontSize(11);
  sheet.setFrozenRows(1);

  sheet.getRange(2, 1, rows.length - 1, 1).setFontWeight("bold").setBackground("#e8f0fe").setFontColor("#174ea6");
  sheet.getRange(2, 2, rows.length - 1, 1).setFontWeight("bold");

  sheet.setColumnWidth(1, 210);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 620);

  SpreadsheetApp.getUi().alert("Instructions Tab Ready!", "The 'Instructions' tab has been created and formatted.", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Clears Judge 1 and Judge 2 assignments (Columns R & S) on 'Judging_master'.
 */
function clearAssignments() {
  const masterSheet = getMasterSheet(SpreadsheetApp.getActiveSpreadsheet());
  if (masterSheet) {
    masterSheet.getRange(2, 18, masterSheet.getMaxRows() - 1, 2).clearContent();
  }
}
