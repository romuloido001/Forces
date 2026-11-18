// Google Apps Script backend for the "Balanced and Unbalanced Forces" lesson.
// Setup: open your Google Sheet > Extensions > Apps Script > paste this file >
// Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
// Copy the Web app URL into SCRIPT_URL in index.html.

const SHEET_NAME = "Forces Results";
const HEADERS = ["Timestamp", "Name", "Section", "Classify (/10)", "Free-body (/7)", "Predict (/6)", "Quick check (/6)", "Total", "Max",
  "Written 1: Balanced forces", "Written 2: Unbalanced example", "Written 3: Moving but balanced", "Written 4: How motion changes", "Raw answers"];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    }
    const w = d.written || [];
    sh.appendRow([
      new Date(d.timestamp), d.name, d.section,
      d.classify, d.fbd, d.predict, d.quiz, d.total, d.max,
      w[0] || "", w[1] || "", w[2] || "", w[3] || "",
      JSON.stringify(d.answers || {})
    ]);
    return ContentService.createTextOutput(JSON.stringify({ok: true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput("Forces lesson backend is running.");
}
