// Google Apps Script — paste into script.google.com attached to a Google Sheet
//
// Deploy as: Web app → Execute as: Me → Who has access: Anyone

const SECRET_PIN = "YOUR_SECRET_PIN"; // set this to your chosen PIN

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.pin !== SECRET_PIN) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: "unauthorized" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dateStr = data.date; // "yyyy-MM-dd"
  const score = data.score != null ? data.score : "";
  const notes = data.notes || "";

  // Find existing row for this date
  const dates = sheet.getRange("A:A").getValues();
  let rowIndex = -1;
  for (let i = 0; i < dates.length; i++) {
    const cell = dates[i][0];
    const cellStr = cell instanceof Date ? Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(cell);
    if (cellStr === dateStr) {
      rowIndex = i + 1;
      break;
    }
  }

  const isEmpty = score === "" && !notes;

  if (rowIndex > 0 && isEmpty) {
    // Delete row if both score and notes are empty
    sheet.deleteRow(rowIndex);
  } else if (rowIndex > 0) {
    // Update existing row
    sheet.getRange(rowIndex, 2).setValue(score);
    sheet.getRange(rowIndex, 3).setValue(notes);
    sheet.getRange(rowIndex, 4).setValue(new Date());
  } else if (!isEmpty) {
    // Append new row only if there's data
    sheet.appendRow([dateStr, score, notes, new Date()]);
  }

  // Sort by date descending so newest entries are on top
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).sort({ column: 1, ascending: false });
  }

  return ContentService.createTextOutput(
    JSON.stringify({ status: "ok" })
  ).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const params = e.parameter;
  if (params.pin !== SECRET_PIN) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: "unauthorized" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const rows = [];

  // Skip header row if present
  const startRow = data.length > 0 && data[0][0] === "date" ? 1 : 0;
  for (let i = startRow; i < data.length; i++) {
    if (data[i][0]) {
      const cell = data[i][0];
      const dateStr = cell instanceof Date ? Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(cell);
      rows.push({
        date: dateStr,
        score: data[i][1],
        notes: data[i][2],
      });
    }
  }

  return ContentService.createTextOutput(
    JSON.stringify(rows)
  ).setMimeType(ContentService.MimeType.JSON);
}
