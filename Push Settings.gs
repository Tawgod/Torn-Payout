// ==========================================
// PUSH SETTINGS — dynamic Dashboard block
// ==========================================

function readPushSettingsFromDashboard_(sheet) {
  if (!sheet) return [];
  const header = String(sheet.getRange("N2").getDisplayValue() || "");
  if (header.indexOf("PUSH SETTINGS") === -1) return [];

  const lastRow = Math.max(4, sheet.getLastRow());
  const values = sheet.getRange(4, 14, lastRow - 3, 3).getValues();
  let lastUsed = -1;
  for (let i = 0; i < values.length; i++) {
    if (values[i].some(v => v !== "" && v !== null)) lastUsed = i;
  }
  return lastUsed >= 0 ? values.slice(0, lastUsed + 1) : [];
}

function buildPushSettingsBlock_(sheet, rows) {
  rows = Array.isArray(rows) ? rows : [];
  const visibleRows = Math.max(3, rows.length + 1);
  const endRow = 3 + visibleRows;

  sheet.getRange("N2:P2").merge().setValue("⏩ PUSH SETTINGS")
    .setBackground("#8e7cc3").setFontColor("#ffffff")
    .setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange("N3:P3").setValues([["Start Time", "Start Date", "Time Limit (Min)"]])
    .setBackground("#d9d2e9").setFontWeight("bold").setHorizontalAlignment("center");

  const body = sheet.getRange(4, 14, visibleRows, 3);
  body.setBackground("#f4edf7")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("N2:P" + endRow)
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  if (rows.length) sheet.getRange(4, 14, rows.length, 3).setValues(rows);

  sheet.getRange(4, 14, visibleRows, 1).setNumberFormat("hh:mm");
  sheet.getRange(4, 15, visibleRows, 1).setNumberFormat("yyyy-mm-dd");
  sheet.getRange(4, 16, visibleRows, 1).setNumberFormat("0");
}

function ensurePushSettingsCapacity_(sheet, throughRow) {
  if (!sheet || throughRow < 4) return;
  const currentLast = Math.max(4, sheet.getLastRow());
  if (throughRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), throughRow - sheet.getMaxRows());
  }
  const start = Math.max(4, currentLast + 1);
  if (throughRow >= start) {
    const r = sheet.getRange(start, 14, throughRow - start + 1, 3);
    r.setBackground("#f4edf7")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(start, 14, throughRow - start + 1, 1).setNumberFormat("hh:mm");
    sheet.getRange(start, 15, throughRow - start + 1, 1).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(start, 16, throughRow - start + 1, 1).setNumberFormat("0");
  }
}

function addPushSetting(startTime, startDate, timeLimitMin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  if (!sheet) throw new Error("Dashboard sheet not found.");

  if (String(sheet.getRange("N2").getDisplayValue() || "").indexOf("PUSH SETTINGS") === -1) {
    buildPushSettingsBlock_(sheet, []);
  }

  let row = 4;
  while (row <= sheet.getMaxRows()) {
    const vals = sheet.getRange(row, 14, 1, 3).getValues()[0];
    if (vals.every(v => v === "" || v === null)) break;
    row++;
  }

  if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 3);
  ensurePushSettingsCapacity_(sheet, row + 1);

  sheet.getRange(row, 14, 1, 3).setValues([[startTime || "", startDate || "", timeLimitMin || ""]]);
  return row;
}

// Manual edits grow the Push Settings lane automatically.
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SETTINGS.dashboardSheet) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < 4 || col < 14 || col > 16) return;
    if (String(sheet.getRange("N2").getDisplayValue() || "").indexOf("PUSH SETTINGS") === -1) return;

    ensurePushSettingsCapacity_(sheet, row + 2);
  } catch (_e) {}
}
