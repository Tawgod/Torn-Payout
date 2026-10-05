// ==========================================
// PUSH SETTINGS — dynamic Dashboard block
// ==========================================

function readPushSettingsFromDashboard_(sheet) {
  if (!sheet) return [];
  let startRow = 17;
  let startCol = 8;
  let header = String(sheet.getRange("H15").getDisplayValue() || "");

  // Migration fallback: prior test layout used N:P; original workbook used K:M.
  if (header.indexOf("PUSH SETTINGS") === -1) {
    if (String(sheet.getRange("N2").getDisplayValue() || "").indexOf("PUSH SETTINGS") !== -1) {
      startRow = 4; startCol = 14;
    } else if (String(sheet.getRange("K2").getDisplayValue() || "").indexOf("PUSH SETTINGS") !== -1) {
      startRow = 4; startCol = 11;
    } else {
      return [];
    }
  }

  const lastRow = Math.max(startRow, sheet.getLastRow());
  const values = sheet.getRange(startRow, startCol, lastRow - startRow + 1, 3).getValues();
  let lastUsed = -1;
  for (let i = 0; i < values.length; i++) {
    if (values[i].some(v => v !== "" && v !== null)) lastUsed = i;
  }
  return lastUsed >= 0 ? values.slice(0, lastUsed + 1) : [];
}

function buildPushSettingsBlock_(sheet, rows) {
  rows = Array.isArray(rows) ? rows : [];
  const visibleRows = Math.max(3, rows.length + 1);
  const startRow = 17;
  const endRow = startRow + visibleRows - 1;

  sheet.getRange("H15:J15").merge().setValue("⏩ PUSH SETTINGS")
    .setBackground("#8e7cc3").setFontColor("#ffffff")
    .setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange("H16:J16").setValues([["Start Time", "Start Date", "Time Limit (Min)"]])
    .setBackground("#d9d2e9").setFontWeight("bold").setHorizontalAlignment("center");

  const body = sheet.getRange(startRow, 8, visibleRows, 3);
  body.setBackground("#f4edf7")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("H15:J" + endRow)
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  if (rows.length) sheet.getRange(startRow, 8, rows.length, 3).setValues(rows);

  sheet.getRange(startRow, 8, visibleRows, 1).setNumberFormat("hh:mm");
  sheet.getRange(startRow, 9, visibleRows, 1).setNumberFormat("yyyy-mm-dd");
  sheet.getRange(startRow, 10, visibleRows, 1).setNumberFormat("0");
}

function ensurePushSettingsCapacity_(sheet, throughRow) {
  if (!sheet || throughRow < 17) return;
  if (throughRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), throughRow - sheet.getMaxRows());
  }
  const r = sheet.getRange(17, 8, throughRow - 16, 3);
  r.setBackground("#f4edf7")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(17, 8, throughRow - 16, 1).setNumberFormat("hh:mm");
  sheet.getRange(17, 9, throughRow - 16, 1).setNumberFormat("yyyy-mm-dd");
  sheet.getRange(17, 10, throughRow - 16, 1).setNumberFormat("0");
}

function addPushSetting(startTime, startDate, timeLimitMin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  if (!sheet) throw new Error("Dashboard sheet not found.");

  if (String(sheet.getRange("H15").getDisplayValue() || "").indexOf("PUSH SETTINGS") === -1) {
    buildPushSettingsBlock_(sheet, []);
  }

  let row = 17;
  while (row <= sheet.getMaxRows()) {
    const vals = sheet.getRange(row, 8, 1, 3).getValues()[0];
    if (vals.every(v => v === "" || v === null)) break;
    row++;
  }

  if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 3);
  ensurePushSettingsCapacity_(sheet, row + 1);

  sheet.getRange(row, 8, 1, 3).setValues([[startTime || "", startDate || "", timeLimitMin || ""]]);
  return row;
}

// Manual edits grow the Push Settings lane automatically.
function handlePushSettingsEdit_(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SETTINGS.dashboardSheet) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < 17 || col < 8 || col > 10) return;
    if (String(sheet.getRange("H15").getDisplayValue() || "").indexOf("PUSH SETTINGS") === -1) return;

    ensurePushSettingsCapacity_(sheet, row + 2);
  } catch (_e) {}
}
