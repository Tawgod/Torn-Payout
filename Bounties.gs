function setupBountyTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS.bountySheet);
  const wantedHeaders = SETTINGS.bountyHeaders || [
    "Date Logged", "Placed By", "Target", "Quantity", "Bounty Amount",
    "Refund Amount", "Status", "Notes", "Source", "Torn News ID"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS.bountySheet, 4);
    sheet.setHiddenGridlines(true);
    sheet.getRange(1, 1, 1, wantedHeaders.length).setValues([wantedHeaders]);
  } else {
    upgradeBountyTrackerSchema_(sheet, wantedHeaders);
  }

  const headerMap = headerMapFromRow_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  const headerRange = sheet.getRange(1, 1, 1, wantedHeaders.length);
  headerRange.setBackground("#4a86e8")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);

  const dateCol = headerIndex_(headerMap, "Date Logged", true) + 1;
  const qtyCol = headerIndex_(headerMap, "Quantity", true) + 1;
  const amountCol = headerIndex_(headerMap, "Bounty Amount", true) + 1;
  const refundCol = headerIndex_(headerMap, "Refund Amount", true) + 1;
  const statusCol = headerIndex_(headerMap, "Status", true) + 1;

  sheet.getRange(2, dateCol, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setNumberFormat("m/d/yyyy h:mm am/pm");
  sheet.getRange(2, qtyCol, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setNumberFormat("0");
  sheet.getRange(2, amountCol, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setNumberFormat('"$ "#,##0');
  sheet.getRange(2, refundCol, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setNumberFormat('"$ "#,##0');

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pending Review", "Approved", "Denied / Invalid", "Paid Out"])
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, statusCol, Math.max(1, sheet.getMaxRows() - 1), 1).setDataValidation(statusRule);

  // Quantity defaults to one for existing/manual rows when other bounty fields are populated.
  const lastRow = Math.max(2, sheet.getLastRow());
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, wantedHeaders.length).getValues();
    const placerIdx = headerIndex_(headerMap, "Placed By", true);
    const targetIdx = headerIndex_(headerMap, "Target", true);
    const qtyIdx = headerIndex_(headerMap, "Quantity", true);
    for (let i = 0; i < data.length; i++) {
      if ((data[i][placerIdx] || data[i][targetIdx]) && !data[i][qtyIdx]) {
        sheet.getRange(i + 2, qtyIdx + 1).setValue(1);
      }
    }
  }

  const widths = {
    "Date Logged": 150, "Placed By": 160, "Target": 160, "Quantity": 85,
    "Bounty Amount": 130, "Refund Amount": 130, "Status": 140,
    "Notes": 250, "Source": 120, "Torn News ID": 120
  };
  wantedHeaders.forEach((name, i) => sheet.setColumnWidth(i + 1, widths[name] || 120));

  for (let r = 2; r <= Math.min(100, sheet.getMaxRows()); r++) {
    sheet.getRange(r, 1, 1, wantedHeaders.length)
      .setBackground(r % 2 === 0 ? "#f8f9fa" : "#ffffff");
  }

  ss.toast("Bounty Tracker ready: quantity + bot/news source fields enabled.", "Success", 5);
}

function upgradeBountyTrackerSchema_(sheet, wantedHeaders) {
  const existingWidth = Math.max(1, sheet.getLastColumn());
  const existingHeaders = sheet.getRange(1, 1, 1, existingWidth).getDisplayValues()[0]
    .map(v => String(v || "").trim());

  if (!existingHeaders.some(Boolean)) {
    sheet.getRange(1, 1, 1, wantedHeaders.length).setValues([wantedHeaders]);
    return;
  }

  const existingMap = {};
  existingHeaders.forEach((h, i) => { if (h) existingMap[normalizeHeader_(h)] = i; });
  const lastRow = Math.max(1, sheet.getLastRow());
  const oldData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, existingWidth).getValues() : [];

  const newRows = oldData.map(row => wantedHeaders.map(header => {
    const oldIdx = existingMap[normalizeHeader_(header)];
    if (oldIdx !== undefined) return row[oldIdx];
    if (header === "Quantity" && (row[existingMap["placed by"]] || row[existingMap["target"]])) return 1;
    if (header === "Source" && (row[existingMap["placed by"]] || row[existingMap["target"]])) return "Sheet";
    return "";
  }));

  sheet.clearContents();
  sheet.getRange(1, 1, 1, wantedHeaders.length).setValues([wantedHeaders]);
  if (newRows.length) sheet.getRange(2, 1, newRows.length, wantedHeaders.length).setValues(newRows);
}
