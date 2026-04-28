function setupBountyTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS.bountySheet);

  if (sheet) {
    ss.toast("Bounty tracker already exists!", "System", 3);
    return;
  }

  sheet = ss.insertSheet(SETTINGS.bountySheet, 4);
  sheet.setHiddenGridlines(true);

  const headers = SETTINGS.bountyHeaders;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground("#4a86e8") 
             .setFontColor("white")
             .setFontWeight("bold")
             .setHorizontalAlignment("center");
  
  sheet.setFrozenRows(1);

  // Formatting
  sheet.getRange("A2:A").setNumberFormat("m/d/yyyy h:mm am/pm"); 
  sheet.getRange("D2:D").setNumberFormat('"$ "#,##0');           // Bounty Amount
  sheet.getRange("E2:E").setNumberFormat('"$ "#,##0');           // Refund Amount

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pending Review", "Approved", "Denied / Invalid", "Paid Out"])
    .setAllowInvalid(true)
    .build();
  sheet.getRange("F2:F").setDataValidation(statusRule);

  for (let r = 2; r <= 100; r++) {
    let bgColor = (r % 2 === 0) ? "#f8f9fa" : "#ffffff";
    sheet.getRange(r, 1, 1, headers.length).setBackground(bgColor);
  }

  sheet.setColumnWidth(1, 150); // Date
  sheet.setColumnWidth(2, 160); // Placed By
  sheet.setColumnWidth(3, 160); // Target
  sheet.setColumnWidth(4, 130); // Bounty Amount
  sheet.setColumnWidth(5, 130); // Refund Amount
  sheet.setColumnWidth(6, 140); // Status
  sheet.setColumnWidth(7, 250); // Notes

  ss.toast("Bounty Tracker Initialized with Refund terminology!", "Success", 5);
}