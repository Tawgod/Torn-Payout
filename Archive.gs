function archiveAndResetWar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  const payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  
  if (!configSheet || !dashSheet || !payoutSheet) {
    ss.toast("Missing necessary tabs to perform archive.", "Error", 5);
    return;
  }

  const archiveId = configSheet.getRange(SETTINGS.archiveIdCell).getValue().toString().trim();
  if (!archiveId) {
    ss.toast("Archive ID missing in Config B3.", "Error", 5);
    return;
  }

  let archiveSS;
  try {
    archiveSS = SpreadsheetApp.openById(archiveId);
  } catch(e) {
    ss.toast("Could not open Archive Spreadsheet.", "Error", 5);
    return;
  }

  ss.toast("Generating permanent archive snapshot...", "System", 3);

  // --- 2. GATHER DATA (Updated for New Row Layout) ---
  const enemyId   = dashSheet.getRange("C3").getValue().toString().trim() || "Unknown";
  const startDate = dashSheet.getRange("C4").getValue().toString().trim() || "Unknown";
  const endDate   = dashSheet.getRange("C5").getValue().toString().trim() || "Ongoing"; // NEW
  const outcome   = dashSheet.getRange("C10").getValue().toString().trim() || "N/A";   // SHIFTED
  const termed    = dashSheet.getRange("C15").getValue().toString().trim() || "N/A";   // SHIFTED

  const apiKey = configSheet.getRange(SETTINGS.apiKeyCell).getValue().toString().trim();
  let enemyName = enemyId;
  if (apiKey && enemyId && enemyId !== "None Set") {
    try {
      const url = `https://api.torn.com/faction/${enemyId}?selections=basic&key=${apiKey}`;
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(response.getContentText());
      if (json && json.name) { enemyName = `${json.name} [${enemyId}]`; }
    } catch (e) {}
  }

  // Master Header now includes End Date
  const masterHeader = `  ⚔️ vs ${enemyName}  |  🏆 ${outcome}  |  🤝 Termed: ${termed}  |  📅 ${startDate} to ${endDate}`;

  // --- 3. EXTRACT VALUES ---
  const pRange = payoutSheet.getDataRange();
  const pData = pRange.getValues(); 
  const pColors = pRange.getBackgrounds();
  const pFontColors = pRange.getFontColors();
  const pFontWeights = pRange.getFontWeights();
  const pFormats = pRange.getNumberFormats();

  // --- 4. CREATE ARCHIVE TAB ---
  const dateString = new Date().toLocaleDateString().replace(/\//g, "-");
  let newSheetName = `${enemyName} (${dateString})`;
  let counter = 1;
  while (archiveSS.getSheetByName(newSheetName)) {
    newSheetName = `${enemyName} (${dateString}) v${counter}`;
    counter++;
  }
  
  const newArchiveSheet = archiveSS.insertSheet(newSheetName, 0);
  newArchiveSheet.setHiddenGridlines(true);

  // --- 5. STYLE HEADER ---
  newArchiveSheet.getRange("B1").setValue(masterHeader).setBackground("#000000").setFontColor("#ffffff").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left").setVerticalAlignment("middle");
  newArchiveSheet.getRange(1, 1, 1, pData[0].length).setBackground("#000000");

  // --- 6. PASTE DATA ---
  const targetRange = newArchiveSheet.getRange(2, 1, pData.length, pData[0].length);
  targetRange.setValues(pData).setBackgrounds(pColors).setFontColors(pFontColors).setFontWeights(pFontWeights).setNumberFormats(pFormats);

  // --- 7. FORMATTING ---
  newArchiveSheet.setColumnWidth(2, 160); 
  newArchiveSheet.setColumnWidth(3, 115); 
  newArchiveSheet.setColumnWidth(4, 130); 
  for (let c = 5; c <= 14; c++) { newArchiveSheet.setColumnWidth(c, 115); }
  newArchiveSheet.setFrozenRows(3); 
  newArchiveSheet.setFrozenColumns(2); 
  newArchiveSheet.hideColumns(1); 
  newArchiveSheet.hideColumns(20); 

  ss.toast("Archive Successful! Scrubbing workspace...", "System", 3);

  // --- 8. RESET EVERYTHING ---
  let rdSheet = ss.getSheetByName(SETTINGS.rdSheet);
  if (rdSheet) {
    rdSheet.clear();
    rdSheet.getRange("A1:R1").setValues([SETTINGS.rdHeaders]).setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
  }

  const lastP = payoutSheet.getLastRow();
  if (lastP >= 3) {
    payoutSheet.getRange(3, 4, lastP - 2, 11).clearContent(); 
    const names = payoutSheet.getRange(1, 2, lastP, 1).getValues();
    for (let i = names.length - 1; i >= 0; i--) {
      if (names[i][0] === "⚠️ NON-ROSTER / LEFT FACTION") { payoutSheet.deleteRow(i + 1); }
    }
  }

  let bountySheet = ss.getSheetByName(SETTINGS.bountySheet);
  if (bountySheet) {
    const bLast = bountySheet.getLastRow();
    if (bLast > 1) bountySheet.getRange(2, 1, bLast - 1, 7).clearContent();
  }

  dashSheet.getRange("C3").setValue("None Set");
  dashSheet.getRange("C4:C5").setValue("No Data"); // Clears Start and End
  dashSheet.getRange("C6").setValue("No Data");    // Clears War ID
  dashSheet.getRange("C10").setValue("Ongoing");
  dashSheet.getRange("C15").setValue("No");
  dashSheet.getRange("C11:C13").clearContent();
  dashSheet.getRange("L3:L8").clearContent(); 
  
  ss.toast("Ready for the next one!", "Success", 5);
}