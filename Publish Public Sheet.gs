// ==========================================
// PUBLISH PAYOUT TO PUBLIC SHEET
// ==========================================
function publishPayoutToPublic() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "📢 PUBLISH TO PUBLIC SHEET", 
    "This will create a static, visual copy of your current Payouts tab and place it at the FRONT of your Public Spreadsheet.\n\nProceed?", 
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Connecting to Public Sheet...", "System", 3);
  
  // --- 1. GET THE PUBLIC SHEET ID FROM CONFIG B8 ---
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  let rawId = configSheet ? configSheet.getRange("B8").getValue().toString().trim() : "";
  
  if (!rawId) {
    ui.alert("⚠️ Missing Public Sheet ID!\n\nPlease enter your external Public Spreadsheet ID or URL in Config tab cell B8.");
    return;
  }

  // AUTO-EXTRACTOR: In case you pasted the full URL
  let publicId = rawId;
  if (rawId.includes("/d/")) {
    let match = rawId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match) publicId = match[1];
  }

  let externalSS;
  try {
    externalSS = SpreadsheetApp.openById(publicId);
  } catch (e) {
    ui.alert(`⚠️ Access Error!\n\nCould not open the public spreadsheet. \n\nMake sure the ID is correct and the account running this script has Editor permissions for it.`);
    return;
  }

  // --- 2. GET ENEMY FACTION NAME ---
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  let dashData = dashSheet.getDataRange().getValues();
  let enemyName = "Public_Payout_" + new Date().toLocaleDateString().replace(/\//g, "-"); 
  
  for (let r = 0; r < dashData.length; r++) {
    for (let c = 0; c < dashData[r].length; c++) {
      let cellText = dashData[r][c] ? dashData[r][c].toString().trim().toLowerCase() : "";
      if (cellText === "enemy faction name" && c + 1 < dashData[r].length) {
        let foundName = dashData[r][c+1] ? dashData[r][c+1].toString().trim() : "";
        if (foundName && foundName !== "Unknown") {
          enemyName = foundName.substring(0, 95).replace(/[\[\]*?:\/\\]/g, "");
        }
      }
    }
  }

  // --- 3. COPY & MOVE THE SHEET ---
  const payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  if (!payoutSheet) {
    ui.alert("Error: Could not find the local Payouts sheet.");
    return;
  }

  ss.toast(`Publishing data for ${enemyName}...`, "System", 5);

  let existingSheet = externalSS.getSheetByName(enemyName);
  if (existingSheet) {
    externalSS.deleteSheet(existingSheet);
  }

  // Create an exact replica of the Payout sheet in the public workbook
  let copiedSheet = payoutSheet.copyTo(externalSS);
  copiedSheet.setName(enemyName);

  // Moves the new sheet to the front
  externalSS.setActiveSheet(copiedSheet);
  externalSS.moveActiveSheet(1);
  
  // --- 4. FLATTEN THE SHEET (THE FIX) ---
  // Using exact matrix dimensions ensures we capture all columns (even hidden ones like Total Points) 
  // and overwrites any live formulas that would break without the Dashboard tab present.
  let maxRows = payoutSheet.getLastRow();
  let maxCols = payoutSheet.getLastColumn();
  
  if (maxRows > 0 && maxCols > 0) {
    let localValues = payoutSheet.getRange(1, 1, maxRows, maxCols).getValues();
    
    // Paste the local values directly over the copied sheet, locking in the math forever
    copiedSheet.getRange(1, 1, maxRows, maxCols).setValues(localValues);
  }

  ui.alert(`✅ Success!\n\nYour payout for [${enemyName}] has been successfully published with locked values!`);
}