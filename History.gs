// ==========================================
// LOG WAR DATA TO HISTORY
// ==========================================
function logWarToHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  
  if (!dashSheet) {
    SpreadsheetApp.getUi().alert("Error: Dashboard not found.");
    return;
  }

  // --- Indestructible Scanner ---
  let allData = dashSheet.getDataRange().getValues();
  const findVal = (label) => {
    for (let r = 0; r < allData.length; r++) {
      for (let c = 0; c < allData[r].length; c++) {
        if (allData[r][c] === label && c + 1 < allData[r].length) return allData[r][c+1];
      }
    }
    return "";
  };

  // Grab all essential data in the new order
  let warId = findVal("War ID") || "Unknown";
  let enemyName = findVal("Enemy Faction Name") || "Unknown";
  let enemyId = findVal("Enemy Faction ID") || "";
  let start = findVal("Official War Start") || "";
  let end = findVal("Official War End") || "";
  let outcome = findVal("Outcome (Result)") || "Unknown";
  let termed = findVal("Termed?") || "No"; // Moved next to outcome
  let score = findVal("War Score") || 0;
  let hits = findVal("Total War Hits") || 0;
  
  // Financials
  let revenue = findVal("Total Revenue") || 0;
  let payoutTotal = findVal("PAYOUT TOTAL") || 0;
  let factionProfit = findVal("Actual Faction Deduction") || 0; // Added Faction Profit
  
  let caches = findVal("Caches / Items Won") || "";
  
  if (warId === "No Data" || warId === "") {
    SpreadsheetApp.getUi().alert("No War ID found on the Dashboard. Fetch official reports first.");
    return;
  }

  // Find the exact History Sheet
  let historySheet = ss.getSheetByName("History");
  
  if (!historySheet) {
    historySheet = ss.insertSheet("History");
    // New Header Layout
    let headers = [
      "Log Date", "War ID", "Enemy Name", "Enemy ID", "Outcome", "Termed", 
      "Score", "Hits", "Start", "End", "Total Revenue", "Payout Total", 
      "Faction Profit", "Caches Won"
    ];
    historySheet.appendRow(headers);
    historySheet.getRange("A1:N1").setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
    historySheet.setFrozenRows(1);
  }

  // Append the row mapping perfectly to the headers
  historySheet.appendRow([
    new Date(), warId, enemyName, enemyId, outcome, termed, score, hits, 
    start, end, revenue, payoutTotal, factionProfit, caches
  ]);

  let lastRow = historySheet.getLastRow();
  
  // Format the 3 currency columns (K, L, and M are columns 11, 12, 13)
  historySheet.getRange(lastRow, 11, 1, 3).setNumberFormat('"$ "#,##0');
  
  // Force text wrapping across the entire data range so caches don't overflow
  historySheet.getDataRange().setWrap(true);

  SpreadsheetApp.getUi().alert(`✅ Success!\n\nWar against [${enemyName}] has been permanently logged to the History tab.`);
}