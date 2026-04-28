function logWarToHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  const payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  let historySheet = ss.getSheetByName(SETTINGS.historySheet);

  if (!dashSheet || !payoutSheet) {
    ss.toast("Missing Dashboard or Payouts tab.", "Error", 5);
    return;
  }

  // 1. Ensure History sheet exists and has headers
  if (!historySheet) {
    historySheet = ss.insertSheet(SETTINGS.historySheet);
    historySheet.getRange(1, 1, 1, SETTINGS.historyHeaders.length)
                .setValues([SETTINGS.historyHeaders])
                .setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
    historySheet.setFrozenRows(1);
  }

  ss.toast("Gathering data and translating faction name...", "System", 3);

  // 2. Gather Variables from the Dashboard
  const warId = dashSheet.getRange("C7").getValue().toString().trim() || "Unknown";
  const targetFactionId = dashSheet.getRange("C3").getValue().toString().trim(); // Here is the missing variable!
  
  // Use the Config settings for Outcome/Termed, fallback to C4/C5 if not set
  const outcomeCell = SETTINGS.dashOutcomeCell || "C4";
  const termedCell = SETTINGS.dashTermedCell || "C5";
  const outcome = dashSheet.getRange(outcomeCell).getValue().toString().trim() || "N/A";
  const termed = dashSheet.getRange(termedCell).getValue().toString().trim() || "N/A";

  const totalRevenue = parseFloat(dashSheet.getRange("L3").getValue()) || 0;
  const medsCost = parseFloat(dashSheet.getRange("L6").getValue()) || 0;
  const bountiesCost = parseFloat(dashSheet.getRange("L7").getValue()) || 0;
  const totalCosts = medsCost + bountiesCost;
  const netProfit = parseFloat(dashSheet.getRange("L9").getValue()) || 0;
  const totalPayout = parseFloat(dashSheet.getRange("L13").getValue()) || 0;

  // 3. Gather Weights from Payouts Tab (Starts at Col E / Index 5)
  let weights = new Array(10).fill(0);
  try {
    weights = payoutSheet.getRange(1, 5, 1, 10).getValues()[0];
  } catch(e) {
    ss.toast("Could not read weights from Payouts tab.", "Warning", 3);
  }

  // 4. API TRANSLATOR: Convert ID to Name
  const apiKey = configSheet ? configSheet.getRange(SETTINGS.apiKeyCell).getValue().toString().trim() : "";
  let enemyFactionName = targetFactionId;

  if (apiKey && targetFactionId && targetFactionId !== "None Set") {
    try {
      const url = `https://api.torn.com/faction/${targetFactionId}?selections=basic&key=${apiKey}`;
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(response.getContentText());

      if (json && json.name) {
        enemyFactionName = `${json.name} [${targetFactionId}]`;
      }
    } catch (e) {
      // Silently fail and fallback to logging the ID if the connection drops
    }
  }

  // 5. Build the Array and Log it
  const today = new Date();
  const logData = [
    today, 
    warId, 
    enemyFactionName, 
    outcome, 
    termed, 
    totalRevenue, 
    totalCosts, 
    netProfit, 
    totalPayout, 
    ...weights
  ];

  historySheet.appendRow(logData);

  // 6. Format the newly appended row
  const lastRow = historySheet.getLastRow();
  historySheet.getRange(lastRow, 1).setNumberFormat("m/d/yyyy"); // Format Date
  historySheet.getRange(lastRow, 6, 1, 4).setNumberFormat('"$ "#,##0'); // Format Revenue through Payout

  ss.toast("War snapshot successfully logged to History!", "Success", 5);
}