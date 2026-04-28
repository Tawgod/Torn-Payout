// ==========================================
// 1. MASTER DASHBOARD BUILDER (With War End & Wide Column I)
// ==========================================
function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rdSheet = ss.getSheetByName(SETTINGS.rdSheet);
  let dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  
  if (!rdSheet) {
    ss.toast("No RD sheet found to build from. Pull data first!", "Error", 5);
    return;
  }

  // --- Extract Existing Data (Safe Refresh) ---
  let targetFactionId = "", officialStartTime = "No Data", officialEndTime = "No Data", activeWarId = "No Data";
  let filterTotal = "", filterWar = "", filterChain = "";
  let custStartDate = "", custStartTime = "00:00", custEndDate = "", custEndTime = "00:00";
  let warResult = "Ongoing", termedEarly = "No"; 
  let cachesWon = "", cacheVal = "", actualCacheVal = ""; 
  let fRev = 0, fTemp = 0, fRevive = 0, fXanax = 0, fOther = 0; 
  let fCutPct = 0.10, fCutMax = ""; 
  
  if (dashSheet) {
    let tempId = dashSheet.getRange("C3").getValue().toString().trim();
    if (tempId !== "None Set" && tempId !== "") targetFactionId = tempId;
    
    officialStartTime = dashSheet.getRange("C4").getValue();
    officialEndTime = dashSheet.getRange("C5").getValue();
    activeWarId = dashSheet.getRange("C6").getValue();
    
    try {
      let tempRes = dashSheet.getRange("C10").getValue(); // Shifted
      if (tempRes !== "") warResult = tempRes;
      
      cachesWon = dashSheet.getRange("C11").getValue();
      cacheVal = dashSheet.getRange("C12").getValue();
      let tempActual = dashSheet.getRange("C13").getValue();
      if (tempActual !== "") actualCacheVal = tempActual;
      
      let tempTerm = dashSheet.getRange("C15").getValue(); // Shifted
      if (tempTerm !== "") termedEarly = tempTerm;
    } catch(e) {}

    filterTotal = dashSheet.getRange("I3").getValue();
    filterWar = dashSheet.getRange("I4").getValue();
    filterChain = dashSheet.getRange("I5").getValue();
    
    try {
      let tsd = dashSheet.getRange("I8").getValue();
      if (tsd !== "") custStartDate = tsd;
      let tst = dashSheet.getRange("I9").getValue();
      if (tst !== "") custStartTime = tst;
      let ted = dashSheet.getRange("I10").getValue();
      if (ted !== "") custEndDate = ted;
      let tet = dashSheet.getRange("I11").getValue();
      if (tet !== "") custEndTime = tet;
    } catch(e) {}

    try {
      fRev = parseFloat(dashSheet.getRange("L3").getValue()) || 0;
      fTemp = parseFloat(dashSheet.getRange("L4").getValue()) || 0;
      fRevive = parseFloat(dashSheet.getRange("L5").getValue()) || 0;
      fXanax = parseFloat(dashSheet.getRange("L6").getValue()) || 0;
      fOther = parseFloat(dashSheet.getRange("L8").getValue()) || 0;
      let tempCutPct = dashSheet.getRange("L10").getValue();
      if (tempCutPct !== "") fCutPct = tempCutPct;
      let tempCutMax = dashSheet.getRange("L11").getValue();
      if (tempCutMax !== "") fCutMax = tempCutMax;
    } catch(e) {}
  }

  // --- Scan RD Tab ---
  const data = rdSheet.getDataRange().getValues();
  let warHits = 0, warRespect = 0, totalHits = 0, totalRespect = 0;
  let lastWarHit = null, firstHit = null, lastHit = null;

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let timestamp = new Date(row[2]); 
    let defFaction = row[8].toString().trim(); 
    let respect = parseFloat(row[10]) || 0; 

    totalHits++; totalRespect += respect;
    if (!firstHit || timestamp < firstHit) firstHit = timestamp;
    if (!lastHit || timestamp > lastHit) lastHit = timestamp;

    if (defFaction === targetFactionId && targetFactionId !== "") {
      warHits++; warRespect += respect;
      if (!lastWarHit || timestamp > lastWarHit) lastWarHit = timestamp;
    }
  }

  if (!dashSheet) {
    dashSheet = ss.insertSheet(SETTINGS.dashboardSheet, 1);
  } else {
    dashSheet.clear(); dashSheet.clearFormats(); dashSheet.getDataRange().clearDataValidations();
  }
  dashSheet.setHiddenGridlines(true);

  // --- 1. WAR INSIGHTS (Shifted for End Time) ---
  const warHeaders = [
    ["⚔️ WAR INSIGHTS", ""],
    ["Enemy Faction ID", targetFactionId || "None Set"],
    ["Official War Start", officialStartTime], 
    ["Official War End", officialEndTime], // NEW ROW
    ["War ID", activeWarId],  
    ["Last War Hit", lastWarHit || "No Data"],
    ["Total War Hits", warHits],
    ["War Respect Gained", warRespect],
    ["Outcome (Result)", warResult], 
    ["Caches Won", cachesWon], 
    ["Est. Cache Value", cacheVal],
    ["Actual Cache Value", actualCacheVal], 
    ["Difference", '=IF(C13="", "", C13 - C12)'],
    ["Termed?", termedEarly] 
  ];
  dashSheet.getRange("B2:C15").setValues(warHeaders);
  dashSheet.getRange("B2:C2").setBackground("#cc0000").setFontColor("white").setFontWeight("bold").merge();
  dashSheet.getRange("B3:C15").setBackground("#f4cccc");
  dashSheet.getRange("B12:C12").setBackground("#ea9999").setFontWeight("bold"); // Cache Val highlight
  dashSheet.getRange("C13").setBackground("#ffffff").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  dashSheet.getRange("C10").setBackground("#ffffff").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID); // Outcome
  dashSheet.getRange("C15").setBackground("#ffffff").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID); // Termed

  const outcomeRule = SpreadsheetApp.newDataValidation().requireValueInList(["Win", "Loss", "Draw", "Ongoing"]).build();
  dashSheet.getRange("C10").setDataValidation(outcomeRule);
  const termedRule = SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"]).build();
  dashSheet.getRange("C15").setDataValidation(termedRule);

  // --- 2. CHAIN INSIGHTS ---
  const chainHeaders = [
    ["🔗 CHAIN INSIGHTS", ""],
    ["First Attack Logged", firstHit || "No Data"],
    ["Latest Attack Logged", lastHit || "No Data"],
    ["Total Attacks Logged", totalHits],
    ["Total Respect Generated", totalRespect]
  ];
  dashSheet.getRange("E2:F6").setValues(chainHeaders);
  dashSheet.getRange("E2:F2").setBackground("#3c78d8").setFontColor("white").setFontWeight("bold").merge();
  dashSheet.getRange("E3:F6").setBackground("#c9daf8");

  // --- 3. PAYOUT FILTERS ---
  const filterHeaders = [
    ["⚙️ PAYOUT FILTERS", ""],
    ["Total Hits (Max Limit)", filterTotal],
    ["Max War Hits", filterWar],
    ["Max Chain Hits", filterChain]     
  ];
  dashSheet.getRange("H2:I5").setValues(filterHeaders);
  dashSheet.getRange("H2:I2").setBackground("#e69138").setFontColor("white").setFontWeight("bold").merge();
  dashSheet.getRange("H3:I5").setBackground("#fce5cd");
  dashSheet.getRange("I3:I5").setBackground("#ffffff").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

  // --- 4. CUSTOM TIME WINDOW ---
  const timeHeaders = [
    ["⏳ CUSTOM TIME WINDOW", ""],
    ["Start Date", custStartDate],
    ["Start Time", custStartTime],
    ["End Date", custEndDate],
    ["End Time", custEndTime]
  ];
  dashSheet.getRange("H7:I11").setValues(timeHeaders);
  dashSheet.getRange("H7:I7").setBackground("#b45f06").setFontColor("white").setFontWeight("bold").merge();
  dashSheet.getRange("H8:I11").setBackground("#f9cb9c");
  dashSheet.getRange("I8:I11").setBackground("#ffffff").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

  // CHANGE: We remove the strict dropdown and just ensure it's formatted as time
  dashSheet.getRange("I9").setDataValidation(null); // Clear old validation
  dashSheet.getRange("I11").setDataValidation(null); // Clear old validation
  
  // Set formatting to show HH:mm:ss
  dashSheet.getRange("I9").setNumberFormat("HH:mm:ss");
  dashSheet.getRange("I11").setNumberFormat("HH:mm:ss");

  // --- 5. FINANCIALS ---
  const bountyFormula = `=IFERROR(SUMIFS('${SETTINGS.bountySheet}'!E:E, '${SETTINGS.bountySheet}'!F:F, "Approved"), 0)`;
  const finHeaders = [
    ["💰 WAR FINANCIALS", ""],
    ["Total Revenue", fRev],
    ["- Temp Cost", fTemp],
    ["- Revives", fRevive],
    ["- Xanax", fXanax],
    ["- Approved Bounties", bountyFormula], 
    ["- Other Cost", fOther],
    ["NET PROFIT", "=IFERROR(L3 - SUM(L4:L8), 0)"], 
    ["Faction Cut %", fCutPct],
    ["Max Faction Cut ($)", fCutMax],
    ["Actual Faction Deduction", "=IF(L9>0, MIN(L9*L10, IF(L11=\"\", 999999999999, L11)), 0)"], 
    ["PAYOUT TOTAL", "=IFERROR(L9 - L12, 0)"] 
  ];
  dashSheet.getRange("K2:L13").setValues(finHeaders);
  dashSheet.getRange("K2:L2").setBackground("#274e13").setFontColor("white").setFontWeight("bold").merge();
  dashSheet.getRange("K3:L13").setBackground("#d9ead3");
  dashSheet.getRange("L3:L8").setBackground("#ffffff").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  dashSheet.getRange("L10:L11").setBackground("#ffffff").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  dashSheet.getRange("K9:L9").setFontWeight("bold").setBackground("#b6d7a8");
  dashSheet.getRange("K12:L13").setFontWeight("bold").setBackground("#b6d7a8");

  // --- 6. LEADERBOARDS ---
  const pSheet = SETTINGS.payoutSheet;
  const fWarHits = `=IFERROR(QUERY('${pSheet}'!B3:N, "SELECT B, E WHERE E > 0 ORDER BY E DESC LIMIT 3", 0), {"No Data", ""})`;
  const fRespect = `=IFERROR(QUERY('${pSheet}'!B3:N, "SELECT B, L WHERE L > 0 ORDER BY L DESC LIMIT 3", 0), {"No Data", ""})`;
  const fChainHits = `=IFERROR(QUERY('${pSheet}'!B3:N, "SELECT B, I WHERE I > 0 ORDER BY I DESC LIMIT 3", 0), {"No Data", ""})`;
  const fChainSaves = `=IFERROR(QUERY('${pSheet}'!B3:N, "SELECT B, J WHERE J > 0 ORDER BY J DESC LIMIT 3", 0), {"No Data", ""})`;
  const fContrib = `=IFERROR(QUERY('${pSheet}'!B3:T, "SELECT B, C WHERE C > 0 ORDER BY C DESC LIMIT 3", 0), {"No Data", ""})`;

  dashSheet.getRange("B17:C17").merge().setValue("🏆 TOP WAR HITS").setBackground("#f1c232").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("B18").setFormula(fWarHits);
  dashSheet.getRange("B22:C22").merge().setValue("⭐ TOP RESPECT").setBackground("#f1c232").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("B23").setFormula(fRespect);
  dashSheet.getRange("E17:F17").merge().setValue("🔗 TOP CHAIN HITS").setBackground("#f1c232").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("E18").setFormula(fChainHits);
  dashSheet.getRange("E22:F22").merge().setValue("🚑 TOP CHAIN SAVES").setBackground("#f1c232").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("E23").setFormula(fChainSaves);
  dashSheet.getRange("K17:L17").merge().setValue("🔥 TOP CONTRIBUTION").setBackground("#f1c232").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("K18").setFormula(fContrib);
  dashSheet.getRange("B18:C20").setBackground("#fff2cc"); dashSheet.getRange("B23:C25").setBackground("#fff2cc"); dashSheet.getRange("E18:F20").setBackground("#fff2cc"); dashSheet.getRange("E23:F25").setBackground("#fff2cc"); dashSheet.getRange("K18:L20").setBackground("#fff2cc");

  // --- Formatting (Strict 24-Hour Torn Time) ---
  // Official War Start, End, and Last War Hit
  dashSheet.getRange("C4:C5").setNumberFormat("yyyy-mm-dd HH:mm:ss"); 
  dashSheet.getRange("C7").setNumberFormat("yyyy-mm-dd HH:mm:ss"); 

  // First/Latest Attack Logged
  dashSheet.getRange("F3:F4").setNumberFormat("yyyy-mm-dd HH:mm:ss"); 

  // Custom Time Window (Dates and Times)
  dashSheet.getRange("I8").setNumberFormat("yyyy-mm-dd"); 
  dashSheet.getRange("I9").setNumberFormat("HH:mm:ss"); 
  dashSheet.getRange("I10").setNumberFormat("yyyy-mm-dd"); 
  dashSheet.getRange("I11").setNumberFormat("HH:mm:ss"); 

  // Financials & Numbers
  dashSheet.getRange("C8:C9").setNumberFormat("#,##0"); 
  dashSheet.getRange("C12:C14").setNumberFormat('"$ "#,##0'); 
  dashSheet.getRange("L3:L9").setNumberFormat('"$ "#,##0');
  dashSheet.getRange("L10").setNumberFormat('0.00%'); 
  dashSheet.getRange("L11:L13").setNumberFormat('"$ "#,##0');

  // Column Width adjustments
  dashSheet.autoResizeColumns(2, 13); 
  dashSheet.setColumnWidth(9, 180); // Width for 24h Date Strings
  dashSheet.setColumnWidth(12, 180);

  ss.toast("Dashboard Rebuilt with War End Time & Wide Column I!", "System", 3);
}

// ==========================================
// 2. UPDATED API FETCHER (Handles End Time)
// ==========================================
function fetchActiveWarDetails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  if (!configSheet || !dashSheet) return;
  const apiKey = configSheet.getRange(SETTINGS.apiKeyCell).getValue().toString().trim();
  if (!apiKey) return;

  try {
    const url = `https://api.torn.com/faction/?selections=basic&key=${apiKey}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    const rankedWars = json.ranked_wars;
    if (!rankedWars || Object.keys(rankedWars).length === 0) return;

    const warId = Object.keys(rankedWars)[0];
    const warData = rankedWars[warId];
    let enemyId = "";
    for (let id in warData.factions) { if (id !== json.ID.toString()) { enemyId = id; break; } }

    let startDate = warData.war.start > 0 ? new Date(warData.war.start * 1000) : "TBD";
    let endDate = warData.war.end > 0 ? new Date(warData.war.end * 1000) : "Ongoing";

    dashSheet.getRange("C3").setValue(enemyId);        
    dashSheet.getRange("C4").setValue(startDate);      
    dashSheet.getRange("C5").setValue(endDate); // NEW       
    dashSheet.getRange("C6").setValue(warId);          
    dashSheet.getRange("C10").setValue("Ongoing");      
    dashSheet.getRange("C15").setValue("No");           
  } catch(e) {}
}

function syncDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  if (!dashSheet) return;
  ss.toast("Syncing live formulas and fetching API...", "System", 3);
  fetchActiveWarDetailsSafe(dashSheet);
}

function fetchActiveWarDetailsSafe(dashSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  const apiKey = configSheet.getRange(SETTINGS.apiKeyCell).getValue().toString().trim();
  if (!apiKey) return;
  try {
    const url = `https://api.torn.com/faction/?selections=basic&key=${apiKey}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    if (json.ranked_wars && Object.keys(json.ranked_wars).length > 0) {
      const warId = Object.keys(json.ranked_wars)[0];
      const warData = json.ranked_wars[warId];
      let enemyId = "";
      for (let id in warData.factions) { if (id !== json.ID.toString()) { enemyId = id; break; } }
      let startDate = warData.war.start > 0 ? new Date(warData.war.start * 1000) : "TBD";
      let endDate = warData.war.end > 0 ? new Date(warData.war.end * 1000) : "Ongoing";
      dashSheet.getRange("C3").setValue(enemyId);
      dashSheet.getRange("C4").setValue(startDate);
      dashSheet.getRange("C5").setValue(endDate); // NEW
      dashSheet.getRange("C6").setValue(warId);
    }
  } catch(e) {}
}