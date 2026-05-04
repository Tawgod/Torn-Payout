function buildPayoutTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(SETTINGS.rosterSheet);
  let payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  const dashboardName = SETTINGS.dashboardSheet;

  if (!rosterSheet) {
    ss.toast("Please run 'Update Faction Roster' first!", "Error", 5);
    return;
  }

  const metricsList = [
    "War Hits", "War Assists", "War Losses", "War Interruptions", 
    "Outside / Chain Hits", "Chain Saves", "Retaliations", 
    "Net Respect", "War Abroad Hits", "War Score"
  ];

  // 1. PRESERVE EXISTING WEIGHTS (Safely)
  let savedWeights = [];
  if (payoutSheet) {
    try {
      let weights = payoutSheet.getRange(1, 5, 1, metricsList.length).getValues()[0];
      for (let i = 0; i < metricsList.length; i++) {
        let w = weights[i];
        savedWeights.push((w !== "" && w !== undefined && !isNaN(w)) ? w : 1.0);
      }
    } catch(e) {}
  }

  if (savedWeights.length === 0 || savedWeights.length !== metricsList.length) {
    savedWeights = new Array(metricsList.length).fill(1.0);
  }

  // 2. GET ACTIVE ROSTER
  const rosterData = rosterSheet.getDataRange().getValues();
  const activeMembers = [];
  for (let i = 1; i < rosterData.length; i++) {
    activeMembers.push([rosterData[i][0], rosterData[i][1]]); 
  }

  // 3. BUILD THE SHEET
  if (!payoutSheet) {
    payoutSheet = ss.insertSheet(SETTINGS.payoutSheet, 3);
  } else {
    payoutSheet.clear();
  }
  
  payoutSheet.setHiddenGridlines(true);
  payoutSheet.setFrozenRows(2); 
  payoutSheet.setFrozenColumns(2); 

  // Row 1: Hit Checkers & Weights
  payoutSheet.getRange("B1").setValue("Total Hits").setFontWeight("bold").setFontColor("#cc0000").setHorizontalAlignment("right");
  
  payoutSheet.getRange("C1").setFormula(`=IFERROR(SUM(E3:E500) + SUM(I3:I500), 0)`)
             .setFontWeight("bold").setFontSize(14).setFontColor("#cc0000")
             .setNumberFormat('#,##0'); 

  payoutSheet.getRange("D1").setValue("Set Point Weight ➡️").setFontStyle("italic").setHorizontalAlignment("right").setFontWeight("bold");
  payoutSheet.getRange(1, 5, 1, metricsList.length).setValues([savedWeights]);

  // Row 2: Main Data Headers 
  let row2 = new Array(20).fill("");
  row2[0] = "Member ID";
  row2[1] = "Name";
  row2[2] = "Contribution %";
  row2[3] = "Payout";
  metricsList.forEach((m, index) => row2[4 + index] = m); 
  row2[19] = "Total Points"; 
  payoutSheet.getRange(2, 1, 1, 20).setValues([row2]);

  if (activeMembers.length > 0) {
    payoutSheet.getRange(3, 1, activeMembers.length, 2).setValues(activeMembers);
  }

  // 4. HEADER & INPUT FORMATTING
  payoutSheet.getRange(2, 1, 1, 14).setBackground("#274e13").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  payoutSheet.getRange(2, 20).setBackground("#274e13").setFontColor("white"); 
  
  payoutSheet.getRange(1, 5, 1, metricsList.length)
             .setBackground("#ffffff")
             .setNumberFormat('0.00')
             .setHorizontalAlignment("center")
             .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  // 5. MASTER PROPORTIONAL FORMULAS (Unbound to Row 500)
  if (activeMembers.length > 0) {
    for (let r = 0; r < activeMembers.length; r++) {
      let rowNum = r + 3;
      
      let scoreFormula = `=IFERROR(SUMPRODUCT($E$1:$N$1, E${rowNum}:N${rowNum}), 0)`;
      
      // FIX: Expanded the sum boundary from maxRow to 500 to automatically capture Left Faction players
      let contribFormula = `=IFERROR(IF(SUM($T$3:$T$500)>0, T${rowNum} / SUM($T$3:$T$500), 0), 0)`;
      let payoutFormula = `=IFERROR(IF(C${rowNum}>0, C${rowNum} * '${dashboardName}'!$I$13, 0), 0)`;
      
      payoutSheet.getRange(rowNum, 20).setFormula(scoreFormula); 
      payoutSheet.getRange(rowNum, 3).setFormula(contribFormula).setNumberFormat('0.00%'); 
      payoutSheet.getRange(rowNum, 4).setFormula(payoutFormula).setNumberFormat('"$ "#,##0'); 
      
      let bgColor = (r % 2 === 0) ? "#ffffff" : "#f1f3f4"; 
      let moneyColor = (r % 2 === 0) ? "#e6f4ea" : "#ceead6"; 
      
      payoutSheet.getRange(rowNum, 1, 1, 14).setBackground(bgColor).setHorizontalAlignment("center");
      payoutSheet.getRange(rowNum, 1, 1, 2).setHorizontalAlignment("left"); 
      payoutSheet.getRange(rowNum, 4).setBackground(moneyColor).setFontWeight("bold"); 
    }
    
    payoutSheet.getRange(2, 1, activeMembers.length + 1, 14).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

    payoutSheet.getRange(3, 5, activeMembers.length, 10).setNumberFormat('#,##0'); 
    payoutSheet.getRange(3, 12, activeMembers.length, 1).setNumberFormat('#,##0.00'); 
    payoutSheet.getRange(3, 14, activeMembers.length, 1).setNumberFormat('#,##0.00'); 
  }

  // 6. CUSTOM COLUMN SIZING
  payoutSheet.setColumnWidth(2, 160); 
  payoutSheet.setColumnWidth(3, 115); 
  payoutSheet.setColumnWidth(4, 130); 
  
  for (let c = 5; c <= 14; c++) { payoutSheet.setColumnWidth(c, 115); }
  payoutSheet.setColumnWidth(8, 140); 
  payoutSheet.setColumnWidth(9, 150); 
  payoutSheet.setColumnWidth(13, 130); 

  payoutSheet.hideColumns(1); 
  payoutSheet.hideColumns(20); 

  ss.toast("Payout Tab Restored! Crunching metrics...", "System", 3);
  SpreadsheetApp.flush();
  
  // Triggers the script we just perfected!
  if (typeof runPayoutMath === "function") {
    runPayoutMath(); 
  }
}