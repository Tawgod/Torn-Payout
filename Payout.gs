function buildPayoutTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(SETTINGS.rosterSheet);
  const dashboardName = SETTINGS.dashboardSheet;
  const dashSheet = ss.getSheetByName(dashboardName);
  let payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);

  if (!rosterSheet) {
    ss.toast("Please run 'Update Faction Roster' first!", "Error", 5);
    return;
  }

  // --- 1. DYNAMICALLY FETCH WATCH TIME LIMITS ---
  let t1 = "Tier 1 Saves", t2 = "Tier 2 Saves", t3 = "Tier 3 Saves";
  
  if (dashSheet) {
    let dashData = dashSheet.getDataRange().getValues();
    let tierRowStart = -1;
    let tierCol = -1;
    
    // Scan Dashboard to find the exact location of the watch limits
    for (let r = 0; r < dashData.length; r++) {
      for (let c = 0; c < dashData[r].length; c++) {
        let val = dashData[r][c] ? dashData[r][c].toString().toLowerCase().trim() : "";
        if (val === "watch time limit") {
          tierRowStart = r + 2; 
          tierCol = c + 1;      
          break;
        }
      }
      if (tierRowStart !== -1) break;
    }

    if (tierRowStart !== -1) {
      // getDisplayValue() grabs the exact visual text (e.g., "3:00") preventing decimal conversion
      let v1 = dashSheet.getRange(tierRowStart, tierCol).getDisplayValue();
      let v2 = dashSheet.getRange(tierRowStart + 1, tierCol).getDisplayValue();
      let v3 = dashSheet.getRange(tierRowStart + 2, tierCol).getDisplayValue();
      
      if (v1) t1 = v1 + " Saves";
      if (v2) t2 = v2 + " Saves";
      if (v3) t3 = v3 + " Saves";
    }
  }

  // ---> UPDATED: Metrics List now dynamically contains the time values natively
  const metricsList = [
    "War Hits", "War Assists", "War Losses", "War Interruptions", 
    "Outside / Chain Hits", "Chain Saves", "Retaliations", 
    "Net Respect", "War Abroad Hits", "War Score",
    t1, t2, t3
  ];

  // 2. PRESERVE EXISTING WEIGHTS (Safely)
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

  // 3. GET ACTIVE ROSTER
  const rosterData = rosterSheet.getDataRange().getValues();
  const activeMembers = [];
  for (let i = 1; i < rosterData.length; i++) {
    activeMembers.push([rosterData[i][0], rosterData[i][1]]); 
  }

  // 4. BUILD THE SHEET
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
  
  // Format O2:Q2 as Plain Text BEFORE pasting, preventing Sheets from changing times into decimals
  payoutSheet.getRange("O2:Q2").setNumberFormat("@");
  payoutSheet.getRange(2, 1, 1, 20).setValues([row2]);

  // Automatically define Named Ranges for your tiers so they can be referenced anywhere
  try {
    ss.setNamedRange("Tier1Label", payoutSheet.getRange("O2"));
    ss.setNamedRange("Tier2Label", payoutSheet.getRange("P2"));
    ss.setNamedRange("Tier3Label", payoutSheet.getRange("Q2"));
  } catch (e) {
    // Ignore if they already exist
  }

  if (activeMembers.length > 0) {
    payoutSheet.getRange(3, 1, activeMembers.length, 2).setValues(activeMembers);
  }

  // 5. HEADER & INPUT FORMATTING (Expanded to Column Q / 17 columns)
  payoutSheet.getRange(2, 1, 1, 17).setBackground("#274e13").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  payoutSheet.getRange(2, 20).setBackground("#274e13").setFontColor("white"); 
  
  payoutSheet.getRange(1, 5, 1, metricsList.length)
             .setBackground("#ffffff")
             .setNumberFormat('0.00')
             .setHorizontalAlignment("center")
             .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  // ---> Disable J1 Weight visual out of the box
  payoutSheet.getRange("J1").clearContent().setBackground("#e6e8eb");

  // 6. MASTER PROPORTIONAL FORMULAS (Unbound to Row 500)
  if (activeMembers.length > 0) {
    for (let r = 0; r < activeMembers.length; r++) {
      let rowNum = r + 3;
      
      let scoreFormula = `=IFERROR(SUMPRODUCT($E$1:$Q$1, E${rowNum}:Q${rowNum}), 0)`;
      let contribFormula = `=IFERROR(IF(SUM($T$3:$T$500)>0, T${rowNum} / SUM($T$3:$T$500), 0), 0)`;
      let payoutFormula = `=IFERROR(IF(C${rowNum}>0, C${rowNum} * '${dashboardName}'!$I$13, 0), 0)`;
      
      payoutSheet.getRange(rowNum, 20).setFormula(scoreFormula); 
      payoutSheet.getRange(rowNum, 3).setFormula(contribFormula).setNumberFormat('0.00%'); 
      payoutSheet.getRange(rowNum, 4).setFormula(payoutFormula).setNumberFormat('"$ "#,##0'); 
      
      let bgColor = (r % 2 === 0) ? "#ffffff" : "#f1f3f4"; 
      let moneyColor = (r % 2 === 0) ? "#e6f4ea" : "#ceead6"; 
      
      payoutSheet.getRange(rowNum, 1, 1, 17).setBackground(bgColor).setHorizontalAlignment("center");
      payoutSheet.getRange(rowNum, 1, 1, 2).setHorizontalAlignment("left"); 
      payoutSheet.getRange(rowNum, 4).setBackground(moneyColor).setFontWeight("bold"); 
    }
    
    payoutSheet.getRange(2, 1, activeMembers.length + 1, 17).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

    payoutSheet.getRange(3, 5, activeMembers.length, 13).setNumberFormat('#,##0'); 
    payoutSheet.getRange(3, 12, activeMembers.length, 1).setNumberFormat('#,##0.00'); // Net Respect
    payoutSheet.getRange(3, 14, activeMembers.length, 1).setNumberFormat('#,##0.00'); // War Score
  }

  // 7. CUSTOM COLUMN SIZING
  payoutSheet.setColumnWidth(2, 160); 
  payoutSheet.setColumnWidth(3, 115); 
  payoutSheet.setColumnWidth(4, 130); 
  
  for (let c = 5; c <= 17; c++) { payoutSheet.setColumnWidth(c, 115); }
  payoutSheet.setColumnWidth(8, 140); 
  payoutSheet.setColumnWidth(9, 150); 
  payoutSheet.setColumnWidth(13, 130); 

  payoutSheet.hideColumns(1); 
  payoutSheet.hideColumns(20); 

  ss.toast("Payout Tab Restored! Crunching metrics...", "System", 3);
  SpreadsheetApp.flush();
  
  // Triggers the Calculator to inject the live dashboard links into O, P, Q
  if (typeof runPayoutMath === "function") {
    runPayoutMath(); 
  }
}