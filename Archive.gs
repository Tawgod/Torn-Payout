// ==========================================
// 1. TEST ARCHIVE 
// ==========================================
function testArchiveOnly() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "🧪 TEST ARCHIVE", 
    "This will push the report to the external Archive using Payouts!C1 for Total Attacks. Proceed?", 
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Connecting to External Archive...", "System", 3);

  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  let rawId = configSheet ? configSheet.getRange("B3").getValue().toString().trim() : "";
  
  let archiveId = rawId;
  if (rawId.includes("/d/")) {
    let match = rawId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match) archiveId = match[1];
  }

  let externalSS;
  try {
    externalSS = SpreadsheetApp.openById(archiveId);
  } catch (e) {
    ui.alert("⚠️ Access Error! Check permissions.");
    return;
  }

  let dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  let payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  
  if (!payoutSheet) {
    ui.alert("Error: Payouts sheet not found.");
    return;
  }

  let allData = dashSheet.getDataRange().getValues();
  const findVal = (label) => {
    let cleanLabel = label.toLowerCase().trim();
    for (let r = 0; r < allData.length; r++) {
      for (let c = 0; c < allData[r].length; c++) {
        let currentCell = allData[r][c].toString().toLowerCase().trim();
        if (currentCell === cleanLabel && c + 1 < allData[r].length) return allData[r][c+1];
      }
    }
    return null;
  };

  let warId = findVal("War ID") || "Test_" + new Date().getTime();
  
  // Header Logic
  let headerInfo = {
    enemy: findVal("Enemy Faction Name") || "Unknown",
    score: Math.round(parseFloat(findVal("War Score")) || 0), 
    outcome: findVal("Outcome (Result)") || "Unknown",
    termed: findVal("Termed?") || "No",
    totalHits: payoutSheet.getRange("C1").getValue(), // FIXED: Pulls from Payouts!C1
    warHits: findVal("Total War Hits") || 0,
    revenue: findVal("Total Revenue") || 0,
    payout: findVal("PAYOUT TOTAL") || 0,
    profit: findVal("Actual Faction Deduction") || 0,
    caches: findVal("Caches / Items Won") || "None"
  };

  let targetSheet = externalSS.getSheetByName(warId.toString()) || externalSS.insertSheet(warId.toString());
  targetSheet.clear(); 

  let numPayoutCols = payoutSheet.getLastColumn();
  let totalArchiveCols = numPayoutCols + 2;

  let lastP = payoutSheet.getLastRow();
  if (lastP >= 3) {
    let pData = payoutSheet.getRange(3, 1, lastP - 2, numPayoutCols).getValues();
    let pHeaders = payoutSheet.getRange(2, 1, 1, numPayoutCols).getValues()[0];
    let activeData = pData.filter(row => row[0] !== ""); 
    
    if (activeData.length > 0) {
      let timeStamp = new Date();
      const pad = (arr) => { while (arr.length < totalArchiveCols) arr.push(""); return arr; };

      let reportHeader = [
        pad([`WAR REPORT: ${headerInfo.enemy} [${warId}]`]),
        pad(["Date Archived:", timeStamp, "Outcome:", headerInfo.outcome, "Termed:", headerInfo.termed]),
        pad(["Total Attacks Logged:", headerInfo.totalHits, "Total War Hits:", headerInfo.warHits, "War Score:", headerInfo.score]),
        pad(["Total Revenue:", headerInfo.revenue, "Total Payouts:", headerInfo.payout, "Faction Profit:", headerInfo.profit]),
        pad(["Caches Won:", headerInfo.caches]),
        pad([""]),
        pad(["Archive Date", "War ID"].concat(pHeaders))
      ];

      targetSheet.getRange(1, 1, reportHeader.length, totalArchiveCols).setValues(reportHeader);
      targetSheet.getRange(1, 1, 1, totalArchiveCols).setBackground("#444444").setFontColor("white").setFontWeight("bold").merge();
      targetSheet.getRange(7, 1, 1, totalArchiveCols).setBackground("#eeeeee").setFontWeight("bold");
      
      // Formatting
      targetSheet.getRange(2, 2).setNumberFormat("yyyy-mm-dd HH:mm");
      targetSheet.getRange(3, 2).setNumberFormat("#,##0");           // Total Hits
      targetSheet.getRange(3, 4).setNumberFormat("#,##0");           // War Hits
      targetSheet.getRange(3, 6).setNumberFormat("#,##0");           // War Score
      
      targetSheet.getRange(4, 2).setNumberFormat('"$ "#,##0'); 
      targetSheet.getRange(4, 4).setNumberFormat('"$ "#,##0'); 
      targetSheet.getRange(4, 6).setNumberFormat('"$ "#,##0'); 

      let archiveData = activeData.map(row => [timeStamp, warId].concat(row));
      targetSheet.getRange(8, 1, archiveData.length, totalArchiveCols).setValues(archiveData);
      
      targetSheet.setColumnWidths(1, totalArchiveCols, 110);
      targetSheet.setFrozenRows(7);
      targetSheet.getDataRange().setWrap(true);

      ui.alert(`🧪 Test Complete! Archive sheet ${warId} created.`);
    }
  }
}

// ==========================================
// 2. ARCHIVE & SMART RESET (PRODUCTION)
// ==========================================
function archiveAndResetWar() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "⚠️ WARNING: FACTORY RESET", 
    "This will archive your report using Payouts!C1 for Hits, uncheck boxes, and WIPE local sheets. Proceed?", 
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  let rawId = configSheet ? configSheet.getRange("B3").getValue().toString().trim() : "";
  
  let archiveId = rawId;
  if (rawId.includes("/d/")) {
    let match = rawId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match) archiveId = match[1];
  }

  let externalSS;
  try {
    externalSS = SpreadsheetApp.openById(archiveId);
  } catch (e) {
    ui.alert("⚠️ Access Error! Check permissions.");
    return;
  }

  let dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  let payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  ss.toast("Archiving and Resetting...", "System", 3);

  let allData = dashSheet.getDataRange().getValues();
  const findVal = (label) => {
    let cleanLabel = label.toLowerCase().trim();
    for (let r = 0; r < allData.length; r++) {
      for (let c = 0; c < allData[r].length; c++) {
        let currentCell = allData[r][c].toString().toLowerCase().trim();
        if (currentCell === cleanLabel && c + 1 < allData[r].length) return allData[r][c+1];
      }
    }
    return null;
  };

  let warId = findVal("War ID") || "War_" + new Date().getTime();
  let targetSheet = externalSS.getSheetByName(warId.toString()) || externalSS.insertSheet(warId.toString());

  let headerInfo = {
    enemy: findVal("Enemy Faction Name") || "Unknown",
    score: Math.round(parseFloat(findVal("War Score")) || 0),
    outcome: findVal("Outcome (Result)") || "Unknown",
    termed: findVal("Termed?") || "No",
    totalHits: payoutSheet.getRange("C1").getValue(), 
    warHits: findVal("Total War Hits") || 0,
    revenue: findVal("Total Revenue") || 0,
    payout: findVal("PAYOUT TOTAL") || 0,
    profit: findVal("Actual Faction Deduction") || 0,
    caches: findVal("Caches / Items Won") || "None"
  };

  let numPayoutCols = payoutSheet.getLastColumn();
  let totalArchiveCols = numPayoutCols + 2;

  let lastP = payoutSheet.getLastRow();
  if (lastP >= 3) {
    let pData = payoutSheet.getRange(3, 1, lastP - 2, numPayoutCols).getValues();
    let pHeaders = payoutSheet.getRange(2, 1, 1, numPayoutCols).getValues()[0];
    let activeData = pData.filter(row => row[0] !== ""); 
    
    if (activeData.length > 0) {
      let timeStamp = new Date();
      const pad = (arr) => { while (arr.length < totalArchiveCols) arr.push(""); return arr; };
      let reportHeader = [
        pad([`WAR REPORT: ${headerInfo.enemy} [${warId}]`]),
        pad(["Date Archived:", timeStamp, "Outcome:", headerInfo.outcome, "Termed:", headerInfo.termed]),
        pad(["Total Attacks Logged:", headerInfo.totalHits, "Total War Hits:", headerInfo.warHits, "War Score:", headerInfo.score]),
        pad(["Total Revenue:", headerInfo.revenue, "Total Payouts:", headerInfo.payout, "Faction Profit:", headerInfo.profit]),
        pad(["Caches Won:", headerInfo.caches]),
        pad([""]),
        pad(["Archive Date", "War ID"].concat(pHeaders))
      ];
      targetSheet.getRange(1, 1, 7, totalArchiveCols).setValues(reportHeader);
      targetSheet.getRange(1, 1, 1, totalArchiveCols).setBackground("#444444").setFontColor("white").setFontWeight("bold").merge();
      targetSheet.getRange(7, 1, 1, totalArchiveCols).setBackground("#eeeeee").setFontWeight("bold");
      
      targetSheet.getRange(2, 2).setNumberFormat("yyyy-mm-dd HH:mm");
      targetSheet.getRange(3, 2).setNumberFormat("#,##0");
      targetSheet.getRange(3, 4).setNumberFormat("#,##0");
      targetSheet.getRange(3, 6).setNumberFormat("#,##0");
      targetSheet.getRange(4, 2).setNumberFormat('"$ "#,##0');
      targetSheet.getRange(4, 4).setNumberFormat('"$ "#,##0');
      targetSheet.getRange(4, 6).setNumberFormat('"$ "#,##0');
      
      let archiveData = activeData.map(row => [timeStamp, warId].concat(row));
      targetSheet.getRange(8, 1, archiveData.length, totalArchiveCols).setValues(archiveData);
      targetSheet.setFrozenRows(7);
      targetSheet.getDataRange().setWrap(true);
    }
  }

  // --- LOCAL WIPE LOGIC ---
  let rdSheet = ss.getSheetByName(SETTINGS.rdSheet);
  if (rdSheet && rdSheet.getLastRow() > 1) {
    rdSheet.getRange(2, 1, rdSheet.getLastRow() - 1, rdSheet.getLastColumn()).clearContent();
  }
  let bountySheet = ss.getSheetByName(SETTINGS.bountySheet);
  if (bountySheet && bountySheet.getLastRow() > 1) {
    bountySheet.getRange(2, 1, bountySheet.getLastRow() - 1, bountySheet.getLastColumn()).clearContent();
  }
  let oWarSheet = ss.getSheetByName("Official War Report");
  if (oWarSheet) oWarSheet.clear();
  let oChainSheet = ss.getSheetByName("Official Chain Report");
  if (oChainSheet) oChainSheet.clear();

  if (payoutSheet && payoutSheet.getLastRow() >= 3) {
    
    // ---> NEW: UNCHECK ALL CHECKBOXES SAFELY <---
    payoutSheet.getRange(3, 1, payoutSheet.getLastRow() - 2, payoutSheet.getLastColumn()).uncheck();
    
    // Clear numerical calculation area
    payoutSheet.getRange(3, 5, payoutSheet.getLastRow() - 2, payoutSheet.getLastColumn() - 4).clearContent();
    
    let pData = payoutSheet.getDataRange().getValues();
    let ghostRowIdx = -1;
    for(let i = 0; i < pData.length; i++) {
      if (pData[i][1] === "⚠️ NON-ROSTER / LEFT FACTION" || pData[i][1].toString().includes("(Left Faction)")) { ghostRowIdx = i + 1; break; }
    }
    if (ghostRowIdx > 0) { payoutSheet.deleteRows(ghostRowIdx, payoutSheet.getLastRow() - ghostRowIdx + 1); }
  }

  if (dashSheet) {
    const setVal = (label, value) => {
      let dashData = dashSheet.getDataRange().getValues();
      for (let r = 0; r < dashData.length; r++) {
        for (let c = 0; c < dashData[r].length; c++) {
          if (dashData[r][c] === label) { dashSheet.getRange(r + 1, c + 2).setValue(value); return; }
        }
      }
    };
    setVal("Enemy Faction ID", ""); setVal("Enemy Faction Name", "Unknown");
    setVal("War Report ID", ""); setVal("Chain Report ID", "");
    setVal("Outcome (Result)", "Ongoing"); setVal("Termed?", "No");
    setVal("Total Revenue", 0); setVal("- Temp Cost", 0);
    setVal("- Revives", 0); setVal("- Xanax", 0); setVal("- Other Cost", 0);
    setVal("Est. Cache Value", ""); setVal("Caches / Items Won", "");
    if (typeof buildDashboard === "function") { buildDashboard(); }
  }

  ui.alert(`🧹 Reset Complete! Created archive sheet: ${warId} and cleared all checkboxes.`);
}