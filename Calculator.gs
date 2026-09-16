// ==========================================
// PAYOUT MATH (Live Weights + Static Headers + Warlord/Pushes)
// ==========================================
function runPayoutMath() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  const rosterSheet = ss.getSheetByName(SETTINGS.rosterSheet);
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  const rdSheet = ss.getSheetByName(SETTINGS.rdSheet);

  const officialWarSheet = ss.getSheetByName("Official War Report");
  const officialChainSheet = ss.getSheetByName("Official Chain Report");

  if (!payoutSheet || !rosterSheet || !configSheet || !dashSheet) return;

  const myFactionId = configSheet.getRange(SETTINGS.factionIdCell || "B7").getValue().toString().trim();

  const cleanId = (val) => (val === null || val === undefined) ? "" : val.toString().replace(/,/g, "").trim();
  const safeStr = (val) => (val === null || val === undefined) ? "" : val.toString().trim();

  // --- 1. GET ENEMY ID, STRICT LIMITS & DASHBOARD TOGGLES ---
  const globalHitLimit = parseInt(dashSheet.getRange("F3").getValue()) || 999999;
  const personalWarLimit = parseInt(dashSheet.getRange("F4").getValue()) || 999999;
  const personalChainLimit = parseInt(dashSheet.getRange("F5").getValue()) || 999999;

  let dashData = dashSheet.getDataRange().getValues();
  let targetFactionId = "";
  let payPostWarStr = "Yes";
  let officialWarEndStr = "";
  let pushLimitStr = ""; // NEW: Push Limit Tracker
  
  for (let r = 0; r < dashData.length; r++) {
    for (let c = 0; c < dashData[r].length; c++) {
      let cellText = safeStr(dashData[r][c]).toLowerCase();
      if (cellText === "enemy faction id" && c + 1 < dashData[r].length) targetFactionId = cleanId(dashData[r][c+1]);
      if (cellText === "pay post-war chain?" && c + 1 < dashData[r].length) payPostWarStr = safeStr(dashData[r][c+1]);
      if (cellText === "official war end" && c + 1 < dashData[r].length) officialWarEndStr = safeStr(dashData[r][c+1]);
      if (cellText === "push time limit" && c + 1 < dashData[r].length) pushLimitStr = safeStr(dashData[r][c+1]); // NEW
    }
  }

  if (!targetFactionId) {
    ss.toast("No Enemy Faction ID found. Running in Chain-Only mode.", "Notice", 4);
  }

  let payPostWar = (payPostWarStr.toLowerCase() === "yes");
  let warEndMs = null;
  if (officialWarEndStr && officialWarEndStr.toLowerCase() !== "ongoing" && officialWarEndStr.toLowerCase() !== "n/a") {
      let d = new Date(officialWarEndStr + " GMT");
      if (!isNaN(d.getTime())) warEndMs = d.getTime();
  }

  // --- 2. PREPARE PAYOUT SHEET ---
  const lastPayoutRow = Math.max(3, payoutSheet.getLastRow());
  const existingIds = payoutSheet.getRange(3, 1, lastPayoutRow - 2, 1).getValues().flat().map(id => cleanId(id));
  
  // Clear E through Q (13 columns) AND R through S (2 columns)
  payoutSheet.getRange(3, 5, lastPayoutRow - 2, 13).clearContent();
  payoutSheet.getRange(3, 18, lastPayoutRow - 2, 2).clearContent();
  let stats = {};

  const initStats = (name) => ({ 
    name: name, wh: 0, wa: 0, wl: 0, wi: 0, ch: 0, cs: 0, ret: 0, res: 0, abr: 0, ws: 0, t0: 0, t1: 0, t2: 0, wlh: 0, ph: 0 
  });

  // --- 3. EXCLUSIVELY LOAD OFFICIAL BASELINE TOTALS ---
  if (officialWarSheet) {
    const warData = officialWarSheet.getDataRange().getValues();
    let startRow = -1;
    for (let i = 0; i < warData.length; i++) {
      if (warData[i][0] === "Faction ID" && warData[i][2] === "Member ID") { startRow = i + 1; break; }
    }
    if (startRow !== -1) {
      for (let i = startRow; i < warData.length; i++) {
        let facId = cleanId(warData[i][0]);
        let memId = cleanId(warData[i][2]);
        if (facId === myFactionId && memId !== "") {
          if (!stats[memId]) stats[memId] = initStats(safeStr(warData[i][3]));
          stats[memId].wh += parseInt(warData[i][4]) || 0;
          stats[memId].ws += parseFloat(warData[i][5]) || 0;
        }
      }
    }
  }

  if (officialChainSheet && officialChainSheet.getLastRow() > 6) {
    const chainData = officialChainSheet.getDataRange().getValues();
    let startRow = -1;
    for(let i = 0; i < chainData.length; i++) {
      if(chainData[i][0] === "Member ID" && chainData[i][1] === "Total Attacks") { startRow = i + 1; break; }
    }
    if (startRow !== -1) {
      for (let i = startRow; i < chainData.length; i++) {
        let memId = cleanId(chainData[i][0]);
        if (memId === "" || memId === "API ERROR:") continue;
        if (!stats[memId]) stats[memId] = initStats(`ID: ${memId}`);
        let s = stats[memId];
        
        s.res += parseFloat(chainData[i][2]) || 0; 
        s.wa += parseInt(chainData[i][6]) || 0;    
        s.abr += parseInt(chainData[i][7]) || 0;   
        
        let chainSuccesses = (parseInt(chainData[i][3]) || 0) + (parseInt(chainData[i][4]) || 0) + (parseInt(chainData[i][5]) || 0);
        s.ch += Math.max(0, chainSuccesses - s.wh); 
      }
    }
  }

  // --- 4. ADVANCED RD PARSING ---
  if (rdSheet) {
    let rdData = rdSheet.getDataRange().getValues();
    let tCol = 2; let aIdCol = 3; let aNameCol = 4; let aFacCol = 5; 
    let dIdCol = 6; let dFacCol = 8; let resltCol = 9; let resCol = 10; 
    let retCol = 13; let cBonusCol = 16; let warlordCol = 19; // Col T 

    let rdEvents = [];

    for (let i = 1; i < rdData.length; i++) {
      let attackerId = cleanId(rdData[i][aIdCol]);
      let aFac = cleanId(rdData[i][aFacCol]);
      if (attackerId !== "" && aFac === myFactionId && !stats[attackerId]) {
        stats[attackerId] = initStats(safeStr(rdData[i][aNameCol]) || `ID: ${attackerId}`);
      }
      rdEvents.push({
        time: new Date(rdData[i][tCol]).getTime(),
        aId: attackerId, aFac: aFac, 
        dId: cleanId(rdData[i][dIdCol]), dFac: cleanId(rdData[i][dFacCol]), 
        result: safeStr(rdData[i][resltCol]).toLowerCase(), 
        respect: parseFloat(rdData[i][resCol]) || 0,
        retMult: parseFloat(rdData[i][retCol]) || 1,
        cBonus: parseFloat(rdData[i][cBonusCol]) || 1,
        warlord: parseFloat(rdData[i][warlordCol]) || 1 // Capture Warlord
      });
    }

    let factionHits = rdEvents.filter(e => e.aFac === myFactionId && e.respect > 0);
    factionHits.sort((a, b) => a.time - b.time);
    let activeChains = [];
    let currentChain = [];
    
    for (let hit of factionHits) {
      if (currentChain.length === 0) { currentChain.push(hit); }
      else {
        if (hit.time - currentChain[currentChain.length - 1].time <= 300000) { currentChain.push(hit); } 
        else {
          if (currentChain.length >= 10) activeChains.push({start: currentChain[0].time, end: currentChain[currentChain.length-1].time});
          currentChain = [hit];
        }
      }
    }
    if (currentChain.length >= 10) activeChains.push({start: currentChain[0].time, end: currentChain[currentChain.length-1].time});

    let totalOfficialHits = 0;
    for (let id in stats) { totalOfficialHits += stats[id].wh + stats[id].ch; }
    let hitsToRemove = totalOfficialHits - globalHitLimit;
    let cutoffTimestamp = Infinity; 

    if (hitsToRemove > 0) {
      rdEvents.sort((a, b) => b.time - a.time); 
      for (let e of rdEvents) {
        if (hitsToRemove <= 0) { cutoffTimestamp = e.time; break; }
        if (e.aFac === myFactionId && e.respect > 0 && stats[e.aId]) {
          let isWarHit = (targetFactionId !== "" && e.dFac === targetFactionId);
          if (isWarHit && stats[e.aId].wh > 0) { 
            stats[e.aId].wh--; hitsToRemove--; 
            stats[e.aId].res = Math.max(0, stats[e.aId].res - e.respect); 
          } 
          else if (!isWarHit && stats[e.aId].ch > 0) { 
            stats[e.aId].ch--; hitsToRemove--; 
            stats[e.aId].res = Math.max(0, stats[e.aId].res - e.respect); 
          } 
        }
      }
    }

    rdEvents.sort((a, b) => a.time - b.time); 
    let pendingInterrupts = {};
    let validHitsForSaves = [];

    for (let e of rdEvents) {
      if (e.time > cutoffTimestamp) continue;

      if (e.aFac === myFactionId && stats[e.aId]) {
        let isTrimmedForPostWar = false;

        if (e.respect > 0 && e.cBonus > 1) {
          stats[e.aId].res -= (e.respect - (e.respect / e.cBonus));
          if (stats[e.aId].res < 0) stats[e.aId].res = 0;
        }

        if (!payPostWar && warEndMs && e.time > warEndMs) {
          let inActiveChain = activeChains.some(c => e.time >= c.start && e.time <= c.end);
          if (inActiveChain && e.respect > 0) {
            stats[e.aId].ch = Math.max(0, stats[e.aId].ch - 1);
            let remainingBaseRes = (e.cBonus > 1) ? (e.respect / e.cBonus) : e.respect;
            stats[e.aId].res = Math.max(0, stats[e.aId].res - remainingBaseRes); 
            isTrimmedForPostWar = true; 
          }
        }

        if (e.respect > 0 && !isNaN(e.time) && !isTrimmedForPostWar) {
          validHitsForSaves.push(e);
        }

        if (targetFactionId !== "" && e.dFac === targetFactionId) {
          if (e.result.includes("lost") || e.result.includes("escape") || e.result.includes("draw") || e.result.includes("timeout") || e.result.includes("stalemate")) {
            stats[e.aId].wl += 1;
          } else if (e.retMult > 1) {
            stats[e.aId].ret += 1;
          }
        }
      }

      if (targetFactionId !== "") {
        if (e.aFac === targetFactionId && e.result.includes("interrupted")) {
          pendingInterrupts[e.aId] = true; 
        }
        if (e.aFac === myFactionId && e.dFac === targetFactionId && pendingInterrupts[e.dId]) {
          if (e.result.includes("hospitalized") || e.result.includes("attacked") || e.result.includes("mugged")) {
            if (stats[e.aId]) stats[e.aId].wi += 1; 
            pendingInterrupts[e.dId] = false; 
          }
        }
      }
    }

    // =========================================
    // D. MULTI-TIER CHAIN SAVES & PUSH WINDOW LOGIC
    // =========================================
    let allTiers = [];
    let saveTiers = [];
    let tierRowStart = -1;
    let tierCol = -1;
    
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

    let dashName = dashSheet.getName();

    if (tierRowStart !== -1) {
      let row1 = tierRowStart;
      let row2 = tierRowStart + 1;
      let row3 = tierRowStart + 2;

      payoutSheet.getRange("O1:Q1").setFormulas([[ 
        `=IFERROR(VALUE('${dashName}'!F${row1}), 0)`, 
        `=IFERROR(VALUE('${dashName}'!F${row2}), 0)`, 
        `=IFERROR(VALUE('${dashName}'!F${row3}), 0)`
      ]]);
      
      let label1 = dashSheet.getRange(row1, tierCol).getDisplayValue();
      let label2 = dashSheet.getRange(row2, tierCol).getDisplayValue();
      let label3 = dashSheet.getRange(row3, tierCol).getDisplayValue();

      payoutSheet.getRange("O2:Q2").setValues([[ 
        label1 ? label1 + " Saves" : "Tier 1",
        label2 ? label2 + " Saves" : "Tier 2",
        label3 ? label3 + " Saves" : "Tier 3"
      ]]);

      for (let i = 0; i < 3; i++) {
        let timeVal = dashSheet.getRange(tierRowStart + i, tierCol).getDisplayValue();
        if (timeVal && timeVal.toString().trim() !== "") {
          let cleanStr = timeVal.toString().replace(/[^\d.:]/g, '');
          let secs = 0;
          let parts = cleanStr.split(":");
          if (parts.length === 3) secs = (parseInt(parts[0], 10)||0)*3600 + (parseInt(parts[1], 10)||0)*60 + (parseInt(parts[2], 10)||0);
          else if (parts.length === 2) secs = (parseInt(parts[0], 10)||0)*60 + (parseInt(parts[1], 10)||0);
          else if (!isNaN(parseFloat(cleanStr))) secs = parseFloat(cleanStr) * 60;
          
          if (secs > 0) saveTiers.push({ requiredGap: 300 - secs, originalIndex: i });
        }
      }
    }

    // ---> NEW: Parse Push Windows from K4:M13 <---
    let pushWindows = [];
    for (let r = 3; r <= 12; r++) { // dashData rows 3-12 correlate to sheet rows 4-13
      let pTimeStr = dashData[r][10] ? dashData[r][10].toString().trim() : ""; // Col K
      let pDateVal = dashData[r][11]; // Col L
      let pLimitStr = dashData[r][12] ? dashData[r][12].toString().trim() : ""; // Col M

      if (pTimeStr && pDateVal && pLimitStr) {
        let pd = new Date(pDateVal);
        let ptParts = pTimeStr.split(":");
        let pHour = parseInt(ptParts[0]) || 0;
        let pMin = parseInt(ptParts[1]) || 0;
        let pSec = parseInt(ptParts[2]) || 0;

        // Convert the Date & Time into a Torn-accurate UTC Timestamp
        let startUnix = Date.UTC(pd.getUTCFullYear(), pd.getUTCMonth(), pd.getUTCDate(), pHour, pMin, pSec);
        let limitMins = parseFloat(pLimitStr) || 0;
        let endUnix = startUnix + (limitMins * 60000); // 60,000 ms in a minute

        pushWindows.push({ start: startUnix, end: endUnix });
      }
    }

    if (payoutSheet.getLastRow() >= 2) {
      payoutSheet.getRange("N1:N2").copyTo(payoutSheet.getRange("O1:Q2"), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      payoutSheet.getRange("O1:Q2").setHorizontalAlignment("center"); 
    }
    
    saveTiers.sort((a, b) => b.requiredGap - a.requiredGap);
    
    let lastHitTime = null;

    for (let hit of validHitsForSaves) {
      // 1. Warlord Tracker
      if (hit.warlord > 1 && stats[hit.aId]) {
         stats[hit.aId].wlh += 1;
      }

      // 2. Push Window Tracker (Independent of gaps, checks exact timestamps)
      if (stats[hit.aId] && targetFactionId !== "" && hit.dFac === targetFactionId) {
        for (let pw of pushWindows) {
           if (hit.time >= pw.start && hit.time <= pw.end) {
              stats[hit.aId].ph += 1; // It's a Push Hit!
              break; 
           }
        }
      }

      // 3. Chain Saves Tracker (Chronological Gap Check)
      if (lastHitTime !== null) {
        let gap = (hit.time - lastHitTime) / 1000;
        if (gap <= 300 && stats[hit.aId]) {
          for (let tier of saveTiers) {
            if (gap >= tier.requiredGap) {
              stats[hit.aId].cs += 1; 
              stats[hit.aId][`t${tier.originalIndex}`] += 1; 
              break; 
            }
          }
        }
      }
      lastHitTime = hit.time;
    }
  }

  for (let id in stats) {
    if (stats[id].wh > personalWarLimit) {
      let overflowHits = stats[id].wh - personalWarLimit;
      stats[id].ch += overflowHits;
      stats[id].wh = personalWarLimit;
    }
    if (stats[id].ch > personalChainLimit) {
      stats[id].ch = personalChainLimit;
    }
  }

  // --- 6. OUTPUT (Split Base + Push/Warlord to save Formulas in O/P/Q) ---
  const outputBase = [];
  const outputNew = [];

  for (let id of existingIds) {
    let s = stats[id] || initStats(id);
    outputBase.push([s.wh, s.wa, s.wl, s.wi, s.ch, s.cs, s.ret, s.res, s.abr, s.ws, s.t0, s.t1, s.t2]);
    outputNew.push([s.wlh, s.ph]); 
  }

  if (outputBase.length > 0) {
    payoutSheet.getRange(3, 5, outputBase.length, 13).setValues(outputBase);
    payoutSheet.getRange(3, 18, outputNew.length, 2).setValues(outputNew); // Write directly to R and S
  }
  
  let missingPlayers = [];
  let missingStatsBase = [];
  let missingStatsNew = [];
  
  for (let id in stats) {
    if (!existingIds.includes(id) && id !== "" && id !== "API ERROR:") {
      missingPlayers.push([id, `${stats[id].name} (Left Faction)`]);
      missingStatsBase.push([stats[id].wh, stats[id].wa, stats[id].wl, stats[id].wi, stats[id].ch, stats[id].cs, stats[id].ret, stats[id].res, stats[id].abr, stats[id].ws, stats[id].t0, stats[id].t1, stats[id].t2]);
      missingStatsNew.push([stats[id].wlh, stats[id].ph]);
    }
  }

  let maxCols = Math.max(19, payoutSheet.getLastColumn()); 

  if (missingPlayers.length > 0) {
    let targetRow = payoutSheet.getLastRow() + 1;

    payoutSheet.getRange(targetRow, 1, missingPlayers.length, 2).setValues(missingPlayers);
    payoutSheet.getRange(targetRow, 5, missingStatsBase.length, 13).setValues(missingStatsBase);
    payoutSheet.getRange(targetRow, 18, missingStatsNew.length, 2).setValues(missingStatsNew);
    payoutSheet.getRange(targetRow, 1, missingPlayers.length, 19).setBackground("#fce8e6").setFontStyle("italic");

    if (payoutSheet.getLastRow() >= 3) {
      let r1c1Formulas = payoutSheet.getRange(3, 1, 1, maxCols).getFormulasR1C1()[0];
      for (let c = 0; c < maxCols; c++) {
        // Exclude raw input cols (E-Q = 4-16, R-S = 17-18)
        let isRawDataCol = (c === 0 || c === 1 || (c >= 4 && c <= 18));
        if (!isRawDataCol && r1c1Formulas[c] !== "") {
          let newFormulaBlock = [];
          for (let r = 0; r < missingPlayers.length; r++) newFormulaBlock.push([r1c1Formulas[c]]);
          payoutSheet.getRange(targetRow, c + 1, missingPlayers.length, 1).setFormulasR1C1(newFormulaBlock);
        }
      }
    }
  }

  let finalLastRow = payoutSheet.getLastRow();

  if (finalLastRow >= 3) {
    payoutSheet.getRange(3, 14, finalLastRow - 2, 1).copyTo(
      payoutSheet.getRange(3, 15, finalLastRow - 2, 3), 
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT, 
      false
    );
    payoutSheet.getRange(3, 15, finalLastRow - 2, 3).setHorizontalAlignment("center").setNumberFormat("0");

    let formulaRange = payoutSheet.getRange(3, 1, finalLastRow - 2, maxCols);
    let formulas = formulaRange.getFormulas();
    let hasChanges = false;
    
    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        let f = formulas[r][c];
        if (f) {
          let orig = f;
          let rowNum = r + 3;
          let newMath = `(($O$1*O${rowNum}) + ($P$1*P${rowNum}) + ($Q$1*Q${rowNum}))`;
          f = f.replace(/\*\s*\$?J\$?1(?!\d)/gi, '');
          f = f.replace(/\$?J\$?1(?!\d)\s*\*/gi, '');
          let jRegex = new RegExp(`([^A-Z])\\$?J\\$?${rowNum}(?!\\d)`, "gi");
          f = f.replace(jRegex, `$1${newMath}`);
          let jRegexStart = new RegExp(`^\\=\\$?J\\$?${rowNum}(?!\\d)`, "gi");
          f = f.replace(jRegexStart, `=${newMath}`);
          if (f !== orig) { formulas[r][c] = f; hasChanges = true; }
        }
      }
    }
    
    if (hasChanges) formulaRange.setFormulas(formulas);
  }

  payoutSheet.getRange("J1").clearContent().setBackground("#e6e8eb");

  let finalMsg = targetFactionId 
    ? (payPostWar ? "Payout calculated! Push hits & Warlords updated." : "Payout calculated! Post-War hits removed & pushes mapped.") 
    : "Chain-Only Payout calculated! (Warlords and pushes mapped)";
  ss.toast(finalMsg, "Success", 5);
}