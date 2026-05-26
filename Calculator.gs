// ==========================================
// PAYOUT MATH (Strict Official Report Engine)
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

  // --- 1. GET ENEMY ID & STRICT LIMITS ---
  const globalHitLimit = parseInt(dashSheet.getRange("F3").getValue()) || 999999;
  const personalWarLimit = parseInt(dashSheet.getRange("F4").getValue()) || 999999;
  const personalChainLimit = parseInt(dashSheet.getRange("F5").getValue()) || 999999;

  let dashData = dashSheet.getDataRange().getValues();
  let targetFactionId = "";
  let limitStr = "3:00"; 
  
  for (let r = 0; r < dashData.length; r++) {
    for (let c = 0; c < dashData[r].length; c++) {
      let cellText = safeStr(dashData[r][c]).toLowerCase();

      if (cellText === "enemy faction id" && c + 1 < dashData[r].length) targetFactionId = cleanId(dashData[r][c+1]);
      if (cellText.includes("chain drop limit") && c + 1 < dashData[r].length) {
        let foundLimit = safeStr(dashData[r][c+1]);
        if (foundLimit) limitStr = foundLimit;
      }
    }
  }

  if (!targetFactionId) {
    ss.toast("No Enemy Faction ID found. Running in Chain-Only mode (War stats ignored).", "Notice", 4);
  }

  // --- 2. PREPARE PAYOUT SHEET ---
  const lastPayoutRow = Math.max(3, payoutSheet.getLastRow());
  const existingIds = payoutSheet.getRange(3, 1, lastPayoutRow - 2, 1).getValues().flat().map(id => cleanId(id));
  payoutSheet.getRange(3, 5, lastPayoutRow - 2, 10).clearContent();
  let stats = {};

  // --- 3. EXCLUSIVELY LOAD OFFICIAL BASELINE TOTALS ---
  
  // A. Parse Official War Report for War Hits (wh) & Score (ws)
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
          if (!stats[memId]) stats[memId] = { name: safeStr(warData[i][3]), wh: 0, wa: 0, wl: 0, wi: 0, ch: 0, cs: 0, ret: 0, res: 0, abr: 0, ws: 0 };
          stats[memId].wh += parseInt(warData[i][4]) || 0;
          stats[memId].ws += parseFloat(warData[i][5]) || 0;
        }
      }
    }
  }

  // B. Parse Official Chain Report for Outside Hits (ch), Respect (res), and Assists (wa)
  let officialChainSheetUsed = false;
  if (officialChainSheet && officialChainSheet.getLastRow() > 6) {
    officialChainSheetUsed = true;
    const chainData = officialChainSheet.getDataRange().getValues();
    let startRow = -1;
    for(let i = 0; i < chainData.length; i++) {
      if(chainData[i][0] === "Member ID" && chainData[i][1] === "Total Attacks") { startRow = i + 1; break; }
    }
    if (startRow !== -1) {
      for (let i = startRow; i < chainData.length; i++) {
        let memId = cleanId(chainData[i][0]);
        if (memId === "" || memId === "API ERROR:") continue;
        if (!stats[memId]) stats[memId] = { name: `ID: ${memId}`, wh: 0, wa: 0, wl: 0, wi: 0, ch: 0, cs: 0, ret: 0, res: 0, abr: 0, ws: 0 };
        let s = stats[memId];
        
        s.res += parseFloat(chainData[i][2]) || 0; // Total Chain Respect
        s.wa += parseInt(chainData[i][6]) || 0;    // Assists
        s.abr += parseInt(chainData[i][7]) || 0;   // Abroad
        
        // OUTSIDE HITS CALC: (Total Successful Chain Hits) minus (Official War Hits)
        // This permanently eliminates ghost hits, as it strictly binds to the Official Chain Report limit.
        let chainSuccesses = (parseInt(chainData[i][3]) || 0) + (parseInt(chainData[i][4]) || 0) + (parseInt(chainData[i][5]) || 0);
        s.ch += Math.max(0, chainSuccesses - s.wh); 
      }
    }
  }

  // --- 4. ADVANCED RD PARSING (For Saves, Retals, Losses, and Respect Corrections) ---
  if (rdSheet) {
    let rdData = rdSheet.getDataRange().getValues();
    
    // RD Column Indices
    let tCol = 2; let aIdCol = 3; let aNameCol = 4; let aFacCol = 5; 
    let dIdCol = 6; let dFacCol = 8; let resltCol = 9; let resCol = 10; 
    let retCol = 13; let cBonusCol = 16; 

    let rdEvents = [];

    // Extract raw events
    for (let i = 1; i < rdData.length; i++) {
      let attackerId = cleanId(rdData[i][aIdCol]);
      let aFac = cleanId(rdData[i][aFacCol]);
      if (attackerId !== "" && aFac === myFactionId && !stats[attackerId]) {
        stats[attackerId] = { name: safeStr(rdData[i][aNameCol]) || `ID: ${attackerId}`, wh: 0, wa: 0, wl: 0, wi: 0, ch: 0, cs: 0, ret: 0, res: 0, abr: 0, ws: 0 };
      }
      rdEvents.push({
        time: new Date(rdData[i][tCol]).getTime(),
        aId: attackerId, aFac: aFac, 
        dId: cleanId(rdData[i][dIdCol]), dFac: cleanId(rdData[i][dFacCol]), 
        result: safeStr(rdData[i][resltCol]).toLowerCase(), 
        respect: parseFloat(rdData[i][resCol]) || 0,
        retMult: parseFloat(rdData[i][retCol]) || 1,
        cBonus: parseFloat(rdData[i][cBonusCol]) || 1
      });
    }

    // A. Identify Global Hit Limit Cutoff Timestamp
    let totalOfficialHits = 0;
    for (let id in stats) { totalOfficialHits += stats[id].wh + stats[id].ch; }
    let hitsToRemove = totalOfficialHits - globalHitLimit;
    let cutoffTimestamp = Infinity; 

    if (hitsToRemove > 0) {
      rdEvents.sort((a, b) => b.time - a.time); // Newest to Oldest
      for (let e of rdEvents) {
        if (hitsToRemove <= 0) { cutoffTimestamp = e.time; break; }
        if (e.aFac === myFactionId && e.respect > 0 && stats[e.aId]) {
          let isWarHit = (targetFactionId !== "" && e.dFac === targetFactionId);
          if (isWarHit && stats[e.aId].wh > 0) { stats[e.aId].wh--; hitsToRemove--; } 
          else if (!isWarHit && stats[e.aId].ch > 0) { stats[e.aId].ch--; hitsToRemove--; } 
        }
      }
    }

    // B. Chronological Processing
    rdEvents.sort((a, b) => a.time - b.time); // Oldest to Newest
    let pendingInterrupts = {};
    let validHitsForSaves = [];

    for (let e of rdEvents) {
      if (e.time > cutoffTimestamp) continue; // Ignore anything past the hit limit

      if (e.aFac === myFactionId && stats[e.aId]) {
        // Track Chain Saves timeline
        if (e.respect > 0 && !isNaN(e.time)) validHitsForSaves.push(e);

        // Deduct Milestone Chain Bonuses to reveal True Base Respect
        if (e.respect > 0 && e.cBonus > 1) {
          stats[e.aId].res -= (e.respect - (e.respect / e.cBonus));
          if (stats[e.aId].res < 0) stats[e.aId].res = 0;
        }

        // War Losses & Retals
        if (targetFactionId !== "" && e.dFac === targetFactionId) {
          if (e.result.includes("lost") || e.result.includes("escape") || e.result.includes("draw") || e.result.includes("timeout") || e.result.includes("stalemate")) {
            stats[e.aId].wl += 1;
          } else if (e.retMult > 1) {
            stats[e.aId].ret += 1;
          }
        }
      }

      // Enemy Interruptions
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

    // C. Chain Saves Math
    let timeRemainingSeconds = 180; 
    let cleanStr = limitStr.replace(/[^\d.:]/g, '');
    let parts = cleanStr.split(":");
    if (parts.length === 3) {
      let h = parseInt(parts[0], 10) || 0; let m = parseInt(parts[1], 10) || 0; let s = parseInt(parts[2], 10) || 0;
      if (h > 0 && m === 0 && s === 0) timeRemainingSeconds = h * 60;
      else timeRemainingSeconds = (h * 3600) + (m * 60) + s;
    } else if (parts.length === 2) {
      let m = parseInt(parts[0], 10) || 0; let s = parseInt(parts[1], 10) || 0; timeRemainingSeconds = (m * 60) + s;
    } else if (!isNaN(parseFloat(cleanStr))) {
      timeRemainingSeconds = parseFloat(cleanStr) * 60;
    }
    
    let requiredGap = 300 - timeRemainingSeconds;
    let lastHitTime = null;

    for (let hit of validHitsForSaves) {
      if (lastHitTime !== null) {
        let gap = (hit.time - lastHitTime) / 1000;
        if (gap >= requiredGap && gap <= 300 && stats[hit.aId]) stats[hit.aId].cs += 1;
      }
      lastHitTime = hit.time;
    }
  }

  // ---> 5. ENFORCE PERSONAL SUB-LIMITS WITH OVERFLOW <---
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

  // --- 6. OUTPUT & "LEFT FACTION" INJECTION ---
  const output = existingIds.map(id => {
    let s = stats[id] || { wh: 0, wa: 0, wl: 0, wi: 0, ch: 0, cs: 0, ret: 0, res: 0, abr: 0, ws: 0 };
    return [s.wh, s.wa, s.wl, s.wi, s.ch, s.cs, s.ret, s.res, s.abr, s.ws];
  });

  if (output.length > 0) payoutSheet.getRange(3, 5, output.length, 10).setValues(output);
  
  let missingPlayers = [];
  let missingStats = [];
  
  for (let id in stats) {
    if (!existingIds.includes(id) && id !== "" && id !== "API ERROR:") {
      missingPlayers.push([id, `${stats[id].name} (Left Faction)`]);
      missingStats.push([stats[id].wh, stats[id].wa, stats[id].wl, stats[id].wi, stats[id].ch, stats[id].cs, stats[id].ret, stats[id].res, stats[id].abr, stats[id].ws]);
    }
  }

  if (missingPlayers.length > 0) {
    let targetRow = payoutSheet.getLastRow() + 1;
    let maxCols = Math.max(15, payoutSheet.getLastColumn());

    payoutSheet.getRange(targetRow, 1, missingPlayers.length, 2).setValues(missingPlayers);
    payoutSheet.getRange(targetRow, 5, missingStats.length, 10).setValues(missingStats);
    payoutSheet.getRange(targetRow, 1, missingPlayers.length, 14).setBackground("#fce8e6").setFontStyle("italic");

    if (payoutSheet.getLastRow() >= 3) {
      let r1c1Formulas = payoutSheet.getRange(3, 1, 1, maxCols).getFormulasR1C1()[0];
      for (let c = 0; c < maxCols; c++) {
        let isRawDataCol = (c === 0 || c === 1 || (c >= 4 && c <= 13));
        if (!isRawDataCol && r1c1Formulas[c] !== "") {
          let newFormulaBlock = [];
          for (let r = 0; r < missingPlayers.length; r++) newFormulaBlock.push([r1c1Formulas[c]]);
          payoutSheet.getRange(targetRow, c + 1, missingPlayers.length, 1).setFormulasR1C1(newFormulaBlock);
        }
      }
    }
  }

  let finalMsg = targetFactionId 
    ? "Payout calculated! Outside Hits mapped directly to Official Chain Report." 
    : "Chain-Only Payout calculated! (Base Respect Adjusted)";
  ss.toast(finalMsg, "Success", 5);
}