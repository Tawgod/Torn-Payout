// ==========================================
// PAYOUT MATH (Reverse Target Trimmer)
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
    ui.alert("⚠️ Enemy Faction ID is missing on your Dashboard. I cannot count War Losses without knowing who the enemy is!");
    return;
  }

  // --- 2. PREPARE PAYOUT SHEET ---
  const lastPayoutRow = Math.max(3, payoutSheet.getLastRow());
  const existingIds = payoutSheet.getRange(3, 1, lastPayoutRow - 2, 1).getValues().flat().map(id => cleanId(id));
  payoutSheet.getRange(3, 5, lastPayoutRow - 2, 10).clearContent();

  let stats = {};

  // --- 3. LOAD OFFICIAL BASELINE TOTALS ---
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

  if (officialChainSheet) {
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
        s.res += parseFloat(chainData[i][2]) || 0;
        s.wa += parseInt(chainData[i][6]) || 0; 
        s.abr += parseInt(chainData[i][7]) || 0;
        
        let success = (parseInt(chainData[i][3]) || 0) + (parseInt(chainData[i][4]) || 0) + (parseInt(chainData[i][5]) || 0);
        s.ch += Math.max(0, success - s.wh); 
      }
    }
  }

  // --- 4. THE EXACT OVERAGE TRIMMER (RD TAB) ---
  if (rdSheet) {
    let rdData = rdSheet.getDataRange().getValues();
    let tCol = 2; let aIdCol = 3; let aNameCol = 4; let aFacCol = 5; let dFacCol = 8; let resltCol = 9; let resCol = 10;   

    let rdEvents = [];

    // Base Roster & Event Gathering
    for (let i = 1; i < rdData.length; i++) {
      let attackerId = cleanId(rdData[i][aIdCol]);
      let defenderFacId = cleanId(rdData[i][dFacCol]);
      let aFac = cleanId(rdData[i][aFacCol]);
      let aName = safeStr(rdData[i][aNameCol]);
      let result = safeStr(rdData[i][resltCol]).toLowerCase();
      let respect = parseFloat(rdData[i][resCol]) || 0;
      let timestamp = new Date(rdData[i][tCol]).getTime();

      if (attackerId !== "" && aFac === myFactionId) {
        if (!stats[attackerId]) stats[attackerId] = { name: aName || `ID: ${attackerId}`, wh: 0, wa: 0, wl: 0, wi: 0, ch: 0, cs: 0, ret: 0, res: 0, abr: 0, ws: 0 };
      }

      if (attackerId === "" || !stats[attackerId]) continue;

      let isLoss = false; let isWarHit = false;

      if (targetFactionId !== "" && defenderFacId === targetFactionId) {
        if (result.includes("lost") || result.includes("escape") || result.includes("draw") || result.includes("timeout") || result.includes("stalemate")) {
          isLoss = true;
        } else {
          isWarHit = true;
        }
      }

      let isChainHit = (respect > 0 && aFac === myFactionId && !isWarHit && !isLoss);
      let savesChain = (respect > 0 && !isNaN(timestamp) && aFac === myFactionId);

      if (isWarHit || isLoss || isChainHit || savesChain) {
        rdEvents.push({ id: attackerId, time: timestamp, isWarHit: isWarHit, isLoss: isLoss, isChainHit: isChainHit, savesChain: savesChain });
      }
    }

    // CALCULATE EXACT OVERAGE
    let totalOfficialHits = 0;
    for (let id in stats) {
      totalOfficialHits += stats[id].wh + stats[id].ch;
    }
    
    let hitsToRemove = totalOfficialHits - globalHitLimit;
    let cutoffTimestamp = Infinity; // Default: include everything

    // REVERSE TRIMMER (Newest to Oldest)
    if (hitsToRemove > 0) {
      // Sort newest events first
      rdEvents.sort((a, b) => b.time - a.time); 

      for (let event of rdEvents) {
        if (hitsToRemove <= 0) {
          cutoffTimestamp = event.time; // The exact second the limit was reached
          break;
        }
        
        let isSuccess = event.isWarHit || event.isChainHit;
        if (isSuccess && stats[event.id]) {
          // Attempt to deduct the hit from the player who made it
          if (event.isWarHit && stats[event.id].wh > 0) {
            stats[event.id].wh--;
            hitsToRemove--;
          } else if (event.isChainHit && stats[event.id].ch > 0) {
            stats[event.id].ch--;
            hitsToRemove--;
          } else if (stats[event.id].ch > 0) { // Fallback if categories misaligned
            stats[event.id].ch--;
            hitsToRemove--;
          } else if (stats[event.id].wh > 0) { // Fallback
            stats[event.id].wh--;
            hitsToRemove--;
          }
        }
      }
    }

    // NOW PROCESS LOSSES & SAVES USING THE OFFICIAL CUTOFF
    let validHitsForSaves = []; 
    for (let event of rdEvents) {
      // Ignore anything that happened after the cutoff time
      if (event.time > cutoffTimestamp) continue;

      if (event.isLoss) stats[event.id].wl += 1;
      if (event.savesChain) validHitsForSaves.push({ time: event.time, id: event.id });
    }

    // Process Chain Saves
    validHitsForSaves.sort((a, b) => a.time - b.time); // Chronological order for gaps
    let timeRemainingSeconds = 180; 
    let cleanStr = limitStr.replace(/[^\d.:]/g, ''); 
    let parts = cleanStr.split(":");
    if (parts.length === 3) {
      let h = parseInt(parts[0], 10) || 0; let m = parseInt(parts[1], 10) || 0; let s = parseInt(parts[2], 10) || 0;
      if (h > 0 && m === 0 && s === 0) timeRemainingSeconds = h * 60; else timeRemainingSeconds = (h * 3600) + (m * 60) + s;
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
        if (gap >= requiredGap && gap <= 300) stats[hit.id].cs += 1;
      }
      lastHitTime = hit.time;
    }
  }

  // ---> 5. OUTPUT & ENFORCE SUB-LIMITS WITH OVERFLOW <---
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

  const output = existingIds.map(id => {
    let s = stats[id] || { wh: 0, wa: 0, wl: 0, wi: 0, ch: 0, cs: 0, ret: 0, res: 0, abr: 0, ws: 0 };
    return [s.wh, s.wa, s.wl, s.wi, s.ch, s.cs, s.ret, s.res, s.abr, s.ws];
  });
  
  if (output.length > 0) payoutSheet.getRange(3, 5, output.length, 10).setValues(output);
  
  // --- 6. INJECT "LEFT FACTION" PLAYERS & ANCHORED FORMULAS ---
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

  ss.toast(`Payout calculated! Exact chronological trim applied for F3 limit.`, "Success", 5);
}