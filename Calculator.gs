function runPayoutMath() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rdSheet = ss.getSheetByName(SETTINGS.rdSheet);
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  const payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  const rosterSheet = ss.getSheetByName(SETTINGS.rosterSheet);
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);

  if (!rdSheet || !dashSheet || !payoutSheet || !rosterSheet || !configSheet) return;

  // 1. GET IDENTITY & LIMITS
  const myFactionId = configSheet.getRange("B7").getValue().toString().trim();
  const targetFactionId = dashSheet.getRange("C3").getValue().toString().trim();
  
  // I3 = Global Faction Cap (The Term)
  // I4/I5 = Personal Caps (Max per player)
  const factionGlobalLimit = parseInt(dashSheet.getRange("I3").getValue()) || 999999;
  const personalWarLimit   = parseInt(dashSheet.getRange("I4").getValue()) || 999999;
  const personalChainLimit = parseInt(dashSheet.getRange("I5").getValue()) || 999999;

  // 2. CLEAR PREVIOUS DATA
  const lastPayoutRow = payoutSheet.getLastRow();
  if (lastPayoutRow >= 3) {
    payoutSheet.getRange(3, 5, lastPayoutRow - 2, 10).clearContent();
  }

  // 3. GET ROSTER MAP
  const rosterData = rosterSheet.getDataRange().getValues();
  const rosterMap = rosterData.slice(1).map(r => r[0].toString().trim());

  // 4. PROCESS RD DATA
  const rdData = rdSheet.getDataRange().getValues();
  const successfulResults = ["hospitalized", "mugged", "attacked"];
  let rowsToProcess = rdData.slice(1);
  rowsToProcess.sort((a, b) => new Date(a[2]) - new Date(b[2]));

  let stats = {};
  let globalTotalCounter = 0; // Faction-wide term tracker
  let ghostWar = 0, ghostChain = 0, ghostRespect = 0;

  for (let i = 0; i < rowsToProcess.length; i++) {
    // Stop everything if the Faction-wide term limit (I3) is hit
    if (globalTotalCounter >= factionGlobalLimit) break;
    
    let row = rowsToProcess[i];
    
    // Faction Shield
    if (row[5].toString().trim() !== myFactionId) continue; 

    let result = row[9] ? row[9].toString().toLowerCase().trim() : "";
    if (!successfulResults.includes(result)) continue;

    let attackerId = row[3].toString().trim();
    let defFaction = row[8].toString().trim();
    let respect = parseFloat(row[10]) || 0;
    let isWar = (defFaction === targetFactionId && targetFactionId !== "" && targetFactionId !== "None Set");

    // Ensure the player object exists in our memory
    if (!stats[attackerId]) {
      stats[attackerId] = { w: 0, c: 0, r: 0 };
    }

    // --- APPLY INDIVIDUAL LIMITS ---
    if (isWar) {
      // Does this PLAYER have room for another War Hit?
      if (stats[attackerId].w < personalWarLimit) {
        stats[attackerId].w++;
        stats[attackerId].r += respect;
        globalTotalCounter++;
      }
    } else {
      // Does this PLAYER have room for another Chain Hit?
      if (stats[attackerId].c < personalChainLimit) {
        stats[attackerId].c++;
        stats[attackerId].r += respect;
        globalTotalCounter++;
      }
    }
  }

  // 5. MAP DATA TO ROSTER
  const payoutIds = payoutSheet.getRange(3, 1, lastPayoutRow - 2, 1).getValues();
  const output = payoutIds.map(idRow => {
    let s = stats[idRow[0].toString().trim()] || { w:0, c:0, r:0 };
    // [WarHits, Assists, Losses, Inter, ChainHits, Saves, Retal, Respect, Abroad, WarScore]
    return [s.w, 0, 0, 0, s.c, 0, 0, s.r, 0, s.r];
  });

  payoutSheet.getRange(3, 5, output.length, 10).setValues(output);

  // 6. GHOST ROW CHECK (In case IDs are missing from roster)
  let ghostRowIdx = -1;
  const names = payoutSheet.getRange(1, 2, payoutSheet.getLastRow(), 1).getValues();
  for(let n=0; n<names.length; n++) {
    if (names[n][0] === "⚠️ NON-ROSTER / LEFT FACTION") { ghostRowIdx = n + 1; break; }
  }

  // Calculate if there are leftovers (IDs not in the current roster)
  let totalCalculated = 0;
  for (let id in stats) {
    if (!rosterMap.includes(id)) {
      ghostWar += stats[id].w;
      ghostChain += stats[id].c;
      ghostRespect += stats[id].r;
    }
  }

  if (ghostWar > 0 || ghostChain > 0) {
    const ghostRow = ghostRowIdx !== -1 ? ghostRowIdx : payoutSheet.getLastRow() + 1;
    payoutSheet.getRange(ghostRow, 2).setValue("⚠️ NON-ROSTER / LEFT FACTION").setFontStyle("italic");
    payoutSheet.getRange(ghostRow, 5).setValue(ghostWar);
    payoutSheet.getRange(ghostRow, 9).setValue(ghostChain);
    payoutSheet.getRange(ghostRow, 12).setValue(ghostRespect);
    payoutSheet.getRange(ghostRow, 1, 1, 14).setBackground("#fce8e6");
  }

  ss.toast(`Personal limits applied! Max Chain: ${personalChainLimit} per player.`, "Success", 5);
}