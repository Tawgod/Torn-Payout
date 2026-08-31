const API_KEY = "K69UD2sZja0BBg3Q"; 

function runBackendTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");
  if (!sheet) return;

  // 1. Initialize Headers if empty
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    const headers = [
      "Attack ID", "Start Time", "End Time",
      "Attacker ID", "Attacker Name", "Attacker Faction ID",
      "Defender ID", "Defender Name", "Defender Faction ID",
      "Result", "Respect",
      "Fair Fight", "War Bonus", "Retaliation", "Group Attack", "Overseas", "Chain Bonus", "Warlord Bonus",
      "Respect Lost"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  // 2. Setup Memory to catch overlaps
  const existingData = sheet.getRange("A:A").getValues();
  const existingIds = new Set();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0]) existingIds.add(existingData[i][0].toString());
  }

  const newRows = [];
  let keepSearching = true;
  let safetyCounter = 0;
  let searchToUnix = "";

  // 3. THE PAGINATION LOOP (Walks Backwards)
  while (keepSearching && safetyCounter < 15) { 
    safetyCounter++;
    
    let url = `https://api.torn.com/v2/faction/attacks?key=${API_KEY}`;
    if (searchToUnix !== "") {
      url += `&to=${searchToUnix}`;
    }

    let response;
    try {
      response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    } catch(e) { break; } 

    const json = JSON.parse(response.getContentText());
    if (json.error || !json.attacks) break;

    const attacks = json.attacks;
    if (!attacks || attacks.length === 0) break; 

    let oldestTimestampInBatch = Infinity;
    let overlapFound = false;

    for (let atk of attacks) {
      let attackId = atk.id || atk.code;

      if (atk.started < oldestTimestampInBatch) {
        oldestTimestampInBatch = atk.started;
      }

      if (existingIds.has(attackId.toString())) {
        overlapFound = true;
        continue; 
      }

      let modifiers = atk.modifiers || {};
      
      newRows.push([
        attackId,
        new Date(atk.started * 1000),
        new Date(atk.ended * 1000),
        atk.attacker?.id || "", 
        atk.attacker?.name || "Unknown/Stealthed", 
        atk.attacker?.faction?.id || "",
        atk.defender?.id || "", 
        atk.defender?.name || "Unknown/Stealthed", 
        atk.defender?.faction?.id || "",
        atk.result || "", 
        atk.respect_gain || 0,
        modifiers.fair_fight || 1, modifiers.war || 1, modifiers.retaliation || 1,
        modifiers.group || 1, modifiers.overseas || 1, modifiers.chain || 1, modifiers.warlord || 0,
        atk.respect_loss || 0
      ]);
      
      existingIds.add(attackId.toString()); 
    }

    if (overlapFound || attacks.length < 100) {
      keepSearching = false;
    } else {
      searchToUnix = oldestTimestampInBatch;
    }
  }

  // 4. Sort chronologically and append to the BOTTOM of the sheet
  if (newRows.length > 0) {
    newRows.sort((a, b) => a[2] - b[2]); 
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
}