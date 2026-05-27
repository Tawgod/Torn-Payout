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
      "Attacker ID", "Attacker Name", "Attacker Faction",
      "Defender ID", "Defender Name", "Defender Faction",
      "Result", "Respect",
      "Fair Fight", "War Bonus", "Retaliation", "Group Attack", "Overseas", "Chain Bonus", "Warlord Bonus"
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

  // We start at the present moment
  let searchToUnix = "";

  // 3. THE PAGINATION LOOP (Walks Backwards)
  while (keepSearching && safetyCounter < 15) { // 15 pages max = 1,500 hits per run to prevent Google timeouts
    safetyCounter++;
    
    let url = `https://api.torn.com/faction/?selections=attacks&key=${API_KEY}`;
    // If we have an older timestamp, tell Torn to look strictly BEFORE it
    if (searchToUnix !== "") {
      url += `&to=${searchToUnix}`;
    }

    let response;
    try {
      response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    } catch(e) { break; } // If Torn crashes, exit safely

    const json = JSON.parse(response.getContentText());
    if (json.error || !json.attacks) break;

    const attacks = json.attacks;
    const attackIds = Object.keys(attacks);
    
    if (attackIds.length === 0) break; // Reached the end of Torn's history

    let oldestTimestampInBatch = Infinity;
    let overlapFound = false;

    for (let id of attackIds) {
      let atk = attacks[id];
      
      // Track the oldest timestamp in this batch so we know where to search next
      if (atk.timestamp_started < oldestTimestampInBatch) {
        oldestTimestampInBatch = atk.timestamp_started;
      }

      // If we find an ID we already have, we have successfully caught up to our database!
      if (existingIds.has(id.toString())) {
        overlapFound = true;
        continue; // Skip logging it, but let the loop finish parsing the rest
      }

      let modifiers = atk.modifiers || {};
      newRows.push([
        id,
        new Date(atk.timestamp_started * 1000),
        new Date(atk.timestamp_ended * 1000),
        atk.attacker_id, atk.attacker_name, atk.attacker_faction,
        atk.defender_id, atk.defender_name, atk.defender_faction,
        atk.result, atk.respect,
        modifiers.fair_fight || 1, modifiers.war || 1, modifiers.retaliation || 1,
        modifiers.group_attack || 1, modifiers.overseas || 1, modifiers.chain_bonus || 1, modifiers.warlord || 1
      ]);
      
      existingIds.add(id.toString()); // Add to live memory
    }

    // Determine if we need to loop again
    if (overlapFound || attackIds.length < 100) {
      // We either found our existing data, or Torn ran out of history
      keepSearching = false; 
    } else {
      // We got exactly 100 brand new hits. We need to go deeper into the past!
      searchToUnix = oldestTimestampInBatch; 
    }
  }

  // 4. Sort chronologically and append to the BOTTOM of the sheet
  if (newRows.length > 0) {
    newRows.sort((a, b) => a[2] - b[2]); // Sort oldest to newest
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
}