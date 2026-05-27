function auditRecentAttacks() {
  const API_KEY = "K69UD2sZja0BBg3Q"; 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");
  
  if (!sheet) return;

  ss.toast("Starting 4-day deep scan for missing hits...", "System", -1);

  // 1. Setup Memory (To know what we already have)
  const existingData = sheet.getRange("A:A").getValues();
  const existingIds = new Set();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0]) existingIds.add(existingData[i][0].toString());
  }

  // 2. Set the 4-Day Boundary limit
  const nowUnix = Math.floor(Date.now() / 1000);
  const fourDaysAgoUnix = nowUnix - (4 * 24 * 60 * 60);

  const missingRows = [];
  let keepSearching = true;
  let safetyCounter = 0; 
  let searchToUnix = "";

  // 3. THE AUDIT LOOP (Forces a walk through the last 96 hours)
  while (keepSearching && safetyCounter < 150) { // Increased limit to 15,000 hits to cover heavy chaining
    safetyCounter++;
    
    let url = `https://api.torn.com/faction/?selections=attacks&key=${API_KEY}`;
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
    const attackIds = Object.keys(attacks);
    
    if (attackIds.length === 0) break; 

    let oldestTimestampInBatch = Infinity;

    for (let id of attackIds) {
      let atk = attacks[id];
      
      if (atk.timestamp_started < oldestTimestampInBatch) {
        oldestTimestampInBatch = atk.timestamp_started;
      }

      // If we already have it, SKIP IT, but DO NOT STOP THE LOOP.
      if (existingIds.has(id.toString())) {
        continue; 
      }

      // We found a missing hit!
      let modifiers = atk.modifiers || {};
      missingRows.push([
        id,
        new Date(atk.timestamp_started * 1000),
        new Date(atk.timestamp_ended * 1000),
        atk.attacker_id, atk.attacker_name, atk.attacker_faction,
        atk.defender_id, atk.defender_name, atk.defender_faction,
        atk.result, atk.respect,
        modifiers.fair_fight || 1, modifiers.war || 1, modifiers.retaliation || 1,
        modifiers.group_attack || 1, modifiers.overseas || 1, modifiers.chain_bonus || 1, modifiers.warlord || 1
      ]);
      
      existingIds.add(id.toString()); 
    }

    // THE KILL SWITCH: Only stop if we've traveled backward past 4 days ago
    if (oldestTimestampInBatch < fourDaysAgoUnix || attackIds.length < 100) {
      keepSearching = false; 
    } else {
      searchToUnix = oldestTimestampInBatch; 
    }
  }

  // 4. Sort chronologically and append
  if (missingRows.length > 0) {
    missingRows.sort((a, b) => a[2] - b[2]); // Sort oldest to newest
    sheet.getRange(sheet.getLastRow() + 1, 1, missingRows.length, missingRows[0].length).setValues(missingRows);
    SpreadsheetApp.getUi().alert(`✅ AUDIT COMPLETE\n\nFound and added ${missingRows.length} missing attacks from the last 4 days!`);
  } else {
    SpreadsheetApp.getUi().alert(`✅ AUDIT COMPLETE\n\nNo missing hits found in the last 4 days. Your data is perfectly synced.`);
  }
}