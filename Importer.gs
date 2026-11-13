// ==========================================
// ROB SHARED ATTACK WAREHOUSE CLIENT
// ==========================================
function robPayoutExportConfig_(configSheet) {
  const properties = PropertiesService.getScriptProperties();
  const secret = String(properties.getProperty("ROB_PAYOUT_EXPORT_SECRET") || "").trim();
  const baseUrl = String(
    properties.getProperty("ROB_PAYOUT_EXPORT_URL") ||
    "https://chain-watch-web-production.up.railway.app/internal/payout-attacks"
  ).trim().replace(/\/+$/, "");

  const factionCell = (typeof SETTINGS !== "undefined" && SETTINGS.factionIdCell)
    ? SETTINGS.factionIdCell : "B7";
  const factionId = String(configSheet.getRange(factionCell).getValue() || "")
    .replace(/,/g, "").trim();

  if (!/^\d+$/.test(factionId)) {
    throw new Error("Missing or invalid Faction ID in Config " + factionCell + ".");
  }
  if (!secret) {
    throw new Error(
      "Missing ROB_PAYOUT_EXPORT_SECRET in Apps Script Script Properties."
    );
  }
  if (!/^https:\/\//i.test(baseUrl)) {
    throw new Error("ROB_PAYOUT_EXPORT_URL must be an HTTPS URL.");
  }
  return { baseUrl: baseUrl, secret: secret, factionId: factionId };
}


function robWarehouseAttackToRdRow_(attack) {
  if (!attack || attack.modifiers_complete !== true) {
    const attackId = attack && attack.attack_id ? attack.attack_id : "unknown";
    throw new Error(
      "ROB attack " + attackId +
      " has incomplete payout modifiers. Import stopped rather than guessing multipliers."
    );
  }

  const started = new Date(attack.started_at);
  const ended = new Date(attack.ended_at);
  if (isNaN(started.getTime()) || isNaN(ended.getTime())) {
    throw new Error("ROB returned an attack with an invalid timestamp.");
  }

  return [
    attack.attack_id,
    started,
    ended,
    attack.attacker_id || "",
    attack.attacker_name || "Unknown/Stealthed",
    attack.attacker_faction_id || "",
    attack.defender_id || "",
    attack.defender_name || "Unknown/Stealthed",
    attack.defender_faction_id || "",
    attack.result || "",
    Number(attack.respect_gain || 0),
    Number(attack.fair_fight),
    Number(attack.war_bonus),
    Number(attack.retaliation),
    Number(attack.group_attack),
    Number(attack.overseas),
    Number(attack.chain_bonus),
    Number(attack.warlord_bonus)
  ];
}


function fetchRobPayoutRows_(configSheet, startMs, endMs) {
  const cfg = robPayoutExportConfig_(configSheet);
  const fromSec = Math.floor(Number(startMs) / 1000);
  const toSec = Math.ceil(Number(endMs) / 1000);
  if (!Number.isFinite(fromSec) || !Number.isFinite(toSec) || toSec <= fromSec) {
    throw new Error("Invalid payout attack time window.");
  }

  const rows = [];
  let cursorStarted = "";
  let cursorId = 0;
  let priorCursor = "";
  const maxPages = 40;

  for (let page = 0; page < maxPages; page++) {
    let url = cfg.baseUrl +
      "?faction_id=" + encodeURIComponent(cfg.factionId) +
      "&from=" + encodeURIComponent(String(fromSec)) +
      "&to=" + encodeURIComponent(String(toSec)) +
      "&limit=5000";

    if (cursorStarted) {
      url += "&cursor_started=" + encodeURIComponent(cursorStarted) +
             "&cursor_id=" + encodeURIComponent(String(cursorId));
    }

    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { Authorization: "Bearer " + cfg.secret }
    });

    const status = response.getResponseCode();
    let payload;
    try {
      payload = JSON.parse(response.getContentText());
    } catch (_e) {
      throw new Error("ROB payout warehouse returned invalid JSON.");
    }
    if (status >= 400 || !payload || payload.ok !== true) {
      const detail = payload && (payload.message || payload.error)
        ? String(payload.message || payload.error) : ("HTTP " + status);
      throw new Error("ROB payout warehouse request failed: " + detail);
    }
    if (String(payload.faction_id) !== cfg.factionId) {
      throw new Error("ROB payout warehouse returned data for the wrong faction.");
    }

    const attacks = Array.isArray(payload.attacks) ? payload.attacks : [];
    for (const attack of attacks) {
      rows.push(robWarehouseAttackToRdRow_(attack));
    }

    const next = payload.next_cursor;
    if (!next) return rows;

    const nextStarted = String(next.started || "").trim();
    const nextId = Number(next.attack_id || 0);
    const nextKey = nextStarted + "|" + nextId;
    if (!nextStarted || !Number.isFinite(nextId) || nextId <= 0 || nextKey === priorCursor) {
      throw new Error("ROB payout warehouse returned an invalid/repeating pagination cursor.");
    }
    priorCursor = nextKey;
    cursorStarted = nextStarted;
    cursorId = nextId;
  }

  throw new Error(
    "ROB payout warehouse exceeded the " + maxPages +
    "-page safety limit. Narrow the payout time window before retrying."
  );
}


// ==========================================
// 1. MASTER IMPORTER (Timezone-Safe + Duplicate Shield)
// ==========================================
function importWarData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  let localRdSheet = ss.getSheetByName(SETTINGS.rdSheet);

  if (!configSheet) {
    ss.toast("Missing Config sheet.", "Error", 5);
    return;
  }
  if (!dashSheet) {
    ss.toast("Please run 'Rebuild Dashboard' first.", "Error", 5);
    return;
  }

  // 1. Read Time Filters from Dashboard
  let startDateVal = dashSheet.getRange("F9").getValue();
  let endDateVal = dashSheet.getRange("F11").getValue();
  let startTimeStr = dashSheet.getRange("F10").getDisplayValue();
  let endTimeStr = dashSheet.getRange("F12").getDisplayValue();

  // --- IRONCLAD TIME FORCER ---
  let sd = new Date(startDateVal);
  let stParts = (startTimeStr ? startTimeStr : "00:00:00").split(":");
  
  // Safely parse start times without overriding zeroes
  let sHour = parseInt(stParts[0]); if (isNaN(sHour)) sHour = 0;
  let sMin = parseInt(stParts[1]); if (isNaN(sMin)) sMin = 0;
  let sSec = parseInt(stParts[2]); if (isNaN(sSec)) sSec = 0;

  let startUnix = Date.UTC(
    sd.getUTCFullYear(), sd.getUTCMonth(), sd.getUTCDate(),
    sHour, sMin, sSec
  );

  let ed = new Date(endDateVal);
  let etParts = (endTimeStr ? endTimeStr : "23:59:59").split(":");
  
  // Safely parse end times without overriding zeroes
  let eHour = parseInt(etParts[0]); if (isNaN(eHour)) eHour = 23;
  let eMin = parseInt(etParts[1]); if (isNaN(eMin)) eMin = 59;
  let eSec = parseInt(etParts[2]); if (isNaN(eSec)) eSec = 59;

  let endUnix = Date.UTC(
    ed.getUTCFullYear(), ed.getUTCMonth(), ed.getUTCDate(),
    eHour, eMin, eSec, 999 
  );

  // 2. Read the shared ROB attack warehouse (no direct Torn API call here).
  ss.toast("Connecting to ROB Attack Warehouse...", "System", 3);
  let rawData;
  try {
    rawData = fetchRobPayoutRows_(configSheet, startUnix, endUnix);
  } catch (e) {
    console.error("ROB payout import failed: " + e.message);
    SpreadsheetApp.getUi().alert(
      "Payout attack import stopped.\n\n" + e.message
    );
    return;
  }

  // 3. Filter the Data (With Duplicate Shield)
  const filteredData = [];
  const seenIds = new Set();
  let duplicatesBlocked = 0;
  
  for (let i = 0; i < rawData.length; i++) {
    let row = rawData[i];
    let attackId = row[0].toString().trim(); 
    
    // Force the Master Sheet's timestamp to be read accurately
    let rowTimeUnix;
    if (row[2] instanceof Date) {
      rowTimeUnix = row[2].getTime();
    } else {
      rowTimeUnix = new Date(row[2] + " UTC").getTime(); // Force string to UTC
    }

    // Check if within our Ironclad time window
    if (rowTimeUnix >= startUnix && rowTimeUnix <= endUnix) {
      
      // Duplicate Shield Check
      if (seenIds.has(attackId)) {
        duplicatesBlocked++;
        continue; 
      }

      filteredData.push(row);
      seenIds.add(attackId);
    }
  }

  // --- SORTING LOGIC ---
  filteredData.sort((a, b) => {
    let timeA = (a[2] instanceof Date) ? a[2].getTime() : new Date(a[2] + " UTC").getTime();
    let timeB = (b[2] instanceof Date) ? b[2].getTime() : new Date(b[2] + " UTC").getTime();
    return timeA - timeB;
  });
  
  // Add headers back to the top
  filteredData.unshift(SETTINGS.rdHeaders);

  // 4. Overwrite Local RD Tab
  if (!localRdSheet) {
    localRdSheet = ss.insertSheet(SETTINGS.rdSheet);
  } else {
    localRdSheet.clear();
  }
  
  if (filteredData.length > 1) {
    localRdSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
    
    // Human-Readable Boundary Receipt
    let startString = new Date(startUnix).toUTCString();
    let endString = new Date(endUnix).toUTCString();
    
    let msg = `IMPORT COMPLETE\n\n` +
              `Imported ${filteredData.length - 1} unique attacks.\n`;
    
    if (duplicatesBlocked > 0) {
      msg += `Blocked ${duplicatesBlocked} duplicates found in Master Data.\n\n`;
    } else {
      msg += `\n`;
    }

    msg += `--- TIME BOUNDARIES USED ---\n` +
           `Start: ${startString}\n` +
           `End:   ${endString}\n\n` +
           `If the number of hits is wrong, check these exact boundary times!`;
           
    SpreadsheetApp.getUi().alert(msg);
    
  } else {
    localRdSheet.getRange(1, 1, 1, SETTINGS.rdHeaders.length).setValues([SETTINGS.rdHeaders]);
    ss.toast("No hits found in that time window.", "System", 5);
  }
}
// ==========================================
// 2. SNIPER AUDIT UTILITY (Diagnose Missing Hits)
// ==========================================
function runSniperAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);

  if (!configSheet || !dashSheet) return;

  // 1. Ask for the specific Member ID
  const promptResponse = ui.prompt(
    "Sniper Audit", 
    "Enter the Torn ID of the member who is missing hits:", 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (promptResponse.getSelectedButton() !== ui.Button.OK) return;
  const targetId = promptResponse.getResponseText().trim();
  if (targetId === "") return;

  // 2. Setup Time Boundaries
  let startDateVal = dashSheet.getRange("F8").getValue();
  let endDateVal = dashSheet.getRange("F10").getValue();
  let startTimeStr = dashSheet.getRange("F9").getDisplayValue();
  let endTimeStr = dashSheet.getRange("F11").getDisplayValue();

  let sd = new Date(startDateVal);
  let stParts = (startTimeStr ? startTimeStr : "00:00:00").split(":");
  let startUnix = Date.UTC(sd.getUTCFullYear(), sd.getUTCMonth(), sd.getUTCDate(), parseInt(stParts[0])||0, parseInt(stParts[1])||0, parseInt(stParts[2])||0);

  let ed = new Date(endDateVal);
  let etParts = (endTimeStr ? endTimeStr : "23:59:59").split(":");
  let endUnix = Date.UTC(ed.getUTCFullYear(), ed.getUTCMonth(), ed.getUTCDate(), parseInt(etParts[0])||23, parseInt(etParts[1])||59, parseInt(etParts[2])||59, 999);

  // 3. Read the shared ROB warehouse, including one hour around the boundary.
  let masterData;
  try {
    masterData = fetchRobPayoutRows_(
      configSheet,
      startUnix - 3600000,
      endUnix + 3600000
    );
  } catch (e) {
    console.error("ROB sniper audit failed: " + e.message);
    ui.alert("Sniper Audit could not load ROB attack data.\n\n" + e.message);
    return;
  }

  // 4. Track exactly what happened
  let log = {
    tooEarly: 0,
    tooLate: 0,
    defending: 0,
    resultsBreakdown: {} // NEW: We will tally every exact word here
  };

  for (let i = 0; i < masterData.length; i++) {
    let row = masterData[i];
    let attackerId = row[3].toString().trim();
    let defenderId = row[6].toString().trim();
    let result = row[9] ? row[9].toString().toLowerCase().trim() : "blank/unknown";
    
    let rowTimeUnix = (row[2] instanceof Date) ? row[2].getTime() : new Date(row[2] + " UTC").getTime();

    // Check if they were the ATTACKER
    if (attackerId === targetId) {
      if (rowTimeUnix < startUnix) {
        if (startUnix - rowTimeUnix < 3600000) log.tooEarly++; 
      } 
      else if (rowTimeUnix > endUnix) {
        if (rowTimeUnix - endUnix < 3600000) log.tooLate++;
      } 
      else {
        // INSIDE THE WINDOW: Itemize exactly what the result was
        if (!log.resultsBreakdown[result]) {
          log.resultsBreakdown[result] = 0;
        }
        log.resultsBreakdown[result]++;
      }
    }
    
    // Check if they were the DEFENDER
    if (defenderId === targetId && rowTimeUnix >= startUnix && rowTimeUnix <= endUnix) {
      log.defending++;
    }
  }

  // 5. Build the Detailed Report
  let report = `🎯 SNIPER REPORT FOR ID: ${targetId} 🎯\n\n`;
  report += `--- EXACT RESULTS INSIDE TIME WINDOW ---\n`;
  
  let totalInWindow = 0;
  let sortedResults = Object.keys(log.resultsBreakdown).sort((a, b) => log.resultsBreakdown[b] - log.resultsBreakdown[a]);
  
  if (sortedResults.length === 0) {
    report += `No attacks found inside the exact time window.\n`;
  } else {
    for (let r of sortedResults) {
      report += `• ${r}: ${log.resultsBreakdown[r]}\n`;
      totalInWindow += log.resultsBreakdown[r];
    }
    report += `\nTotal hits logged in window: ${totalInWindow}\n\n`;
  }
  
  report += `--- OUTSIDE WINDOW / OTHER ---\n`;
  report += `❌ Hits just BEFORE start time: ${log.tooEarly}\n`;
  report += `❌ Hits just AFTER end time: ${log.tooLate}\n`;
  report += `🛡️ Times they were attacked: ${log.defending}\n\n`;
  
  report += `Compare the 'Total hits logged' above to what your Payout Sheet says!`;

  ui.alert(report);
}

