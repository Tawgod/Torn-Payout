/**
 * Master Audit for RPG Att Data
 * Focus: Identify why the hit count isn't 5,000.
 */
function runMasterDataAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("RPG Att Data") || ss.getSheetByName("RD"); // Adjust name as needed
  const configSheet = ss.getSheetByName("Config");
  
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert("Could not find the Master Data sheet. Check the name!");
    return;
  }

  // 1. SETTINGS
  const myFactionId = "40664"; // Your Faction ID
  const startWindow = new Date("2026-04-25T15:00:03").getTime();
  const endWindow = new Date("2026-04-27T12:42:22").getTime();
  const hitResults = ["hospitalized", "mugged", "attacked"];

  // 2. DATA LOAD
  const data = masterSheet.getDataRange().getValues();
  let stats = {};
  
  // LOG COUNTERS
  let audit = {
    totalRows: 0,
    wrongTime: 0,
    incomingAttacks: 0, // THE "MORE HITS" CULPRIT
    assists: 0,
    losses: 0,
    validOutgoingHits: 0
  };

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let timestamp = new Date(row[2]).getTime();
    let attackerFacId = row[5].toString().trim();
    let result = row[9] ? row[9].toString().toLowerCase().trim() : "";
    let attackerId = row[3].toString().trim();
    let attackerName = row[4].toString().trim();

    // FILTER 1: Time Window
    if (timestamp < startWindow || timestamp > endWindow) {
      audit.wrongTime++;
      continue;
    }

    audit.totalRows++;

    // FILTER 2: Faction Shield (Incoming vs Outgoing)
    // If this ID is NOT 40664, it's an enemy hitting you!
    if (attackerFacId !== myFactionId) {
      audit.incomingAttacks++;
      continue;
    }

    // FILTER 3: Result Type (Hits vs Non-Hits)
    if (result === "assist") {
      audit.assists++;
      continue;
    }
    if (result === "lost" || result === "interrupted") {
      audit.losses++;
      continue;
    }

    // If it passed all filters, it is a VALID OUTGOING HIT
    if (hitResults.includes(result)) {
      audit.validOutgoingHits++;
      if (!stats[attackerId]) stats[attackerId] = { name: attackerName, count: 0 };
      stats[attackerId].count++;
    }
  }

  // 3. GENERATE AUDIT REPORT
  let report = `--- DATA RECONCILIATION REPORT ---\n`;
  report += `Window: 15:00:03 (25th) to 12:42:22 (27th)\n\n`;
  report += `✅ VALID OUTGOING HITS: ${audit.validOutgoingHits}\n`;
  report += `❌ INCOMING (ENEMY) HITS SKIPPED: ${audit.incomingAttacks}\n`;
  report += `❌ ASSISTS SKIPPED: ${audit.assists}\n`;
  report += `❌ LOSSES SKIPPED: ${audit.losses}\n`;
  report += `❌ HITS OUTSIDE TIME RANGE: ${audit.wrongTime}\n\n`;
  
  if (audit.validOutgoingHits < 5000) {
    report += `⚠️ DISCREPANCY: You are missing ${5000 - audit.validOutgoingHits} hits from your log. You likely need to pull more API pages.\n`;
  }

  // Optional: List top 5 members to check for "overcounting"
  report += `\n--- TOP MEMBER COUNTS (VERIFY THESE) ---\n`;
  let sorted = Object.values(stats).sort((a,b) => b.count - a.count).slice(0, 5);
  sorted.forEach(m => report += `${m.name}: ${m.count} hits\n`);

  SpreadsheetApp.getUi().alert(report);
}