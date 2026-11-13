function auditRecentAttacks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");

  if (!sheet) return;

  ss.toast("Starting 4-day deep scan for missing hits...", "System", -1);

  const existingData = sheet.getRange("A:A").getValues();
  const existingIds = new Set();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0]) existingIds.add(existingData[i][0].toString());
  }

  const nowUnix = Math.floor(Date.now() / 1000);
  const fourDaysAgoUnix = nowUnix - (4 * 24 * 60 * 60);

  const missingRows = [];
  let keepSearching = true;
  let safetyCounter = 0;
  let searchToUnix = "";

  while (keepSearching && safetyCounter < 150) {
    safetyCounter++;

    let json;
    try {
      json = fetchRawDataAttacksPage_(searchToUnix);
    } catch (e) {
      console.error("Payout raw-data deep audit failed: " + e.message);
      SpreadsheetApp.getUi().alert("Raw-data audit stopped because the Torn API request failed.\n\n" + e.message);
      return;
    }

    const attacks = Array.isArray(json.attacks) ? json.attacks : [];
    if (attacks.length === 0) break;

    let oldestTimestampInBatch = Infinity;

    for (const atk of attacks) {
      const attackId = atk.id || atk.code;
      if (!attackId) continue;

      const started = Number(atk.started || 0);
      if (started > 0 && started < oldestTimestampInBatch) oldestTimestampInBatch = started;

      if (existingIds.has(attackId.toString())) continue;

      missingRows.push(rawAttackToWarehouseRow_(atk));
      existingIds.add(attackId.toString());
    }

    if (
      !Number.isFinite(oldestTimestampInBatch) ||
      oldestTimestampInBatch < fourDaysAgoUnix ||
      attacks.length < 100
    ) {
      keepSearching = false;
    } else {
      searchToUnix = oldestTimestampInBatch;
    }
  }

  if (missingRows.length > 0) {
    missingRows.sort((a, b) => a[2] - b[2]);
    sheet.getRange(sheet.getLastRow() + 1, 1, missingRows.length, missingRows[0].length).setValues(missingRows);
    SpreadsheetApp.getUi().alert(
      `✅ AUDIT COMPLETE\n\nFound and added ${missingRows.length} missing attacks from the last 4 days!`
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "✅ AUDIT COMPLETE\n\nNo missing hits found in the last 4 days."
    );
  }
}
