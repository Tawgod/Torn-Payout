function rawDataApiKey_() {
  const key = String(PropertiesService.getScriptProperties().getProperty("TORN_API_KEY") || "").trim();
  if (!key) throw new Error("Missing TORN_API_KEY in Apps Script Script Properties.");
  return key;
}

function fetchRawDataAttacksPage_(toUnix) {
  let url = "https://api.torn.com/v2/faction/attacks?limit=100&sort=DESC&comment=" +
            encodeURIComponent("Payout Raw Data");
  if (toUnix !== "" && toUnix !== null && toUnix !== undefined) {
    url += "&to=" + encodeURIComponent(String(toUnix));
  }
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { Authorization: "ApiKey " + rawDataApiKey_() }
  });
  const status = response.getResponseCode();
  let json;
  try {
    json = JSON.parse(response.getContentText());
  } catch (_e) {
    throw new Error("Torn API returned invalid JSON while loading payout raw data.");
  }
  if (status >= 400 || json.error) {
    const detail = json && json.error ? JSON.stringify(json.error) : ("HTTP " + status);
    throw new Error("Torn API raw-data request failed: " + detail);
  }
  return json;
}

function rawAttackToWarehouseRow_(atk) {
  const modifiers = (atk && atk.modifiers) || {};
  // IMPORTANT: payout Chain Bonus is the respect multiplier in modifiers.chain.
  // The top-level attack.chain value is a chain counter and may be null starting
  // Jan 1, 2027; it is intentionally not used for payout bonus calculations.
  return [
    atk.id || atk.code || "",
    new Date(Number(atk.started || 0) * 1000),
    new Date(Number(atk.ended || 0) * 1000),
    atk.attacker?.id || "",
    atk.attacker?.name || "Unknown/Stealthed",
    atk.attacker?.faction?.id || "",
    atk.defender?.id || "",
    atk.defender?.name || "Unknown/Stealthed",
    atk.defender?.faction?.id || "",
    atk.result || "",
    Number(atk.respect_gain || 0),
    Number(modifiers.fair_fight || 1),
    Number(modifiers.war || 1),
    Number(modifiers.retaliation || 1),
    Number(modifiers.group || 1),
    Number(modifiers.overseas || 1),
    Number(modifiers.chain || 1),
    Number(modifiers.warlord || 0),
    Number(atk.respect_loss || 0)
  ];
}

function runBackendTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");
  if (!sheet) return;

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

  const existingData = sheet.getRange("A:A").getValues();
  const existingIds = new Set();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0]) existingIds.add(existingData[i][0].toString());
  }

  const newRows = [];
  let keepSearching = true;
  let safetyCounter = 0;
  let searchToUnix = "";

  while (keepSearching && safetyCounter < 15) {
    safetyCounter++;

    let json;
    try {
      json = fetchRawDataAttacksPage_(searchToUnix);
    } catch (e) {
      console.error("Payout raw-data sync failed: " + e.message);
      break;
    }

    const attacks = Array.isArray(json.attacks) ? json.attacks : [];
    if (attacks.length === 0) break;

    let oldestTimestampInBatch = Infinity;
    let overlapFound = false;

    for (const atk of attacks) {
      const attackId = atk.id || atk.code;
      if (!attackId) continue;

      const started = Number(atk.started || 0);
      if (started > 0 && started < oldestTimestampInBatch) oldestTimestampInBatch = started;

      if (existingIds.has(attackId.toString())) {
        overlapFound = true;
        continue;
      }

      newRows.push(rawAttackToWarehouseRow_(atk));
      existingIds.add(attackId.toString());
    }

    if (overlapFound || attacks.length < 100 || !Number.isFinite(oldestTimestampInBatch)) {
      keepSearching = false;
    } else {
      searchToUnix = oldestTimestampInBatch;
    }
  }

  if (newRows.length > 0) {
    newRows.sort((a, b) => a[2] - b[2]);
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
}
