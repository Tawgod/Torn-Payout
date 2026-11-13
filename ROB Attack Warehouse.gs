const ROB_ATTACK_WAREHOUSE_HEADERS = [
  "Attack ID", "Start Time", "End Time",
  "Attacker ID", "Attacker Name", "Attacker Faction",
  "Defender ID", "Defender Name", "Defender Faction",
  "Result", "Respect",
  "Fair Fight", "War Bonus", "Retaliation", "Group Attack", "Overseas", "Chain Bonus", "Warlord Bonus",
  "Respect Lost"
];

function robAttackWarehouseEnabled_() {
  const props = PropertiesService.getScriptProperties();
  const override = String(props.getProperty("ROB_ATTACK_WAREHOUSE_ENABLED") || "")
    .trim().toLowerCase();
  if (["true", "1", "yes", "on"].includes(override)) return true;
  if (["false", "0", "no", "off"].includes(override)) return false;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = ss.getSheetByName(SETTINGS.configSheet || "Config");
    if (!config) return false;
    const value = String(
      config.getRange(SETTINGS.robWarehouseEnabledCell || "B9").getDisplayValue() || ""
    ).trim().toLowerCase();
    return ["true", "1", "yes", "on", "enabled"].includes(value);
  } catch (_e) {
    return false;
  }
}

function robFactionIdForArchiveKey_(factionKey) {
  const key = String(factionKey || "").trim().toLowerCase();
  if (key === "ironsides") return "40664";
  if (key === "resolute") return "46442";
  return "";
}

function robAttackWarehouseRow_(attack) {
  attack = attack || {};
  if (attack.modifiers_complete !== true) {
    throw new Error(
      "ROB attack warehouse returned attack " + String(attack.attack_id || "unknown") +
      " without complete payout modifier data."
    );
  }
  // chain_value is the Torn chain counter. Payout Chain Bonus is the
  // respect multiplier chain_bonus; never substitute one for the other.
  return [
    attack.attack_id || "",
    attack.started_at ? new Date(attack.started_at) : "",
    attack.ended_at ? new Date(attack.ended_at) : "",
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
    Number(attack.warlord_bonus),
    Number(attack.respect_loss || 0)
  ];
}

function loadWarDataFromRobWarehouse_(factionId, startUnixMs, endUnixMs) {
  if (!robAttackWarehouseEnabled_()) return null;

  const faction = String(factionId || "").replace(/,/g, "").trim();
  if (!/^\d+$/.test(faction)) {
    throw new Error("A numeric faction ID is required for the ROB attack warehouse.");
  }

  const archiveCfg = getArchiveDatabaseConfig_();
  const expectedFaction = robFactionIdForArchiveKey_(archiveCfg.factionKey);
  if (!expectedFaction || expectedFaction !== faction) {
    throw new Error("Payout faction ID does not match the configured Railway archive namespace.");
  }

  const startSec = Math.floor(Number(startUnixMs) / 1000);
  const endSec = Math.ceil(Number(endUnixMs) / 1000);
  if (!Number.isFinite(startSec) || !Number.isFinite(endSec) || startSec <= 0 || endSec <= startSec) {
    throw new Error("Invalid payout time window for the ROB attack warehouse.");
  }

  const rows = [ROB_ATTACK_WAREHOUSE_HEADERS.slice()];
  const seen = new Set();
  let cursorStarted = new Date(startSec * 1000).toISOString();
  let cursorId = 0;
  const seenCursors = new Set();
  let pages = 0;

  while (pages < 100) {
    const cursorKey = cursorStarted + "|" + String(cursorId);
    if (seenCursors.has(cursorKey)) {
      throw new Error("ROB attack warehouse pagination cursor repeated.");
    }
    seenCursors.add(cursorKey);
    pages++;
    const query = [
      "faction_id=" + encodeURIComponent(faction),
      "from=" + encodeURIComponent(String(startSec)),
      "to=" + encodeURIComponent(String(endSec)),
      "limit=5000",
      "cursor_started=" + encodeURIComponent(String(cursorStarted)),
      "cursor_id=" + encodeURIComponent(String(cursorId))
    ].join("&");
    const body = archiveApiRequest_(
      "/api/rob-attack-warehouse?faction=" + encodeURIComponent(archiveCfg.factionKey) + "&" + query
    );
    if (!body || body.ok !== true || !Array.isArray(body.attacks)) {
      const detail = body && body.error ? String(body.error) : "Invalid warehouse response";
      throw new Error("ROB attack warehouse request failed: " + detail);
    }
    if (body.source !== "rob_shared_attack_warehouse") {
      throw new Error("ROB attack warehouse returned an unexpected source identifier.");
    }

    for (const attack of body.attacks) {
      const id = String(attack && attack.attack_id || "").trim();
      if (!id || seen.has(id)) continue;
      rows.push(robAttackWarehouseRow_(attack));
      seen.add(id);
    }

    if (!body.next_cursor) return rows;
    const nextStarted = String(body.next_cursor.started || "").trim();
    const nextId = Number(body.next_cursor.attack_id);
    if (!nextStarted || !Number.isFinite(Date.parse(nextStarted)) || !Number.isFinite(nextId)) {
      throw new Error("ROB attack warehouse returned an invalid pagination cursor.");
    }
    const nextKey = nextStarted + "|" + String(nextId);
    if (nextKey === cursorKey || seenCursors.has(nextKey)) {
      throw new Error("ROB attack warehouse pagination cursor did not advance.");
    }
    // Do not compare Date.parse values for ordering here: Apps Script truncates
    // sub-millisecond precision. The ROB endpoint validates cursor ordering;
    // the client only detects repeated/cyclic cursor pairs.
    cursorStarted = nextStarted;
    cursorId = nextId;
  }

  throw new Error("ROB attack warehouse exceeded the 100-page safety limit.");
}
