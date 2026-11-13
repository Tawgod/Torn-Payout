const ROB_ATTACK_WAREHOUSE_HEADERS = [
  "Attack ID", "Start Time", "End Time",
  "Attacker ID", "Attacker Name", "Attacker Faction",
  "Defender ID", "Defender Name", "Defender Faction",
  "Result", "Respect",
  "Fair Fight", "War Bonus", "Retaliation", "Group Attack", "Overseas", "Chain Bonus", "Warlord Bonus",
  "Respect Lost"
];

function robAttackWarehouseConfig_() {
  const props = PropertiesService.getScriptProperties();
  const enabled = String(props.getProperty("ROB_ATTACK_WAREHOUSE_ENABLED") || "")
    .trim().toLowerCase() === "true";
  return {
    enabled: enabled,
    url: String(props.getProperty("ROB_ATTACK_WAREHOUSE_URL") || "").trim().replace(/\/+$/, ""),
    secret: String(props.getProperty("ROB_PAYOUT_EXPORT_SECRET") || "").trim()
  };
}

function robAttackWarehouseEnabled_() {
  return robAttackWarehouseConfig_().enabled;
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
  const cfg = robAttackWarehouseConfig_();
  if (!cfg.enabled) return null;
  if (!cfg.url || !cfg.secret) {
    throw new Error(
      "ROB attack warehouse is enabled but ROB_ATTACK_WAREHOUSE_URL or ROB_PAYOUT_EXPORT_SECRET is missing."
    );
  }

  const faction = String(factionId || "").replace(/,/g, "").trim();
  if (!/^\d+$/.test(faction)) {
    throw new Error("A numeric faction ID is required for the ROB attack warehouse.");
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
  let pages = 0;

  while (pages < 100) {
    pages++;
    const query = [
      "faction_id=" + encodeURIComponent(faction),
      "from=" + encodeURIComponent(String(startSec)),
      "to=" + encodeURIComponent(String(endSec)),
      "limit=5000",
      "cursor_started=" + encodeURIComponent(String(cursorStarted)),
      "cursor_id=" + encodeURIComponent(String(cursorId))
    ].join("&");
    const response = UrlFetchApp.fetch(cfg.url + "/internal/payout-attacks?" + query, {
      method: "get",
      muteHttpExceptions: true,
      headers: { Authorization: "Bearer " + cfg.secret }
    });
    const status = response.getResponseCode();
    let body;
    try {
      body = JSON.parse(response.getContentText());
    } catch (_e) {
      throw new Error("ROB attack warehouse returned invalid JSON.");
    }
    if (status >= 400 || !body || body.ok !== true || !Array.isArray(body.attacks)) {
      const detail = body && body.error ? String(body.error) : ("HTTP " + status);
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
    const nextStartedMs = Date.parse(nextStarted);
    const cursorStartedMs = Date.parse(cursorStarted);
    const nextId = Number(body.next_cursor.attack_id);
    if (!nextStarted || !Number.isFinite(nextStartedMs) || !Number.isFinite(nextId)) {
      throw new Error("ROB attack warehouse returned an invalid pagination cursor.");
    }
    if (
      nextStartedMs < cursorStartedMs ||
      (nextStartedMs === cursorStartedMs && nextId <= cursorId)
    ) {
      throw new Error("ROB attack warehouse pagination cursor did not advance.");
    }
    cursorStarted = nextStarted;
    cursorId = nextId;
  }

  throw new Error("ROB attack warehouse exceeded the 100-page safety limit.");
}
