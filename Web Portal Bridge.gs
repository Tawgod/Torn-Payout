// ==========================================
// PAYOUT WEB PORTAL BRIDGE
// Bound Apps Script endpoint for controlled web <-> Sheet synchronization.
// This file is safe to deploy with the Sheet project; it does nothing until
// the Apps Script project is explicitly deployed as a web app and the shared
// bridge secret is configured in Script Properties.
// ==========================================

const PAYOUT_WEB_ALLOWED_DASHBOARD_LABELS = new Set([
  "War Report ID",
  "Chain Report ID",
  "Start Date",
  "Start Time",
  "End Date",
  "End Time",
  "Total Hits (Max Limit)",
  "Max War Hits",
  "Max Chain Hits (Faction Total)",
  "Pay Post-War Chain?",
  "Payout Preset",
  "Faction Cut %",
  "Max Faction Cut ($)",
  "Total Revenue",
  "- Temp Cost",
  "- Revives",
  "- Xanax",
  "- Other Cost",
  "Termed?"
]);

function payoutWebBridgeConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    secret: props.getProperty("PAYOUT_WEB_BRIDGE_SECRET") || "",
    factionKey: (props.getProperty("ARCHIVE_FACTION_KEY") || "").trim().toLowerCase()
  };
}

function payoutWebBridgeCanonical_(request) {
  return [
    String(request.ts || ""),
    String(request.nonce || ""),
    String(request.action || ""),
    String(request.faction || ""),
    JSON.stringify(request.payload || {})
  ].join(".");
}

function payoutWebBridgeSignature_(secret, canonical) {
  const bytes = Utilities.computeHmacSha256Signature(canonical, secret);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, "");
}

function payoutWebBridgeVerify_(request) {
  const cfg = payoutWebBridgeConfig_();
  if (!cfg.secret) throw new Error("PAYOUT_WEB_BRIDGE_SECRET is not configured.");
  if (!request || !request.sig) throw new Error("Missing bridge signature.");

  const ts = Number(request.ts || 0);
  const now = Math.floor(Date.now() / 1000);
  if (!Number.isFinite(ts) || Math.abs(now - ts) > 300) {
    throw new Error("Bridge request expired.");
  }

  const faction = String(request.faction || "").trim().toLowerCase();
  if (!faction || (cfg.factionKey && faction !== cfg.factionKey)) {
    throw new Error("Faction mismatch.");
  }

  const expected = payoutWebBridgeSignature_(cfg.secret, payoutWebBridgeCanonical_(request));
  if (expected !== String(request.sig || "")) throw new Error("Invalid bridge signature.");
  return cfg;
}

function payoutWebJson_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const raw = e && e.postData ? e.postData.contents : "";
    const request = JSON.parse(raw || "{}");
    payoutWebBridgeVerify_(request);

    const result = payoutWebDispatch_(String(request.action || ""), request.payload || {});
    return payoutWebJson_({ ok: true, result: result });
  } catch (err) {
    return payoutWebJson_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function payoutWebDispatch_(action, payload) {
  switch (action) {
    case "snapshot":
      return buildPayoutWebSnapshot_();

    case "war_live_snapshot":
      return buildWarLiveSnapshot_();

    case "set_dashboard_value":
      return payoutWebSetDashboardValue_(payload);

    case "set_dashboard_values":
      return payoutWebSetDashboardValues_(payload);

    case "add_push":
      return payoutWebAddPush_(payload);

    case "refresh_dashboard":
      refreshDashboard();
      return { refreshed: true };

    case "refresh_awards":
      if (typeof refreshAwards !== "function") throw new Error("Awards module is not available.");
      refreshAwards();
      return { refreshed: true };

    case "fetch_reports_and_rd":
      fetchOfficialReports();
      return { fetched: true };

    case "build_payout":
      buildPayoutTab();
      return { built: true };

    case "calculate_payout":
      runPayoutMath();
      return { calculated: true };

    case "build_final_payout":
      buildFinalPayoutTab();
      return { built: true };

    default:
      throw new Error("Unsupported payout web action: " + action);
  }
}

function payoutWebSetDashboardValue_(payload) {
  const label = String(payload.label || "").trim();
  if (!PAYOUT_WEB_ALLOWED_DASHBOARD_LABELS.has(label)) {
    throw new Error("Dashboard field is not web-editable: " + label);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SETTINGS.dashboardSheet);
  if (!dash) throw new Error("Dashboard sheet not found.");

  if (!setLabelValue_(dash, label, payload.value)) {
    throw new Error("Dashboard field not found: " + label);
  }

  SpreadsheetApp.flush();
  return { label: label, value: labelValue_(dash, label, "") };
}

function payoutWebSetDashboardValues_(payload) {
  const values = payload && payload.values && typeof payload.values === "object" ? payload.values : {};
  const changed = [];
  Object.keys(values).forEach(label => {
    changed.push(payoutWebSetDashboardValue_({ label: label, value: values[label] }));
  });
  return { changed: changed };
}

function payoutWebAddPush_(payload) {
  if (typeof addPushSetting !== "function") throw new Error("Push Settings module is not available.");
  const row = addPushSetting(payload.start_time || "", payload.start_date || "", payload.time_limit_min || "");
  SpreadsheetApp.flush();
  return { row: row };
}

function payoutWebSheetValues_(sheetName, maxRows, maxCols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;

  const rows = Math.min(sheet.getLastRow(), maxRows || 500);
  const cols = Math.min(sheet.getLastColumn(), maxCols || 40);
  if (rows < 1 || cols < 1) return [];
  return sheet.getRange(1, 1, rows, cols).getDisplayValues();
}

function payoutWebDashboardMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SETTINGS.dashboardSheet);
  if (!dash) return {};

  const data = dash.getDataRange().getDisplayValues();
  const out = {};
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length - 1; c++) {
      const label = String(data[r][c] || "").trim();
      if (!label) continue;
      if (
        PAYOUT_WEB_ALLOWED_DASHBOARD_LABELS.has(label) ||
        [
          "Enemy Faction Name", "Enemy Faction ID", "Official War Start",
          "Official War End", "War ID", "Total War Hits", "War Score",
          "Outcome (Result)", "Caches / Items Won", "Est. Cache Value",
          "NET PROFIT", "Actual Faction Deduction", "PAYOUT TOTAL",
          "First Attack Logged", "Latest Attack Logged", "Total Attacks Logged",
          "Total Respect Generated"
        ].indexOf(label) !== -1
      ) {
        out[label] = data[r][c + 1];
      }
    }
  }
  return out;
}

function buildPayoutWebSnapshot_() {
  const cfg = payoutWebBridgeConfig_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  return {
    faction_key: cfg.factionKey,
    spreadsheet_id: ss.getId(),
    generated_at: new Date().toISOString(),
    dashboard: payoutWebDashboardMap_(),
    payouts: payoutWebSheetValues_(SETTINGS.payoutSheet || "Payouts", 500, 40),
    final_payout: payoutWebSheetValues_("Final Payout", 500, 30),
    awards: payoutWebSheetValues_("Awards", 500, 30)
  };
}


// ==========================================
// CURRENT WAR LIVE DASHBOARD FEED
// Operational data only — intentionally excludes payout/financial data.
// ==========================================

function payoutWebCleanId_(value) {
  return value === null || value === undefined ? "" : String(value).replace(/,/g, "").trim();
}

function payoutWebSafeNumber_(value) {
  const n = Number(value);
  return isFinite(n) ? n : 0;
}

function payoutWebWarLiveBounties_(currentWarId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS.bountySheet || "Bounties");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const map = headerMapFromRow_(headers);

  const optionalIndex = name => {
    try { return headerIndex_(map, name, false); } catch (_e) { return -1; }
  };

  const idx = {
    date: optionalIndex("Date Logged"),
    placedBy: optionalIndex("Placed By"),
    target: optionalIndex("Target"),
    qty: optionalIndex("Quantity"),
    amount: optionalIndex("Bounty Amount"),
    refund: optionalIndex("Refund Amount"),
    status: optionalIndex("Status"),
    notes: optionalIndex("Notes"),
    source: optionalIndex("Source"),
    warId: optionalIndex("War ID"),
    verification: optionalIndex("War Verification")
  };

  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const status = idx.status >= 0 ? String(row[idx.status] || "").trim() : "";
    if (!row.some(v => v !== "" && v !== null)) continue;
    if (/denied|invalid/i.test(status)) continue;

    const warId = idx.warId >= 0 ? payoutWebCleanId_(row[idx.warId]) : "";
    if (currentWarId && warId && warId !== currentWarId) continue;

    rows.push({
      date_logged: idx.date >= 0 && row[idx.date] instanceof Date
        ? row[idx.date].toISOString()
        : (idx.date >= 0 ? String(row[idx.date] || "") : ""),
      placed_by: idx.placedBy >= 0 ? String(row[idx.placedBy] || "") : "",
      target: idx.target >= 0 ? String(row[idx.target] || "") : "",
      quantity: idx.qty >= 0 ? payoutWebSafeNumber_(row[idx.qty]) : 1,
      bounty_amount: idx.amount >= 0 ? payoutWebSafeNumber_(row[idx.amount]) : 0,
      refund_amount: idx.refund >= 0 ? payoutWebSafeNumber_(row[idx.refund]) : 0,
      status: status,
      notes: idx.notes >= 0 ? String(row[idx.notes] || "") : "",
      source: idx.source >= 0 ? String(row[idx.source] || "") : "",
      war_id: warId,
      verification: idx.verification >= 0 ? String(row[idx.verification] || "") : ""
    });
  }

  rows.sort((a, b) => {
    const ta = Date.parse(a.date_logged || "") || 0;
    const tb = Date.parse(b.date_logged || "") || 0;
    return tb - ta;
  });
  return rows.slice(0, 100);
}

function buildWarLiveSnapshot_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = payoutWebBridgeConfig_();
  const dash = ss.getSheetByName(SETTINGS.dashboardSheet);
  const rd = ss.getSheetByName(SETTINGS.rdSheet || "RD");

  const factionId = payoutWebCleanId_(
    ss.getSheetByName(SETTINGS.configSheet)
      .getRange(SETTINGS.factionIdCell || "B7")
      .getValue()
  );
  const targetFactionId = dash ? payoutWebCleanId_(labelValue_(dash, "Enemy Faction ID", "")) : "";
  const currentWarId = dash ? payoutWebCleanId_(labelValue_(dash, "War ID", labelValue_(dash, "War Report ID", ""))) : "";

  const snapshot = {
    faction_key: cfg.factionKey,
    generated_at: new Date().toISOString(),
    war: {
      war_id: currentWarId,
      enemy_name: dash ? String(labelValue_(dash, "Enemy Faction Name", "") || "") : "",
      enemy_faction_id: targetFactionId,
      official_start: dash ? String(labelValue_(dash, "Official War Start", "") || "") : "",
      official_end: dash ? String(labelValue_(dash, "Official War End", "") || "") : "",
      war_score: dash ? payoutWebSafeNumber_(labelValue_(dash, "War Score", 0)) : 0,
      outcome: dash ? String(labelValue_(dash, "Outcome (Result)", "") || "") : ""
    },
    totals: {
      attacks_logged: 0,
      successful_hits: 0,
      war_hits: 0,
      chain_hits: 0,
      assists: 0,
      losses: 0,
      interruptions: 0,
      retaliations: 0,
      overseas_hits: 0,
      base_respect: 0,
      raw_respect: 0,
      chain_bonus_hits: 0
    },
    members: [],
    recent_attacks: [],
    chain_bonus_hits: [],
    bounties: payoutWebWarLiveBounties_(currentWarId)
  };

  if (!rd || rd.getLastRow() < 2) return snapshot;

  const data = rd.getDataRange().getValues();
  const hm = headerMapFromRow_(data[0]);
  const col = name => headerIndex_(hm, name, true);
  const opt = name => headerIndex_(hm, name, false);

  const tCol = headerIndex_(hm, ["End Time", "Timestamp", "End"], true);
  const attackIdCol = opt("Attack ID");
  const aIdCol = col("Attacker ID");
  const aNameCol = col("Attacker Name");
  const aFacCol = headerIndex_(hm, ["Attacker Faction", "Attacker Faction ID"], true);
  const dIdCol = col("Defender ID");
  const dNameCol = opt("Defender Name");
  const dFacCol = headerIndex_(hm, ["Defender Faction", "Defender Faction ID"], true);
  const resultCol = col("Result");
  const respectCol = col("Respect");
  const cBonusCol = opt("Chain Bonus");
  const retaliationCol = opt("Retaliation");
  const overseasCol = opt("Overseas");

  const members = {};
  const pendingInterrupts = {};

  function memberFor(id, name) {
    if (!members[id]) {
      members[id] = {
        member_id: id,
        name: name || ("ID: " + id),
        attacks: 0,
        successful_hits: 0,
        war_hits: 0,
        chain_hits: 0,
        assists: 0,
        losses: 0,
        interruptions: 0,
        retaliations: 0,
        overseas_hits: 0,
        base_respect: 0,
        raw_respect: 0,
        chain_bonus_hits: 0
      };
    }
    return members[id];
  }

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const aFac = payoutWebCleanId_(row[aFacCol]);
    const dFac = payoutWebCleanId_(row[dFacCol]);
    const aId = payoutWebCleanId_(row[aIdCol]);
    const dId = payoutWebCleanId_(row[dIdCol]);
    const resultText = String(row[resultCol] || "");
    const result = resultText.toLowerCase();
    const respect = payoutWebSafeNumber_(row[respectCol]);
    const cBonus = cBonusCol >= 0 ? (payoutWebSafeNumber_(row[cBonusCol]) || 1) : 1;
    const baseRespect = respect / (cBonus > 1 ? cBonus : 1);
    const retaliation = retaliationCol >= 0 ? (payoutWebSafeNumber_(row[retaliationCol]) || 1) : 1;
    const overseas = overseasCol >= 0 ? (payoutWebSafeNumber_(row[overseasCol]) || 1) : 1;
    const timeValue = row[tCol] instanceof Date ? row[tCol].toISOString() : String(row[tCol] || "");

    if (aFac === targetFactionId && result.includes("interrupted") && aId) {
      pendingInterrupts[aId] = true;
    }

    if (aFac !== factionId || !aId) continue;

    const m = memberFor(aId, String(row[aNameCol] || ""));
    m.attacks++;
    snapshot.totals.attacks_logged++;

    const isWarOpponent = Boolean(targetFactionId && dFac === targetFactionId);
    const isAssist = isWarOpponent && result.includes("assist");
    const isLoss = isWarOpponent && (
      result.includes("lost") || result.includes("escape") || result.includes("draw") ||
      result.includes("timeout") || result.includes("stalemate")
    );
    const successfulHit = respect > 0 && (
      result.includes("hospitalized") || result.includes("attacked") ||
      result.includes("mugged") || (!result.includes("assist") && !isLoss)
    );

    if (successfulHit) {
      m.successful_hits++;
      snapshot.totals.successful_hits++;
      m.raw_respect += respect;
      m.base_respect += baseRespect;
      snapshot.totals.raw_respect += respect;
      snapshot.totals.base_respect += baseRespect;

      if (isWarOpponent) {
        m.war_hits++;
        snapshot.totals.war_hits++;
      } else {
        m.chain_hits++;
        snapshot.totals.chain_hits++;
      }

      if (retaliation > 1) {
        m.retaliations++;
        snapshot.totals.retaliations++;
      }
      if (isWarOpponent && overseas > 1) {
        m.overseas_hits++;
        snapshot.totals.overseas_hits++;
      }
      if (cBonus > 1) {
        m.chain_bonus_hits++;
        snapshot.totals.chain_bonus_hits++;
        snapshot.chain_bonus_hits.push({
          time: timeValue,
          attacker_id: aId,
          attacker_name: m.name,
          multiplier: cBonus,
          raw_respect: respect,
          base_respect: baseRespect
        });
      }
    }

    if (isAssist) {
      m.assists++;
      snapshot.totals.assists++;
    }
    if (isLoss) {
      m.losses++;
      snapshot.totals.losses++;
    }
    if (isWarOpponent && pendingInterrupts[dId] && successfulHit) {
      m.interruptions++;
      snapshot.totals.interruptions++;
      pendingInterrupts[dId] = false;
    }

    snapshot.recent_attacks.push({
      attack_id: attackIdCol >= 0 ? String(row[attackIdCol] || "") : "",
      time: timeValue,
      attacker_id: aId,
      attacker_name: m.name,
      defender_id: dId,
      defender_name: dNameCol >= 0 ? String(row[dNameCol] || "") : "",
      defender_faction_id: dFac,
      result: resultText,
      raw_respect: respect,
      base_respect: baseRespect,
      chain_bonus: cBonus,
      retaliation: retaliation,
      overseas: overseas
    });
  }

  snapshot.members = Object.keys(members).map(k => members[k])
    .sort((a, b) => {
      if (b.war_hits !== a.war_hits) return b.war_hits - a.war_hits;
      if (b.base_respect !== a.base_respect) return b.base_respect - a.base_respect;
      return a.name.localeCompare(b.name);
    });

  snapshot.recent_attacks.sort((a, b) => (Date.parse(b.time) || 0) - (Date.parse(a.time) || 0));
  snapshot.recent_attacks = snapshot.recent_attacks.slice(0, 50);

  snapshot.chain_bonus_hits.sort((a, b) => (Date.parse(b.time) || 0) - (Date.parse(a.time) || 0));
  snapshot.chain_bonus_hits = snapshot.chain_bonus_hits.slice(0, 250);

  return snapshot;
}
