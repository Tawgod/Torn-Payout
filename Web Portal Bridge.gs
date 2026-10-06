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
