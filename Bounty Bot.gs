function configureBountyBotBridge() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const currentUrl = props.getProperty("BOUNTY_BOT_URL") || "";
  const urlPrompt = ui.prompt(
    "Bounty Bot Bridge",
    "Enter the faction-tools web base URL." + (currentUrl ? "\n\nA URL is already configured." : ""),
    ui.ButtonSet.OK_CANCEL
  );
  if (urlPrompt.getSelectedButton() !== ui.Button.OK) return;
  const url = urlPrompt.getResponseText().trim().replace(/\/+$/, "");
  if (!/^https:\/\//i.test(url)) {
    ui.alert("The bounty bot URL must start with https://");
    return;
  }

  const tokenPrompt = ui.prompt(
    "Bounty Bot Token",
    "Enter the BOUNTY_TRIGGER_SECRET. It is stored in Document Properties, not a visible sheet cell.",
    ui.ButtonSet.OK_CANCEL
  );
  if (tokenPrompt.getSelectedButton() !== ui.Button.OK) return;
  const token = tokenPrompt.getResponseText().trim();
  if (!token) {
    ui.alert("A token is required.");
    return;
  }

  props.setProperties({
    BOUNTY_BOT_URL: url,
    BOUNTY_BOT_TOKEN: token,
    BOUNTY_LOOKBACK_MINUTES: props.getProperty("BOUNTY_LOOKBACK_MINUTES") || "30"
  });
  ui.alert("✅ Bounty bot bridge configured. Default pre-war lookback: 30 minutes.");
}

function getBountyBotConfig_() {
  const props = PropertiesService.getDocumentProperties();
  const cfg = {
    url: (props.getProperty("BOUNTY_BOT_URL") || "").replace(/\/+$/, ""),
    token: props.getProperty("BOUNTY_BOT_TOKEN") || "",
    lookbackMinutes: parseInt(props.getProperty("BOUNTY_LOOKBACK_MINUTES") || "30", 10)
  };
  if (!cfg.url || !cfg.token) {
    throw new Error("Bounty bot bridge is not configured. Run Faction Tools → Bounties → Configure Bounty Bot Bridge.");
  }
  if (!Number.isFinite(cfg.lookbackMinutes)) cfg.lookbackMinutes = 30;
  cfg.lookbackMinutes = Math.max(0, Math.min(180, cfg.lookbackMinutes));
  return cfg;
}

function parseDashboardUtc_(value, allowBlank) {
  if (value === null || value === undefined || value === "") {
    if (allowBlank) return null;
    throw new Error("Missing war date/time on Dashboard.");
  }
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  let raw = String(value).trim();
  if (!raw || raw.toLowerCase() === "ongoing" || raw.toLowerCase() === "n/a") {
    if (allowBlank) return null;
    throw new Error("Missing war date/time on Dashboard.");
  }

  let d = new Date(raw);
  if (isNaN(d.getTime())) d = new Date(raw + " GMT");
  if (isNaN(d.getTime())) throw new Error("Could not parse Dashboard date/time: " + raw);
  return d;
}

function triggerCurrentWarBountyScan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SETTINGS.dashboardSheet);
  const config = ss.getSheetByName(SETTINGS.configSheet);
  if (!dash || !config) throw new Error("Dashboard and Config tabs are required.");

  const cfg = getBountyBotConfig_();
  const warIdRaw = labelValue_(dash, "War ID", "") || labelValue_(dash, "War Report ID", "");
  const warId = parseInt(String(warIdRaw || "").replace(/,/g, ""), 10);
  if (!warId) throw new Error("A valid War ID is required before scanning bounties.");

  const factionId = parseInt(String(config.getRange(SETTINGS.factionIdCell || "B7").getValue() || "").replace(/,/g, ""), 10);
  if (!factionId) throw new Error("Faction ID is missing from Config.");

  const factionName = String(labelValue_(dash, "Our Faction Name", "") || "").trim();
  const warStart = parseDashboardUtc_(labelValue_(dash, "Official War Start", ""), false);
  const warEnd = parseDashboardUtc_(labelValue_(dash, "Official War End", ""), true);

  const payload = {
    faction_id: factionId,
    faction_name: factionName,
    war_id: warId,
    war_start: warStart.toISOString(),
    war_end: warEnd ? warEnd.toISOString() : null,
    lookback_minutes: cfg.lookbackMinutes
  };

  const response = UrlFetchApp.fetch(cfg.url + "/internal/bounty-scan", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + cfg.token }
  });

  const status = response.getResponseCode();
  let body = {};
  try { body = JSON.parse(response.getContentText() || "{}"); }
  catch (e) { body = { error: response.getContentText() }; }

  if (status < 200 || status >= 300) {
    throw new Error("Bounty scan trigger failed (" + status + "): " + (body.error || "Unknown error"));
  }

  ss.toast(
    "Bounty scan queued for War " + warId + " using a " + cfg.lookbackMinutes +
      "-minute pre-war lookback" + (body.deduplicated ? " (existing scan reused)." : "."),
    "Bounties",
    6
  );
}

function checkCurrentWarBountyScanStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SETTINGS.dashboardSheet);
  const config = ss.getSheetByName(SETTINGS.configSheet);
  const cfg = getBountyBotConfig_();

  const warIdRaw = labelValue_(dash, "War ID", "") || labelValue_(dash, "War Report ID", "");
  const warId = parseInt(String(warIdRaw || "").replace(/,/g, ""), 10);
  const factionId = parseInt(String(config.getRange(SETTINGS.factionIdCell || "B7").getValue() || "").replace(/,/g, ""), 10);
  if (!warId || !factionId) throw new Error("War ID and Faction ID are required.");

  const response = UrlFetchApp.fetch(
    cfg.url + "/internal/bounty-scan/status?faction_id=" + encodeURIComponent(factionId) +
      "&war_id=" + encodeURIComponent(warId),
    {
      muteHttpExceptions: true,
      headers: { Authorization: "Bearer " + cfg.token }
    }
  );

  const code = response.getResponseCode();
  let body = {};
  try { body = JSON.parse(response.getContentText() || "{}"); }
  catch (e) { body = {}; }

  if (code === 404) {
    SpreadsheetApp.getUi().alert("No bounty scan has been requested for this war yet.");
    return;
  }
  if (code < 200 || code >= 300) {
    throw new Error("Could not read bounty scan status (" + code + ").");
  }

  SpreadsheetApp.getUi().alert(
    "Bounty Scan — War " + warId,
    "Status: " + (body.status || "unknown") + "\n" +
    "Lookback: " + (body.lookback_minutes || cfg.lookbackMinutes) + " minutes before war start\n" +
    "News rows scanned: " + (body.scanned_news_rows || 0) + "\n" +
    "Bounties recognized: " + (body.recognized_bounties || 0) + "\n" +
    "New bounties added: " + (body.inserted_bounties || 0) +
    (body.error_text ? "\n\nError: " + body.error_text : ""),
    ui.ButtonSet.OK
  );
}
