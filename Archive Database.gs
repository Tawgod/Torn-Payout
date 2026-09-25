// ==========================================
// RAILWAY ARCHIVE DATABASE
// ==========================================
// Stores configuration in Document Properties so database credentials do not
// need to live in visible spreadsheet cells.

function configureArchiveDatabase() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const currentUrl = props.getProperty("ARCHIVE_API_URL") || "";
  const urlPrompt = ui.prompt(
    "Railway Archive API",
    "Enter the Railway API base URL (example: https://service.up.railway.app)" +
      (currentUrl ? "\n\nCurrent URL is already configured." : ""),
    ui.ButtonSet.OK_CANCEL
  );
  if (urlPrompt.getSelectedButton() !== ui.Button.OK) return;

  const apiUrl = urlPrompt.getResponseText().trim().replace(/\/+$/, "");
  if (!/^https:\/\//i.test(apiUrl)) {
    ui.alert("The archive API URL must start with https://");
    return;
  }

  const tokenPrompt = ui.prompt(
    "Archive API Token",
    "Enter the ARCHIVE_API_TOKEN configured on Railway. It will be stored in this spreadsheet's Document Properties, not in a visible cell.",
    ui.ButtonSet.OK_CANCEL
  );
  if (tokenPrompt.getSelectedButton() !== ui.Button.OK) return;
  const token = tokenPrompt.getResponseText().trim();
  if (!token) {
    ui.alert("A token is required.");
    return;
  }

  const currentFaction = props.getProperty("ARCHIVE_FACTION_KEY") || "";
  const factionPrompt = ui.prompt(
    "Archive Namespace",
    "Enter a short faction key such as ironsides or resolute." +
      (currentFaction ? "\n\nCurrent key: " + currentFaction : ""),
    ui.ButtonSet.OK_CANCEL
  );
  if (factionPrompt.getSelectedButton() !== ui.Button.OK) return;

  const factionKey = factionPrompt.getResponseText().trim().toLowerCase().replace(/[^a-z0-9_-]/g, "-");
  if (!factionKey) {
    ui.alert("A faction key is required.");
    return;
  }

  props.setProperties({
    ARCHIVE_API_URL: apiUrl,
    ARCHIVE_API_TOKEN: token,
    ARCHIVE_FACTION_KEY: factionKey
  });

  try {
    const response = UrlFetchApp.fetch(
      apiUrl + "/api/wars?faction=" + encodeURIComponent(factionKey) + "&limit=1",
      {
        muteHttpExceptions: true,
        headers: { Authorization: "Bearer " + token }
      }
    );
    const code = response.getResponseCode();
    if (code < 200 || code >= 300) {
      throw new Error("HTTP " + code + ": " + response.getContentText());
    }
    ui.alert("✅ Railway archive database connection verified for this spreadsheet.");
  } catch (e) {
    props.deleteProperty("ARCHIVE_API_TOKEN");
    ui.alert(
      "❌ Railway connection failed.\n\n" +
      "The URL/faction were saved, but the token was not kept because authentication failed.\n\n" +
      e.message
    );
  }
}

function testRailwayTornConnection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const cfg = getArchiveDatabaseConfig_();

    // Authenticated archive API test.
    const archive = archiveApiRequest_("/api/wars?faction=" + encodeURIComponent(cfg.factionKey) + "&limit=1");

    // Read-only Torn API test through Railway.
    const torn = tornApiRequest_("faction", "", "basic");
    if (torn.error) throw new Error(torn.error.error || torn.error);

    const memberCount = torn.members ? Object.keys(torn.members).length : 0;
    const factionName = torn.name || torn.faction_name || cfg.factionKey;
    const factionId = torn.ID || torn.id || torn.faction_id || "";

    ui.alert(
      "✅ Test environment connection passed",
      "Railway database API: connected\n" +
      "Torn API via Railway: connected\n" +
      "Faction: " + factionName + (factionId ? " [" + factionId + "]" : "") + "\n" +
      "Members returned: " + memberCount + "\n" +
      "Archive records currently visible in staging: " + ((archive.wars || []).length ? "1+" : "0"),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert("❌ Test connection failed\n\n" + e.message);
  }
}

function getArchiveDatabaseConfig_() {
  const props = PropertiesService.getDocumentProperties();
  const cfg = {
    url: (props.getProperty("ARCHIVE_API_URL") || "").replace(/\/+$/, ""),
    token: props.getProperty("ARCHIVE_API_TOKEN") || "",
    factionKey: props.getProperty("ARCHIVE_FACTION_KEY") || ""
  };

  if (!cfg.url || !cfg.token || !cfg.factionKey) {
    throw new Error("Archive database is not configured. Run Faction Tools → Database Archive → Configure Database first.");
  }
  return cfg;
}

function archiveApiRequest_(path, options) {
  const cfg = getArchiveDatabaseConfig_();
  const params = options || {};
  params.muteHttpExceptions = true;
  params.headers = Object.assign({}, params.headers || {}, {
    Authorization: "Bearer " + cfg.token
  });

  const response = UrlFetchApp.fetch(cfg.url + path, params);
  const code = response.getResponseCode();
  const text = response.getContentText();
  let body = {};
  try { body = text ? JSON.parse(text) : {}; } catch (e) { body = { raw: text }; }

  if (code < 200 || code >= 300) {
    throw new Error("Archive API error " + code + ": " + (body.error || body.raw || "Unknown error"));
  }
  return body;
}

function tornApiRequest_(scope, id, selections) {
  const cfg = getArchiveDatabaseConfig_();
  const params = [
    "faction=" + encodeURIComponent(cfg.factionKey),
    "scope=" + encodeURIComponent(scope),
    "selections=" + encodeURIComponent(selections)
  ];
  if (id !== null && id !== undefined && String(id).trim() !== "") {
    params.push("id=" + encodeURIComponent(String(id).trim()));
  }
  return archiveApiRequest_("/api/torn?" + params.join("&"));
}

function dashboardLabelMap_(sheet) {
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length - 1; c++) {
      const label = data[r][c];
      if (typeof label === "string" && label.trim() !== "") {
        map[label.trim().toLowerCase()] = {
          label: label.trim(),
          row: r + 1,
          col: c + 2,
          value: data[r][c + 1]
        };
      }
    }
  }
  return map;
}

function dashboardValue_(map, label, fallback) {
  const hit = map[label.toLowerCase()];
  return hit ? hit.value : fallback;
}

function setDashboardValueByLabel_(sheet, label, value) {
  const data = sheet.getDataRange().getValues();
  const wanted = label.toLowerCase().trim();
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length - 1; c++) {
      const cellText = data[r][c] === null || data[r][c] === undefined
        ? ""
        : data[r][c].toString().toLowerCase().trim();
      if (cellText === wanted) {
        sheet.getRange(r + 1, c + 2).setValue(value === null || value === undefined ? "" : value);
        return true;
      }
    }
  }
  return false;
}

function serializeSheetValue_(value) {
  if (value instanceof Date) return value.toISOString();
  return value;
}

function archiveCurrentWarToDatabase_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  const payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  if (!dashSheet || !payoutSheet) throw new Error("Dashboard and Payouts sheets are required.");

  const cfg = getArchiveDatabaseConfig_();
  const dashMap = dashboardLabelMap_(dashSheet);

  let warId = dashboardValue_(dashMap, "war id", "");
  if (!warId) warId = dashboardValue_(dashMap, "war report id", "");
  warId = parseInt(String(warId).replace(/,/g, ""), 10);
  if (!warId) throw new Error("No valid War ID was found on the Dashboard.");

  const lastCol = Math.max(1, payoutSheet.getLastColumn());
  const lastRow = Math.max(2, payoutSheet.getLastRow());

  const topRow = payoutSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(serializeSheetValue_);
  const headers = payoutSheet.getRange(2, 1, 1, lastCol).getDisplayValues()[0];
  const memberRows = lastRow >= 3
    ? payoutSheet.getRange(3, 1, lastRow - 2, lastCol).getValues()
        .filter(row => row[0] !== "" && row[0] !== null)
        .map(row => row.map(serializeSheetValue_))
    : [];

  const dashboardValues = {};
  Object.keys(dashMap).forEach(key => {
    const item = dashMap[key];
    dashboardValues[item.label] = serializeSheetValue_(item.value);
  });

  const payload = {
    faction_key: cfg.factionKey,
    war_id: warId,
    enemy_name: String(dashboardValue_(dashMap, "enemy faction name", "Unknown") || "Unknown"),
    enemy_faction_id: dashboardValue_(dashMap, "enemy faction id", null),
    archived_at: new Date().toISOString(),
    source: "apps-script",
    summary: {
      outcome: dashboardValue_(dashMap, "outcome (result)", "Unknown"),
      termed: dashboardValue_(dashMap, "termed?", "No"),
      total_hits: payoutSheet.getRange("C1").getValue() || 0,
      total_war_hits: dashboardValue_(dashMap, "total war hits", 0),
      war_score: dashboardValue_(dashMap, "war score", 0),
      total_revenue: dashboardValue_(dashMap, "total revenue", 0),
      total_payout: dashboardValue_(dashMap, "payout total", 0),
      faction_profit: dashboardValue_(dashMap, "actual faction deduction", 0),
      caches: dashboardValue_(dashMap, "caches / items won", ""),
      official_war_start: dashboardValue_(dashMap, "official war start", ""),
      official_war_end: dashboardValue_(dashMap, "official war end", ""),
      dashboard_values: dashboardValues
    },
    payout: {
      top_row: topRow,
      headers: headers,
      rows: memberRows
    }
  };

  return archiveApiRequest_("/api/wars", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  });
}

function archiveCurrentWarToDatabase() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = archiveCurrentWarToDatabase_();
    ui.alert("✅ War saved to Railway database.\n\nMember payout rows stored: " + (result.members || 0));
  } catch (e) {
    ui.alert("Archive failed:\n\n" + e.message);
  }
}

function archiveAndResetWarDatabase() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    "⚠️ ARCHIVE TO DATABASE & RESET",
    "This will save the current war to Railway/Postgres and only reset the local workbook after the database confirms the archive was saved. Proceed?",
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  try {
    const result = archiveCurrentWarToDatabase_();
    resetAfterDatabaseArchive_();
    ui.alert("✅ Archived to Railway and reset successfully.\n\nMember rows saved: " + (result.members || 0));
  } catch (e) {
    ui.alert("Nothing was reset because the database archive failed.\n\n" + e.message);
  }
}

function resetAfterDatabaseArchive_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const dataTabs = [
    SETTINGS.rdSheet || "RD",
    SETTINGS.bountySheet || "Bounties",
    "Bonus", "Adjustments", "Outside Hits"
  ];
  dataTabs.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();
    }
  });

  ["Official War Report", "Official Chain Report"].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      sheet.clear();
      sheet.clearFormats();
    }
  });

  const payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  if (payoutSheet && payoutSheet.getLastRow() >= 3) {
    payoutSheet.getRange(3, 1, payoutSheet.getLastRow() - 2, payoutSheet.getLastColumn()).uncheck();
    payoutSheet.getRange(3, 5, payoutSheet.getLastRow() - 2, Math.max(1, payoutSheet.getLastColumn() - 4)).clearContent();

    const data = payoutSheet.getDataRange().getValues();
    let ghostRow = -1;
    for (let i = 0; i < data.length; i++) {
      const name = data[i][1] ? data[i][1].toString() : "";
      if (name === "⚠️ NON-ROSTER / LEFT FACTION" || name.includes("(Left Faction)")) {
        ghostRow = i + 1;
        break;
      }
    }
    if (ghostRow > 0) payoutSheet.deleteRows(ghostRow, payoutSheet.getLastRow() - ghostRow + 1);
  }

  const finalSheet = ss.getSheetByName(SETTINGS.finalSheet || "Final Payout");
  if (finalSheet && finalSheet.getLastRow() > 1) {
    finalSheet.getRange(2, 1, finalSheet.getMaxRows() - 1, finalSheet.getMaxColumns()).clearContent();
  }

  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  if (dashSheet) {
    [
      ["Enemy Faction ID", ""],
      ["Enemy Faction Name", "Unknown"],
      ["War Report ID", ""],
      ["Chain Report ID", ""],
      ["Outcome (Result)", "Ongoing"],
      ["Termed?", "No"],
      ["Total Revenue", 0],
      ["- Temp Cost", 0],
      ["- Revives", 0],
      ["- Xanax", 0],
      ["- Other Cost", 0],
      ["Est. Cache Value", ""],
      ["Actual Cache Value", ""],
      ["Caches / Items Won", ""]
    ].forEach(pair => setDashboardValueByLabel_(dashSheet, pair[0], pair[1]));
  }

  if (typeof buildDashboard === "function") buildDashboard();
}

function refreshWarArchiveIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("War Archive");
  if (!sheet) sheet = ss.insertSheet("War Archive");

  try {
    const cfg = getArchiveDatabaseConfig_();
    const data = archiveApiRequest_("/api/wars?faction=" + encodeURIComponent(cfg.factionKey) + "&limit=500");
    const wars = data.wars || [];

    sheet.clear();
    sheet.setHiddenGridlines(true);
    sheet.setFrozenRows(2);

    sheet.getRange("A1:J1").merge()
      .setValue("🗄️ RAILWAY WAR ARCHIVE")
      .setBackground("#444444").setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center");

    const headers = ["War ID / Legacy Key", "Archived", "Enemy", "Outcome", "Termed", "Total Hits", "War Score", "Total Payout", "Faction Profit", "Record ID"];
    sheet.getRange(2, 1, 1, headers.length).setValues([headers])
      .setBackground("#274e13").setFontColor("white").setFontWeight("bold");

    if (wars.length) {
      const rows = wars.map(w => [
        w.war_id || w.legacy_key || "",
        w.archived_at ? new Date(w.archived_at) : "",
        w.enemy_name || "",
        w.outcome || "",
        w.termed === true ? "Yes" : (w.termed === false ? "No" : ""),
        Number(w.total_hits || 0),
        Number(w.war_score || 0),
        Number(w.total_payout || 0),
        Number(w.faction_profit || 0),
        Number(w.record_id || 0)
      ]);
      sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
      sheet.getRange(3, 2, rows.length, 1).setNumberFormat("yyyy-mm-dd HH:mm");
      sheet.getRange(3, 8, rows.length, 2).setNumberFormat('"$ "#,##0');
      sheet.getRange(3, 1, rows.length, headers.length).setBorder(true, true, true, true, true, true);
    }

    sheet.autoResizeColumns(1, headers.length);
    try { sheet.hideColumns(10); } catch (e) {}
    ss.setActiveSheet(sheet);
    ss.toast("Archive index refreshed: " + wars.length + " wars.", "Database", 4);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Could not refresh archive index:\n\n" + e.message);
  }
}

function loadSelectedArchivedWar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const active = ss.getActiveSheet();
  if (active.getName() !== "War Archive") {
    SpreadsheetApp.getUi().alert("Open the War Archive tab, select a war row, then run this command again.");
    return;
  }
  const row = active.getActiveRange().getRow();
  if (row < 3) {
    SpreadsheetApp.getUi().alert("Select a war row first.");
    return;
  }
  const recordId = active.getRange(row, 10).getValue();
  if (!recordId) {
    SpreadsheetApp.getUi().alert("That row does not contain a database record ID.");
    return;
  }
  loadArchivedRecordById_(recordId);
}

function loadArchivedRecordById_(recordId) {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    "Load Archived Record?",
    "This will replace the current Dashboard archive fields and Payouts data with the selected historical snapshot. Current RD/raw attack data will not be restored.",
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  try {
    const cfg = getArchiveDatabaseConfig_();
    const data = archiveApiRequest_("/api/records/" + encodeURIComponent(recordId) + "?faction=" + encodeURIComponent(cfg.factionKey));
    restoreArchivedWar_(data);
    ui.alert("✅ Archived payout snapshot restored into the workbook.");
  } catch (e) {
    ui.alert("Could not load archived record:\n\n" + e.message);
  }
}

function loadArchivedWarByIdPrompt() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt("Load Prior War", "Enter the War ID to restore from Railway:", ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  const warId = prompt.getResponseText().trim();
  if (!/^\d+$/.test(warId)) {
    ui.alert("Please enter a numeric War ID.");
    return;
  }
  loadArchivedWarById_(warId);
}

function loadArchivedWarById_(warId) {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    "Load Archived War " + warId + "?",
    "This will replace the current Dashboard archive fields and Payouts data with the stored historical snapshot. Current RD/raw attack data will not be restored.",
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  try {
    const cfg = getArchiveDatabaseConfig_();
    const data = archiveApiRequest_("/api/wars/" + encodeURIComponent(warId) + "?faction=" + encodeURIComponent(cfg.factionKey));
    restoreArchivedWar_(data);
    ui.alert("✅ Archived war " + warId + " restored into the workbook.");
  } catch (e) {
    ui.alert("Could not load archived war:\n\n" + e.message);
  }
}

function restoreArchivedWar_(archive) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  let payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);

  if (!payoutSheet) {
    if (typeof buildPayoutTab === "function") buildPayoutTab();
    payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  }
  if (!dashSheet || !payoutSheet) throw new Error("Dashboard or Payouts sheet is missing.");

  const payout = archive.payout || {};
  const topRow = Array.isArray(payout.top_row) ? payout.top_row : [];
  const headers = Array.isArray(payout.headers) ? payout.headers : [];
  const rows = Array.isArray(payout.rows) ? payout.rows : [];
  const colCount = Math.max(topRow.length, headers.length, rows.length ? rows[0].length : 0, 1);

  if (payoutSheet.getMaxColumns() < colCount) {
    payoutSheet.insertColumnsAfter(payoutSheet.getMaxColumns(), colCount - payoutSheet.getMaxColumns());
  }
  if (payoutSheet.getMaxRows() < rows.length + 2) {
    payoutSheet.insertRowsAfter(payoutSheet.getMaxRows(), rows.length + 2 - payoutSheet.getMaxRows());
  }

  payoutSheet.getRange(1, 1, payoutSheet.getMaxRows(), payoutSheet.getMaxColumns()).clearContent();

  if (topRow.length) {
    const padded = topRow.concat(new Array(Math.max(0, colCount - topRow.length)).fill(""));
    payoutSheet.getRange(1, 1, 1, colCount).setValues([padded]);
  }
  if (headers.length) {
    const padded = headers.concat(new Array(Math.max(0, colCount - headers.length)).fill(""));
    payoutSheet.getRange(2, 1, 1, colCount).setValues([padded]);
  }
  if (rows.length) {
    const paddedRows = rows.map(row => row.concat(new Array(Math.max(0, colCount - row.length)).fill("")));
    payoutSheet.getRange(3, 1, paddedRows.length, colCount).setValues(paddedRows);
  }

  payoutSheet.setFrozenRows(2);
  payoutSheet.setFrozenColumns(2);
  if (colCount >= 4 && rows.length) {
    payoutSheet.getRange(3, 3, rows.length, 1).setNumberFormat("0.00%");
    payoutSheet.getRange(3, 4, rows.length, 1).setNumberFormat('"$ "#,##0');
  }
  try { payoutSheet.hideColumns(1); } catch (e) {}
  if (colCount >= 20) {
    try { payoutSheet.hideColumns(20); } catch (e) {}
  }

  const summary = archive.summary || {};
  const directValues = [
    ["Enemy Faction Name", archive.enemy_name || "Unknown"],
    ["Enemy Faction ID", archive.enemy_faction_id || ""],
    ["War ID", archive.war_id || ""],
    ["War Report ID", archive.war_id || ""],
    ["Outcome (Result)", summary.outcome || "Unknown"],
    ["Termed?", summary.termed === true ? "Yes" : (summary.termed === false ? "No" : summary.termed || "")],
    ["Total War Hits", summary.total_war_hits || 0],
    ["War Score", summary.war_score || 0],
    ["Total Revenue", summary.total_revenue || 0],
    ["Caches / Items Won", summary.caches || ""],
    ["Official War Start", summary.official_war_start || ""],
    ["Official War End", summary.official_war_end || ""]
  ];
  directValues.forEach(pair => setDashboardValueByLabel_(dashSheet, pair[0], pair[1]));

  SpreadsheetApp.flush();
}

function extractSpreadsheetId_(raw) {
  const text = String(raw || "").trim();
  const match = text.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : text;
}

function migrateLegacyGoogleArchivesToDatabase() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    "Migrate Legacy Google Archives",
    "This imports the existing War Archive workbook into Railway. If the Public Payout workbook is configured, its saved weight row is merged into each matching war. Existing War IDs are updated rather than duplicated. Proceed?",
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  if (!configSheet) {
    ui.alert("Config sheet not found.");
    return;
  }

  try {
    const cfg = getArchiveDatabaseConfig_();
    const archiveId = extractSpreadsheetId_(configSheet.getRange("B3").getValue());
    const publicId = extractSpreadsheetId_(configSheet.getRange("B8").getValue());
    if (!archiveId) throw new Error("Legacy archive spreadsheet ID is missing from Config!B3.");

    const archiveSS = SpreadsheetApp.openById(archiveId);
    let publicSS = null;
    if (publicId) {
      try { publicSS = SpreadsheetApp.openById(publicId); } catch (e) {}
    }

    const publicTopRows = {};
    if (publicSS) {
      publicSS.getSheets().forEach(sheet => {
        const name = sheet.getName();
        if (name.toLowerCase().includes("summary")) return;
        const cols = Math.max(1, sheet.getLastColumn());
        publicTopRows[name.toLowerCase().trim()] = sheet.getRange(1, 1, Math.min(2, sheet.getMaxRows()), cols).getValues();
      });
    }

    let migrated = 0;
    let failed = [];

    archiveSS.getSheets().forEach(sheet => {
      const sheetName = sheet.getName().trim();
      if (!/^\d+$/.test(sheetName)) return;

      try {
        const data = sheet.getDataRange().getValues();
        if (data.length < 8) return;

        const title = String(data[0][0] || "");
        const titleMatch = title.match(/^WAR REPORT:\s*(.*?)\s*\[(\d+)\]/i);
        const enemyName = titleMatch ? titleMatch[1].trim() : "Unknown";
        const warId = parseInt(sheetName, 10);

        const findPair = (rowIndex, label) => {
          const row = data[rowIndex] || [];
          for (let c = 0; c < row.length - 1; c++) {
            if (String(row[c] || "").trim().toLowerCase() === label.toLowerCase()) return row[c + 1];
          }
          return "";
        };

        const archivedRaw = findPair(1, "Date Archived:");
        let archivedDate;
        if (archivedRaw instanceof Date) {
          archivedDate = archivedRaw;
        } else if (typeof archivedRaw === "number") {
          archivedDate = new Date(Date.UTC(1899, 11, 30) + archivedRaw * 86400000);
        } else {
          archivedDate = new Date(archivedRaw);
        }
        const archivedAt = isNaN(archivedDate.getTime()) ? null : archivedDate.toISOString();

        const archiveHeaders = (data[6] || []).slice(2);
        while (archiveHeaders.length && archiveHeaders[archiveHeaders.length - 1] === "") archiveHeaders.pop();

        const rows = [];
        for (let r = 7; r < data.length; r++) {
          if (data[r][2] === "" || data[r][2] === null) continue;
          rows.push(data[r].slice(2, 2 + archiveHeaders.length).map(serializeSheetValue_));
        }

        let topRow = new Array(archiveHeaders.length).fill("");
        const publicRows = publicTopRows[enemyName.toLowerCase()];
        if (publicRows && publicRows[0]) {
          topRow = publicRows[0].slice(0, archiveHeaders.length).map(serializeSheetValue_);
        } else {
          if (topRow.length > 1) topRow[1] = "Total Hits";
          if (topRow.length > 2) topRow[2] = findPair(2, "Total Attacks Logged:");
          if (topRow.length > 3) topRow[3] = "Set Point Weight ➡️";
        }

        const payload = {
          faction_key: cfg.factionKey,
          war_id: warId,
          enemy_name: enemyName,
          archived_at: archivedAt,
          source: "legacy-google-archive",
          summary: {
            outcome: findPair(1, "Outcome:"),
            termed: findPair(1, "Termed:"),
            total_hits: findPair(2, "Total Attacks Logged:"),
            total_war_hits: findPair(2, "Total War Hits:"),
            war_score: findPair(2, "War Score:"),
            total_revenue: findPair(3, "Total Revenue:"),
            total_payout: findPair(3, "Total Payouts:"),
            faction_profit: findPair(3, "Faction Profit:"),
            caches: findPair(4, "Caches Won:")
          },
          payout: {
            top_row: topRow,
            headers: archiveHeaders,
            rows: rows
          }
        };

        const saveResult = archiveApiRequest_("/api/wars", {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(payload)
        });

        if (publicRows && saveResult && saveResult.id) {
          archiveApiRequest_("/api/records/" + encodeURIComponent(saveResult.id) + "/publish", {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify({
              faction_key: cfg.factionKey,
              public_title: enemyName
            })
          });
        }

        migrated++;
      } catch (e) {
        failed.push(sheetName + ": " + e.message);
      }
    });

    refreshWarArchiveIndex();

    let message = "✅ Migration complete.\n\nWars migrated/updated: " + migrated;
    if (failed.length) message += "\n\nFailures:\n" + failed.slice(0, 10).join("\n");
    ui.alert(message);
  } catch (e) {
    ui.alert("Migration failed:\n\n" + e.message);
  }
}
