function setupBountyTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS.bountySheet);
  const wantedHeaders = SETTINGS.bountyHeaders || [
    "Date Logged", "Placed By", "Target", "Quantity", "Bounty Amount",
    "Refund Amount", "Status", "Notes", "Source", "Torn News ID",
    "Refundable War Bounty", "War ID", "War Verification", "Discord Message ID",
    "Added To Payout", "Payout Reference"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS.bountySheet, 4);
    sheet.setHiddenGridlines(true);
    sheet.getRange(1, 1, 1, wantedHeaders.length).setValues([wantedHeaders]);
  } else {
    upgradeBountyTrackerSchema_(sheet, wantedHeaders);
  }

  const headerMap = headerMapFromRow_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  const headerRange = sheet.getRange(1, 1, 1, wantedHeaders.length);
  headerRange.setBackground("#4a86e8")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);

  const dateCol = headerIndex_(headerMap, "Date Logged", true) + 1;
  const qtyCol = headerIndex_(headerMap, "Quantity", true) + 1;
  const amountCol = headerIndex_(headerMap, "Bounty Amount", true) + 1;
  const refundCol = headerIndex_(headerMap, "Refund Amount", true) + 1;
  const statusCol = headerIndex_(headerMap, "Status", true) + 1;

  sheet.getRange(2, dateCol, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setNumberFormat("m/d/yyyy h:mm am/pm");
  sheet.getRange(2, qtyCol, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setNumberFormat("0");
  sheet.getRange(2, amountCol, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setNumberFormat('"$ "#,##0');
  sheet.getRange(2, refundCol, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setNumberFormat('"$ "#,##0');

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pending Review", "Approved", "Denied / Invalid", "Paid Out"])
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, statusCol, Math.max(1, sheet.getMaxRows() - 1), 1).setDataValidation(statusRule);

  // Quantity defaults to one for existing/manual rows when other bounty fields are populated.
  const lastRow = Math.max(2, sheet.getLastRow());
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, wantedHeaders.length).getValues();
    const placerIdx = headerIndex_(headerMap, "Placed By", true);
    const targetIdx = headerIndex_(headerMap, "Target", true);
    const qtyIdx = headerIndex_(headerMap, "Quantity", true);
    for (let i = 0; i < data.length; i++) {
      if ((data[i][placerIdx] || data[i][targetIdx]) && !data[i][qtyIdx]) {
        sheet.getRange(i + 2, qtyIdx + 1).setValue(1);
      }
    }
  }

  const widths = {
    "Date Logged": 150, "Placed By": 160, "Target": 160, "Quantity": 85,
    "Bounty Amount": 130, "Refund Amount": 130, "Status": 140,
    "Notes": 300, "Source": 135, "Torn News ID": 170,
    "Refundable War Bounty": 140, "War ID": 110, "War Verification": 150,
    "Discord Message ID": 170, "Added To Payout": 125, "Payout Reference": 150
  };
  wantedHeaders.forEach((name, i) => sheet.setColumnWidth(i + 1, widths[name] || 120));

  for (let r = 2; r <= Math.min(100, sheet.getMaxRows()); r++) {
    sheet.getRange(r, 1, 1, wantedHeaders.length)
      .setBackground(r % 2 === 0 ? "#f8f9fa" : "#ffffff");
  }

  ss.toast("Bounty Tracker ready: quantity + bot/news source fields enabled.", "Success", 5);
}

function upgradeBountyTrackerSchema_(sheet, wantedHeaders) {
  const existingWidth = Math.max(1, sheet.getLastColumn());
  const existingHeaders = sheet.getRange(1, 1, 1, existingWidth).getDisplayValues()[0]
    .map(v => String(v || "").trim());

  if (!existingHeaders.some(Boolean)) {
    sheet.getRange(1, 1, 1, wantedHeaders.length).setValues([wantedHeaders]);
    return;
  }

  const existingMap = {};
  existingHeaders.forEach((h, i) => { if (h) existingMap[normalizeHeader_(h)] = i; });
  const lastRow = Math.max(1, sheet.getLastRow());
  const oldData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, existingWidth).getValues() : [];

  const newRows = oldData.map(row => wantedHeaders.map(header => {
    const oldIdx = existingMap[normalizeHeader_(header)];
    if (oldIdx !== undefined) return row[oldIdx];
    if (header === "Quantity" && (row[existingMap["placed by"]] || row[existingMap["target"]])) return 1;
    if (header === "Source" && (row[existingMap["placed by"]] || row[existingMap["target"]])) return "Sheet";
    return "";
  }));

  sheet.clearContents();
  sheet.getRange(1, 1, 1, wantedHeaders.length).setValues([wantedHeaders]);
  if (newRows.length) sheet.getRange(2, 1, newRows.length, wantedHeaders.length).setValues(newRows);
}


function bountySheetReference_(rowNumber, referenceCol) {
  const letters = [];
  let n = referenceCol;
  while (n > 0) {
    n--;
    letters.unshift(String.fromCharCode(65 + (n % 26)));
    n = Math.floor(n / 26);
  }
  return SETTINGS.bountySheet + "!" + letters.join("") + rowNumber;
}

function acknowledgeWarBountyExport_(cfg, item, sheetReference) {
  return archiveApiRequest_("/api/war-bounty-payouts/" + encodeURIComponent(item.export_id) + "/ack", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      faction_key: cfg.factionKey,
      sheet_reference: sheetReference
    })
  });
}

function recordWarBountyExportFailure_(cfg, item, errorText) {
  try {
    archiveApiRequest_("/api/war-bounty-payouts/" + encodeURIComponent(item.export_id) + "/fail", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        faction_key: cfg.factionKey,
        error: String(errorText || "Unknown Apps Script export error").slice(0, 1800)
      })
    });
  } catch (_ignored) {}
}

function syncEligibleWarBountiesFromRailway() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return { ok: false, busy: true };

  try {
    setupBountyTracker();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SETTINGS.bountySheet);
    const cfg = getArchiveDatabaseConfig_();
    if (!sheet) throw new Error("Bounties sheet is unavailable.");

    const data = archiveApiRequest_(
      "/api/war-bounty-payouts?faction=" + encodeURIComponent(cfg.factionKey) +
      "&status=pending&limit=100"
    );
    const items = data.exports || [];
    if (!items.length) {
      ss.toast("No verified RFC War Bounties are waiting for payout.", "Bounties", 4);
      return { ok: true, imported: 0, recovered: 0, failed: 0 };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const map = headerMapFromRow_(headers);
    const required = [
      "Date Logged", "Placed By", "Target", "Quantity", "Bounty Amount",
      "Refund Amount", "Status", "Notes", "Source", "Torn News ID",
      "Refundable War Bounty", "War ID", "War Verification",
      "Discord Message ID", "Added To Payout", "Payout Reference"
    ];
    required.forEach(name => headerIndex_(map, name, true));

    const refCol = headerIndex_(map, "Payout Reference", true) + 1;
    const existingByRef = {};
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const refs = sheet.getRange(2, refCol, lastRow - 1, 1).getDisplayValues();
      refs.forEach((row, idx) => {
        const ref = String(row[0] || "").trim();
        if (ref) existingByRef[ref] = idx + 2;
      });
    }

    let imported = 0;
    let recovered = 0;
    let failed = 0;

    items.forEach(item => {
      const ref = String(item.payout_reference || "").trim();
      if (!ref) {
        failed++;
        recordWarBountyExportFailure_(cfg, item, "Missing payout reference from RFC");
        return;
      }

      try {
        if (existingByRef[ref]) {
          acknowledgeWarBountyExport_(
            cfg,
            item,
            bountySheetReference_(existingByRef[ref], refCol)
          );
          recovered++;
          return;
        }

        const rowValues = new Array(headers.length).fill("");
        const put = (name, value) => {
          const idx = headerIndex_(map, name, true);
          rowValues[idx] = value === null || value === undefined ? "" : value;
        };

        const eligibleAt = item.payout_eligible_at ? new Date(item.payout_eligible_at) : new Date();
        put("Date Logged", isNaN(eligibleAt.getTime()) ? new Date() : eligibleAt);
        put("Placed By", item.placer_name || String(item.placer_torn_id || ""));
        put(
          "Target",
          (item.target_name || "Target") +
          (item.target_torn_id ? " [" + item.target_torn_id + "]" : "")
        );
        put("Quantity", Number(item.quantity || 1));
        put("Bounty Amount", Number(item.bounty_amount || 0));
        put("Refund Amount", Number(item.refund_amount || 0));
        put("Status", "Approved");
        put(
          "Notes",
          "RFC verified reimbursement. Actual placer ID " + item.placer_torn_id +
          "; placement " + item.placement_id + "; request " + item.request_id + "."
        );
        put("Source", "RFC War Bounty");
        put("Torn News ID", item.torn_bounty_ref || "");
        put("Refundable War Bounty", true);
        put("War ID", item.war_id || "");
        put("War Verification", item.war_verification || "verified");
        put("Discord Message ID", item.discord_message_id || "");
        put("Added To Payout", true);
        put("Payout Reference", ref);

        const rowNumber = sheet.getLastRow() + 1;
        sheet.getRange(rowNumber, 1, 1, rowValues.length).setValues([rowValues]);
        sheet.getRange(rowNumber, headerIndex_(map, "Bounty Amount", true) + 1).setNumberFormat('"$ "#,##0');
        sheet.getRange(rowNumber, headerIndex_(map, "Refund Amount", true) + 1).setNumberFormat('"$ "#,##0');
        sheet.getRange(rowNumber, headerIndex_(map, "Date Logged", true) + 1).setNumberFormat("m/d/yyyy h:mm am/pm");

        existingByRef[ref] = rowNumber;
        acknowledgeWarBountyExport_(
          cfg,
          item,
          bountySheetReference_(rowNumber, refCol)
        );
        imported++;
      } catch (e) {
        failed++;
        recordWarBountyExportFailure_(cfg, item, e && e.message ? e.message : e);
      }
    });

    SpreadsheetApp.flush();
    ss.toast(
      "RFC War Bounties: " + imported + " added, " + recovered +
      " recovered, " + failed + " failed.",
      "Bounties",
      6
    );
    return { ok: failed === 0, imported: imported, recovered: recovered, failed: failed };
  } finally {
    lock.releaseLock();
  }
}

function showWarBountyPayoutReconciliation() {
  const cfg = getArchiveDatabaseConfig_();
  const data = archiveApiRequest_(
    "/api/war-bounty-payouts-reconciliation?faction=" + encodeURIComponent(cfg.factionKey)
  );
  const s = data.summary || {};
  const missing = data.eligible_not_queued || [];
  const lines = [
    "Pending export: " + Number(s.pending || 0),
    "Exported / added to payout: " + Number(s.exported || 0),
    "Failed: " + Number(s.failed || 0),
    "Blocked: " + Number(s.blocked || 0),
    "Eligible but not queued: " + missing.length
  ];
  if (missing.length) {
    lines.push("", "Eligible but not queued:");
    missing.slice(0, 10).forEach(row => {
      lines.push(
        "• " + (row.placer_name || row.placer_torn_id) +
        " → " + (row.target_name || row.target_torn_id) +
        " ($" + Number(row.total_cost || 0).toLocaleString() + ")"
      );
    });
  }
  SpreadsheetApp.getUi().alert(
    "RFC War Bounty Payout Reconciliation",
    lines.join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function enableAutomaticRfcBountySync() {
  disableAutomaticRfcBountySync_(false);
  ScriptApp.newTrigger("syncEligibleWarBountiesFromRailway")
    .timeBased()
    .everyMinutes(5)
    .create();
  SpreadsheetApp.getUi().alert("Automatic RFC War Bounty sync enabled every 5 minutes.");
}

function disableAutomaticRfcBountySync() {
  const removed = disableAutomaticRfcBountySync_(true);
  SpreadsheetApp.getUi().alert(
    removed
      ? "Automatic RFC War Bounty sync disabled."
      : "No automatic RFC War Bounty sync trigger was installed."
  );
}

function disableAutomaticRfcBountySync_(returnCount) {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "syncEligibleWarBountiesFromRailway") {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  return returnCount ? removed : removed;
}
