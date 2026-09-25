// ==========================================
// RAILWAY PUBLIC PAYOUT SITE
// ==========================================

function publishCurrentPayoutToWeb() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    "🌐 PUBLISH PAYOUT TO WEB",
    "This will save the current payout snapshot to Railway and make that war visible on the public payout website. The existing Google public payout sheet will NOT be changed. Proceed?",
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  try {
    const cfg = getArchiveDatabaseConfig_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
    if (!dashSheet) throw new Error("Dashboard sheet not found.");

    const map = dashboardLabelMap_(dashSheet);
    let warId = dashboardValue_(map, "war id", "");
    if (!warId) warId = dashboardValue_(map, "war report id", "");
    warId = parseInt(String(warId).replace(/,/g, ""), 10);
    if (!warId) throw new Error("No valid War ID was found on the Dashboard.");

    const enemyName = String(dashboardValue_(map, "enemy faction name", "Unknown") || "Unknown");

    const saveResult = archiveCurrentWarToDatabase_();

    archiveApiRequest_("/api/wars/" + encodeURIComponent(warId) + "/publish", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        faction_key: cfg.factionKey,
        public_title: enemyName
      })
    });

    const publicUrl = cfg.url + "/public/" + encodeURIComponent(cfg.factionKey);
    ui.alert(
      "✅ Published to Railway web payout history.\n\n" +
      "War ID: " + warId + "\n" +
      "Member rows saved: " + (saveResult.members || 0) + "\n\n" +
      "Public history:\n" + publicUrl
    );
  } catch (e) {
    ui.alert("Web publish failed:\n\n" + e.message);
  }
}

function showPublicPayoutWebUrl() {
  try {
    const cfg = getArchiveDatabaseConfig_();
    SpreadsheetApp.getUi().alert(
      "Public Payout History\n\n" +
      cfg.url + "/public/" + encodeURIComponent(cfg.factionKey)
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(e.message);
  }
}

function publishSelectedArchivedWarToWeb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== "War Archive") {
    ui.alert("Open the War Archive tab and select the war you want to publish.");
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row < 3) {
    ui.alert("Select a war row first.");
    return;
  }

  const recordId = sheet.getRange(row, 10).getValue();
  const label = sheet.getRange(row, 1).getDisplayValue();
  const enemy = sheet.getRange(row, 3).getDisplayValue();

  if (!recordId) {
    ui.alert("That row does not contain a database record ID.");
    return;
  }

  const answer = ui.alert(
    "Publish Archived Payout?",
    "Publish " + enemy + " (" + label + ") to the Railway public payout site?\n\nThe backup Google sheets will not be changed.",
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  try {
    const cfg = getArchiveDatabaseConfig_();
    archiveApiRequest_("/api/records/" + encodeURIComponent(recordId) + "/publish", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        faction_key: cfg.factionKey,
        public_title: enemy
      })
    });

    ui.alert(
      "✅ Archived payout published.\n\n" +
      cfg.url + "/public/" + encodeURIComponent(cfg.factionKey)
    );
  } catch (e) {
    ui.alert("Publish failed:\n\n" + e.message);
  }
}
