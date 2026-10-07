function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const environment = String(props.getProperty("PAYOUT_ENVIRONMENT") || "").trim().toLowerCase();
  const isTest = environment === "test" || environment === "staging" || environment === "development";

  // ==========================================
  // CURRENT WAR WORKFLOW
  // Ordered roughly in the sequence leadership uses it.
  // ==========================================
  const bountiesMenu = ui.createMenu("🎯 Bounties")
    .addItem("Scan Current War Bounties", "triggerCurrentWarBountyScan")
    .addItem("Check Bounty Scan Status", "checkCurrentWarBountyScanStatus")
    .addItem("Sync Eligible War Bounties", "syncEligibleWarBountiesFromRailway")
    .addItem("Bounty Payout Reconciliation", "showWarBountyPayoutReconciliation");

  const reviewMenu = ui.createMenu("✅ Review & Payment Checks")
    .addItem("Run Payout Auditor", "runPayoutAudit")
    .addItem("Refresh Awards", "refreshAwards")
    .addSeparator()
    .addItem("Check Payout Reconciliation", "buildPayoutReconciliationReport");

  const archiveMenu = ui.createMenu("🗄️ Archive & Web")
    .addItem("Save Current War to Database", "archiveCurrentWarToDatabase")
    .addItem("Publish Current Payout to Web", "publishCurrentPayoutToWeb")
    .addItem("Show Public Web URL", "showPublicPayoutWebUrl")
    .addSeparator()
    .addItem("Refresh War Archive", "refreshWarArchiveIndex")
    .addItem("Load Selected Prior War", "loadSelectedArchivedWar")
    .addItem("Load Prior War by ID", "loadArchivedWarByIdPrompt")
    .addItem("Publish Selected Archived War to Web", "publishSelectedArchivedWarToWeb")
    .addSeparator()
    .addItem("Archive to Database & Reset", "archiveAndResetWarDatabase");

  const setupMenu = ui.createMenu("⚙️ Setup & Maintenance")
    .addItem("Update Faction Roster", "updateRoster")
    .addItem("Initialize / Repair Bounty Tracker", "setupBountyTracker")
    .addItem("Rebuild Dashboard UI", "buildDashboard")
    .addSeparator()
    .addItem("Configure Railway Database", "configureArchiveDatabase")
    .addItem("Test Railway + Torn Connection", "testRailwayTornConnection")
    .addSeparator()
    .addItem("Configure Bounty Bot Bridge", "configureBountyBotBridge")
    .addItem("Enable Automatic Bounty Sync", "enableAutomaticRfcBountySync")
    .addItem("Disable Automatic Bounty Sync", "disableAutomaticRfcBountySync");

  const legacyMenu = ui.createMenu("🧰 Legacy / Manual")
    .addItem("Publish to Legacy Public Google Sheet", "publishPayoutToPublic")
    .addSeparator()
    .addItem("Fetch Official Reports + RD Manually", "fetchOfficialReports")
    .addItem("Pull Raw Attack Data Only", "importWarData")
    .addSeparator()
    .addItem("Log War to Legacy History Tab", "logWarToHistory")
    .addItem("Legacy Google Sheet Archive & Reset", "archiveAndResetWar")
    .addItem("Test Legacy Archive (No Reset)", "testArchiveOnly")
    .addItem("Clean Sweep / Reset Sheet", "cleanSweep")
    .addSeparator()
    .addItem("Migrate Legacy Google Archives", "migrateLegacyGoogleArchivesToDatabase")
    .addItem("Migrate Public-Only Payout Tabs", "migrateLegacyPublicOnlyWarsToDatabase");

  if (isTest) {
    legacyMenu
      .addSeparator()
      .addItem("TEST: Clear Current Payout Snapshot", "clearCurrentArchiveSnapshotForTest");
  }

  ui.createMenu("⚔️ Faction Tools")
    .addItem("1. Refresh Current War Data", "refreshDashboard")
    .addItem("2. Select Official Reports", "showOfficialReportPicker")
    .addSubMenu(bountiesMenu)
    .addSeparator()
    .addItem("3. Generate Payout Tab", "buildPayoutTab")
    .addItem("4. Apply Payout Preset", "applySelectedPayoutPreset")
    .addItem("5. Calculate Payout Metrics", "runPayoutMath")
    .addSubMenu(reviewMenu)
    .addItem("6. Generate Final Payouts", "buildFinalPayoutTab")
    .addSeparator()
    .addSubMenu(archiveMenu)
    .addSubMenu(setupMenu)
    .addSubMenu(legacyMenu)
    .addToUi();
}
