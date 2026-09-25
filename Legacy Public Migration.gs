// ==========================================
// LEGACY PUBLIC PAYOUT MIGRATION
// ==========================================
// Imports payout-only tabs that do not have a matching numeric war sheet in the
// legacy archive workbook. These are retained with a legacy key and can still
// be restored from the War Archive index.

function migrateLegacyPublicOnlyWarsToDatabase() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    "Migrate Public-Only Payout Tabs",
    "This imports payout tabs that are present in the legacy Public Payout workbook but do not already exist in Railway by enemy name. These records may not have a numeric War ID. Proceed?",
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
    const publicId = extractSpreadsheetId_(configSheet.getRange("B8").getValue());
    if (!publicId) throw new Error("Legacy public payout spreadsheet ID is missing from Config!B8.");

    const publicSS = SpreadsheetApp.openById(publicId);
    const existing = archiveApiRequest_("/api/wars?faction=" + encodeURIComponent(cfg.factionKey) + "&limit=500");
    const existingEnemies = {};
    (existing.wars || []).forEach(w => {
      if (w.enemy_name) existingEnemies[String(w.enemy_name).toLowerCase().trim()] = true;
    });

    let imported = 0;
    let skipped = 0;
    let failed = [];

    publicSS.getSheets().forEach(sheet => {
      const name = sheet.getName().trim();
      if (!name || name.toLowerCase().includes("summary")) return;

      const enemyKey = name.toLowerCase();
      if (existingEnemies[enemyKey]) {
        skipped++;
        return;
      }

      try {
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        if (lastRow < 2 || lastCol < 2) return;

        const topRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(serializeSheetValue_);
        const headers = sheet.getRange(2, 1, 1, lastCol).getDisplayValues()[0];

        const memberRows = lastRow >= 3
          ? sheet.getRange(3, 1, lastRow - 2, lastCol).getValues()
              .filter(row => row[0] !== "" && row[0] !== null)
              .map(row => row.map(serializeSheetValue_))
          : [];

        let totalPayout = 0;
        for (let i = 0; i < memberRows.length; i++) {
          const payout = Number(memberRows[i][3] || 0);
          if (isFinite(payout)) totalPayout += payout;
        }

        const payload = {
          faction_key: cfg.factionKey,
          legacy_key: "public:" + name,
          enemy_name: name,
          archived_at: null,
          source: "legacy-public-payout",
          summary: {
            outcome: "",
            termed: null,
            total_hits: Number(topRow[2] || 0),
            total_war_hits: null,
            war_score: null,
            total_revenue: null,
            total_payout: totalPayout,
            faction_profit: null,
            caches: ""
          },
          payout: {
            top_row: topRow,
            headers: headers,
            rows: memberRows
          }
        };

        archiveApiRequest_("/api/wars", {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(payload)
        });

        imported++;
        existingEnemies[enemyKey] = true;
      } catch (e) {
        failed.push(name + ": " + e.message);
      }
    });

    refreshWarArchiveIndex();

    let message =
      "✅ Public payout migration complete.\n\n" +
      "Public-only records imported: " + imported + "\n" +
      "Already represented by archive records: " + skipped;

    if (failed.length) message += "\n\nFailures:\n" + failed.slice(0, 10).join("\n");
    ui.alert(message);
  } catch (e) {
    ui.alert("Public payout migration failed:\n\n" + e.message);
  }
}
