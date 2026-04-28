function runPayoutAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  const finalSheet = ss.getSheetByName(SETTINGS.finalSheet);

  if (!finalSheet) {
    ss.toast("No Final Payout tab found to audit.", "Error", 5);
    return;
  }

  const apiKey = configSheet.getRange(SETTINGS.apiKeyCell).getValue().toString().trim();
  if (!apiKey) return;

  ss.toast("Scanning Faction Bank Logs...", "System", 3);

  // 1. Fetch Faction "Funds News" (The Bank Logs)
  const url = `https://api.torn.com/faction/?selections=fundsnews&key=${apiKey}`;
  let response;
  try {
    response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch(e) {
    ss.toast("API Connection Failed.", "Error", 5);
    return;
  }

  const json = JSON.parse(response.getContentText());
  if (json.error || !json.fundsnews) {
    ss.toast("Could not read bank logs. Ensure your API Key has bank viewing permissions.", "Error", 5);
    return;
  }

  const newsLog = json.fundsnews;
  const recentTransfers = [];

  // 2. Parse the HTML strings in the logs into a clean Javascript array
  // We are looking for strings like: "Name sent $X,XXX of faction funds to TargetName."
  const transferRegex = /sent\s\$([0-9,]+)\sof\sfaction\sfunds\sto\s<a\shref\s=\s"http:\/\/www\.torn\.com\/profiles\.php\?XID=(\d+)"/i;

  for (let id in newsLog) {
    let entry = newsLog[id].news;
    let match = entry.match(transferRegex);
    if (match) {
      recentTransfers.push({
        amount: parseInt(match[1].replace(/,/g, '')), // The money sent
        targetId: match[2]                            // The ID of the member who got it
      });
    }
  }

  // 3. Compare the logs against the checkboxes on the Final Payout sheet
  const lastRow = finalSheet.getLastRow();
  if (lastRow < 2) return;

  // Grab IDs, Total Expected Payouts, and Checkbox Status
  const ledgerData = finalSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  let checksMade = 0;

  for (let i = 0; i < ledgerData.length; i++) {
    let memberId = ledgerData[i][0].toString().trim();
    let expectedPayout = parseInt(ledgerData[i][5]) || 0; // Col F (Index 5)
    let isChecked = ledgerData[i][7];                     // Col H (Index 7)

    // Skip if they are already checked off or are owed nothing
    if (isChecked === true || expectedPayout <= 0) continue;

    // Search the Torn logs to see if we successfully sent them that exact amount recently
    let matchedTransfer = recentTransfers.find(t => t.targetId === memberId && t.amount === expectedPayout);

    if (matchedTransfer) {
      // It matches! Check the box automatically.
      finalSheet.getRange(i + 2, 8).setValue(true);
      checksMade++;
    }
  }

  if (checksMade > 0) {
    ss.toast(`Audit Complete! ${checksMade} payouts automatically verified.`, "Success", 5);
  } else {
    ss.toast("Audit Complete. No new matching transfers found in logs.", "System", 3);
  }
}