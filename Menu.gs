function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚔️ Faction Tools')
    .addItem('1. 🔄 Sync Dashboard Data', 'syncDashboard') // <-- NEW: Safe daily refresh
    .addItem('2. Update Faction Roster', 'updateRoster')
    .addItem('3. Initialize Bounty Tracker', 'setupBountyTracker') 
    .addSeparator()
    .addItem('4. Pull Raw Attack Data', 'importWarData') 
    .addItem('5. Generate Payout Tab', 'buildPayoutTab') 
    .addItem('6. Calculate Payout Metrics', 'runPayoutMath') 
    .addItem('7. Generate Final Payouts', 'buildFinalPayoutTab')
    .addItem('8. 🔍 Run Auto-Auditor', 'runPayoutAudit') 
    .addSeparator()
    .addItem('9. Log War to History', 'logWarToHistory') 
    .addItem('10. ARCHIVE & RESET', 'archiveAndResetWar') 
    .addSeparator()
    .addItem('Rebuild Dashboard UI', 'buildDashboard') // <-- Moved safely to the bottom
    .addToUi();
}