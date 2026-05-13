function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚔️ Faction Tools')
    .addItem(' Quick Refresh Dashboard Data', 'refreshDashboard')
    .addItem('2. Update Faction Roster', 'updateRoster')
    .addItem('3. Initialize Bounty Tracker', 'setupBountyTracker') 
    .addSeparator()
    .addItem('4. Fetch Official Torn Reports', 'fetchOfficialReports') // <-- NEW: Pulls the API reports
    .addItem('5. Pull Raw Attack Data', 'importWarData') 
    .addItem('6. Generate Payout Tab', 'buildPayoutTab') 
    .addItem('7. Calculate Payout Metrics', 'runPayoutMath') 
    .addItem('8. Generate Final Payouts', 'buildFinalPayoutTab')
    .addItem('9. Publish Payout to Public Sheet', 'publishPayoutToPublic')
    .addSeparator()
    .addItem('10. Log War to History', 'logWarToHistory') 
    .addItem('11. ARCHIVE & RESET', 'archiveAndResetWar') 
    .addItem('🧪 Test Archive (No Reset)', 'testArchiveOnly')
    .addItem('🧹 Clean Sweep (Reset Sheet)', 'cleanSweep')
    .addSeparator()
    .addItem('🔍 Run Auto-Auditor', 'runPayoutAudit') 
    .addItem('🎯 Run Sniper Audit (Check missing hits)', 'runSniperAudit') // <-- NEW: Player-specific debugger
    .addItem('Rebuild Dashboard UI', 'buildDashboard') 
    .addToUi();
}