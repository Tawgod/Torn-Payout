function fetchActiveWarDetails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  if (!dashSheet) {
    ss.toast("Please run 'Rebuild Dashboard' first.", "System", 5);
    return;
  }

  let myFactionId;
  try {
    const userRes = tornApiRequest_("user", "", "profile");
    if (userRes.error) throw new Error(userRes.error.error);
    myFactionId = userRes.faction.faction_id.toString();
  } catch (e) {
    ss.toast("Error: Could not retrieve your Faction ID.", "System", 5);
    return;
  }

  let factionRes;
  try {
    factionRes = tornApiRequest_("faction", "", "basic");
  } catch (e) {
    ss.toast("Railway/Torn API error: " + e.message, "System", 5);
    return;
  }

  if (factionRes.error) {
    ss.toast(`Torn API Error: ${factionRes.error.error}`, "System", 5);
    return;
  }

  let enemyId = null, warStartUnix = null, activeWarId = null;

  if (factionRes.ranked_wars && Object.keys(factionRes.ranked_wars).length > 0) {
    activeWarId = Object.keys(factionRes.ranked_wars)[0]; 
    const warData = factionRes.ranked_wars[activeWarId];
    
    const factionIds = Object.keys(warData.factions);
    enemyId = factionIds.find(id => id !== myFactionId);
    warStartUnix = warData.war.start;
  }

  if (enemyId && warStartUnix && activeWarId) {
    dashSheet.getRange("C3").setValue(enemyId);
    
    dashSheet.getRange("C4").setValue(new Date(warStartUnix * 1000));
    dashSheet.getRange("C4").setNumberFormat("m/d/yyyy h:mm:ss am/pm");
    
    // Write War ID to Dashboard instead of Config
    dashSheet.getRange("C5").setValue(activeWarId);
    
    // Clear out old End-of-War stats
    dashSheet.getRange("C9:C11").setValue(""); 
    
    ss.toast(`Setup Complete! Target: ${enemyId}`, "War Tracker", 5);
  } else {
    ss.toast("No active Ranked War found.", "War Tracker", 5);
  }
}