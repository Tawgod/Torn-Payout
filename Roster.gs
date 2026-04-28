function updateRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  const apiKey = configSheet.getRange(SETTINGS.apiKeyCell).getValue();

  if (!apiKey) {
    ss.toast("Error: No API Key found in Config!B1", "System", 5);
    return;
  }

  // 1. Fetch Faction Basic Data (Contains the member list)
  const factionUrl = `https://api.torn.com/faction/?selections=basic&key=${apiKey}`;
  const response = UrlFetchApp.fetch(factionUrl, { muteHttpExceptions: true });
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    ss.toast(`Torn API Error: ${json.error.error}`, "System", 5);
    return;
  }

  const members = json.members;
  if (!members) {
    ss.toast("Could not find member data.", "System", 5);
    return;
  }

  // 2. Prepare the Roster Sheet
  let rosterSheet = ss.getSheetByName(SETTINGS.rosterSheet);
  if (!rosterSheet) {
    rosterSheet = ss.insertSheet(SETTINGS.rosterSheet, 2); // Put it after Dashboard
  }
  
  rosterSheet.clear();
  
  // Set headers and format them
  rosterSheet.getRange(1, 1, 1, SETTINGS.rosterHeaders.length)
             .setValues([SETTINGS.rosterHeaders])
             .setBackground("#4c1130") // Dark red/purple header
             .setFontColor("white")
             .setFontWeight("bold");

  // 3. Process the Member Data
  const rows = [];
  
  for (let id in members) {
    let m = members[id];
    
    // Convert Torn's "last_action" timestamp to a readable date
    let lastActionDate = m.last_action.timestamp ? new Date(m.last_action.timestamp * 1000) : "Unknown";
    
    rows.push([
      id,
      m.name || "Unknown",
      m.level || 0,
      m.position || "Member",
      m.days_in_faction || 0,
      m.status.state || "Unknown",
      lastActionDate
    ]);
  }

  // 4. Write data to the sheet
  if (rows.length > 0) {
    // Sort alphabetically by Name (Column 2) before inserting
    rows.sort((a, b) => a[1].localeCompare(b[1])); 
    
    rosterSheet.getRange(2, 1, rows.length, SETTINGS.rosterHeaders.length).setValues(rows);
    
    // Format the "Last Action" column to look nice
    rosterSheet.getRange(2, 7, rows.length, 1).setNumberFormat("m/d/yyyy h:mm am/pm");
    
    // Freeze the top row and auto-resize columns
    rosterSheet.setFrozenRows(1);
    rosterSheet.autoResizeColumns(1, SETTINGS.rosterHeaders.length);
  }

  ss.toast(`Successfully updated roster with ${rows.length} members.`, "System", 5);
}