const url = `https://api.torn.com/faction/?selections=basic&key=${API_KEY}`;

function updateHistoricalRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Historical Roster");

  // 1. Initialize Sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet("Historical Roster");
    const headers = ["Member ID", "Name", "Level", "Days in Faction", "Status", "First Seen", "Last Seen Active"];
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setFontWeight("bold")
         .setBackground("#4a86e8")
         .setFontColor("white");
    sheet.setFrozenRows(1);
  }

  // 2. Fetch current faction members from Torn API
  const url = `https://api.torn.com/faction/?selections=basic&key=${API_KEY}`;
  let response;
  try {
    response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch(e) {
    return; // Fail silently if Torn API is down
  }

  const json = JSON.parse(response.getContentText());
  if (json.error || !json.members) return;

  const currentMembers = json.members;
  const currentMemberIds = Object.keys(currentMembers);
  const today = new Date();

  // 3. Load existing historical data
  const lastRow = sheet.getLastRow();
  let existingData = [];
  let memberMap = {}; // Maps ID to row index

  if (lastRow > 1) {
    existingData = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (let i = 0; i < existingData.length; i++) {
      let id = existingData[i][0].toString().trim();
      memberMap[id] = i;
    }
  }

  let newRows = [];
  let updatedData = [...existingData];

  // 4. Temporarily mark everyone as "Left Faction". 
  // We will overwrite the active ones in the next step.
  for (let i = 0; i < updatedData.length; i++) {
     if (updatedData[i][4] === "Active") {
         updatedData[i][4] = "Left Faction"; 
     }
  }

  // 5. Process the live API data
  for (let id of currentMemberIds) {
    let mem = currentMembers[id];
    
    if (memberMap.hasOwnProperty(id)) {
      // Member exists: Update their dynamic stats
      let rowIndex = memberMap[id];
      updatedData[rowIndex][1] = mem.name;
      updatedData[rowIndex][2] = mem.level;
      updatedData[rowIndex][3] = mem.days_in_faction;
      updatedData[rowIndex][4] = "Active";
      updatedData[rowIndex][6] = today; // Update "Last Seen"
    } else {
      // Brand new member: Log them for the first time
      newRows.push([
        id, 
        mem.name, 
        mem.level, 
        mem.days_in_faction, 
        "Active", 
        today, 
        today  
      ]);
    }
  }

  // 6. Write everything back to the sheet
  if (updatedData.length > 0) {
    sheet.getRange(2, 1, updatedData.length, 7).setValues(updatedData);
  }
  
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
  }
  
  // Format dates and columns
  const finalRow = sheet.getLastRow();
  if (finalRow > 1) {
    sheet.getRange(2, 6, finalRow - 1, 2).setNumberFormat("m/d/yyyy");
  }
  sheet.autoResizeColumns(1, 7);
}