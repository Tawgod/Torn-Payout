// ==========================================
// ARCHIVE ROSTER ANALYTICS & TREND DASHBOARD
// ==========================================

function buildRosterAnalytics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  ss.toast("Scanning archived wars...", "System", 2);

  // 1. Get all sheets and prepare the roster dictionary
  const sheets = ss.getSheets();
  let roster = {}; // { memberId: { name, totalHits, totalScore, warsParticipated, totalPayout } }
  let totalWarsScanned = 0;

  // 2. Loop through every sheet (skip the analytics dashboard if it already exists)
  for (let s = 0; s < sheets.length; s++) {
    let sheet = sheets[s];
    let sheetName = sheet.getName();
    
    if (sheetName === "Roster Analytics" || sheetName === "Instructions") continue;

    // Based on your archive setup, data headers usually live on Row 7
    let lastRow = sheet.getLastRow();
    let lastCol = sheet.getLastColumn();
    if (lastRow < 8 || lastCol < 3) continue; // Skip empty or malformed sheets

    let data = sheet.getRange(7, 1, lastRow - 6, lastCol).getValues();
    let headers = data[0].map(h => h.toString().toLowerCase().trim());
    
    // 3. Dynamic Header Mapping - Find where columns live in THIS specific sheet
    let idxId = headers.findIndex(h => h === "member id" || h === "id");
    let idxName = headers.findIndex(h => h === "member name" || h === "name");
    let idxHits = headers.findIndex(h => h.includes("attacks") || h.includes("hits") || h.includes("war hits"));
    let idxScore = headers.findIndex(h => h.includes("war score") || h.includes("score"));
    let idxPayout = headers.findIndex(h => h.includes("payout") || h.includes("net payout"));

    // If we can't find basic identifiers, skip this tab
    if (idxId === -1 || idxName === -1) continue;

    totalWarsScanned++;

    // 4. Aggregate the data for this war
    for (let r = 1; r < data.length; r++) { // Start at 1 to skip headers
      let row = data[r];
      let memId = row[idxId];
      if (!memId || memId === "") continue;

      let memName = row[idxName] || "Unknown";
      let hits = idxHits !== -1 ? (parseFloat(row[idxHits]) || 0) : 0;
      let score = idxScore !== -1 ? (parseFloat(row[idxScore]) || 0) : 0;
      let payout = idxPayout !== -1 ? (parseFloat(row[idxPayout]) || 0) : 0;

      // Ignore standard non-player rows
      if (memName.toString().includes("LEFT FACTION") || memName.toString().includes("TOTALS")) continue;

      if (!roster[memId]) {
        roster[memId] = {
          name: memName,
          totalHits: 0,
          totalScore: 0,
          totalPayout: 0,
          warsParticipated: 0
        };
      }

      // Add to running totals
      roster[memId].totalHits += hits;
      roster[memId].totalScore += score;
      roster[memId].totalPayout += payout;
      roster[memId].warsParticipated += 1;
    }
  }

  // 5. Setup the Analytics Sheet
  let dashName = "Roster Analytics";
  let dashSheet = ss.getSheetByName(dashName);
  if (!dashSheet) {
    dashSheet = ss.insertSheet(dashName, 0); // Put it at the front
  } else {
    dashSheet.clear();
    if (dashSheet.getFilter()) dashSheet.getFilter().remove();
  }

  // 6. Build the Output Array
  let output = [];
  let outHeaders = [
    "Member ID", 
    "Member Name", 
    "Wars Attended", 
    "Missed Wars (Zero Hits)", 
    "Total Hits", 
    "Avg Hits / War", 
    "Total Score", 
    "Total Net Payout"
  ];
  output.push(outHeaders);

  for (let id in roster) {
    let m = roster[id];
    let missedWars = totalWarsScanned - m.warsParticipated;
    let avgHits = m.warsParticipated > 0 ? (m.totalHits / m.warsParticipated) : 0;
    
    output.push([
      id, 
      m.name, 
      m.warsParticipated, 
      missedWars, 
      m.totalHits, 
      avgHits, 
      m.totalScore, 
      m.totalPayout
    ]);
  }

  // 7. Write Data to Sheet and Format
  if (output.length > 1) {
    dashSheet.getRange(1, 1, output.length, outHeaders.length).setValues(output);
    
    // Styling
    let headerRange = dashSheet.getRange(1, 1, 1, outHeaders.length);
    headerRange.setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
    
    // Number Formats
    dashSheet.getRange(2, 6, output.length - 1, 1).setNumberFormat("0.0"); // Avg Hits
    dashSheet.getRange(2, 7, output.length - 1, 1).setNumberFormat("#,##0"); // Score
    dashSheet.getRange(2, 8, output.length - 1, 1).setNumberFormat('"$ "#,##0'); // Payout
    
    // Auto-Resize and Freeze Headers
    dashSheet.autoResizeColumns(1, outHeaders.length);
    dashSheet.setFrozenRows(1);
    
    // Turn on Native Filtering for easy sorting
    dashSheet.getDataRange().createFilter();
    
    ui.alert(`✅ Dashboard Built!\n\nSuccessfully scanned ${totalWarsScanned} archived wars and generated analytics for ${output.length - 1} members.`);
  } else {
    ui.alert("⚠️ No data found to analyze.");
  }
}