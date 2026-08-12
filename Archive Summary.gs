// ==========================================
// ARCHIVE SUMMARY BUILDER
// ==========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Analytics & 📚 Archive Tools')
    .addItem('Search Member History', 'searchMemberHistory')
    .addItem('Build/Run Comparison', 'runComparison')
    .addItem('Refresh Member Dropdowns', 'refreshManualDropdownsMenu')
    .addItem('🔄 Rebuild Summary Page', 'buildArchiveSummary')
    .addToUi();
}
    
function buildArchiveSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summaryName = "📋 Archive Summary";
  
  ss.toast("Scanning archive tabs...", "System", 3);
  
  // 1. Get or Create the Summary Sheet at the front
  let summarySheet = ss.getSheetByName(summaryName);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(summaryName, 0); 
  } else {
    ss.setActiveSheet(summarySheet);
    ss.moveActiveSheet(1);
  }
  
  const sheets = ss.getSheets();
  let summaryData = [];
  
  // 2. Loop through all sheets and extract data
  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    let sheetName = sheet.getName();
    
    // Skip the summary sheet itself
    if (sheetName === summaryName) continue;
    
    // Grab the standardized 5-row header block pushed by the Calculator
    let headerData = sheet.getRange("A1:F5").getValues();
    let titleRow = headerData[0][0] ? headerData[0][0].toString() : "";
    
    // Ensure this is actually an archived war/chain tab by checking A1
    if (!titleRow.includes("WAR REPORT") && !titleRow.includes("REPORT")) {
      continue; 
    }
    
    // Clean up the title to extract just the Enemy Name
    let enemyName = titleRow.replace(/WAR REPORT:\s*/i, "").replace(/\[.*?\]/, "").trim();
    if (!enemyName) enemyName = "Unknown Faction / Chain";

    // Extract core stats
    let dateArchived = headerData[1][1];                 
    let outcome = headerData[1][3] || "-";   
    let termed = headerData[1][5] ? headerData[1][5].toString().trim() : "No";             
    let totalHits = parseInt(headerData[2][1]) || 0;     
    let score = parseInt(headerData[2][5]) || 0;         
    let payout = parseFloat(headerData[3][3]) || 0;      
    let profit = parseFloat(headerData[3][5]) || 0;      
    
    // Create a clickable hyperlink to jump to the specific tab
    let sheetId = sheet.getSheetId();
    let linkFormula = `=HYPERLINK("#gid=${sheetId}", "${sheetName}")`;
    
    // Push to our array
    summaryData.push([
      dateArchived,
      linkFormula,
      enemyName,
      outcome,
      termed,    
      totalHits,
      score,
      payout,
      profit
    ]);
  }
  
  // 3. Sort the data by Date Archived (Newest at the top)
  summaryData.sort((a, b) => {
    let dateA = new Date(a[0]).getTime();
    let dateB = new Date(b[0]).getTime();
    if (isNaN(dateA)) dateA = 0;
    if (isNaN(dateB)) dateB = 0;
    return dateB - dateA; 
  });
  
  // 4. Print to the Summary Sheet
  summarySheet.clear();
  
  let headers = [["Date Archived", "War ID (Link)", "Enemy Faction", "Outcome", "Termed", "Total Hits", "War Score", "Total Payout", "Faction Profit"]];
  
  // Format Headers
  summarySheet.getRange(1, 1, 1, headers[0].length).setValues(headers)
    .setBackground("#2f4f4f")
    .setFontColor("white")
    .setFontWeight("bold");
    
  if (summaryData.length > 0) {
    let dataRange = summarySheet.getRange(2, 1, summaryData.length, headers[0].length);
    dataRange.setValues(summaryData);
    
    // Apply Formatting 
    summarySheet.getRange(2, 1, summaryData.length, 1).setNumberFormat("yyyy-mm-dd");     // Date
    summarySheet.getRange(2, 6, summaryData.length, 2).setNumberFormat("#,##0");          // Hits & Score
    summarySheet.getRange(2, 8, summaryData.length, 2).setNumberFormat('"$ "#,##0');      // Money
    
    // Center Align outcome, termed status, hits, and score
    summarySheet.getRange(2, 4, summaryData.length, 4).setHorizontalAlignment("center");
  } else {
    summarySheet.getRange(2, 1).setValue("No archived wars found yet.");
  }
  
  // 5. Final Polish
  summarySheet.autoResizeColumns(1, headers[0].length); // Auto-resize everything first
  summarySheet.setColumnWidth(2, 120); // Force War ID link column width
  
  // ---> THE FIX: Force the Enemy Faction column to be larger <---
  summarySheet.setColumnWidth(3, 250); 
  
  summarySheet.setFrozenRows(1);
  
  ss.toast("Archive Summary successfully rebuilt!", "Success", 5);
}