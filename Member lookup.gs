// ==========================================
// COMPREHENSIVE INDIVIDUAL MEMBER LOOKUP (INTERACTIVE CHARTS - FIXED)
// ==========================================

function searchMemberHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const tabName = "Member Lookup";
  let lookupSheet = ss.getSheetByName(tabName);

  // 1. Setup the UI if it doesn't exist
  if (!lookupSheet) {
    lookupSheet = ss.insertSheet(tabName, 0);
    lookupSheet.getRange("A1").setValue("🔍 COMPREHENSIVE MEMBER LOOKUP").setFontWeight("bold").setFontSize(14);
    lookupSheet.getRange("A2").setValue("Enter Member ID or Name:").setFontWeight("bold");
    lookupSheet.getRange("B2").setBackground("#fff2cc").setBorder(true, true, true, true, false, false);
    
    ui.alert("✅ Lookup tab created!\n\nPlease type a Member Name or ID into cell B2, then run this script again from the menu.");
    return;
  }

  // 2. Get the search term
  let searchTerm = lookupSheet.getRange("B2").getValue().toString().toLowerCase().trim();
  if (!searchTerm) {
    ui.alert("⚠️ Please enter a Member ID or Name in cell B2.");
    return;
  }

  ss.toast(`Searching archives for: ${searchTerm}...`, "Scanning", 3);

  let sheets = ss.getSheets();
  let results = [];
  let masterHeaders = ["War ID"]; 

  // 3. Scan all archived sheets
  for (let s = 0; s < sheets.length; s++) {
    let sheet = sheets[s];
    let sheetName = sheet.getName();
    
    if (sheetName === tabName || sheetName === "Instructions" || sheetName === "Roster Analytics") continue;

    let warIdNum = parseInt(sheetName.replace(/\D/g, '')) || 0;
    let lastRow = sheet.getLastRow();
    let lastCol = sheet.getLastColumn();
    
    if (lastRow < 8 || lastCol < 3) continue; 

    let data = sheet.getRange(7, 1, lastRow - 6, lastCol).getValues();
    let headers = data[0];
    
    let idxId = headers.findIndex(h => h.toString().toLowerCase().trim() === "member id" || h.toString().toLowerCase().trim() === "id");
    let idxName = headers.findIndex(h => h.toString().toLowerCase().trim() === "member name" || h.toString().toLowerCase().trim() === "name");

    if (idxId === -1 && idxName === -1) continue;

    // 4. Search the rows for the member
    for (let r = 1; r < data.length; r++) {
      let rowId = idxId !== -1 ? data[r][idxId].toString().toLowerCase().trim() : "";
      let rowName = idxName !== -1 ? data[r][idxName].toString().toLowerCase().trim() : "";

      if (rowId === searchTerm || rowName.includes(searchTerm)) {
        let rowData = { "War ID": sheetName, "_sortVal": warIdNum };

        // 5. Dynamically grab every column
        for (let c = 0; c < headers.length; c++) {
          let colHeader = headers[c] ? headers[c].toString().trim() : "";
          if (colHeader === "") continue; 
          if (colHeader.toLowerCase() === "points") continue;

          let cellValue = data[r][c];

          // Round Respect to the nearest whole number
          if (colHeader.toLowerCase().includes("respect")) {
            cellValue = Math.round(parseFloat(cellValue)) || 0;
          }

          rowData[colHeader] = cellValue;
          if (!masterHeaders.includes(colHeader)) {
            masterHeaders.push(colHeader);
          }
        }
        results.push(rowData);
        break; 
      }
    }
  }

  // 6. Clear Old Data (starting at row 25 for data)
  let startRow = 25;
  let maxRows = lookupSheet.getMaxRows();
  let maxCols = 25; // FIX: Stop clearing at Column Y to protect the dropdown in Column Z!
  
  if (maxRows >= startRow) {
    lookupSheet.getRange(startRow, 1, maxRows - startRow + 1, maxCols).clearContent();
    lookupSheet.getRange(startRow, 1, maxRows - startRow + 1, maxCols).clearFormat();
  }

  // Unfreeze rows to prevent Google Sheets from pushing the charts down
  lookupSheet.setFrozenRows(0);

  if (results.length > 0) {
    results.sort((a, b) => a._sortVal - b._sortVal);

    let outputArray = [masterHeaders]; 
    for (let i = 0; i < results.length; i++) {
      let resultRow = [];
      for (let h = 0; h < masterHeaders.length; h++) {
        let headerName = masterHeaders[h];
        resultRow.push(results[i].hasOwnProperty(headerName) ? results[i][headerName] : "");
      }
      outputArray.push(resultRow);
    }

    let targetRange = lookupSheet.getRange(startRow, 1, outputArray.length, masterHeaders.length);
    targetRange.setValues(outputArray);

    let headerRange = lookupSheet.getRange(startRow, 1, 1, masterHeaders.length);
    headerRange.setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
    
    for (let h = 0; h < masterHeaders.length; h++) {
      let headerName = masterHeaders[h].toLowerCase();
      let colRange = lookupSheet.getRange(startRow + 1, h + 1, outputArray.length - 1, 1);
      
      if (headerName.includes("contribution")) colRange.setNumberFormat("0.00%");
      else if (headerName.includes("respect") || headerName.includes("score") || headerName.includes("attacks") || headerName.includes("hits")) colRange.setNumberFormat('#,##0');
      else if (headerName.includes("payout") || headerName.includes("deduction") || headerName.includes("cost")) colRange.setNumberFormat('"$ "#,##0');
    }
    lookupSheet.autoResizeColumns(1, masterHeaders.length);
    
    // ==========================================
    // 7. BUILD CHART FILTERS UI
    // ==========================================
    
    let filterOptions = masterHeaders.filter(h => h !== "War ID"); 
    let defaultChart1 = filterOptions.find(h => h.toLowerCase().includes("contribution")) || filterOptions[0];
    let defaultChart2 = filterOptions.find(h => h.toLowerCase().includes("respect")) || filterOptions.find(h => h.toLowerCase().includes("attacks")) || filterOptions[1] || filterOptions[0];

    lookupSheet.getRange("A4").setValue("📈 CHART 1 METRIC").setFontWeight("bold");
    lookupSheet.getRange("A7").setValue("📉 CHART 2 METRIC").setFontWeight("bold");

    let rule = SpreadsheetApp.newDataValidation().requireValueInList(filterOptions).build();
    lookupSheet.getRange("A5:B5").merge().setDataValidation(rule).setValue(defaultChart1).setBackground("#eaf2f8").setBorder(true, true, true, true, false, false);
    lookupSheet.getRange("A8:B8").merge().setDataValidation(rule).setValue(defaultChart2).setBackground("#eaf2f8").setBorder(true, true, true, true, false, false);

    // Call the chart drawer
    redrawCharts(lookupSheet);

    ui.alert(`✅ Success!\n\nLoaded ${results.length} records. Use the dropdowns in A5 and A8 to change the graphs.`);
  } else {
    ui.alert("⚠️ No records found for that ID or Name.");
  }
}

// ==========================================
// ON-EDIT TRIGGER (MAKES DROPDOWNS INTERACTIVE)
// ==========================================
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  
  if (sheet.getName() === "Member Lookup") {
    let row = e.range.getRow();
    let col = e.range.getColumn();
    
    if (col <= 2 && (row === 5 || row === 8)) {
      redrawCharts(sheet);
    }
  }
}

// ==========================================
// DYNAMIC CHART DRAWER
// ==========================================
function redrawCharts(sheet) {
  // 1. SAFELY Remove old charts
  try {
    let existingCharts = sheet.getCharts();
    for (let i = 0; i < existingCharts.length; i++) {
      sheet.removeChart(existingCharts[i]);
    }
  } catch (e) {
    console.log("Chart fetch bypass triggered: " + e.toString());
  }

  let startRow = 25;
  let lastRow = sheet.getLastRow();
  if (lastRow <= startRow) return; 
  
  let numRows = lastRow - startRow + 1;
  let headers = sheet.getRange(startRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  let colWarId = headers.indexOf("War ID") + 1;
  if (colWarId === 0) colWarId = 1; 

  let metric1 = sheet.getRange("A5").getValue();
  let metric2 = sheet.getRange("A8").getValue();

  let col1 = headers.indexOf(metric1) + 1;
  let col2 = headers.indexOf(metric2) + 1;

  // 2. SAFELY Build Chart 1 in D4
  if (col1 > 0) {
    try {
      let chart1 = sheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(sheet.getRange(startRow, colWarId, numRows, 1)) 
        .addRange(sheet.getRange(startRow, col1, numRows, 1))     
        .setPosition(4, 4, 0, 0)
        .setOption('title', `${metric1} Over Time`)
        .setOption('legend', {position: 'none'})
        .setOption('width', 450)
        .setOption('height', 360) // Adjusted height to fit nicely in the gap
        .build();
      sheet.insertChart(chart1);
    } catch (e) {
      console.log("Error building Chart 1: " + e.toString());
    }
  }

  // 3. SAFELY Build Chart 2 in J4
  if (col2 > 0) {
    try {
      let chart2 = sheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(sheet.getRange(startRow, colWarId, numRows, 1)) 
        .addRange(sheet.getRange(startRow, col2, numRows, 1))     
        .setPosition(4, 10, 0, 0)
        .setOption('title', `${metric2} Over Time`)
        .setOption('legend', {position: 'none'})
        .setOption('width', 450)
        .setOption('height', 360) // Adjusted height to fit nicely in the gap
        .build();
      sheet.insertChart(chart2);
    } catch (e) {
      console.log("Error building Chart 2: " + e.toString());
    }
  }
}

// ==========================================
// BUILD SEARCH-AS-YOU-TYPE AUTOCOMPLETE
// ==========================================

function buildSearchDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let lookupSheet = ss.getSheetByName("Member Lookup");
  
  if (!lookupSheet) {
    ui.alert("⚠️ Please run the Member Search at least once to build the lookup tab.");
    return;
  }

  ss.toast("Harvesting names from archives...", "System", 2);

  let sheets = ss.getSheets();
  let uniqueMembers = new Set(); // Using a Set prevents duplicates

  // 1. Scan all archives to build a master list of names and IDs
  for (let s = 0; s < sheets.length; s++) {
    let sheet = sheets[s];
    let sheetName = sheet.getName();
    
    if (sheetName === "Member Lookup" || sheetName === "Instructions" || sheetName === "Roster Analytics") continue;

    let lastRow = sheet.getLastRow();
    let lastCol = sheet.getLastColumn();
    if (lastRow < 8 || lastCol < 3) continue;

    let data = sheet.getRange(7, 1, lastRow - 6, lastCol).getValues();
    let headers = data[0];
    
    let idxId = headers.findIndex(h => h.toString().toLowerCase().trim() === "member id" || h.toString().toLowerCase().trim() === "id");
    let idxName = headers.findIndex(h => h.toString().toLowerCase().trim() === "member name" || h.toString().toLowerCase().trim() === "name");

    if (idxId === -1 && idxName === -1) continue;

    // Add every name and ID we find to the Set
    for (let r = 1; r < data.length; r++) {
      if (idxName !== -1 && data[r][idxName]) uniqueMembers.add(data[r][idxName].toString().trim());
    }
  }

  // 2. Convert to an alphabetical array
  let memberArray = Array.from(uniqueMembers).sort().map(m => [m]);

  if (memberArray.length === 0) {
    ui.alert("⚠️ No members found in the archives.");
    return;
  }

  // 3. Store the list in a hidden column (Column Z)
  lookupSheet.getRange("Z:Z").clearContent();
  lookupSheet.getRange(1, 26, memberArray.length, 1).setValues(memberArray);
  lookupSheet.hideColumns(26); // Hide column Z to keep things clean

  // 4. Apply the Dropdown to cell B2
  let range = lookupSheet.getRange("Z1:Z" + memberArray.length);
  // requireValueInRange(range, true) turns on the dropdown arrow and autocomplete
  let rule = SpreadsheetApp.newDataValidation().requireValueInRange(range, true).setAllowInvalid(true).build();
  
  lookupSheet.getRange("B2").setDataValidation(rule);
  
  ui.alert(`✅ Autocomplete created!\n\nHarvested ${memberArray.length} unique names and IDs. You can now search while typing in cell B2.`);
}