// ==========================================
// FACTION ROSTER COMPARISON ENGINE (V7 - WITH TREND LINES)
// ==========================================

function runComparison() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const tabName = "Roster Comparison";
  let compSheet = ss.getSheetByName(tabName);

  // ==========================================
  // 🛑 PASTE YOUR MAIN PAYOUT SHEET ID HERE 🛑
  // ==========================================
  const PAYOUT_SHEET_ID = "14KBGkYc5UagtCYyYhjJuqBJ_xGp1sfPNPw8ylOSnr50"; 

  // 1. BUILD THE UI IF IT DOESN'T EXIST
  if (!compSheet) {
    compSheet = ss.insertSheet(tabName, 1);
    
    compSheet.getRange("A1").setValue("⚖️ ROSTER COMPARISON ENGINE").setFontWeight("bold").setFontSize(14);
    
    let labels = [
      ["Rank Mode:", "Top Performers"],
      ["Rank Metric:", "Net Respect"],
      ["Lookback Start (1 = Newest War):", 1],
      ["Lookback End (e.g. 5 = 5th Oldest):", 5],
      ["Members to Show (1-15):", 10]
    ];
    compSheet.getRange("A2:B6").setValues(labels);
    compSheet.getRange("A2:A6").setFontWeight("bold").setBackground("#eaf2f8");
    compSheet.getRange("B2:B6").setBackground("#fff2cc").setBorder(true, true, true, true, false, false);

    let modeRule = SpreadsheetApp.newDataValidation().requireValueInList(["Top Performers", "Bottom Performers", "Manual Selection"]).build();
    compSheet.getRange("B2").setDataValidation(modeRule);
    
    let metricRule = SpreadsheetApp.newDataValidation().requireValueInList(["Total Hits", "Net Respect", "War Score", "Total Payout", "Average Contribution"]).build();
    compSheet.getRange("B3").setDataValidation(metricRule);

    compSheet.getRange("A8").setValue("👇 MANUAL SELECTION (Select up to 10 Members)").setFontWeight("bold");
    for(let i = 0; i < 10; i++) {
      compSheet.getRange(9 + i, 1).setValue(`Manual Member ${i + 1}:`).setFontWeight("bold").setBackground("#eaf2f8");
      compSheet.getRange(9 + i, 2).setBackground("#fff2cc").setBorder(true, true, true, true, false, false);
    }

    compSheet.autoResizeColumns(1, 2);
    updateManualDropdowns(ss, compSheet);

    ui.alert("✅ Comparison tab created!\n\nSet your parameters and run 'Build/Run Comparison' again.");
    return;
  }

  // 2. READ PARAMETERS
  let mode = compSheet.getRange("B2").getValue().toString().trim();
  let metric = compSheet.getRange("B3").getValue().toString().trim();
  let startIdx = parseInt(compSheet.getRange("B4").getValue()) || 1;
  let endIdx = parseInt(compSheet.getRange("B5").getValue()) || 5;
  let numMembers = parseInt(compSheet.getRange("B6").getValue()) || 10;

  let manualInputs = compSheet.getRange("B9:B18").getValues()
                      .map(row => row[0].toString().toLowerCase().trim())
                      .filter(val => val !== "");

  if (mode === "Manual Selection" && manualInputs.length === 0) {
    ui.alert("⚠️ You selected 'Manual Selection' but left the slots blank.");
    return;
  }
  if (startIdx > endIdx) {
    ui.alert("⚠️ Lookback Start cannot be greater than Lookback End.");
    return;
  }

  // 3. FETCH API KEY
  let apiKey = "";
  if (mode !== "Manual Selection") {
    if (PAYOUT_SHEET_ID === "PASTE_YOUR_PAYOUT_SHEET_ID_HERE") {
       ui.alert("⚠️ Paste your Payout Sheet ID into the script editor first!");
       return;
    }
    ss.toast("Connecting to Payout Sheet to grab API Key...", "System", 2);
    try {
      let externalSS = SpreadsheetApp.openById(PAYOUT_SHEET_ID);
      let configSheet = externalSS.getSheetByName("Config");
      if (!configSheet) {
        ui.alert("⚠️ Could not find a tab named 'Config' in your Payout Sheet.");
        return;
      }
      let configData = configSheet.getDataRange().getValues();
      for (let r = 0; r < configData.length; r++) {
        if (configData[r][0] && configData[r][0].toString().toLowerCase().includes("api")) {
          apiKey = configData[r][1].toString().trim();
          break;
        }
      }
    } catch (e) {
       ui.alert("⚠️ Error connecting to Payout Sheet.\nDetails: " + e.message);
       return;
    }
    if (!apiKey) {
      ui.alert("⚠️ Could not find 'API Key' on the Config tab.");
      return;
    }
  }

  // 4. FETCH LIVE FACTION ROSTER VIA API
  let activeFactionIDs = [];
  if (mode !== "Manual Selection") { 
    ss.toast(`Fetching Live Faction Roster from Torn...`, "API Check", 3);
    try {
      let response = UrlFetchApp.fetch(`https://api.torn.com/faction/?selections=basic&key=${apiKey}`);
      let json = JSON.parse(response.getContentText());
      if (json.error) throw new Error(json.error.error);
      if (json.members) activeFactionIDs = Object.keys(json.members); 
    } catch (e) {
      ui.alert("⚠️ API Error: Could not fetch faction roster.\nDetails: " + e.message);
      return;
    }
  }

  ss.toast(`Aggregating data for wars ${startIdx} to ${endIdx}...`, "Processing", 3);

  // 5. FETCH AND SORT ARCHIVED WARS
  let sheets = ss.getSheets();
  let warTabs = [];
  
  for (let s = 0; s < sheets.length; s++) {
    let sheetName = sheets[s].getName();
    if (sheetName === tabName || sheetName === "Member Lookup" || sheetName === "Instructions" || sheetName === "Roster Analytics" || sheetName === "Config") continue;
    let warIdNum = parseInt(sheetName.replace(/\D/g, '')) || 0;
    if (warIdNum > 0) warTabs.push({ sheet: sheets[s], id: warIdNum, name: sheetName });
  }

  warTabs.sort((a, b) => b.id - a.id);
  let selectedWars = warTabs.slice(startIdx - 1, endIdx);
  
  if (selectedWars.length === 0) {
    ui.alert("⚠️ No wars found in that range.");
    return;
  }

  // 6. AGGREGATE DATA BY MEMBER ID & TRACK TRENDS
  let roster = {};

  for (let w = 0; w < selectedWars.length; w++) {
    let sheet = selectedWars[w].sheet;
    let warName = selectedWars[w].name;
    let lastRow = sheet.getLastRow();
    let lastCol = sheet.getLastColumn();
    if (lastRow < 8 || lastCol < 3) continue; 

    let data = sheet.getRange(7, 1, lastRow - 6, lastCol).getValues();
    let headers = data[0].map(h => h.toString().toLowerCase().trim());
    
    let idxId = headers.findIndex(h => h === "member id" || h === "id");
    let idxName = headers.findIndex(h => h === "member name" || h === "name");
    
    if (idxId === -1 || idxName === -1) continue;

    let idxHits = headers.findIndex(h => h.includes("attacks") || h.includes("hits") || h.includes("war hits"));
    let idxScore = headers.findIndex(h => h.includes("war score") || h.includes("score"));
    let idxRespect = headers.findIndex(h => h.includes("respect"));
    let idxPayout = headers.findIndex(h => h.includes("payout") || h.includes("net payout"));
    let idxContrib = headers.findIndex(h => h.includes("contribution"));

    for (let r = 1; r < data.length; r++) {
      let memId = data[r][idxId].toString().trim();
      let rawName = data[r][idxName].toString().trim();
      
      if (!memId || rawName.includes("TOTALS")) continue;
      let cleanName = rawName.replace(/\(Left Faction\)/ig, "").trim();

      if (!roster[memId]) {
        roster[memId] = {
          id: memId, name: cleanName, totalHits: 0, totalScore: 0,
          totalRespect: 0, totalPayout: 0, totalContrib: 0, warsTracked: 0,
          history: {} // <-- NEW: Tracks data per war for the line chart
        };
      }

      let hits = idxHits !== -1 ? (parseFloat(data[r][idxHits]) || 0) : 0;
      let score = idxScore !== -1 ? (parseFloat(data[r][idxScore]) || 0) : 0;
      let respect = idxRespect !== -1 ? (parseFloat(data[r][idxRespect]) || 0) : 0;
      let payout = idxPayout !== -1 ? (parseFloat(data[r][idxPayout]) || 0) : 0;
      let contrib = idxContrib !== -1 ? (parseFloat(data[r][idxContrib]) || 0) : 0;

      roster[memId].totalHits += hits;
      roster[memId].totalScore += score;
      roster[memId].totalRespect += respect;
      roster[memId].totalPayout += payout;
      roster[memId].totalContrib += contrib;
      roster[memId].warsTracked += 1;

      // Save the specific metric for the line graph
      let metricVal = 0;
      if (metric === "Total Hits") metricVal = hits;
      else if (metric === "War Score") metricVal = score;
      else if (metric === "Net Respect") metricVal = respect;
      else if (metric === "Total Payout") metricVal = payout;
      else if (metric === "Average Contribution") metricVal = contrib;
      
      roster[memId].history[warName] = metricVal;
    }
  }

  // 7. SORT AND FILTER RESULTS
  let results = Object.values(roster);
  results.forEach(m => { m.avgContrib = m.warsTracked > 0 ? (m.totalContrib / m.warsTracked) : 0; });

  if (mode === "Manual Selection") {
    results = results.filter(m => {
      let lowerId = m.id.toLowerCase();
      let lowerName = m.name.toLowerCase();
      return manualInputs.includes(lowerId) || manualInputs.some(term => lowerName.includes(term));
    });
  } else {
    results = results.filter(m => activeFactionIDs.includes(m.id.toString()));
  }

  results.sort((a, b) => {
    let valA, valB;
    if (metric === "Total Hits") { valA = a.totalHits; valB = b.totalHits; }
    else if (metric === "War Score") { valA = a.totalScore; valB = b.totalScore; }
    else if (metric === "Net Respect") { valA = a.totalRespect; valB = b.totalRespect; }
    else if (metric === "Total Payout") { valA = a.totalPayout; valB = b.totalPayout; }
    else if (metric === "Average Contribution") { valA = a.avgContrib; valB = b.avgContrib; }
    else { valA = a.totalRespect; valB = b.totalRespect; } 

    if (mode === "Bottom Performers") return valA - valB;
    return valB - valA; 
  });

  let finalData = mode === "Manual Selection" ? results : results.slice(0, numMembers);

 // 8. OUTPUT TO SHEET (Safely clearing data zones)
  let startRow = 24; 
  let maxRows = compSheet.getMaxRows();
  let maxCols = 40; 
  
  if (maxRows >= startRow) {
    compSheet.getRange(startRow, 1, maxRows - startRow + 1, maxCols).clearContent().clearFormat();
  }
  
  try {
    let existingCharts = compSheet.getCharts();
    existingCharts.forEach(c => compSheet.removeChart(c));
  } catch(e) {} 

  if (finalData.length === 0) {
    ui.alert("⚠️ No matching data found for the selected members.");
    return;
  }

  // A. Build Primary Summary Table
  let outHeaders = ["Member ID", "Name", "Wars Included", "Total Hits", "War Score", "Net Respect", "Total Payout", "Avg Contribution"];
  let outputArray = [outHeaders];

  finalData.forEach(m => {
    outputArray.push([
      m.id, m.name, m.warsTracked, m.totalHits, m.totalScore,
      Math.round(m.totalRespect), m.totalPayout, m.avgContrib
    ]);
  });

  compSheet.getRange(startRow, 1, outputArray.length, outHeaders.length).setValues(outputArray);
  compSheet.getRange(startRow, 1, 1, outHeaders.length).setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
  compSheet.getRange(startRow + 1, 4, finalData.length, 3).setNumberFormat('#,##0'); 
  compSheet.getRange(startRow + 1, 7, finalData.length, 1).setNumberFormat('"$ "#,##0'); 
  compSheet.getRange(startRow + 1, 8, finalData.length, 1).setNumberFormat('0.00%'); 
  compSheet.autoResizeColumns(1, outHeaders.length);

  // B. Build Secondary Trend Table
  let trendStartCol = outHeaders.length + 2; 
  let trendHeaders = ["War Label"].concat(finalData.map(m => m.name));
  let trendOutput = [trendHeaders];

  let chronoWars = [...selectedWars].reverse();
  for (let w = 0; w < chronoWars.length; w++) {
    let warIdStr = chronoWars[w].name;
    let row = ["War " + chronoWars[w].id]; 
    
    for (let i = 0; i < finalData.length; i++) {
      row.push(finalData[i].history[warIdStr] || 0); 
    }
    trendOutput.push(row);
  }

  compSheet.getRange(startRow, trendStartCol, trendOutput.length, trendHeaders.length).setValues(trendOutput);
  compSheet.autoResizeColumns(trendStartCol, trendHeaders.length);

  // ==========================================
  // CUSTOM COLOR CODING & FORMATTING
  // ==========================================
  
  // 1. Define 20 distinct, vibrant colors to completely avoid overlap
  const chartColors = [
    '#3366cc', '#dc3912', '#ff9900', '#109618', '#990099', 
    '#0099c6', '#dd4477', '#66aa00', '#b82e2e', '#316395', 
    '#994499', '#22aa99', '#aaaa11', '#6633cc', '#e67300',
    '#8b0707', '#329262', '#5574a6', '#3b3eac', '#b56308'
  ];

  // 2. Format the "War Label" header
  compSheet.getRange(startRow, trendStartCol).setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");

  // 3. Apply exact matching colors to the Member Name headers to act as an indestructible legend!
  let headerBackgrounds = [[]];
  for(let i = 0; i < finalData.length; i++) {
    headerBackgrounds[0].push(chartColors[i % chartColors.length]);
  }
  compSheet.getRange(startRow, trendStartCol + 1, 1, finalData.length)
           .setBackgrounds(headerBackgrounds)
           .setFontColor("white")
           .setFontWeight("bold");

  // 4. Force Percentage formatting if the metric is Contribution
  let isPercentage = metric.toLowerCase().includes("contribution");
  let trendDataRange = compSheet.getRange(startRow + 1, trendStartCol + 1, trendOutput.length - 1, finalData.length);
  
  if (isPercentage) {
    trendDataRange.setNumberFormat('0.00%');
  } else {
    trendDataRange.setNumberFormat('#,##0');
  }

  // 9. DRAW DUAL CHARTS (Skipping the header row to prevent text from rendering)
  let metricColMap = { "Total Hits": 4, "War Score": 5, "Net Respect": 6, "Total Payout": 7, "Average Contribution": 8 };
  let colIdx = metricColMap[metric] || 6;
  
  // Dynamically switch the Y-Axis graph scale from numbers to percentages
  let vAxisFormat = isPercentage ? '0.00%' : '#,##0';

  // Chart 1: Bar Chart (Overall Performance) - Start at startRow + 1
  let chart1 = compSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(compSheet.getRange(startRow + 1, 2, finalData.length, 1)) // X-Axis (Names)
    .addRange(compSheet.getRange(startRow + 1, colIdx, finalData.length, 1)) // Y-Axis (Metric)
    .setPosition(2, 4, 0, 0)
    .setOption('title', `${mode}: Overall ${metric}`)
    .setOption('legend', {position: 'none'})
    .setOption('vAxis', {format: vAxisFormat})
    .setOption('width', 450)
    .setOption('height', 400)
    .build();
  compSheet.insertChart(chart1);

  // Chart 2: Line Chart (Trend Over Time) - Start at startRow + 1
  let chart2 = compSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(compSheet.getRange(startRow + 1, trendStartCol, trendOutput.length - 1, trendHeaders.length)) 
    .setPosition(2, 10, 0, 0)
    .setOption('title', `${metric} Trajectory Over Time (See colored headers for Legend)`)
    .setOption('legend', {position: 'none'}) 
    .setOption('useFirstColumnAsDomain', true)
    .setOption('colors', chartColors) 
    .setOption('vAxis', {format: vAxisFormat}) 
    .setOption('width', 650)
    .setOption('height', 400)
    .build();
  compSheet.insertChart(chart2);
}
// ==========================================
// HARVEST MEMBERS FOR DROPDOWN AUTOCOMPLETE
// ==========================================
function refreshManualDropdownsMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let compSheet = ss.getSheetByName("Roster Comparison");
  if (!compSheet) {
    SpreadsheetApp.getUi().alert("⚠️ Please run the comparison script first to create the tab.");
    return;
  }
  updateManualDropdowns(ss, compSheet);
  SpreadsheetApp.getUi().alert("✅ Member dropdowns refreshed!");
}

function updateManualDropdowns(ss, compSheet) {
  let sheets = ss.getSheets();
  let uniqueMembers = new Set();

  for (let s = 0; s < sheets.length; s++) {
    let sheet = sheets[s];
    let sheetName = sheet.getName();
    
    if (sheetName === "Member Lookup" || sheetName === "Roster Comparison" || sheetName === "Instructions" || sheetName === "Roster Analytics" || sheetName === "Config") continue;

    let lastRow = sheet.getLastRow();
    let lastCol = sheet.getLastColumn();
    if (lastRow < 8 || lastCol < 3) continue;

    let data = sheet.getRange(7, 1, lastRow - 6, lastCol).getValues();
    let headers = data[0];
    
    let idxId = headers.findIndex(h => h.toString().toLowerCase().trim() === "member id" || h.toString().toLowerCase().trim() === "id");
    let idxName = headers.findIndex(h => h.toString().toLowerCase().trim() === "member name" || h.toString().toLowerCase().trim() === "name");

    if (idxId === -1 && idxName === -1) continue;

    for (let r = 1; r < data.length; r++) {
      if (idxName !== -1 && data[r][idxName]) {
        let clean = data[r][idxName].toString().replace(/\(Left Faction\)/ig, "").trim();
        if (clean && !clean.includes("TOTALS")) uniqueMembers.add(clean);
      }
      if (idxId !== -1 && data[r][idxId]) {
        let idVal = data[r][idxId].toString().trim();
        if (idVal) uniqueMembers.add(idVal);
      }
    }
  }

  let memberArray = Array.from(uniqueMembers).sort().map(m => [m]);
  if (memberArray.length > 0) {
    compSheet.getRange("Z:Z").clearContent();
    compSheet.getRange(1, 26, memberArray.length, 1).setValues(memberArray);
    compSheet.hideColumns(26);

    let range = compSheet.getRange("Z1:Z" + memberArray.length);
    let rule = SpreadsheetApp.newDataValidation().requireValueInRange(range, true).setAllowInvalid(true).build();
    compSheet.getRange("B9:B18").setDataValidation(rule);
  }
}