// ==========================================
// 1. MASTER IMPORTER (Timezone-Safe + Duplicate Shield)
// ==========================================
function importWarData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  let localRdSheet = ss.getSheetByName(SETTINGS.rdSheet);

  if (!dashSheet) {
    ss.toast("Please run 'Rebuild Dashboard' first.", "Error", 5);
    return;
  }

  const backendId = configSheet.getRange(SETTINGS.backendIdCell).getValue().toString().trim();
  if (!backendId) {
    ss.toast("Missing Backend Sheet ID in Config B2!", "Error", 5);
    return;
  }

  // 1. Read Time Filters from Dashboard
  let startDateVal = dashSheet.getRange("I8").getValue();
  let endDateVal = dashSheet.getRange("I10").getValue();
  
  // USE GET DISPLAY VALUE: Grabs exactly what is on the screen (e.g., "15:00:00")
  let startTimeStr = dashSheet.getRange("I9").getDisplayValue();
  let endTimeStr = dashSheet.getRange("I11").getDisplayValue();

  // --- IRONCLAD TIME FORCER ---
  let sd = new Date(startDateVal);
  let stParts = (startTimeStr ? startTimeStr : "00:00:00").split(":");
  
  // Safely parse start times without overriding zeroes
  let sHour = parseInt(stParts[0]); if (isNaN(sHour)) sHour = 0;
  let sMin = parseInt(stParts[1]); if (isNaN(sMin)) sMin = 0;
  let sSec = parseInt(stParts[2]); if (isNaN(sSec)) sSec = 0;

  let startUnix = Date.UTC(
    sd.getUTCFullYear(), sd.getUTCMonth(), sd.getUTCDate(),
    sHour, sMin, sSec
  );

  let ed = new Date(endDateVal);
  let etParts = (endTimeStr ? endTimeStr : "23:59:59").split(":");
  
  // Safely parse end times without overriding zeroes
  let eHour = parseInt(etParts[0]); if (isNaN(eHour)) eHour = 23;
  let eMin = parseInt(etParts[1]); if (isNaN(eMin)) eMin = 59;
  let eSec = parseInt(etParts[2]); if (isNaN(eSec)) eSec = 59;

  let endUnix = Date.UTC(
    ed.getUTCFullYear(), ed.getUTCMonth(), ed.getUTCDate(),
    eHour, eMin, eSec, 999 
  );

  // 2. Connect to Backend Warehouse
  ss.toast("Connecting to Data Warehouse...", "System", 3);
  let backendSS;
  try {
    backendSS = SpreadsheetApp.openById(backendId);
  } catch(e) {
    ss.toast("Could not access Backend Sheet.", "Error", 5);
    return;
  }

  const backendDataSheet = backendSS.getSheetByName("Data"); 
  if (!backendDataSheet) return;

  const rawData = backendDataSheet.getDataRange().getValues();
  if (rawData.length <= 1) return;

  // 3. Filter the Data (With Duplicate Shield)
  const filteredData = [];
  const seenIds = new Set();
  let duplicatesBlocked = 0;
  
  for (let i = 1; i < rawData.length; i++) {
    let row = rawData[i];
    let attackId = row[0].toString().trim(); 
    
    // Force the Master Sheet's timestamp to be read accurately
    let rowTimeUnix;
    if (row[2] instanceof Date) {
      rowTimeUnix = row[2].getTime();
    } else {
      rowTimeUnix = new Date(row[2] + " UTC").getTime(); // Force string to UTC
    }

    // Check if within our Ironclad time window
    if (rowTimeUnix >= startUnix && rowTimeUnix <= endUnix) {
      
      // Duplicate Shield Check
      if (seenIds.has(attackId)) {
        duplicatesBlocked++;
        continue; 
      }

      filteredData.push(row);
      seenIds.add(attackId);
    }
  }

  // --- SORTING LOGIC ---
  filteredData.sort((a, b) => {
    let timeA = (a[2] instanceof Date) ? a[2].getTime() : new Date(a[2] + " UTC").getTime();
    let timeB = (b[2] instanceof Date) ? b[2].getTime() : new Date(b[2] + " UTC").getTime();
    return timeA - timeB;
  });
  
  // Add headers back to the top
  filteredData.unshift(SETTINGS.rdHeaders);

  // 4. Overwrite Local RD Tab
  if (!localRdSheet) {
    localRdSheet = ss.insertSheet(SETTINGS.rdSheet);
  } else {
    localRdSheet.clear();
  }
  
  if (filteredData.length > 1) {
    localRdSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
    
    // Human-Readable Boundary Receipt
    let startString = new Date(startUnix).toUTCString();
    let endString = new Date(endUnix).toUTCString();
    
    let msg = `IMPORT COMPLETE\n\n` +
              `Imported ${filteredData.length - 1} unique attacks.\n`;
    
    if (duplicatesBlocked > 0) {
      msg += `Blocked ${duplicatesBlocked} duplicates found in Master Data.\n\n`;
    } else {
      msg += `\n`;
    }

    msg += `--- TIME BOUNDARIES USED ---\n` +
           `Start: ${startString}\n` +
           `End:   ${endString}\n\n` +
           `If the number of hits is wrong, check these exact boundary times!`;
           
    SpreadsheetApp.getUi().alert(msg);
    
  } else {
    localRdSheet.getRange(1, 1, 1, SETTINGS.rdHeaders.length).setValues([SETTINGS.rdHeaders]);
    ss.toast("No hits found in that time window.", "System", 5);
  }
}
// ==========================================
// 2. SNIPER AUDIT UTILITY (Diagnose Missing Hits)
// ==========================================
function runSniperAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);

  if (!configSheet || !dashSheet) return;

  // 1. Ask for the specific Member ID
  const promptResponse = ui.prompt(
    "Sniper Audit", 
    "Enter the Torn ID of the member who is missing hits:", 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (promptResponse.getSelectedButton() !== ui.Button.OK) return;
  const targetId = promptResponse.getResponseText().trim();
  if (targetId === "") return;

  // 2. Setup Time Boundaries
  let startDateVal = dashSheet.getRange("I8").getValue();
  let endDateVal = dashSheet.getRange("I10").getValue();
  let startTimeStr = dashSheet.getRange("I9").getDisplayValue();
  let endTimeStr = dashSheet.getRange("I11").getDisplayValue();

  let sd = new Date(startDateVal);
  let stParts = (startTimeStr ? startTimeStr : "00:00:00").split(":");
  let startUnix = Date.UTC(sd.getUTCFullYear(), sd.getUTCMonth(), sd.getUTCDate(), parseInt(stParts[0])||0, parseInt(stParts[1])||0, parseInt(stParts[2])||0);

  let ed = new Date(endDateVal);
  let etParts = (endTimeStr ? endTimeStr : "23:59:59").split(":");
  let endUnix = Date.UTC(ed.getUTCFullYear(), ed.getUTCMonth(), ed.getUTCDate(), parseInt(etParts[0])||23, parseInt(etParts[1])||59, parseInt(etParts[2])||59, 999);

  // 3. Connect to Master Backend
  const backendId = configSheet.getRange(SETTINGS.backendIdCell).getValue().toString().trim();
  let backendSS = SpreadsheetApp.openById(backendId);
  let masterData = backendSS.getSheetByName("Data").getDataRange().getValues();

  // 4. Track exactly what happened
  let log = {
    tooEarly: 0,
    tooLate: 0,
    defending: 0,
    resultsBreakdown: {} // NEW: We will tally every exact word here
  };

  for (let i = 1; i < masterData.length; i++) {
    let row = masterData[i];
    let attackerId = row[3].toString().trim();
    let defenderId = row[6].toString().trim();
    let result = row[9] ? row[9].toString().toLowerCase().trim() : "blank/unknown";
    
    let rowTimeUnix = (row[2] instanceof Date) ? row[2].getTime() : new Date(row[2] + " UTC").getTime();

    // Check if they were the ATTACKER
    if (attackerId === targetId) {
      if (rowTimeUnix < startUnix) {
        if (startUnix - rowTimeUnix < 3600000) log.tooEarly++; 
      } 
      else if (rowTimeUnix > endUnix) {
        if (rowTimeUnix - endUnix < 3600000) log.tooLate++;
      } 
      else {
        // INSIDE THE WINDOW: Itemize exactly what the result was
        if (!log.resultsBreakdown[result]) {
          log.resultsBreakdown[result] = 0;
        }
        log.resultsBreakdown[result]++;
      }
    }
    
    // Check if they were the DEFENDER
    if (defenderId === targetId && rowTimeUnix >= startUnix && rowTimeUnix <= endUnix) {
      log.defending++;
    }
  }

  // 5. Build the Detailed Report
  let report = `🎯 SNIPER REPORT FOR ID: ${targetId} 🎯\n\n`;
  report += `--- EXACT RESULTS INSIDE TIME WINDOW ---\n`;
  
  let totalInWindow = 0;
  let sortedResults = Object.keys(log.resultsBreakdown).sort((a, b) => log.resultsBreakdown[b] - log.resultsBreakdown[a]);
  
  if (sortedResults.length === 0) {
    report += `No attacks found inside the exact time window.\n`;
  } else {
    for (let r of sortedResults) {
      report += `• ${r}: ${log.resultsBreakdown[r]}\n`;
      totalInWindow += log.resultsBreakdown[r];
    }
    report += `\nTotal hits logged in window: ${totalInWindow}\n\n`;
  }
  
  report += `--- OUTSIDE WINDOW / OTHER ---\n`;
  report += `❌ Hits just BEFORE start time: ${log.tooEarly}\n`;
  report += `❌ Hits just AFTER end time: ${log.tooLate}\n`;
  report += `🛡️ Times they were attacked: ${log.defending}\n\n`;
  
  report += `Compare the 'Total hits logged' above to what your Payout Sheet says!`;

  ui.alert(report);
}

