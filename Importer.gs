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
  let startDateVal = dashSheet.getRange("F9").getValue();
  let endDateVal = dashSheet.getRange("F11").getValue();
  let startTimeStr = dashSheet.getRange("F10").getDisplayValue();
  let endTimeStr = dashSheet.getRange("F12").getDisplayValue();

  // --- IRONCLAD TIME FORCER ---
  let sd = new Date(startDateVal);
  let stParts = (startTimeStr ? startTimeStr : "00:00:00").split(":");
  
  let sHour = parseInt(stParts[0]); if (isNaN(sHour)) sHour = 0;
  let sMin = parseInt(stParts[1]); if (isNaN(sMin)) sMin = 0;
  let sSec = parseInt(stParts[2]); if (isNaN(sSec)) sSec = 0;

  let startUnix = Date.UTC(
    sd.getUTCFullYear(), sd.getUTCMonth(), sd.getUTCDate(),
    sHour, sMin, sSec
  );

  let ed = new Date(endDateVal);
  let etParts = (endTimeStr ? endTimeStr : "23:59:59").split(":");
  
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

  // ---> THE FIX: Dynamically grab the new API v2 headers <---
  const masterHeaders = rawData[0];

  // 3. Filter the Data (With Duplicate Shield)
  const filteredData = [];
  const seenIds = new Set();
  let duplicatesBlocked = 0;
  
  for (let i = 1; i < rawData.length; i++) {
    let row = rawData[i];
    let attackId = row[0].toString().trim(); 
    
    let rowTimeUnix;
    if (row[2] instanceof Date) {
      rowTimeUnix = row[2].getTime();
    } else {
      rowTimeUnix = new Date(row[2] + " UTC").getTime();
    }

    if (rowTimeUnix >= startUnix && rowTimeUnix <= endUnix) {
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
  
  // ---> THE FIX: Push the dynamic headers back on top <---
  filteredData.unshift(masterHeaders);

  // 4. Overwrite Local RD Tab
  if (!localRdSheet) {
    localRdSheet = ss.insertSheet(SETTINGS.rdSheet);
  } else {
    localRdSheet.clear();
  }
  
  if (filteredData.length > 1) {
    localRdSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
    
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
    // ---> THE FIX: Write dynamic headers even if empty <---
    localRdSheet.getRange(1, 1, 1, masterHeaders.length).setValues([masterHeaders]);
    ss.toast("No hits found in that time window.", "System", 5);
  }
}