// ==========================================
// 3. OFFICIAL TORN REPORT FETCHER (MULTI-CHAIN)
// ==========================================
function fetchOfficialReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  
  if (!dashSheet || !configSheet) return;
  
  const apiKey = configSheet.getRange(SETTINGS.apiKeyCell).getValue().toString().trim();
  const rawConfigFactionId = configSheet.getRange(SETTINGS.factionIdCell || "B7").getValue().toString().trim();

  if (!apiKey) {
    SpreadsheetApp.getUi().alert("Missing API Key in Config Sheet.");
    return;
  }
  
  let allData = dashSheet.getDataRange().getValues();
  const findVal = (label) => {
    for (let r = 0; r < allData.length; r++) {
      for (let c = 0; c < allData[r].length; c++) {
        if (allData[r][c] === label && c + 1 < allData[r].length) return allData[r][c+1];
      }
    }
    return "";
  };

  const warId = findVal("War Report ID").toString().trim();
  const rawChainId = findVal("Chain Report ID").toString().trim();
  
  if (!warId && !rawChainId) {
    SpreadsheetApp.getUi().alert("Please enter a War ID or a Chain ID in the Official Reports box.");
    return;
  }
  
  // --- CACHE VALUATION ---
  let itemValueMap = {};
  try {
    let itemsUrl = `https://api.torn.com/torn/?selections=items&key=${apiKey}`;
    let itemsRes = UrlFetchApp.fetch(itemsUrl, { muteHttpExceptions: true });
    let itemsJson = JSON.parse(itemsRes.getContentText());
    if (itemsJson.items) {
      for (let id in itemsJson.items) {
        itemValueMap[itemsJson.items[id].name] = itemsJson.items[id].market_value || 0;
      }
    }
  } catch(e) {}
  
  // --- OFFICIAL WAR REPORT ---
  if (warId) {
    let warUrl = `https://api.torn.com/torn/${warId}?selections=rankedwarreport&key=${apiKey}`;
    let res = UrlFetchApp.fetch(warUrl, { muteHttpExceptions: true });
    let json = JSON.parse(res.getContentText());
    
    let warSheetName = "Official War Report";
    let warSheet = ss.getSheetByName(warSheetName);
    if (!warSheet) { warSheet = ss.insertSheet(warSheetName); }
    else { warSheet.clear(); warSheet.clearFormats(); }
    
    if (json.rankedwarreport) {
      let rw = json.rankedwarreport;
      let startUnix = rw.war ? rw.war.start : rw.start;
      let endUnix = rw.war ? rw.war.end : rw.end;
      let start = new Date(startUnix * 1000).toUTCString();
      let end = endUnix ? new Date(endUnix * 1000).toUTCString() : "Ongoing";
      
      let output = [];
      output.push(["⚔️ OFFICIAL WAR REPORT", ""]);
      output.push(["War ID", warId]);
      output.push(["Start (UTC)", start]);
      output.push(["End (UTC)", end]);
      output.push(["(Note: API does not provide Hosp/Mug splits for Wars)", ""]);
      output.push(["", ""]);
      
      output.push(["🏆 FACTION REWARDS", ""]);
      
      let facKeys = Object.keys(rw.factions || {});
      let myFacIdStr = rawConfigFactionId;
      
      if (facKeys.length > 0 && (!myFacIdStr || !facKeys.includes(myFacIdStr))) {
        let dashEnemyId = findVal("Enemy Faction ID").toString().trim(); 
        for (let f of facKeys) {
          if (f !== dashEnemyId) { myFacIdStr = f; break; }
        }
      }

      let outcomeStr = "Ongoing";
      let warEnded = (rw.war && rw.war.end && rw.war.end > 0);
      if (warEnded) {
        if (facKeys.length === 2 && myFacIdStr) {
          let myScore = rw.factions[myFacIdStr].score;
          let enemyId = facKeys[0] === myFacIdStr ? facKeys[1] : facKeys[0];
          let enemyScore = rw.factions[enemyId].score;
          if (myScore > enemyScore) outcomeStr = "Win";
          else if (myScore < enemyScore) outcomeStr = "Loss";
          else outcomeStr = "Draw";
        }
      }
      output.push(["Outcome", outcomeStr]);

      let exactLogString = "";
      try {
        let newsUrl = `https://api.torn.com/faction/?selections=news&key=${apiKey}`;
        let newsRes = UrlFetchApp.fetch(newsUrl, { muteHttpExceptions: true });
        let newsJson = JSON.parse(newsRes.getContentText());
        if (newsJson.news) {
          for (let n in newsJson.news) {
            let text = newsJson.news[n].news || "";
            if (text.includes("received") && text.includes("respect") && (text.includes("ranked") || text.includes("retained"))) {
              exactLogString = text.replace(/(<([^>]+)>)/ig, ''); 
              break; 
            }
          }
        }
      } catch(e) {}
      
      output.push(["Official Faction Log", exactLogString || "Log not found"]);
      
      let cacheList = [];
      let totalCacheValue = 0;
      if (rw.factions && rw.factions[myFacIdStr] && rw.factions[myFacIdStr].rewards) {
        let r = rw.factions[myFacIdStr].rewards;
        output.push(["Points Won", r.points || 0]);
        output.push(["Respect Won", r.respect || 0]);
        
        if (r.items) {
          for (let i in r.items) { 
            let qty = r.items[i].quantity;
            let name = r.items[i].name;
            cacheList.push(`${qty}x ${name}`); 
            if (itemValueMap[name]) {
              totalCacheValue += (qty * itemValueMap[name]);
            }
          }
        }
      }
      output.push(["Caches / Items Won", cacheList.length > 0 ? cacheList.join("\n") : "None"]);
      output.push(["Estimated Cache Value", totalCacheValue > 0 ? totalCacheValue : ""]);
      output.push(["", ""]);
      
      let memberRows = [];
      memberRows.push(["Faction ID", "Faction Name", "Member ID", "Member Name", "Total Attacks", "War Score"]);
      if (rw.factions) {
        for (let f in rw.factions) {
          let fac = rw.factions[f];
          output.push([`Faction: ${fac.name} [${f}]`, `Total Score: ${fac.score}`]);
          if (fac.members) {
            for (let m in fac.members) {
              memberRows.push([f, fac.name, m, fac.members[m].name, fac.members[m].attacks, fac.members[m].score]);
            }
          }
        }
      }
      
      output.push(["", ""]);
      output = output.concat(memberRows);
      
      const maxColsWar = 6;
      for (let r = 0; r < output.length; r++) { while (output[r].length < maxColsWar) { output[r].push(""); } }
      
      warSheet.getRange(1, 1, output.length, maxColsWar).setValues(output);
      warSheet.getRange(1, 1, 1, 2).setBackground("#cc0000").setFontColor("white").setFontWeight("bold");
      warSheet.getRange(7, 1, 1, 2).setBackground("#f1c232").setFontWeight("bold"); 
      warSheet.getRange(output.length - memberRows.length + 1, 1, 1, maxColsWar).setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
      warSheet.autoResizeColumns(1, maxColsWar);
    } else {
      warSheet.getRange("A1").setValue("Error: Could not load War Report. Check ID or API Key.");
    }
  }
  
  // --- OFFICIAL CHAIN REPORT (MULTI-CHAIN) ---
  let chainIds = rawChainId.split(',').map(id => id.trim()).filter(id => id !== "");
  
  if (chainIds.length > 0) {
    let chainSheetName = "Official Chain Report";
    let chainSheet = ss.getSheetByName(chainSheetName);
    if (!chainSheet) { chainSheet = ss.insertSheet(chainSheetName); }
    else { chainSheet.clear(); chainSheet.clearFormats(); }

    if (chainIds.length > 1) {
      ss.toast(`Fetching ${chainIds.length} chains. This may take a moment...`, "Processing", 5);
    }

    let combinedStats = {};
    let totalChainHits = 0;
    let totalChainRespect = 0;
    let firstStart = null;
    let lastEnd = null;
    let apiErrorFlag = false;

    for (let i = 0; i < chainIds.length; i++) {
      let chainId = chainIds[i];
      let chainUrl = `https://api.torn.com/torn/${chainId}?selections=chainreport&key=${apiKey}`;
      
      try {
        let res = UrlFetchApp.fetch(chainUrl, { muteHttpExceptions: true });
        let json = JSON.parse(res.getContentText());
        
        if (json.chainreport) {
          let cr = json.chainreport;
          
          if (firstStart === null || cr.start < firstStart) firstStart = cr.start;
          if (lastEnd === null || cr.end > lastEnd) lastEnd = cr.end;
          
          totalChainHits += (cr.hits || 0);
          totalChainRespect += (cr.respect || 0);

          if (cr.members) {
            for (let m in cr.members) {
              if (!combinedStats[m]) {
                combinedStats[m] = { attacks: 0, respect: 0, leave: 0, mug: 0, hosp: 0, assist: 0, overseas: 0, draw: 0, escape: 0, loss: 0 };
              }
              let mem = cr.members[m];
              combinedStats[m].attacks += mem.attacks || 0;
              combinedStats[m].respect += mem.respect || 0;
              combinedStats[m].leave += mem.leave || 0;
              combinedStats[m].mug += mem.mug || 0;
              combinedStats[m].hosp += mem.hosp || 0;
              combinedStats[m].assist += mem.assist || 0;
              combinedStats[m].overseas += mem.overseas || 0;
              combinedStats[m].draw += mem.draw || 0;
              combinedStats[m].escape += mem.escape || 0;
              combinedStats[m].loss += mem.loss || 0;
            }
          } else {
            apiErrorFlag = true;
          }
        }
      } catch (e) {
        ss.toast(`Error fetching chain ${chainId}`, "API Error", 5);
      }

      // Pause to respect Torn API limits if pulling multiple chains
      if (i < chainIds.length - 1) {
        Utilities.sleep(350); 
      }
    }

    // Build the output for the Chain tab
    let output = [];
    output.push(["⛓️ OFFICIAL CHAIN REPORT", ""]);
    output.push(["Chain ID(s)", chainIds.join(", ")]);
    output.push(["Earliest Start (UTC)", firstStart ? new Date(firstStart * 1000).toUTCString() : "N/A"]);
    output.push(["Latest End (UTC)", lastEnd ? new Date(lastEnd * 1000).toUTCString() : "N/A"]);
    output.push(["Total Hits", totalChainHits]);
    output.push(["Total Respect", totalChainRespect]);
    output.push(["", ""]);
    
    let memberRows = [];
    memberRows.push(["Member ID", "Total Attacks", "Respect", "Leave", "Mug", "Hosp", "Assist", "Overseas", "Draw", "Escape", "Loss"]);
    
    if (apiErrorFlag && Object.keys(combinedStats).length === 0) {
      memberRows.push(["API ERROR:", "Torn hid member data. You MUST use a 'Limited Access' API key.", "", "", "", "", "", "", "", "", ""]);
    } else {
      for (let m in combinedStats) {
        let mem = combinedStats[m];
        memberRows.push([
          m, mem.attacks, mem.respect, mem.leave, mem.mug, mem.hosp, 
          mem.assist, mem.overseas, mem.draw, mem.escape, mem.loss
        ]);
      }
    }
    
    output = output.concat(memberRows);
    
    const maxColsChain = 11;
    for (let r = 0; r < output.length; r++) { while (output[r].length < maxColsChain) { output[r].push(""); } }
    
    chainSheet.getRange(1, 1, output.length, maxColsChain).setValues(output);
    chainSheet.getRange(1, 1, 1, 2).setBackground("#3c78d8").setFontColor("white").setFontWeight("bold");
    chainSheet.getRange(8, 1, 1, maxColsChain).setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
    chainSheet.autoResizeColumns(1, maxColsChain);
  }
  
  if (typeof refreshDashboard === "function") { refreshDashboard(); }
  SpreadsheetApp.getUi().alert("✅ Official Reports Fetched & Combined!");
}

// ==========================================
// DASHBOARD ENGINE (Smart Auto-Fill)
// ==========================================
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);

  if (!dashSheet || !configSheet) {
    SpreadsheetApp.getUi().alert("⚠️ Missing Dashboard or Config sheet.");
    return;
  }

  // Grab API Key and your Faction ID from Config
  const apiKey = configSheet.getRange(SETTINGS.apiKeyCell || "B3").getValue().toString().trim();
  const myFactionId = configSheet.getRange(SETTINGS.factionIdCell || "B7").getValue().toString().trim();

  if (!apiKey) {
    ss.toast("API Key missing. Cannot refresh dashboard.", "Error", 5);
    return;
  }

  ss.toast("Pinging Torn API...", "System", 3);

  let dashData = dashSheet.getDataRange().getValues();

  // --- DYNAMIC DASHBOARD SCANNERS ---
  const findVal = (label) => {
    let cleanLabel = label.toLowerCase().trim();
    for (let r = 0; r < dashData.length; r++) {
      for (let c = 0; c < dashData[r].length; c++) {
        if (dashData[r][c] && dashData[r][c].toString().toLowerCase().trim() === cleanLabel) {
          return dashData[r][c + 1] ? dashData[r][c + 1].toString().trim() : "";
        }
      }
    }
    return "";
  };

  const setVal = (label, value) => {
    let cleanLabel = label.toLowerCase().trim();
    for (let r = 0; r < dashData.length; r++) {
      for (let c = 0; c < dashData[r].length; c++) {
        if (dashData[r][c] && dashData[r][c].toString().toLowerCase().trim() === cleanLabel) {
          dashSheet.getRange(r + 1, c + 2).setValue(value);
          return;
        }
      }
    }
  };

  // Check if you manually typed in a War ID
  let manualWarId = findVal("War Report ID");

  try {
    if (manualWarId) {
      // --- SCENARIO A: USER ENTERED A SPECIFIC WAR ID ---
      let url = `https://api.torn.com/torn/${manualWarId}?selections=rankedwarreport&key=${apiKey}`;
      let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      let json = JSON.parse(response.getContentText());

      if (json.error) {
        SpreadsheetApp.getUi().alert("Torn API Error on War ID: " + json.error.error);
        return;
      }

      if (json.rankedwarreport && json.rankedwarreport.factions) {
        let factions = json.rankedwarreport.factions;
        let enemyFound = false;

        // Find the faction that ISN'T you
        for (let id in factions) {
          if (id !== myFactionId) {
            setVal("Enemy Faction ID", id);
            setVal("Enemy Faction Name", factions[id].name);
            enemyFound = true;
          } else {
            setVal("War Score", factions[id].score);
          }
        }

        // --- NEW: DETERMINE WAR OUTCOME (Hardcoded to C10) ---
        if (json.rankedwarreport.war && json.rankedwarreport.war.winner !== undefined) {
          let winnerId = json.rankedwarreport.war.winner.toString();
          
          if (winnerId === myFactionId) {
            dashSheet.getRange("C10").setValue("Win");
          } else if (winnerId === "0") {
            dashSheet.getRange("C10").setValue("Draw");
          } else {
            dashSheet.getRange("C10").setValue("Loss");
          }
        } else {
          dashSheet.getRange("C10").setValue("Finished"); // Fallback if API acts weird
        }

        if (enemyFound) {
          ss.toast("✅ Enemy details & War Outcome auto-filled from War ID!", "Success", 5);
        } else {
          SpreadsheetApp.getUi().alert(`⚠️ Could not find an enemy faction in War [${manualWarId}]. Make sure Config B7 is your correct Faction ID.`);
        }
      } else {
        SpreadsheetApp.getUi().alert(`⚠️ War ID [${manualWarId}] is invalid or API did not return war data.`);
      }

    } else {
      // --- SCENARIO B: NO WAR ID (LOOK FOR ACTIVE WAR) ---
      let url = `https://api.torn.com/faction/?selections=basic&key=${apiKey}`;
      let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      let json = JSON.parse(response.getContentText());

      if (json.error) {
         SpreadsheetApp.getUi().alert("Torn API Error: " + json.error.error);
         return;
      }

      // Check if you are currently in an active Ranked War
      if (json.ranked_wars && Object.keys(json.ranked_wars).length > 0) {
        let activeWarId = Object.keys(json.ranked_wars)[0];
        let warData = json.ranked_wars[activeWarId];

        setVal("War Report ID", activeWarId);
        setVal("War ID", activeWarId); 
        
        // --- NEW: MARK ACTIVE WARS AS ONGOING (Hardcoded to C10) ---
        dashSheet.getRange("C10").setValue("Ongoing");

        let factions = warData.factions;
        for (let id in factions) {
          if (id !== myFactionId) {
            setVal("Enemy Faction ID", id);
            setVal("Enemy Faction Name", factions[id].name);
          } else {
            setVal("War Score", factions[id].score);
          }
        }
        ss.toast("✅ Dashboard updated with live active Ranked War!", "Success", 5);
      } else {
        ss.toast("No War ID entered and no active Ranked War found.", "System", 3);
      }
    }
  } catch (e) {
    ss.toast("Failed to connect to Torn API.", "Error", 5);
  }
}

// (Dummy script to catch broken menu links)
function buildDashboard() {
  SpreadsheetApp.getUi().alert("Your dashboard layout is already built! Just use 'Refresh Dashboard' to pull the latest API data.");
}