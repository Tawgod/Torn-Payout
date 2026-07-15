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
    let cleanLabel = label.toLowerCase().trim();
    for (let r = 0; r < allData.length; r++) {
      for (let c = 0; c < allData[r].length; c++) {
        let cellText = allData[r][c] ? allData[r][c].toString().toLowerCase().trim() : "";
        if (cellText === cleanLabel && c + 1 < allData[r].length) {
          return allData[r][c+1];
        }
      }
    }
    return "";
  };

  const setVal = (label, value) => {
    let cleanLabel = label.toLowerCase().trim();
    for (let r = 0; r < allData.length; r++) {
      for (let c = 0; c < allData[r].length; c++) {
        let cellText = allData[r][c] ? allData[r][c].toString().toLowerCase().trim() : "";
        if (cellText === cleanLabel) {
          dashSheet.getRange(r + 1, c + 2).setValue(value);
          return;
        }
      }
    }
  };

  const formatTornDate = (unix) => {
    if (!unix || unix === 0) return "Ongoing";
    let d = new Date(unix * 1000);
    return Utilities.formatDate(d, "GMT", "yyyy-MM-dd HH:mm"); 
  };

  const warId = findVal("war report id").toString().trim();
  const rawChainId = findVal("chain report id").toString().trim();
  
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
    ss.toast("Fetching War Report...", "System", 3);
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
      
      let startStr = formatTornDate(startUnix);
      let endStr = formatTornDate(endUnix);
      
      let output = [];
      output.push(["⚔️ OFFICIAL WAR REPORT", ""]);
      output.push(["War ID", warId]);
      output.push(["Start (UTC)", startStr]);
      output.push(["End (UTC)", endStr]);
      output.push(["(Note: API does not provide Hosp/Mug splits for Wars)", ""]);
      output.push(["", ""]);
      
      output.push(["🏆 FACTION REWARDS", ""]);
      let facKeys = Object.keys(rw.factions || {});
      let myFacIdStr = rawConfigFactionId;
      
      if (facKeys.length > 0 && (!myFacIdStr || !facKeys.includes(myFacIdStr))) {
        let dashEnemyId = findVal("enemy faction id").toString().trim();
        for (let f of facKeys) {
          if (f !== dashEnemyId) { myFacIdStr = f; break; }
        }
      }

      let outcomeStr = "Ongoing";
      let warEnded = (rw.war && rw.war.end && rw.war.end > 0);
      let myTotalWarHits = 0; 
      
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
      if (rw.factions && rw.factions[myFacIdStr]) {
        if (rw.factions[myFacIdStr].members) {
           for (let m in rw.factions[myFacIdStr].members) {
              myTotalWarHits += (rw.factions[myFacIdStr].members[m].attacks || 0);
           }
        }

        if (rw.factions[myFacIdStr].rewards) {
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
      }
      
      let cacheString = cacheList.length > 0 ? cacheList.join(", ") : "None";
      output.push(["Caches / Items Won", cacheString]);
      output.push(["Estimated Cache Value", totalCacheValue > 0 ? totalCacheValue : ""]);
      
      dashSheet.getRange("C5").setNumberFormat("@").setValue(startStr);       
      dashSheet.getRange("C6").setNumberFormat("@").setValue(endStr);   
      dashSheet.getRange("C7").setNumberFormat("@").setValue(warId);      
      
      setVal("caches / items won", cacheString);
      setVal("total war hits", myTotalWarHits);
      
      if (totalCacheValue > 0) {
        setVal("est. cache value", totalCacheValue);
        setVal("actual cache value", totalCacheValue);
      }

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
      warSheet.getRange("B3:B4").setNumberFormat("@");
    }
  }
  
  // --- OFFICIAL CHAIN REPORT & RD TAB EXTRACTION ---
  let chainIds = rawChainId.split(',').map(id => id.trim()).filter(id => id !== "");
  let usedOfficialChain = false;

  let finalChainStart = "N/A";
  let finalChainEnd = "N/A";
  let finalChainHits = 0;
  let finalChainRespect = 0;
  
  if (chainIds.length > 0) {
    let chainSheetName = "Official Chain Report";
    let chainSheet = ss.getSheetByName(chainSheetName);
    if (!chainSheet) { chainSheet = ss.insertSheet(chainSheetName); }
    else { chainSheet.clear(); chainSheet.clearFormats(); }

    if (chainIds.length > 1) {
      ss.toast(`Fetching ${chainIds.length} chains...`, "Processing", 5);
    }

    let combinedStats = {};
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
          
          finalChainRespect += (cr.respect || 0);

          if (cr.members) {
            for (let m in cr.members) {
              if (!combinedStats[m]) combinedStats[m] = { attacks: 0, respect: 0, leave: 0, mug: 0, hosp: 0, assist: 0, overseas: 0, draw: 0, escape: 0, loss: 0, retal: 0 };
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
              combinedStats[m].retal += mem.retal || 0; 
              
              finalChainHits += mem.attacks || 0;
            }
          } else {
            apiErrorFlag = true;
          }
        }
      } catch (e) {}
      if (i < chainIds.length - 1) Utilities.sleep(350);
    }

    if (finalChainHits > 0) {
      usedOfficialChain = true;
      finalChainStart = formatTornDate(firstStart);
      finalChainEnd = formatTornDate(lastEnd);

      // ---> THE FIX: Cleanse the API's 'Total Respect Generated' display by subtracting Chain Bonuses
      let rdSheetForBonus = ss.getSheetByName(SETTINGS.rdSheet || "RD");
      if (rdSheetForBonus && rdSheetForBonus.getLastRow() > 1) {
        let rdData = rdSheetForBonus.getDataRange().getValues();
        let totalBonusRespectToStrip = 0;
        for (let i = 1; i < rdData.length; i++) {
          let aFac = rdData[i][5] ? rdData[i][5].toString().replace(/,/g, "").trim() : "";
          if (aFac === rawConfigFactionId) {
            let respect = parseFloat(rdData[i][10]) || 0;
            let cBonus = parseFloat(rdData[i][16]) || 1;
            if (cBonus > 1) {
               totalBonusRespectToStrip += (respect - (respect / cBonus));
            }
          }
        }
        finalChainRespect -= totalBonusRespectToStrip;
        if (finalChainRespect < 0) finalChainRespect = 0;
      }

      let output = [];
      output.push(["⛓️ OFFICIAL CHAIN REPORT", ""]);
      output.push(["Chain ID(s)", chainIds.join(", ")]);
      output.push(["Earliest Start (UTC)", finalChainStart]);
      output.push(["Latest End (UTC)", finalChainEnd]);
      output.push(["Total Hits", "=SUM(B9:B)"]); 
      output.push(["Total Respect", finalChainRespect]);
      output.push(["", ""]);
      
      let memberRows = [];
      memberRows.push(["Member ID", "Total Attacks", "Respect", "Leave", "Mug", "Hosp", "Assist", "Overseas", "Draw", "Escape", "Loss", "Retal"]);
      
      if (apiErrorFlag && Object.keys(combinedStats).length === 0) {
        memberRows.push(["API ERROR:", "Torn hid member data.", "", "", "", "", "", "", "", "", "", ""]);
      } else {
        for (let m in combinedStats) {
          let mem = combinedStats[m];
          memberRows.push([m, mem.attacks, mem.respect, mem.leave, mem.mug, mem.hosp, mem.assist, mem.overseas, mem.draw, mem.escape, mem.loss, mem.retal]);
        }
      }
      
      output = output.concat(memberRows);
      
      const maxColsChain = 12;
      for (let r = 0; r < output.length; r++) { while (output[r].length < maxColsChain) { output[r].push(""); } }
      
      chainSheet.getRange(1, 1, output.length, maxColsChain).setValues(output);
      chainSheet.getRange(1, 1, 1, 2).setBackground("#3c78d8").setFontColor("white").setFontWeight("bold");
      chainSheet.getRange(8, 1, 1, maxColsChain).setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
      chainSheet.autoResizeColumns(1, maxColsChain);
      chainSheet.getRange("B3:B4").setNumberFormat("@");
    }
  }
  
  // FALLBACK: Read directly from the RD Tab
  if (!usedOfficialChain) {
    let rdSheet = ss.getSheetByName(SETTINGS.rdSheet);
    if (rdSheet && rdSheet.getLastRow() > 1) {
      let rdData = rdSheet.getDataRange().getValues();
      let rdFirst = null; let rdLast = null; 
      let rdTotalAttacks = 0; let rdTotalRespect = 0;
      
      let cleanId = (val) => (val === null || val === undefined) ? "" : val.toString().replace(/,/g, "").trim();
      
      for (let i = 1; i < rdData.length; i++) {
        let tVal = rdData[i][2]; // Column C (Timestamp)
        let aFac = cleanId(rdData[i][5]); // Column F (Attacker Faction ID)
        let respect = parseFloat(rdData[i][10]) || 0; // Column K (Respect)
        let cBonus = parseFloat(rdData[i][16]) || 1; // Column Q (Chain Bonus)
        
        // Strip the Chain Bonus mathematically 
        let adjRespect = respect / (cBonus > 1 ? cBonus : 1);

        if (aFac !== "" && aFac === rawConfigFactionId) {
          rdTotalAttacks++;
          rdTotalRespect += adjRespect;
          let tMs = new Date(tVal).getTime();
          if (!isNaN(tMs)) {
            if (rdFirst === null || tMs < rdFirst) rdFirst = tMs;
            if (rdLast === null || tMs > rdLast) rdLast = tMs;
          }
        }
      }
      
      const formatTornDateMs = (ms) => {
        if (!ms) return "N/A";
        return Utilities.formatDate(new Date(ms), "GMT", "yyyy-MM-dd HH:mm");
      };
      
      if (rdTotalAttacks > 0) {
        finalChainStart = rdFirst ? formatTornDateMs(rdFirst) : "N/A";
        finalChainEnd = rdLast ? formatTornDateMs(rdLast) : "N/A";
        finalChainHits = rdTotalAttacks;
        finalChainRespect = rdTotalRespect;
      }
    }
  }

  // HARDCODED DASHBOARD PUSH
  dashSheet.getRange("C16").setNumberFormat("@").setValue(finalChainStart);       
  dashSheet.getRange("C17").setNumberFormat("@").setValue(finalChainEnd);         
  dashSheet.getRange("C18").setValue(finalChainHits);        
  dashSheet.getRange("C19").setValue(finalChainRespect);     
  
  if (typeof refreshDashboard === "function") { refreshDashboard(); }
  SpreadsheetApp.getUi().alert("✅ Official Reports Fetched & Insights Populated!");
}

// ==========================================
// DASHBOARD ENGINE (Smart Auto-Fill)
// ==========================================
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);

  if (!dashSheet || !configSheet) return;

  const apiKey = configSheet.getRange(SETTINGS.apiKeyCell || "B3").getValue().toString().trim();
  const myFactionId = configSheet.getRange(SETTINGS.factionIdCell || "B7").getValue().toString().trim();

  if (!apiKey) return;

  ss.toast("Pinging Torn API & Scanning Data...", "System", 3);

  let dashData = dashSheet.getDataRange().getValues();

  const findVal = (label) => {
    let cleanLabel = label.toLowerCase().trim();
    for (let r = 0; r < dashData.length; r++) {
      for (let c = 0; c < dashData[r].length; c++) {
        let cellText = dashData[r][c] ? dashData[r][c].toString().toLowerCase().trim() : "";
        if (cellText === cleanLabel && c + 1 < dashData[r].length) {
          return dashData[r][c+1];
        }
      }
    }
    return "";
  };

  const setVal = (label, value) => {
    let cleanLabel = label.toLowerCase().trim();
    for (let r = 0; r < dashData.length; r++) {
      for (let c = 0; c < dashData[r].length; c++) {
        let cellText = dashData[r][c] ? dashData[r][c].toString().toLowerCase().trim() : "";
        if (cellText === cleanLabel) {
          dashSheet.getRange(r + 1, c + 2).setValue(value);
          return;
        }
      }
    }
  };

  const formatTornDate = (unix) => {
    if (!unix || unix === 0) return "Ongoing";
    let d = new Date(unix * 1000);
    return Utilities.formatDate(d, "GMT", "yyyy-MM-dd HH:mm"); 
  };

  let manualWarId = findVal("war report id");

  // 1. UPDATE ACTIVE WAR DATA
  try {
    if (manualWarId) {
      let url = `https://api.torn.com/torn/${manualWarId}?selections=rankedwarreport&key=${apiKey}`;
      let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      let json = JSON.parse(response.getContentText());

      if (json.rankedwarreport && json.rankedwarreport.factions) {
        let factions = json.rankedwarreport.factions;
        for (let id in factions) {
          if (id !== myFactionId) {
            setVal("enemy faction id", id);
            setVal("enemy faction name", factions[id].name);
          } else {
            setVal("war score", factions[id].score);
          }
        }

        let startUnix = json.rankedwarreport.war ? json.rankedwarreport.war.start : json.rankedwarreport.start;
        let endUnix = json.rankedwarreport.war ? json.rankedwarreport.war.end : json.rankedwarreport.end;
        
        dashSheet.getRange("C5").setNumberFormat("@").setValue(formatTornDate(startUnix));
        dashSheet.getRange("C6").setNumberFormat("@").setValue(formatTornDate(endUnix));
        dashSheet.getRange("C7").setNumberFormat("@").setValue(manualWarId);

        if (factions[myFactionId] && factions[myFactionId].rewards && factions[myFactionId].rewards.items) {
          let myItems = factions[myFactionId].rewards.items;
          let cacheArr = [];
          for (let i in myItems) { cacheArr.push(`${myItems[i].quantity}x ${myItems[i].name}`); }
          setVal("caches / items won", cacheArr.length > 0 ? cacheArr.join(", ") : "None");
        }

        if (json.rankedwarreport.war && json.rankedwarreport.war.winner !== undefined) {
          let winnerId = json.rankedwarreport.war.winner.toString();
          if (winnerId === myFactionId) dashSheet.getRange("C10").setValue("Win");
          else if (winnerId === "0") dashSheet.getRange("C10").setValue("Draw");
          else dashSheet.getRange("C10").setValue("Loss");
        } else {
          dashSheet.getRange("C10").setValue("Finished");
        }
      }
    } else {
      let url = `https://api.torn.com/faction/?selections=basic&key=${apiKey}`;
      let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      let json = JSON.parse(response.getContentText());

      if (json.ranked_wars && Object.keys(json.ranked_wars).length > 0) {
        let activeWarId = Object.keys(json.ranked_wars)[0];
        let warData = json.ranked_wars[activeWarId];

        setVal("war report id", activeWarId);
        dashSheet.getRange("C7").setNumberFormat("@").setValue(activeWarId); 
        dashSheet.getRange("C10").setValue("Ongoing");

        let factions = warData.factions;
        for (let id in factions) {
          if (id !== myFactionId) {
            setVal("enemy faction id", id);
            setVal("enemy faction name", factions[id].name);
          } else {
            setVal("war score", factions[id].score);
          }
        }
        
        if (warData.war && warData.war.start) {
            dashSheet.getRange("C5").setNumberFormat("@").setValue(formatTornDate(warData.war.start));
        }
      }
    }
  } catch (e) {}

  // 2. CHAIN INSIGHTS: Pull directly from Official Chain Report Tab first
  let oChainSheet = ss.getSheetByName("Official Chain Report");
  let chainInsightsSet = false;

  if (oChainSheet && oChainSheet.getLastRow() >= 6) {
    let b3 = oChainSheet.getRange("B3").getDisplayValue().trim();
    let b4 = oChainSheet.getRange("B4").getDisplayValue().trim();
    let b5 = oChainSheet.getRange("B5").getValue();
    let b6 = oChainSheet.getRange("B6").getValue();

    if (b3 !== "" && b3 !== "N/A" && b5 !== "") {
      let displayRespect = parseFloat(b6) || 0;
      let rdSheetForAdjust = ss.getSheetByName(SETTINGS.rdSheet || "RD");
      if (rdSheetForAdjust && rdSheetForAdjust.getLastRow() > 1) {
        let rdData = rdSheetForAdjust.getDataRange().getValues();
        for (let r = 1; r < rdData.length; r++) {
          let aFac = rdData[r][5] ? rdData[r][5].toString().replace(/,/g, "").trim() : "";
          if (aFac === myFactionId) {
            let respect = parseFloat(rdData[r][10]) || 0;
            let cBonus = parseFloat(rdData[r][16]) || 1;
            if (cBonus > 1) {
              displayRespect -= (respect - (respect / cBonus));
            }
          }
        }
      }
      if (displayRespect < 0) displayRespect = 0;
      dashSheet.getRange("C16").setNumberFormat("@").setValue(b3);
      dashSheet.getRange("C17").setNumberFormat("@").setValue(b4);
      dashSheet.getRange("C18").setValue(b5);
      dashSheet.getRange("C19").setValue(displayRespect);
      chainInsightsSet = true;
    }
  }

  // 3. FALLBACK: Live extraction from RD tab 
  if (!chainInsightsSet) {
    let rdSheet = ss.getSheetByName(SETTINGS.rdSheet);
    if (rdSheet && rdSheet.getLastRow() > 1) {
      let rdData = rdSheet.getDataRange().getValues();
      let rdFirst = null; let rdLast = null; 
      let rdTotalAttacks = 0; let rdTotalRespect = 0;
      let cleanId = (val) => (val === null || val === undefined) ? "" : val.toString().replace(/,/g, "").trim();
      
      for (let i = 1; i < rdData.length; i++) {
        let tVal = rdData[i][2]; // Timestamp
        let aFac = cleanId(rdData[i][5]); 
        let respect = parseFloat(rdData[i][10]) || 0;
        let cBonus = parseFloat(rdData[i][16]) || 1;
        
        let adjRespect = respect / (cBonus > 1 ? cBonus : 1);
        if (aFac !== "" && aFac === myFactionId) {
          rdTotalAttacks++;
          rdTotalRespect += adjRespect;
          let tMs = new Date(tVal).getTime();
          if (!isNaN(tMs)) {
            if (rdFirst === null || tMs < rdFirst) rdFirst = tMs;
            if (rdLast === null || tMs > rdLast) rdLast = tMs;
          }
        }
      }
      
      const formatTornDateMs = (ms) => {
        if (!ms) return "N/A";
        return Utilities.formatDate(new Date(ms), "GMT", "yyyy-MM-dd HH:mm");
      };
      
      if (rdTotalAttacks > 0) {
        dashSheet.getRange("C16").setNumberFormat("@").setValue(rdFirst ? formatTornDateMs(rdFirst) : "N/A");
        dashSheet.getRange("C17").setNumberFormat("@").setValue(rdLast ? formatTornDateMs(rdLast) : "N/A"); 
        dashSheet.getRange("C18").setValue(rdTotalAttacks); 
        dashSheet.getRange("C19").setValue(rdTotalRespect); 
      }
    }
  }

  // ==========================================
  // 4. LEADERBOARDS (Avg Respect & Top Contribution)
  // ==========================================
  let payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  if (payoutSheet && payoutSheet.getLastRow() >= 3) {
    let pData = payoutSheet.getRange(3, 1, payoutSheet.getLastRow() - 2, 15).getValues();
    let respectData = [];
    let contribData = [];

    for (let row of pData) {
      let name = row[1];
      let contrib = parseFloat(row[2]) || 0;
      let wh = parseFloat(row[4]) || 0;
      let ch = parseFloat(row[8]) || 0;
      let respect = parseFloat(row[11]) || 0;
      let totalHits = wh + ch;
      let safeName = name ? name.toString().toLowerCase().trim() : "";
      
      if (safeName && !safeName.includes("left faction") && safeName !== "totals" && safeName !== "total") {
        contribData.push({name: name, val: contrib});
        if (totalHits > 0) {
          respectData.push({name: name, val: respect / totalHits});
        }
      }
    }

    respectData.sort((a, b) => b.val - a.val);
    contribData.sort((a, b) => b.val - a.val);

    // --- Print Top 4 Avg Respect per Hit ---
    dashSheet.getRange("H16:I19").clearContent(); // Safely clear content, leave borders alone
    let respectOutput = [];
    for (let i = 0; i < Math.min(4, respectData.length); i++) {
      respectOutput.push([respectData[i].name, respectData[i].val]);
    }
    if (respectOutput.length > 0) {
      dashSheet.getRange(16, 8, respectOutput.length, 2).setValues(respectOutput);
    }
    dashSheet.getRange("I16:I19").setNumberFormat("0.00");

    // --- Print Top 8 Contribution (Shifted to row 24 for new layout) ---
    dashSheet.getRange("H24:I31").clearContent(); 
    let contribOutput = [];
    for (let i = 0; i < Math.min(8, contribData.length); i++) {
      contribOutput.push([contribData[i].name, contribData[i].val]);
    }
    if (contribOutput.length > 0) {
      dashSheet.getRange(24, 8, contribOutput.length, 2).setValues(contribOutput);
    }
    dashSheet.getRange("I24:I31").setNumberFormat("0.00%");
  }

  // ==========================================
  // 5. BONUS CHAIN HITS LEADERBOARD (FIXED MAP & SAFE FORMAT)
  // ==========================================
  let rdSheetBonus = ss.getSheetByName(SETTINGS.rdSheet || "RD");
  let bonusMap = new Map();
  
  if (rdSheetBonus && rdSheetBonus.getLastRow() > 1) {
    let rdBonusData = rdSheetBonus.getDataRange().getValues();
    for (let r = 1; r < rdBonusData.length; r++) {
      let aFac = rdBonusData[r][5] ? rdBonusData[r][5].toString().replace(/,/g, "").trim() : "";
      let aName = rdBonusData[r][4] ? rdBonusData[r][4].toString().trim() : "Unknown";
      let cBonus = parseFloat(rdBonusData[r][16]) || 0;
      if (aFac === myFactionId && cBonus >= 10) {
        if (!bonusMap.has(aName) || cBonus > bonusMap.get(aName)) {
            bonusMap.set(aName, cBonus);
        }
      }
    }
  }

  let bonusDataList = Array.from(bonusMap.entries()).map(([name, val]) => ({name, val}));
  bonusDataList.sort((a, b) => b.val - a.val);

  let maxRows = dashSheet.getMaxRows();
  if (maxRows >= 16) {
    dashSheet.getRange(16, 11, maxRows - 15, 2).clearContent().clearFormat();
  }

  let numBonusHits = Math.max(1, bonusDataList.length);
  let targetRange = dashSheet.getRange(16, 11, numBonusHits, 2);
  
  // Reapply the gold format to however many rows were generated
  targetRange.setBackground("#fff2cc").setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  let bonusOutput = [];
  for (let i = 0; i < bonusDataList.length; i++) {
    bonusOutput.push([bonusDataList[i].name, bonusDataList[i].val]);
  }
  
  if (bonusOutput.length > 0) {
    targetRange.setValues(bonusOutput);
  }
  dashSheet.getRange(16, 12, numBonusHits, 1).setNumberFormat("#,##0");

  // ==========================================
  // RESTORE STATIC LEADERBOARD FORMULAS (ALIGNED)
  // ==========================================
  let targetPayoutSheet = (typeof SETTINGS !== "undefined" && SETTINGS.payoutSheet) ? SETTINGS.payoutSheet : "Payouts";
  dashSheet.getRange("B24").setFormula(`=IFERROR(QUERY('${targetPayoutSheet}'!B3:E, "SELECT B, E WHERE E > 0 ORDER BY E DESC LIMIT 3", 0), "")`);
  dashSheet.getRange("B29").setFormula(`=IFERROR(QUERY('${targetPayoutSheet}'!B3:L, "SELECT B, L WHERE L > 0 ORDER BY L DESC LIMIT 3", 0), "")`);
  dashSheet.getRange("E24").setFormula(`=IFERROR(QUERY('${targetPayoutSheet}'!B3:I, "SELECT B, I WHERE I > 0 ORDER BY I DESC LIMIT 3", 0), "")`);

  // ==========================================
  // TOP CHAIN SAVES (PULLED FROM PAYOUT TAB)
  // ==========================================
  let payoutSheetForSaves = ss.getSheetByName(SETTINGS.payoutSheet || "Payout");
  if (payoutSheetForSaves) {
    let payoutData = payoutSheetForSaves.getDataRange().getValues();
    let savesLeaderboard = [];
    for (let i = 2; i < payoutData.length; i++) {
      let playerName = payoutData[i][1];
      let playerSaves = parseFloat(payoutData[i][9]) || 0; // Parses standard & decimal tiers

      if (playerName && playerSaves > 0) {
        savesLeaderboard.push({name: playerName, val: playerSaves});
      }
    }

    savesLeaderboard.sort((a, b) => b.val - a.val);
    let savesOutput = [];
    for (let i = 0; i < 3; i++) {
      if (i < savesLeaderboard.length) {
        savesOutput.push([savesLeaderboard[i].name, savesLeaderboard[i].val]);
      } else {
        savesOutput.push(["", ""]);
      }
    }

    // Print the data directly into E29:F31
    dashSheet.getRange("E29:F31").setValues(savesOutput);
  }
   // Format Top Respect to two decimal places
  dashSheet.getRange("C29:C31").setNumberFormat("#,##0.00");
}
// ==========================================
// DASHBOARD BUILDER (MASTER 1:1 VISUAL REPLICA)
// ==========================================
function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = (typeof SETTINGS !== "undefined" && SETTINGS.dashboardSheet) ? SETTINGS.dashboardSheet : "Dashboard";
  let dashSheet = ss.getSheetByName(sheetName);
  
// --- 1. STATE PRESERVATION (Save data before wiping) ---
  let sd = {}; // Saved Data object
  if (dashSheet && dashSheet.getLastRow() > 0) {
    try {
      sd.enemyName = dashSheet.getRange("C3").getValue();
      sd.enemyId = dashSheet.getRange("C4").getValue();
      sd.termed = dashSheet.getRange("C13").getValue();
      
      sd.totLim = dashSheet.getRange("F3").getValue();
      sd.warLim = dashSheet.getRange("F4").getValue();
      sd.chainLim = dashSheet.getRange("F5").getValue();
      sd.payPost = dashSheet.getRange("F6").getValue();
      
      // Use getDisplayValue() for dates/times to prevent formatting corruption
      sd.sDate = dashSheet.getRange("F9").getDisplayValue(); 
      sd.sTime = dashSheet.getRange("F10").getDisplayValue();
      sd.eDate = dashSheet.getRange("F11").getDisplayValue();
      sd.eTime = dashSheet.getRange("F12").getDisplayValue();
      
      sd.warRep = dashSheet.getRange("F15").getValue();
      sd.chainRep = dashSheet.getRange("F16").getValue();
      
      sd.tiers = dashSheet.getRange("E19:F21").getValues();
    } catch(e) {}
  }

  // Helper to fallback to default if saved data is blank
  const def = (val, fallback) => (val !== undefined && val !== "") ? val : fallback;

  // --- 2. WIPE SHEET ---
  if (!dashSheet) {
    dashSheet = ss.insertSheet(sheetName);
  } else {
    dashSheet.clear();
    dashSheet.clearFormats();
    dashSheet.getDataRange().clearDataValidations();
  }

  // --- 3. EXACT COLUMN WIDTHS ---
  dashSheet.setColumnWidth(1, 15);   // A: Spacer
  dashSheet.setColumnWidth(2, 170);  // B: War Labels
  dashSheet.setColumnWidth(3, 200);  // C: War Values
  dashSheet.setColumnWidth(4, 15);   // D: Spacer
  dashSheet.setColumnWidth(5, 170);  // E: Filter Labels
  dashSheet.setColumnWidth(6, 170);  // F: Filter Values
  dashSheet.setColumnWidth(7, 15);   // G: Spacer
  dashSheet.setColumnWidth(8, 170);  // H: Finance Labels
  dashSheet.setColumnWidth(9, 130);  // I: Finance Values
  dashSheet.setColumnWidth(10, 15);  // J: Spacer
  dashSheet.setColumnWidth(11, 150); // K: Bonus Labels
  dashSheet.setColumnWidth(12, 60);  // L: Bonus Values

  // --- COLOR PALETTE (From your file) ---
  const colors = {
    warHeader: "#cc0000", warBg: "#f4cccc",
    chainHeader: "#3c78d8", chainBg: "#cfe2f3",
    orangeHeader: "#e69138", orangeBg: "#fce5cd",
    purpleHeader: "#8e7cc3", purpleBg: "#d9d2e9",
    greenHeader: "#38761d", greenBg: "#d9ead3",
    goldHeader: "#f1c232", goldBg: "#fff2cc",
    valBg: "#ffffff", valText: "#000000", headerText: "#ffffff"
  };

  // Helper to build a section
  function buildBlock(rangeStr, headerRange, title, headerColor, labelColor) {
    let range = dashSheet.getRange(rangeStr);
    range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    dashSheet.getRange(headerRange).merge().setValue(title)
      .setBackground(headerColor).setFontColor(colors.headerText).setFontWeight("bold").setHorizontalAlignment("center");
    
    let rowStart = range.getRow() + 1;
    let rowCount = range.getNumRows() - 1;
    let col = range.getColumn();
    if (rowCount > 0) {
      dashSheet.getRange(rowStart, col, rowCount, 1).setBackground(labelColor).setFontWeight("bold").setHorizontalAlignment("right");
      dashSheet.getRange(rowStart, col + 1, rowCount, 1).setBackground(colors.valBg).setHorizontalAlignment("left");
    }
  }

  const yesNoRule = SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"], true).build();

  // ==========================================
  // LEFT COLUMN: WAR & CHAIN INSIGHTS
  // ==========================================
  buildBlock("B2:C13", "B2:C2", "⚔️ WAR INSIGHTS", colors.warHeader, colors.warBg);
  const warLabels = [
    ["Enemy Faction Name", ""], 
    ["Enemy Faction ID", ""],            
    ["Official War Start", ""],               
    ["Official War End", ""],                 
    ["War ID", ""],                
    ["Total War Hits", ""],                   
    ["War Score", ""],                        
    ["Outcome (Result)", ""],          
    ["Caches / Items Won", ""],         
    ["Est. Cache Value", ""],                 
    ["Termed?", "No"]                         
  ];
  dashSheet.getRange("B3:C13").setValues(warLabels);
  dashSheet.getRange("C11").setWrap(true); 
  dashSheet.getRange("C13").setDataValidation(yesNoRule).setHorizontalAlignment("left");

  buildBlock("B15:C19", "B15:C15", "🔗 CHAIN INSIGHTS", colors.chainHeader, colors.chainBg);
  const chainLabels = [
    ["First Attack Logged", ""],              
    ["Latest Attack Logged", ""],             
    ["Total Attacks Logged", ""],             
    ["Total Respect Generated", ""]           
  ];
  dashSheet.getRange("B16:C19").setValues(chainLabels);

  // ==========================================
  // MIDDLE COLUMN: FILTERS, TIME, REPORTS
  // ==========================================
  buildBlock("E2:F6", "E2:F2", "⚙️ PAYOUT FILTERS", colors.orangeHeader, colors.orangeBg);
  dashSheet.getRange("E3:F6").setValues([
    ["Total Hits (Max Limit)", ""], 
    ["Max War Hits", ""], 
    ["Max Chain Hits", ""],
    ["Pay Post-War Chain?", "Yes"]
  ]);
  dashSheet.getRange("F6").setDataValidation(yesNoRule);

  buildBlock("E8:F12", "E8:F8", "⏳ CUSTOM TIME WINDOW", "#1a73e8", colors.orangeBg);
  dashSheet.getRange("E9:F12").setValues([
    ["Start Date", ""], ["Start Time", ""], ["End Date", ""], ["End Time", ""]
  ]);

  buildBlock("E14:F16", "E14:F14", "📋 OFFICIAL REPORTS", colors.purpleHeader, colors.purpleBg);
  dashSheet.getRange("E15:F16").setValues([
    ["War Report ID", ""],               
    ["Chain Report ID", ""]                   
  ]);

  // ---> MULTI-TIER CHAIN WATCH <---
  buildBlock("E18:F21", "E18:F18", "⏱️ CHAIN WATCH", colors.purpleHeader, colors.purpleBg);
  // Row 19 is the Time/Weight Input. E18 is actually the Header in buildBlock. 
  // Let's manually set E18 to Sub-headers to fit your exact request
  dashSheet.getRange("E18:F18").breakApart().setValues([["Watch Time Limit", "Weight"]]).setBackground(colors.purpleBg).setFontColor("#000000");
  
  dashSheet.getRange("E19:F21").setValues([
    ["3:00", ".5"],
    ["", ""],
    ["", ""]
  ]);

  // ==========================================
  // RIGHT COLUMN: WAR FINANCIALS
  // ==========================================
  buildBlock("H2:I13", "H2:I2", "💰 WAR FINANCIALS", colors.greenHeader, colors.greenBg);
  const financeLabels = [
    ["Total Revenue", "0"],                   
    ["- Temp Cost", "0"],                     
    ["- Revives", "0"],                       
    ["- Xanax", "0"],         
    ["- Approved Bounties", "0"],             
    ["- Other Cost", "0"],                    
    ["NET PROFIT", "0"],                      
    ["Faction Cut %", "6%"],  
    ["Max Faction Cut ($)", ""],              
    ["Actual Faction Deduction", "0"],        
    ["PAYOUT TOTAL", "0"]                     
  ];
  dashSheet.getRange("H3:I13").setValues(financeLabels);
  
  dashSheet.getRange("H9:I9").setFontWeight("bold");
  dashSheet.getRange("H13:I13").setFontWeight("bold").setBackground(colors.greenBg);

  let bountySheetName = (typeof SETTINGS !== "undefined" && SETTINGS.bountySheet) ? SETTINGS.bountySheet : "Bounties";
  dashSheet.getRange("I7").setFormula(`=IFERROR(SUMIFS('${bountySheetName}'!E:E, '${bountySheetName}'!F:F, "Approved"), 0)`);
  dashSheet.getRange("I9").setFormula("=IFERROR(I3 - SUM(I4:I8), 0)");
  dashSheet.getRange("I12").setFormula("=IFERROR(I9*I10, 0)");
  dashSheet.getRange("I13").setFormula("=IFERROR(I9-I12, 0)"); 

  dashSheet.getRange("C12").setNumberFormat('"$ "#,##0');
  dashSheet.getRange("I3:I9").setNumberFormat('"$ "#,##0');
  dashSheet.getRange("I10").setNumberFormat('0%');
  dashSheet.getRange("I11:I13").setNumberFormat('"$ "#,##0');

  // ==========================================
  // BOTTOM LEADERBOARDS (From your original file)
  // ==========================================
  dashSheet.getRange("B23:C23").merge().setValue("🏆 TOP WAR HITS").setBackground(colors.goldHeader).setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("B24:C26").setBackground(colors.goldBg).setBorder(true, true, true, true, true, true);
  dashSheet.getRange("B24").setFormula(`=IFERROR(QUERY(Payouts!B3:E, "SELECT B, E WHERE E > 0 ORDER BY E DESC LIMIT 3", 0), "")`);

  dashSheet.getRange("B28:C28").merge().setValue("⭐ TOP RESPECT").setBackground(colors.goldHeader).setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("B29:C31").setBackground(colors.goldBg).setBorder(true, true, true, true, true, true);
  dashSheet.getRange("B29").setFormula(`=IFERROR(QUERY(Payouts!B3:L, "SELECT B, L WHERE L > 0 ORDER BY L DESC LIMIT 3", 0), "")`);

  dashSheet.getRange("E23:F23").merge().setValue("🔗 TOP CHAIN HITS").setBackground(colors.goldHeader).setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("E24:F26").setBackground(colors.goldBg).setBorder(true, true, true, true, true, true);
  dashSheet.getRange("E24").setFormula(`=IFERROR(QUERY(Payouts!B3:I, "SELECT B, I WHERE I > 0 ORDER BY I DESC LIMIT 3", 0), "")`);

  dashSheet.getRange("E28:F28").merge().setValue("🛡️ TOP CHAIN SAVES").setBackground(colors.goldHeader).setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("E29:F31").setBackground(colors.goldBg).setBorder(true, true, true, true, true, true);
  dashSheet.getRange("E29").setFormula(`=IFERROR(QUERY(Payouts!B3:J, "SELECT B, J WHERE J > 0 ORDER BY J DESC LIMIT 3", 0), "")`);

  dashSheet.getRange("H15:I15").merge().setValue("☑️ Avg Respect/Hit").setBackground(colors.goldHeader).setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("H16:I19").setBackground(colors.goldBg).setBorder(true, true, true, true, true, true);

  dashSheet.getRange("H23:I23").merge().setValue("🏆 Top Contribution").setBackground(colors.goldHeader).setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("H24:I31").setBackground(colors.goldBg).setBorder(true, true, true, true, true, true);

  dashSheet.getRange("K15:L15").merge().setValue("🎯 Bonus Chain Hits").setBackground(colors.goldHeader).setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("K16:L28").setBackground(colors.goldBg).setBorder(true, true, true, true, true, true);

  // ==========================================
  // CUSTOM COLOR OVERRIDES & AUTO-FILL SHADING
  // ==========================================
  // Shade Auto-Fill / Read-Only boxes light gray
  dashSheet.getRange("C4:C7").setBackground("#e6e8eb"); 
  dashSheet.getRange("C9:C12").setBackground("#e6e8eb"); 
  dashSheet.getRange("C16:C19").setBackground("#e6e8eb"); 

  // Make manual entry boxes bright white
  dashSheet.getRange("C4").setBackground("#ffffff"); // Faction ID
  dashSheet.getRange("C13").setBackground("#ffffff"); // Termed?
  dashSheet.getRange("F3:F6").setBackground("#ffffff"); // Payout Filters
  dashSheet.getRange("F9:F12").setBackground("#ffffff"); // Time Window
  dashSheet.getRange("F15:F16").setBackground("#ffffff"); // Report IDs
  dashSheet.getRange("E19:F21").setBackground("#F0FFFF").setHorizontalAlignment("center"); // Multi-Tier times & weights

  dashSheet.getRange("E6").setBackground(dashSheet.getRange("E3").getBackground());
  let colorI13 = dashSheet.getRange("I13").getBackground();
  dashSheet.getRange("I9").setBackground(colorI13);
  dashSheet.getRange("I11").setBackground(colorI13);

  // Formatting specific sections (From your file)
  dashSheet.getRange("E8:F8").setBackground("#1a73e8").setFontColor("#ffffff"); // Vibrant Blue Time Window
  dashSheet.getRange("E18:F18").setBackground("#1a73e8").setFontColor("#ffffff"); // Vibrant Blue Chain Watch

  dashSheet.setHiddenGridlines(true);
  dashSheet.setFrozenRows(1);
  
  if (typeof SpreadsheetApp !== "undefined" && SpreadsheetApp.getUi) {
    SpreadsheetApp.getUi().alert("✅ Full UI Rebuilt! Multi-tier saves added and all custom shading/leaderboards restored.");
  }
}