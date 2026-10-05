// ==========================================
// 3. OFFICIAL TORN REPORT FETCHER (MULTI-CHAIN)
// ==========================================
function fetchOfficialReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);
  
  if (!dashSheet || !configSheet) return;
  
  const rawConfigFactionId = configSheet.getRange(SETTINGS.factionIdCell || "B7").getValue().toString().trim();
  
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
    let itemsJson = tornApiRequest_("torn", "", "items");
    if (itemsJson.items) {
      for (let id in itemsJson.items) {
        itemValueMap[itemsJson.items[id].name] = itemsJson.items[id].market_value || 0;
      }
    }
  } catch(e) {}
  
  // --- OFFICIAL WAR REPORT ---
  if (warId) {
    ss.toast("Fetching War Report...", "System", 3);
    let json = tornApiRequest_("torn", warId, "rankedwarreport");
    
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
        let newsJson = tornApiRequest_("faction", "", "news");
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
      
      setVal("official war start", startStr);
      setVal("official war end", endStr);
      setVal("war id", warId);
      
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
      try {
        let json = tornApiRequest_("torn", chainId, "chainreport");
        
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
        const rdHeaderMap = headerMapFromRow_(rdData[0] || []);
        const aFacCol = headerIndex_(rdHeaderMap, ["Attacker Faction", "Attacker Faction ID"], true);
        const respectCol = headerIndex_(rdHeaderMap, "Respect", true);
        const chainBonusCol = headerIndex_(rdHeaderMap, "Chain Bonus", false);
        let totalBonusRespectToStrip = 0;
        for (let i = 1; i < rdData.length; i++) {
          let aFac = rdData[i][aFacCol] ? rdData[i][aFacCol].toString().replace(/,/g, "").trim() : "";
          if (aFac === rawConfigFactionId) {
            let respect = parseFloat(rdData[i][respectCol]) || 0;
            let cBonus = chainBonusCol >= 0 ? (parseFloat(rdData[i][chainBonusCol]) || 1) : 1;
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
      
      const rdHeaderMap = headerMapFromRow_(rdData[0] || []);
      const timeCol = headerIndex_(rdHeaderMap, ["End Time", "Timestamp", "End"], true);
      const aFacCol = headerIndex_(rdHeaderMap, ["Attacker Faction", "Attacker Faction ID"], true);
      const respectCol = headerIndex_(rdHeaderMap, "Respect", true);
      const chainBonusCol = headerIndex_(rdHeaderMap, "Chain Bonus", false);

      for (let i = 1; i < rdData.length; i++) {
        let tVal = rdData[i][timeCol];
        let aFac = cleanId(rdData[i][aFacCol]);
        let respect = parseFloat(rdData[i][respectCol]) || 0;
        let cBonus = chainBonusCol >= 0 ? (parseFloat(rdData[i][chainBonusCol]) || 1) : 1;
        
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

  // Push chain insights by label so Dashboard layout changes are safe.
  setVal("first attack logged", finalChainStart);
  setVal("latest attack logged", finalChainEnd);
  setVal("total attacks logged", finalChainHits);
  setVal("total respect generated", finalChainRespect);
  
  if (typeof importWarData === "function") { importWarData(true); }
  if (typeof refreshDashboard === "function") { refreshDashboard(); }
  SpreadsheetApp.getUi().alert("✅ Official Reports Fetched, Raw Attack Data Pulled & Dashboard Refreshed!");
}

// ==========================================
// DASHBOARD ENGINE (Smart Auto-Fill)
// ==========================================
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  const configSheet = ss.getSheetByName(SETTINGS.configSheet);

  if (!dashSheet || !configSheet) return;

  const myFactionId = configSheet.getRange(SETTINGS.factionIdCell || "B7").getValue().toString().trim();

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
      let json = tornApiRequest_("torn", manualWarId, "rankedwarreport");

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
        
        setVal("official war start", formatTornDate(startUnix));
        setVal("official war end", formatTornDate(endUnix));
        setVal("war id", manualWarId);

        if (factions[myFactionId] && factions[myFactionId].rewards && factions[myFactionId].rewards.items) {
          let myItems = factions[myFactionId].rewards.items;
          let cacheArr = [];
          for (let i in myItems) { cacheArr.push(`${myItems[i].quantity}x ${myItems[i].name}`); }
          setVal("caches / items won", cacheArr.length > 0 ? cacheArr.join(", ") : "None");
        }

        if (json.rankedwarreport.war && json.rankedwarreport.war.winner !== undefined) {
          let winnerId = json.rankedwarreport.war.winner.toString();
          if (winnerId === myFactionId) setVal("outcome (result)", "Win");
          else if (winnerId === "0") setVal("outcome (result)", "Draw");
          else setVal("outcome (result)", "Loss");
        } else {
          setVal("outcome (result)", "Finished");
        }
      }
    } else {
      let json = tornApiRequest_("faction", "", "basic");

      if (json.ranked_wars && Object.keys(json.ranked_wars).length > 0) {
        let activeWarId = Object.keys(json.ranked_wars)[0];
        let warData = json.ranked_wars[activeWarId];

        setVal("war report id", activeWarId);
        setVal("war id", activeWarId);
        setVal("outcome (result)", "Ongoing");

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
            setVal("official war start", formatTornDate(warData.war.start));
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
  // 4. AWARDS TAB
  // ==========================================
  if (typeof refreshAwards === "function") {
    refreshAwards();
  }
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
      
      sd.totLim = labelValue_(dashSheet, "Total Hits (Max Limit)", "");
      sd.warLim = labelValue_(dashSheet, "Max War Hits", "");
      sd.chainLim = labelValue_(dashSheet, "Max Chain Hits (Faction Total)", labelValue_(dashSheet, "Max Chain Hits", ""));
      sd.payPost = labelValue_(dashSheet, "Pay Post-War Chain?", "Yes");

      // Preserve report/time values by label so layout changes remain safe.
      sd.sDate = labelValue_(dashSheet, "Start Date", "");
      sd.sTime = labelValue_(dashSheet, "Start Time", "");
      sd.eDate = labelValue_(dashSheet, "End Date", "");
      sd.eTime = labelValue_(dashSheet, "End Time", "");
      sd.warRep = labelValue_(dashSheet, "War Report ID", "");
      sd.chainRep = labelValue_(dashSheet, "Chain Report ID", "");

      sd.factionCut = labelValue_(dashSheet, "Faction Cut %", 0.06);
      sd.maxFactionCut = labelValue_(dashSheet, "Max Faction Cut ($)", "");

      // Preserve Push Settings from the dynamic block, with legacy/current layout fallback.
      try {
        sd.pushSettings = readPushSettingsFromDashboard_(dashSheet);
      } catch (e) {
        try { sd.pushSettings = dashSheet.getRange("K4:M13").getValues().filter(r => r.some(v => v !== "")); }
        catch (_e) { sd.pushSettings = []; }
      }

      sd.preset = labelValue_(dashSheet, "Payout Preset", "Custom");

      try {
        const cw = findLabelCell_(dashSheet, "Watch Time Limit");
        if (cw) sd.tiers = dashSheet.getRange(cw.labelRange.getRow() + 1, cw.labelRange.getColumn(), 3, 2).getValues();
        else sd.tiers = [];
      } catch (e) { sd.tiers = []; }
    } catch(e) {}
  }

  // Helper to fallback to default if saved data is blank
  const def = (val, fallback) => (val !== undefined && val !== "") ? val : fallback;

  // --- 2. WIPE SHEET ---
  if (!dashSheet) {
    dashSheet = ss.insertSheet(sheetName);
  } else {
    // Clear validations across the full grid BEFORE clearing contents.
    // getDataRange() can collapse after clear(), leaving stale validations behind.
    dashSheet.getRange(1, 1, dashSheet.getMaxRows(), dashSheet.getMaxColumns()).clearDataValidations();
    dashSheet.clear();
    dashSheet.clearFormats();
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
  dashSheet.setColumnWidth(11, 170); // K: Financial labels
  dashSheet.setColumnWidth(12, 130); // L: Financial values
  dashSheet.setColumnWidth(13, 15);  // M: Spacer
  dashSheet.setColumnWidth(14, 95);  // N: Push Start Time
  dashSheet.setColumnWidth(15, 105); // O: Push Start Date
  dashSheet.setColumnWidth(16, 120); // P: Push Time Limit

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
  // SETTINGS GROUP 1: REPORTS & TIME
  // ==========================================
  buildBlock("E2:F8", "E2:F2", "📋 REPORTS & TIME", colors.purpleHeader, colors.purpleBg);
  dashSheet.getRange("E3:F8").setValues([
    ["War Report ID", def(sd.warRep, "")],
    ["Chain Report ID", def(sd.chainRep, "")],
    ["Start Date", def(sd.sDate, "")],
    ["Start Time", def(sd.sTime, "")],
    ["End Date", def(sd.eDate, "")],
    ["End Time", def(sd.eTime, "")]
  ]);

  // ==========================================
  // SETTINGS GROUP 2: CHAIN WATCH
  // ==========================================
  buildBlock("E10:F14", "E10:F10", "⏱️ CHAIN WATCH", "#1a73e8", colors.purpleBg);
  dashSheet.getRange("E11:F11").setValues([["Watch Time Limit", "Weight"]])
    .setBackground(colors.purpleBg).setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange("E12:F14").setValues([
    ["3:00", ".5"],
    ["", ""],
    ["", ""]
  ]);
  if (sd.tiers && sd.tiers.length > 0) dashSheet.getRange("E12:F14").setValues(sd.tiers);

  // ==========================================
  // SETTINGS GROUP 3: PAYOUT SETTINGS
  // ==========================================
  buildBlock("H2:I12", "H2:I2", "⚙️ PAYOUT SETTINGS", colors.orangeHeader, colors.orangeBg);
  const presetCfg = (typeof getPresetConfig_ === "function") ? getPresetConfig_() : {defaultCut: 0.06, defaultMaxCut: ""};
  dashSheet.getRange("H3:I12").setValues([
    ["Total Hits (Max Limit)", def(sd.totLim, "")],
    ["Max War Hits", def(sd.warLim, "")],
    ["Max Chain Hits (Faction Total)", def(sd.chainLim, "")],
    ["Pay Post-War Chain?", def(sd.payPost, "Yes")],
    ["Payout Preset", def(sd.preset, "Custom")],
    ["Preset Cut", presetCfg.defaultCut],
    ["Preset Max Cut", presetCfg.defaultMaxCut],
    ["Weights", "TBD / Current"],
    ["Faction Cut %", def(sd.factionCut, presetCfg.defaultCut)],
    ["Max Faction Cut ($)", def(sd.maxFactionCut, presetCfg.defaultMaxCut)]
  ]);
  dashSheet.getRange("I6").setDataValidation(yesNoRule);

  const presetRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Custom", "Termed Win", "Termed Loss", "Real War"], true)
    .setAllowInvalid(false)
    .build();
  dashSheet.getRange("I7").setDataValidation(presetRule);
  dashSheet.getRange("I8").setNumberFormat("0%");
  dashSheet.getRange("I9").setNumberFormat('"$ "#,##0');
  dashSheet.getRange("I11").setNumberFormat("0%");
  dashSheet.getRange("I12").setNumberFormat('"$ "#,##0');

  // ==========================================
  // SETTINGS GROUP 4: WAR FINANCIALS
  // ==========================================
  buildBlock("K2:L13", "K2:L2", "💰 WAR FINANCIALS", colors.greenHeader, colors.greenBg);
  const financeLabels = [
    ["Total Revenue", "0"],
    ["- Temp Cost", "0"],
    ["- Revives", "0"],
    ["- Xanax", "0"],
    ["- Approved Bounties", "0"],
    ["- Other Cost", "0"],
    ["NET PROFIT", "0"],
    ["Actual Faction Deduction", "0"],
    ["PAYOUT TOTAL", "0"],
    ["", ""],
    ["", ""]
  ];
  dashSheet.getRange("K3:L13").setValues(financeLabels);

  let bountySheetName = (typeof SETTINGS !== "undefined" && SETTINGS.bountySheet) ? SETTINGS.bountySheet : "Bounties";
  dashSheet.getRange("L7").setFormula(`=IFERROR(SUMIFS('${bountySheetName}'!E:E, '${bountySheetName}'!F:F, "Approved"), 0)`);
  dashSheet.getRange("L9").setFormula("=IFERROR(L3 - SUM(L4:L8), 0)");
  dashSheet.getRange("L10").setFormula('=IFERROR(MIN(L9*I11, IF(I12="", L9*I11, I12)), 0)');
  dashSheet.getRange("L11").setFormula("=IFERROR(L9-L10, 0)");
  dashSheet.getRange("L9:L11").setFontWeight("bold");
  dashSheet.getRange("L3:L11").setNumberFormat('"$ "#,##0');

  // ==========================================
  // SETTINGS GROUP 5: PUSH SETTINGS
  // ==========================================
  const pushRows = (sd.pushSettings && sd.pushSettings.length) ? sd.pushSettings : [];
  buildPushSettingsBlock_(dashSheet, pushRows);

  // ==========================================
  // INPUT / AUTO-FILL SHADING
  // ==========================================
  dashSheet.getRange("C4:C7").setBackground("#e6e8eb");
  dashSheet.getRange("C9:C12").setBackground("#e6e8eb");
  dashSheet.getRange("C16:C19").setBackground("#e6e8eb");
  dashSheet.getRange("C4").setBackground("#ffffff");
  dashSheet.getRange("C13").setBackground("#ffffff").setDataValidation(yesNoRule);

  dashSheet.getRange("F3:F8").setBackground("#ffffff");
  dashSheet.getRange("E12:F14").setBackground("#F0FFFF").setHorizontalAlignment("center");

  dashSheet.getRange("I3:I7").setBackground("#ffffff");
  dashSheet.getRange("I8:I10").setBackground("#e6e8eb");
  dashSheet.getRange("I11:I12").setBackground("#ffffff");

  dashSheet.getRange("L3:L8").setBackground("#ffffff");
  dashSheet.getRange("L9:L11").setBackground(colors.greenBg);

  dashSheet.setHiddenGridlines(true);
  dashSheet.setFrozenRows(1);

  if (typeof buildAwardsSheet === "function") {
    buildAwardsSheet();
  }
  
  if (typeof SpreadsheetApp !== "undefined" && SpreadsheetApp.getUi) {
    SpreadsheetApp.getUi().alert("✅ Dashboard rebuilt and Awards tab refreshed.");
  }
}