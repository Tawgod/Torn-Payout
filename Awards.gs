// ==========================================
// AWARDS / LEADERBOARDS
// ==========================================

function getOrCreateAwardsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Awards");
  if (!sheet) sheet = ss.insertSheet("Awards");
  return sheet;
}

function buildAwardsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateAwardsSheet_();
  sheet.clear();
  sheet.clearFormats();
  sheet.getDataRange().clearDataValidations();

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(2);

  [1,4,7,10,13].forEach(c => sheet.setColumnWidth(c, 165));
  [2,5,8,11,14].forEach(c => sheet.setColumnWidth(c, 95));
  [3,6,9,12,15].forEach(c => sheet.setColumnWidth(c, 24));

  sheet.getRange("A1:N1").merge().setValue("🏆 FACTION AWARDS")
    .setBackground("#f1c232").setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center");

  sheet.getRange("A2:N2").merge().setValue("War, chain, respect, and activity leaders for the currently loaded payout window.")
    .setBackground("#fff2cc").setHorizontalAlignment("center");

  refreshAwards();
}

function writeAwardBlock_(sheet, startRow, startCol, title, rows, numberFormat) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const count = Math.max(1, safeRows.length);
  sheet.getRange(startRow, startCol, 1, 2).merge().setValue(title)
    .setBackground("#f1c232").setFontWeight("bold").setHorizontalAlignment("center");
  const body = sheet.getRange(startRow + 1, startCol, count, 2);
  body.clearContent().clearFormat()
    .setBackground("#fff2cc")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  if (safeRows.length) body.offset(0,0,safeRows.length,2).setValues(safeRows);
  if (numberFormat) sheet.getRange(startRow + 1, startCol + 1, count, 1).setNumberFormat(numberFormat);
}

function refreshAwards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const awards = getOrCreateAwardsSheet_();
  const payout = ss.getSheetByName(SETTINGS.payoutSheet);
  const rd = ss.getSheetByName(SETTINGS.rdSheet || "RD");
  const config = ss.getSheetByName(SETTINGS.configSheet);
  const myFactionId = config ? String(config.getRange(SETTINGS.factionIdCell || "B7").getValue()).replace(/,/g,"").trim() : "";

  if (awards.getRange("A1").getDisplayValue().indexOf("FACTION AWARDS") === -1) {
    buildAwardsSheet();
    return;
  }

  let pRows = [];
  if (payout && payout.getLastRow() >= 3) {
    pRows = payout.getRange(3,1,payout.getLastRow()-2,Math.max(17,payout.getLastColumn())).getValues();
  }

  const valid = row => {
    const name = String(row[1] || "").trim();
    const low = name.toLowerCase();
    return name && !low.includes("left faction") && low !== "totals" && low !== "total";
  };

  function topFromPayout(idx, n, transform) {
    return pRows.filter(valid).map(r => [r[1], transform ? transform(r[idx], r) : (parseFloat(r[idx]) || 0)])
      .filter(r => Number(r[1]) > 0)
      .sort((a,b)=>b[1]-a[1]).slice(0,n);
  }

  const topWarHits = topFromPayout(4, 10);
  const topChainHits = topFromPayout(8, 10);
  const topChainSaves = topFromPayout(9, 10);
  const topRespect = topFromPayout(11, 10);
  const topContribution = topFromPayout(2, 10);

  const avgRespect = pRows.filter(valid).map(r => {
    const wh = parseFloat(r[4]) || 0;
    const ch = parseFloat(r[8]) || 0;
    const res = parseFloat(r[11]) || 0;
    const hits = wh + ch;
    return [r[1], hits > 0 ? res / hits : 0];
  }).filter(r=>r[1]>0).sort((a,b)=>b[1]-a[1]).slice(0,10);

  let bonusMap = new Map();
  let warlordMap = new Map();
  let retalMap = new Map();
  let overseasMap = new Map();
  let energyMap = new Map();
  let bestRespectHit = [];

  if (rd && rd.getLastRow() > 1) {
    const data = rd.getDataRange().getValues();
    const hm = headerMapFromRow_(data[0] || []);
    const aFacCol = headerIndex_(hm, ["Attacker Faction", "Attacker Faction ID"], true);
    const aNameCol = headerIndex_(hm, "Attacker Name", true);
    const respectCol = headerIndex_(hm, "Respect", true);
    const cBonusCol = headerIndex_(hm, "Chain Bonus", false);
    const retCol = headerIndex_(hm, "Retaliation", false);
    const overseasCol = headerIndex_(hm, "Overseas", false);
    const warlordCol = headerIndex_(hm, ["Warlord Bonus", "Warlord"], false);

    for (let r=1;r<data.length;r++) {
      const fac = String(data[r][aFacCol] || "").replace(/,/g,"").trim();
      if (!fac || fac !== myFactionId) continue;
      const name = String(data[r][aNameCol] || "Unknown").trim();
      const respect = parseFloat(data[r][respectCol]) || 0;
      const cBonus = cBonusCol >= 0 ? (parseFloat(data[r][cBonusCol]) || 0) : 0;
      const ret = retCol >= 0 ? (parseFloat(data[r][retCol]) || 1) : 1;
      const overseas = overseasCol >= 0 ? (parseFloat(data[r][overseasCol]) || 1) : 1;
      const warlord = warlordCol >= 0 ? (parseFloat(data[r][warlordCol]) || 1) : 1;

      energyMap.set(name, (energyMap.get(name) || 0) + 25);
      if (cBonus >= 10) bonusMap.set(name, Math.max(bonusMap.get(name) || 0, cBonus));
      if (ret > 1) retalMap.set(name, (retalMap.get(name) || 0) + 1);
      if (overseas > 1) overseasMap.set(name, (overseasMap.get(name) || 0) + 1);
      if (warlord > 1) warlordMap.set(name, (warlordMap.get(name) || 0) + 1);
      if (respect > 0) bestRespectHit.push([name, respect]);
    }
  }

  const fromMap = (m,n=10) => Array.from(m.entries()).sort((a,b)=>b[1]-a[1]).slice(0,n);
  bestRespectHit.sort((a,b)=>b[1]-a[1]);

  writeAwardBlock_(awards, 4, 1, "⚔️ Top War Hits", topWarHits, "#,##0");
  writeAwardBlock_(awards, 4, 4, "🔗 Top Chain Hits", topChainHits, "#,##0");
  writeAwardBlock_(awards, 4, 7, "🛡️ Top Chain Saves", topChainSaves, "#,##0.00");
  writeAwardBlock_(awards, 4, 10, "⭐ Top Respect", topRespect, "#,##0.00");
  writeAwardBlock_(awards, 4, 13, "🏆 Top Contribution", topContribution, "0.00%");

  writeAwardBlock_(awards, 18, 1, "☑️ Best Respect / Hit", avgRespect, "#,##0.00");
  writeAwardBlock_(awards, 18, 4, "🎯 Bonus Chain Hits", fromMap(bonusMap), "#,##0");
  writeAwardBlock_(awards, 18, 7, "🔥 Est. Energy Spent", fromMap(energyMap), "#,##0");
  writeAwardBlock_(awards, 18, 10, "🗡️ Warlord Hits", fromMap(warlordMap), "#,##0");
  writeAwardBlock_(awards, 18, 13, "↩️ Retaliations", fromMap(retalMap), "#,##0");

  writeAwardBlock_(awards, 32, 1, "🌍 Overseas Hits", fromMap(overseasMap), "#,##0");
  writeAwardBlock_(awards, 32, 4, "💥 Highest Respect Hit", bestRespectHit.slice(0,10), "#,##0.00");

  awards.autoResizeRows(1, Math.min(60, awards.getMaxRows()));
}
