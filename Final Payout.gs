function buildFinalPayoutTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payoutSheet = ss.getSheetByName(SETTINGS.payoutSheet);
  const bountySheet = ss.getSheetByName(SETTINGS.bountySheet);
  let finalSheet = ss.getSheetByName(SETTINGS.finalSheet);

  if (!payoutSheet) {
    ss.toast("Please generate the Payouts tab first!", "Error", 5);
    return;
  }

  const lastRow = payoutSheet.getLastRow();
  if (lastRow < 3) {
    ss.toast("No members found on Payouts tab.", "Error", 5);
    return;
  }

  const rosterData = payoutSheet.getRange(3, 1, lastRow - 2, 2).getValues();

  let bountyFormulaParts = null;
  if (bountySheet && bountySheet.getLastColumn() > 0) {
    const bountyHeaders = headerMapFromRow_(bountySheet.getRange(1, 1, 1, bountySheet.getLastColumn()).getValues()[0]);
    const placedByCol = headerIndex_(bountyHeaders, "Placed By", false);
    const refundCol = headerIndex_(bountyHeaders, "Refund Amount", false);
    const statusCol = headerIndex_(bountyHeaders, "Status", false);
    if (placedByCol >= 0 && refundCol >= 0 && statusCol >= 0) {
      const toA1Col = n => {
        let s = "";
        n++;
        while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
        return s;
      };
      bountyFormulaParts = {
        placed: toA1Col(placedByCol),
        refund: toA1Col(refundCol),
        status: toA1Col(statusCol)
      };
    }
  }

  if (!finalSheet) {
    finalSheet = ss.insertSheet(SETTINGS.finalSheet, 5);
  } else {
    finalSheet.clear();
  }

  finalSheet.setHiddenGridlines(true);
  finalSheet.setFrozenRows(1);
  finalSheet.setFrozenColumns(2);

  const headers = SETTINGS.finalHeaders;
  finalSheet.getRange(1, 1, 1, headers.length).setValues([headers])
            .setBackground("#b45f06") 
            .setFontColor("white")
            .setFontWeight("bold")
            .setHorizontalAlignment("center");

  if (rosterData.length > 0) {
    finalSheet.getRange(2, 1, rosterData.length, 2).setValues(rosterData);
    
    for (let r = 0; r < rosterData.length; r++) {
      let rowNum = r + 2;     
      let payoutRow = r + 3;  
      
      let warFormula = `='${SETTINGS.payoutSheet}'!D${payoutRow}`;
      let bountyFormula = bountyFormulaParts
        ? `=SUMIFS('${SETTINGS.bountySheet}'!${bountyFormulaParts.refund}:${bountyFormulaParts.refund}, '${SETTINGS.bountySheet}'!${bountyFormulaParts.placed}:${bountyFormulaParts.placed}, B${rowNum}, '${SETTINGS.bountySheet}'!${bountyFormulaParts.status}:${bountyFormulaParts.status}, "Approved")`
        : `=0`;
      let totalFormula = `=C${rowNum} + D${rowNum} + E${rowNum}`;
      let linkFormula = `=IF(F${rowNum}>0, HYPERLINK("https://www.torn.com/factions.php?step=your#/tab=controls&addMoneyTo=" & A${rowNum} & "&money=" & INT(F${rowNum}), "💸 Pay " & B${rowNum}), "No Payout")`;

      finalSheet.getRange(rowNum, 3).setFormula(warFormula);
      finalSheet.getRange(rowNum, 4).setFormula(bountyFormula);
      finalSheet.getRange(rowNum, 6).setFormula(totalFormula);
      finalSheet.getRange(rowNum, 7).setFormula(linkFormula);
      
      let bgColor = (r % 2 === 0) ? "#f8f9fa" : "#ffffff";
      finalSheet.getRange(rowNum, 1, 1, headers.length).setBackground(bgColor).setHorizontalAlignment("center");
      finalSheet.getRange(rowNum, 1, 1, 2).setHorizontalAlignment("left"); 
      
      finalSheet.getRange(rowNum, 5).setBackground("#fff2cc").setBorder(true, true, true, true, false, false, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
      finalSheet.getRange(rowNum, 6).setBackground((r % 2 === 0) ? "#e6f4ea" : "#ceead6").setFontWeight("bold");
    }
    
    // Formatting Financials
    finalSheet.getRange(2, 3, rosterData.length, 4).setNumberFormat('"$ "#,##0');
    finalSheet.getRange(1, 1, rosterData.length + 1, headers.length).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
    
    // NEW: Insert Checkboxes into Column H (Payment Status)
    finalSheet.getRange(2, 8, rosterData.length, 1).insertCheckboxes();

    // NEW: Add Conditional Formatting so checked rows dim and strike-through
    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2=TRUE')
      .setBackground('#d9d9d9')
      .setFontColor('#999999')
      .setStrikethrough(true)
      .setRanges([finalSheet.getRange(2, 1, rosterData.length, headers.length)])
      .build();
    let rules = finalSheet.getConditionalFormatRules();
    rules.push(rule);
    finalSheet.setConditionalFormatRules(rules);
  }

  // Formatting Column Widths
  finalSheet.setColumnWidth(2, 160); 
  finalSheet.setColumnWidth(3, 140); 
  finalSheet.setColumnWidth(4, 140); 
  finalSheet.setColumnWidth(5, 160); 
  finalSheet.setColumnWidth(6, 160); 
  finalSheet.setColumnWidth(7, 180); 
  finalSheet.setColumnWidth(8, 140); // Checkbox Column
  
  finalSheet.hideColumns(1); 

  ss.toast("Final Payout Tab Generated with Status Checkboxes!", "System", 3);
}