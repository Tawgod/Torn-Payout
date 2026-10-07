// ==========================================
// PAYOUT RECONCILIATION
// Compares Final Payout expectations to Torn faction giveFunds news.
// Read-only audit: never changes balances or payout amounts.
// ==========================================

function payoutReconciliationWindow_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SETTINGS.dashboardSheet);
  let start = null;

  if (dash) {
    const rawEnd = labelValue_(dash, "Official War End", "");
    if (rawEnd && String(rawEnd).toLowerCase() !== "ongoing") {
      let d = new Date(rawEnd);
      if (isNaN(d.getTime())) d = new Date(String(rawEnd) + " GMT");
      if (!isNaN(d.getTime())) start = d;
    }
  }

  if (!start) start = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  return {
    from: Math.floor(start.getTime() / 1000),
    to: Math.floor(Date.now() / 1000),
    start: start,
    end: new Date()
  };
}

function payoutNewsRows_(body) {
  if (!body) return [];
  if (Array.isArray(body.news)) return body.news;
  if (body.news && typeof body.news === "object") return Object.keys(body.news).map(k => body.news[k]);
  if (Array.isArray(body)) return body;
  return [];
}

function payoutNewsText_(row) {
  return String(
    (row && (row.text || row.message || row.news || row.description || row.details)) || ""
  );
}

function payoutNewsTimestamp_(row) {
  const raw = row && (row.timestamp || row.time || row.created_at || row.date);
  if (raw instanceof Date) return raw;
  if (typeof raw === "number") return new Date(raw * (raw < 20000000000 ? 1000 : 1));
  const d = new Date(raw || 0);
  return isNaN(d.getTime()) ? null : d;
}

function payoutNewsMemberId_(row) {
  const candidates = [
    row && row.user && row.user.id,
    row && row.member && row.member.id,
    row && row.target && row.target.id,
    row && row.user_id,
    row && row.member_id,
    row && row.target_id
  ];
  for (let i = 0; i < candidates.length; i++) {
    const v = parseInt(String(candidates[i] || "").replace(/,/g, ""), 10);
    if (v) return String(v);
  }

  const text = payoutNewsText_(row);
  let m = text.match(/[(d{3,12})]/);
  if (m) return m[1];

  m = text.match(/(?:user|member|player)s*(?:id)?s*[:#]?s*(d{3,12})/i);
  return m ? m[1] : "";
}

function payoutNewsAmount_(row) {
  const candidates = [
    row && row.amount,
    row && row.money,
    row && row.value,
    row && row.details && row.details.amount
  ];
  for (let i = 0; i < candidates.length; i++) {
    const n = Number(String(candidates[i] || "").replace(/[$,]/g, ""));
    if (isFinite(n) && n > 0) return n;
  }

  const text = payoutNewsText_(row);
  const matches = text.match(/$[d,]+(?:.d+)?/g) || [];
  if (!matches.length) return 0;
  return Math.max.apply(null, matches.map(v => Number(v.replace(/[$,]/g, "")) || 0));
}

function fetchFactionGiveFundsNews_(window) {
  const cfg = getArchiveDatabaseConfig_();
  return archiveApiRequest_(
    "/api/torn-news?faction=" + encodeURIComponent(cfg.factionKey) +
    "&category=giveFunds" +
    "&from=" + encodeURIComponent(window.from) +
    "&to=" + encodeURIComponent(window.to)
  );
}

function buildPayoutReconciliationReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finalSheet = ss.getSheetByName(SETTINGS.finalSheet || "Final Payout");
  if (!finalSheet || finalSheet.getLastRow() < 2) {
    throw new Error("Final Payout must be generated before reconciliation.");
  }

  SpreadsheetApp.flush();

  const headers = finalSheet.getRange(1, 1, 1, finalSheet.getLastColumn()).getDisplayValues()[0];
  const map = headerMapFromRow_(headers);
  const idCol = headerIndex_(map, "Member ID", true);
  const nameCol = headerIndex_(map, "Name", true);
  const totalCol = headerIndex_(map, "Total Final Payout ($)", true);
  const paidCol = headerIndex_(map, "Payment Status", false);

  const rows = finalSheet.getRange(2, 1, finalSheet.getLastRow() - 1, finalSheet.getLastColumn()).getValues()
    .filter(r => payoutWebCleanId_(r[idCol]) && Number(r[totalCol] || 0) > 0);

  const window = payoutReconciliationWindow_();
  const newsBody = fetchFactionGiveFundsNews_(window);
  const newsRows = payoutNewsRows_(newsBody);

  const paidByMember = {};
  const evidenceByMember = {};

  newsRows.forEach(row => {
    const memberId = payoutNewsMemberId_(row);
    const amount = payoutNewsAmount_(row);
    if (!memberId || !(amount > 0)) return;

    paidByMember[memberId] = (paidByMember[memberId] || 0) + amount;
    if (!evidenceByMember[memberId]) evidenceByMember[memberId] = [];
    evidenceByMember[memberId].push({
      amount: amount,
      timestamp: payoutNewsTimestamp_(row),
      text: payoutNewsText_(row)
    });
  });

  const results = rows.map(row => {
    const memberId = payoutWebCleanId_(row[idCol]);
    const name = String(row[nameCol] || "");
    const expected = Number(row[totalCol] || 0);
    const observed = Number(paidByMember[memberId] || 0);
    const diff = observed - expected;
    const markedPaid = paidCol >= 0 ? row[paidCol] === true : false;

    let status = "MATCH";
    if (observed === 0 && !markedPaid) status = "NOT PAID YET";
    else if (observed === 0 && markedPaid) status = "MISSING PAYMENT";
    else if (diff < 0) status = "UNDERPAID";
    else if (diff > 0) status = "OVERPAID";

    const evidence = evidenceByMember[memberId] || [];
    const latest = evidence.length
      ? evidence.map(x => x.timestamp).filter(Boolean).sort((a,b)=>b-a)[0]
      : null;

    return [
      memberId,
      name,
      expected,
      observed,
      diff,
      markedPaid ? "Yes" : "No",
      status,
      latest || "",
      evidence.length
    ];
  });

  let audit = ss.getSheetByName("Payout Reconciliation");
  if (!audit) audit = ss.insertSheet("Payout Reconciliation");
  audit.clear();
  audit.clearFormats();
  audit.setHiddenGridlines(true);
  audit.setFrozenRows(2);

  audit.getRange("A1:I1").merge()
    .setValue("🔎 PAYOUT RECONCILIATION")
    .setBackground("#783f04").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center");

  const outHeaders = [
    "Member ID", "Name", "Expected Payout", "Bank Log Total", "Difference",
    "Sheet Marked Paid", "Status", "Latest Payment", "Matching Log Entries"
  ];
  audit.getRange(2, 1, 1, outHeaders.length).setValues([outHeaders])
    .setBackground("#b45f06").setFontColor("white").setFontWeight("bold");

  if (results.length) {
    audit.getRange(3, 1, results.length, outHeaders.length).setValues(results);
    audit.getRange(3, 3, results.length, 3).setNumberFormat('"$ "#,##0');
    audit.getRange(3, 8, results.length, 1).setNumberFormat("yyyy-mm-dd hh:mm");

    const statusCol = 7;
    results.forEach((row, i) => {
      const status = row[6];
      let bg = "#d9ead3";
      if (status === "NOT PAID YET") bg = "#fff2cc";
      if (status === "MISSING PAYMENT" || status === "UNDERPAID" || status === "OVERPAID") bg = "#f4cccc";
      audit.getRange(i + 3, statusCol).setBackground(bg).setFontWeight("bold");
    });
  }

  audit.autoResizeColumns(1, outHeaders.length);
  audit.setColumnWidth(2, 180);
  audit.setColumnWidth(7, 150);

  const errors = results.filter(r => ["MISSING PAYMENT", "UNDERPAID", "OVERPAID"].indexOf(r[6]) !== -1).length;
  ss.setActiveSheet(audit);
  ss.toast(
    "Reconciliation complete: " + errors + " payout error" + (errors === 1 ? "" : "s") + " found.",
    "Payout Audit",
    6
  );

  return {
    checked: results.length,
    errors: errors,
    rows: results,
    from: window.start.toISOString(),
    to: window.end.toISOString()
  };
}
