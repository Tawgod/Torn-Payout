// ==========================================
// OFFICIAL REPORT PICKER
// Recent ranked wars + multi-chain selection
// ==========================================

function showOfficialReportPicker() {
  const html = HtmlService.createHtmlOutput(`
<!doctype html>
<html>
<head>
  <base target="_top">
  <style>
    body{font-family:Arial,sans-serif;margin:0;padding:16px;color:#202124;background:#f8f9fa}
    h2{margin:0 0 6px} h3{margin:18px 0 8px;font-size:15px}
    .muted{color:#5f6368;font-size:12px}.box{background:white;border:1px solid #dadce0;border-radius:8px;padding:10px;max-height:210px;overflow:auto}
    label.row{display:block;padding:7px;border-bottom:1px solid #eee;cursor:pointer}.row:last-child{border-bottom:0}
    .meta{display:block;margin-left:24px;color:#5f6368;font-size:11px}
    select{width:100%;padding:7px;margin-top:5px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
    .actions{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}
    button{padding:8px 14px;border:0;border-radius:6px;cursor:pointer}.primary{background:#1a73e8;color:white}
    .status{margin-top:10px;font-size:12px;color:#5f6368}.error{color:#b3261e}
  </style>
</head>
<body>
  <h2>Select Official Reports</h2>
  <div class="muted">Choose one ranked war and any number of completed chains. Times shown are UTC.</div>

  <h3>Recent Ranked Wars</h3>
  <div id="wars" class="box">Loading…</div>

  <h3>Recent Chains</h3>
  <div id="chains" class="box">Loading…</div>

  <h3>Custom Time Window</h3>
  <div class="grid">
    <div><span class="muted">Start boundary</span><select id="startBoundary"></select></div>
    <div><span class="muted">End boundary</span><select id="endBoundary"></select></div>
  </div>
  <div class="muted" style="margin-top:6px">Auto start prioritizes War Start. Auto end prioritizes the latest selected Chain End.</div>

  <div id="status" class="status"></div>
  <div class="actions">
    <button onclick="google.script.host.close()">Cancel</button>
    <button class="primary" id="applyBtn" onclick="applySelection()">Apply Selection</button>
  </div>

<script>
let DATA={wars:[],chains:[]};

function fmt(ts){
  if(!ts) return "Ongoing / unknown";
  const d=new Date(ts*1000);
  return d.toISOString().replace("T"," ").slice(0,16)+" UTC";
}
function esc(s){return String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}

function selectedWar(){
  const el=document.querySelector('input[name="war"]:checked');
  if(!el||!el.value) return null;
  return DATA.wars.find(x=>String(x.id)===el.value)||null;
}
function selectedChains(){
  const ids=[...document.querySelectorAll('input[name="chain"]:checked')].map(x=>x.value);
  return DATA.chains.filter(x=>ids.includes(String(x.id)));
}
function rebuildBoundaries(){
  const war=selectedWar(), chains=selectedChains();
  const s=document.getElementById("startBoundary"), e=document.getElementById("endBoundary");
  const oldS=s.value||"auto", oldE=e.value||"auto";
  let so=[{v:"auto",t:"Auto — War Start, then earliest Chain Start"}];
  let eo=[{v:"auto",t:"Auto — latest Chain End, then War End"}];
  if(war){
    if(war.start) so.push({v:"war:start",t:"War Start — "+fmt(war.start)});
    if(war.end) eo.push({v:"war:end",t:"War End — "+fmt(war.end)});
  }
  chains.forEach(c=>{
    if(c.start) so.push({v:"chain:"+c.id+":start",t:"Chain "+c.id+" Start — "+fmt(c.start)});
    if(c.end) eo.push({v:"chain:"+c.id+":end",t:"Chain "+c.id+" End — "+fmt(c.end)});
  });
  s.innerHTML=so.map(o=>'<option value="'+o.v+'">'+esc(o.t)+'</option>').join("");
  e.innerHTML=eo.map(o=>'<option value="'+o.v+'">'+esc(o.t)+'</option>').join("");
  if(so.some(o=>o.v===oldS)) s.value=oldS;
  if(eo.some(o=>o.v===oldE)) e.value=oldE;
}
function render(data){
  DATA=data||{wars:[],chains:[]};
  const w=document.getElementById("wars"), c=document.getElementById("chains");
  w.innerHTML='<label class="row"><input type="radio" name="war" value="" checked> No war report</label>'+
    DATA.wars.map(x=>'<label class="row"><input type="radio" name="war" value="'+esc(x.id)+'"> <b>War '+esc(x.id)+'</b> — '+esc(x.opponent||"Unknown opponent")+
      '<span class="meta">'+fmt(x.start)+' → '+fmt(x.end)+'</span></label>').join("");
  c.innerHTML=DATA.chains.length?DATA.chains.map(x=>'<label class="row"><input type="checkbox" name="chain" value="'+esc(x.id)+'"> <b>Chain '+esc(x.id)+'</b> — '+esc(x.length||"?")+' hits'+
      '<span class="meta">'+fmt(x.start)+' → '+fmt(x.end)+'</span></label>').join(""):'No completed chains returned.';
  document.querySelectorAll('input[name="war"],input[name="chain"]').forEach(x=>x.addEventListener("change",rebuildBoundaries));
  rebuildBoundaries();
  document.getElementById("status").textContent="";
}
function loadFail(err){
  document.getElementById("status").className="status error";
  document.getElementById("status").textContent=(err&&err.message)||String(err);
}
function applySelection(){
  const war=selectedWar(), chains=selectedChains();
  if(!war && !chains.length){loadFail("Select at least one war or chain.");return;}
  const btn=document.getElementById("applyBtn"); btn.disabled=true;
  document.getElementById("status").className="status";
  document.getElementById("status").textContent="Applying selection…";
  google.script.run.withSuccessHandler(msg=>{document.getElementById("status").textContent=msg;setTimeout(()=>google.script.host.close(),500);})
    .withFailureHandler(err=>{btn.disabled=false;loadFail(err);})
    .applyOfficialReportSelection({
      war:war,
      chains:chains,
      startBoundary:document.getElementById("startBoundary").value,
      endBoundary:document.getElementById("endBoundary").value
    });
}
google.script.run.withSuccessHandler(render).withFailureHandler(loadFail).getRecentOfficialReportOptions();
</script>
</body></html>`);
  html.setWidth(720).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, "Official Torn Reports");
}

function getRecentOfficialReportOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfgSheet = ss.getSheetByName(SETTINGS.configSheet);
  const factionId = cfgSheet ? String(cfgSheet.getRange(SETTINGS.factionIdCell || "B7").getValue()).replace(/,/g, "").trim() : "";

  const warsJson = tornApiRequest_("faction", "", "rankedwars");
  const chainsJson = tornApiRequest_("faction", "", "chains");

  const wars = normalizeRecentReportList_(warsJson.rankedwars || warsJson.wars || warsJson, "war")
    .map(x => ({
      id: x.id,
      start: x.start,
      end: x.end,
      opponent: recentWarOpponent_(x.raw, factionId)
    }))
    .filter(x => x.id)
    .sort((a,b) => (b.end || b.start || 0) - (a.end || a.start || 0))
    .slice(0, 20);

  const chains = normalizeRecentReportList_(chainsJson.chains || chainsJson, "chain")
    .map(x => ({
      id: x.id,
      start: x.start,
      end: x.end,
      length: reportNum_(x.raw.chain || x.raw.length || x.raw.hits || x.raw.current || x.raw.max)
    }))
    .filter(x => x.id)
    .sort((a,b) => (b.end || b.start || 0) - (a.end || a.start || 0))
    .slice(0, 30);

  return {wars:wars, chains:chains};
}

function normalizeRecentReportList_(src, kind) {
  let rows = [];
  if (Array.isArray(src)) rows = src;
  else if (src && typeof src === "object") {
    rows = Object.keys(src).map(k => {
      const v = src[k];
      if (v && typeof v === "object" && v.id === undefined) v.__key = k;
      return v;
    }).filter(v => v && typeof v === "object");
  }

  return rows.map(r => {
    const id = reportNum_(r.id || r.war_id || r.warId || r.chain_id || r.chainId || r.__key);
    const warObj = r.war && typeof r.war === "object" ? r.war : {};
    const start = reportUnix_(r.start || r.started || r.start_time || r.created_at || warObj.start);
    const end = reportUnix_(r.end || r.ended || r.end_time || r.finished_at || warObj.end);
    return {id:id, start:start, end:end, raw:r, kind:kind};
  });
}

function recentWarOpponent_(war, factionId) {
  const factions = war && war.factions ? war.factions : null;
  if (factions && typeof factions === "object") {
    const arr = Array.isArray(factions) ? factions : Object.keys(factions).map(k => {
      const v = factions[k] || {};
      if (v.id === undefined) v.id = k;
      return v;
    });
    for (const f of arr) {
      const id = String(f.id || f.faction_id || "").replace(/,/g, "").trim();
      if (id && id !== factionId) return String(f.name || f.faction_name || ("Faction " + id));
    }
  }
  return String(war.opponent_name || war.enemy_name || war.name || "");
}

function reportNum_(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(String(v).replace(/,/g, ""));
  return isFinite(n) ? Math.trunc(n) : null;
}

function reportUnix_(v) {
  if (!v) return null;
  if (typeof v === "number") return v > 20000000000 ? Math.floor(v / 1000) : Math.floor(v);
  const n = Number(v);
  if (isFinite(n) && n > 0) return n > 20000000000 ? Math.floor(n / 1000) : Math.floor(n);
  const ms = new Date(v).getTime();
  return isNaN(ms) ? null : Math.floor(ms / 1000);
}

function applyOfficialReportSelection(selection) {
  selection = selection || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SETTINGS.dashboardSheet);
  if (!dash) throw new Error("Dashboard sheet not found.");

  const war = selection.war || null;
  const chains = Array.isArray(selection.chains) ? selection.chains : [];
  const chainIds = chains.map(c => reportNum_(c.id)).filter(Boolean);

  setLabelValue_(dash, "War Report ID", war && war.id ? reportNum_(war.id) : "");
  setLabelValue_(dash, "Chain Report ID", chainIds.join(", "));

  const startUnix = resolveReportBoundary_(selection.startBoundary, war, chains, "start");
  const endUnix = resolveReportBoundary_(selection.endBoundary, war, chains, "end");

  if (startUnix) writeDashboardBoundary_(dash, "Start Date", "Start Time", startUnix);
  if (endUnix) writeDashboardBoundary_(dash, "End Date", "End Time", endUnix);

  ss.toast("Official reports selected. Fetching reports…", "Reports", 5);
  fetchOfficialReports();
  return "Selection applied, custom time window updated, and official reports fetched.";
}

function resolveReportBoundary_(key, war, chains, side) {
  const isStart = side === "start";
  key = String(key || "auto");

  if (key === "auto") {
    if (isStart) {
      const w = war ? reportUnix_(war.start) : null;
      if (w) return w;
      const starts = chains.map(c => reportUnix_(c.start)).filter(Boolean);
      return starts.length ? Math.min.apply(null, starts) : null;
    }
    const ends = chains.map(c => reportUnix_(c.end)).filter(Boolean);
    if (ends.length) return Math.max.apply(null, ends);
    return war ? reportUnix_(war.end) : null;
  }

  if (key === "war:start") return war ? reportUnix_(war.start) : null;
  if (key === "war:end") return war ? reportUnix_(war.end) : null;

  const m = key.match(/^chain:(\\d+):(start|end)$/);
  if (m) {
    const c = chains.find(x => String(x.id) === m[1]);
    return c ? reportUnix_(c[m[2]]) : null;
  }
  return null;
}

function writeDashboardBoundary_(dash, dateLabel, timeLabel, unix) {
  const d = new Date(unix * 1000);
  const y = Number(Utilities.formatDate(d, "GMT", "yyyy"));
  const m = Number(Utilities.formatDate(d, "GMT", "MM"));
  const day = Number(Utilities.formatDate(d, "GMT", "dd"));
  const time = Utilities.formatDate(d, "GMT", "HH:mm:ss");

  const dateHit = findLabelCell_(dash, dateLabel);
  const timeHit = findLabelCell_(dash, timeLabel);
  if (!dateHit || !timeHit) throw new Error("Custom time window fields were not found on Dashboard.");

  dateHit.valueRange.setValue(new Date(Date.UTC(y, m - 1, day))).setNumberFormat("yyyy-MM-dd");
  timeHit.valueRange.setValue(time).setNumberFormat("@");
}
