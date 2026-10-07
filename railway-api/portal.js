const express = require("express");
const crypto = require("crypto");

const ACCESS_RANK = { limited: 1, view: 2, edit: 3 };
const LIMITED_MEMBER_FIELDS = new Set([
  "Name", "Payout", "War Hits", "Assists", "Losses", "Interruptions",
  "Outside/Chain Hits", "Chain Hits", "Chain Saves", "Retaliations",
  "Net Respect", "War Abroad Hits", "War Score"
]);
const LIMITED_SUMMARY_FIELDS = new Set([
  "outcome", "termed", "total_hits", "total_war_hits", "war_score",
  "caches", "official_war_start", "official_war_end"
]);

function b64urlDecode(value) {
  return Buffer.from(String(value || "").replace(/-/g, "+").replace(/_/g, "/"), "base64");
}

function verifyPortalToken(token, secret) {
  if (!token || !secret) return null;
  try {
    const [body, signature] = String(token).split(".", 2);
    if (!body || !signature) return null;
    const expected = crypto.createHmac("sha256", secret).update(body).digest();
    const actual = b64urlDecode(signature);
    if (expected.length !== actual.length || !crypto.timingSafeEqual(expected, actual)) return null;
    const payload = JSON.parse(b64urlDecode(body).toString("utf8"));
    if (!payload || Number(payload.exp || 0) <= Math.floor(Date.now() / 1000)) return null;
    if (!ACCESS_RANK[payload.access]) return null;
    if (!payload.guild_id || !payload.discord_user_id) return null;
    const factions = Array.isArray(payload.factions) ? payload.factions : [];
    payload.factions = factions.map(v => String(v).trim().toLowerCase()).filter(Boolean);
    return payload;
  } catch (_e) {
    return null;
  }
}

function parseCookies(req) {
  const out = {};
  const raw = req.get("cookie") || "";
  for (const part of raw.split(";")) {
    const idx = part.indexOf("=");
    if (idx <= 0) continue;
    out[part.slice(0, idx).trim()] = decodeURIComponent(part.slice(idx + 1).trim());
  }
  return out;
}

function tokenFromRequest(req) {
  const bearer = (req.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (bearer) return bearer;
  const header = req.get("x-portal-session");
  if (header) return header;
  return parseCookies(req).payout_portal_session || "";
}

function factionAllowed(identity, faction) {
  const key = String(faction || "").trim().toLowerCase();
  return identity && (identity.factions.includes("both") || identity.factions.includes(key));
}

function accessAtLeast(identity, required) {
  return Boolean(identity && ACCESS_RANK[identity.access] >= ACCESS_RANK[required]);
}

function limitedSummary(summary) {
  const out = {};
  for (const [key, value] of Object.entries(summary || {})) {
    if (LIMITED_SUMMARY_FIELDS.has(key)) out[key] = value;
  }
  return out;
}

function memberRows(headers, rows, access) {
  const sourceHeaders = Array.isArray(headers) ? headers : [];
  const sourceRows = Array.isArray(rows) ? rows : [];
  if (access !== "limited") {
    return { headers: sourceHeaders, rows: sourceRows };
  }
  const keep = sourceHeaders.map((h, i) => ({ h: String(h || ""), i }))
    .filter(x => LIMITED_MEMBER_FIELDS.has(x.h));
  return {
    headers: keep.map(x => x.h),
    rows: sourceRows.map(row => keep.map(x => Array.isArray(row) ? row[x.i] : null))
  };
}

function redactRecord(record, access) {
  const payout = memberRows(record.payout_headers || [], record.member_rows || [], access);
  const base = {
    id: record.id,
    faction_key: record.faction_key,
    war_id: record.war_id,
    legacy_key: record.legacy_key,
    enemy_name: record.enemy_name,
    enemy_faction_id: record.enemy_faction_id,
    archived_at: record.archived_at,
    published_at: record.published_at,
    public_title: record.public_title,
    outcome: record.outcome,
    termed: record.termed,
    total_hits: record.total_hits,
    total_war_hits: record.total_war_hits,
    war_score: record.war_score,
    caches: record.caches,
    official_war_start: record.official_war_start,
    official_war_end: record.official_war_end,
    summary: access === "limited" ? limitedSummary(record.summary) : (record.summary || {}),
    payout_headers: payout.headers,
    member_rows: payout.rows
  };
  if (access !== "limited") {
    Object.assign(base, {
      total_revenue: record.total_revenue,
      total_payout: record.total_payout,
      faction_profit: record.faction_profit,
      source: record.source,
      updated_at: record.updated_at
    });
  }
  return base;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function portalShell(identity, body) {
  const access = escapeHtml(identity.access);
  const who = escapeHtml(identity.discord_name || identity.torn_name || identity.discord_user_id);
  return `<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Resolute Family Payout Portal</title>
<style>
:root{color-scheme:dark;--bg:#0d1117;--card:#161b22;--line:#30363d;--muted:#8b949e;--gold:#d29922}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:#f0f3f6;font:15px/1.45 system-ui,-apple-system,Segoe UI,sans-serif}
.wrap{max-width:1200px;margin:auto;padding:28px 20px}.top{display:flex;justify-content:space-between;gap:16px;align-items:center;flex-wrap:wrap}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px;margin:14px 0}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}
a{color:#f2cc60}.muted{color:var(--muted)}.badge{display:inline-block;padding:4px 9px;border:1px solid var(--line);border-radius:999px;text-transform:uppercase;font-size:12px}
table{width:100%;border-collapse:collapse}th,td{padding:9px;border-bottom:1px solid var(--line);text-align:left}.scroll{overflow:auto}
button{background:var(--gold);border:0;border-radius:8px;padding:9px 12px;font-weight:700;color:#111}.disabled{opacity:.55}
</style></head><body><main class="wrap"><div class="top"><div><h1>Faction Payout Portal</h1>
<div class="muted">Signed in as ${who} · <span class="badge">${access}</span></div></div>
<div><a href="/portal/logout">Sign out</a></div></div>${body}</main></body></html>`;
}

function sheetBridgeCanonical(request) {
  return [
    String(request.ts || ""),
    String(request.nonce || ""),
    String(request.action || ""),
    String(request.faction || ""),
    JSON.stringify(request.payload || {})
  ].join(".");
}

async function callSheetBridge(bridge, faction, action, payload) {
  if (!bridge || !bridge.url || !bridge.secret) {
    const err = new Error("Sheet bridge is not configured for this faction.");
    err.status = 503;
    throw err;
  }
  const request = {
    ts: Math.floor(Date.now() / 1000),
    nonce: crypto.randomUUID(),
    action,
    faction,
    payload: payload || {}
  };
  request.sig = crypto.createHmac("sha256", bridge.secret)
    .update(sheetBridgeCanonical(request)).digest("base64url");

  const response = await fetch(bridge.url, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(request),
    redirect: "follow"
  });
  const text = await response.text();
  let body;
  try { body = JSON.parse(text); } catch (_e) { body = { ok: false, error: text || "Invalid Sheet bridge response" }; }
  if (!response.ok || !body.ok) {
    const err = new Error(body.error || ("Sheet bridge returned HTTP " + response.status));
    err.status = response.status >= 400 ? response.status : 502;
    throw err;
  }
  return body.result;
}

const LIMITED_DASHBOARD_FIELDS = new Set([
  "Enemy Faction Name", "Enemy Faction ID", "Official War Start", "Official War End",
  "War ID", "Total War Hits", "War Score", "Outcome (Result)",
  "First Attack Logged", "Latest Attack Logged", "Total Attacks Logged", "Total Respect Generated"
]);

function redactSnapshot(snapshot, access) {
  if (access !== "limited") return snapshot;
  const dashboard = {};
  for (const [key, value] of Object.entries((snapshot && snapshot.dashboard) || {})) {
    if (LIMITED_DASHBOARD_FIELDS.has(key)) dashboard[key] = value;
  }
  const safe = {
    faction_key: snapshot && snapshot.faction_key,
    generated_at: snapshot && snapshot.generated_at,
    dashboard,
    awards: snapshot && snapshot.awards ? snapshot.awards : []
  };
  return safe;
}

const liveCache = new Map();

async function fetchChainWatch(upstream, factionId) {
  if (!upstream || !upstream.url || !upstream.token) return null;
  const qs = new URLSearchParams({ faction_id: String(factionId) });
  const response = await fetch(upstream.url.replace(/\/+$/, "") + "/internal/payout-live?" + qs.toString(), {
    headers: { Authorization: "Bearer " + upstream.token }
  });
  const text = await response.text();
  let body;
  try { body = JSON.parse(text); } catch (_e) { body = { error: text || "Invalid Chain Watch response" }; }
  if (!response.ok) throw new Error(body.error || ("Chain Watch HTTP " + response.status));
  return body.chain_watch || null;
}

function factionIdForKey(faction) {
  const key = String(faction || "").toLowerCase();
  if (key === "ironsides") return 40664;
  if (key === "resolute") return 46442;
  return 0;
}

function createPortalRouter({ pool, sessionSecret, sheetBridges = {}, chainWatchUpstream = {} }) {
  const router = express.Router();

  function identity(req) {
    return verifyPortalToken(tokenFromRequest(req), sessionSecret);
  }

  router.get("/portal/login", (req, res) => {
    if (!sessionSecret) return res.status(503).send("Portal authentication is not configured.");
    const token = String(req.query.t || "");
    const id = verifyPortalToken(token, sessionSecret);
    if (!id) return res.status(401).send("Invalid or expired portal link.");
    res.setHeader("Set-Cookie", `payout_portal_session=${encodeURIComponent(token)}; Path=/portal; HttpOnly; Secure; SameSite=Lax; Max-Age=43200`);
    return res.redirect("/portal");
  });

  router.get("/portal/logout", (_req, res) => {
    res.setHeader("Set-Cookie", "payout_portal_session=; Path=/portal; HttpOnly; Secure; SameSite=Lax; Max-Age=0");
    res.redirect("/");
  });

  router.get("/portal", async (req, res) => {
    const id = identity(req);
    if (!id) return res.status(401).send(portalShell({ access: "limited", discord_user_id: "Guest" },
      '<div class="card"><h2>Private portal</h2><p>This page requires a signed Discord access link.</p></div>'));

    const factions = id.factions.includes("both") ? ["ironsides", "resolute"] : id.factions;
    const rows = await pool.query(
      `SELECT id,faction_key,war_id,legacy_key,enemy_name,archived_at,published_at,outcome,termed,total_payout,updated_at
       FROM archive_wars
       WHERE faction_key = ANY($1::text[])
         AND ($2::boolean = FALSE OR published_at IS NOT NULL)
       ORDER BY COALESCE(archived_at::timestamptz, updated_at) DESC NULLS LAST
       LIMIT 40`,
      [factions, id.access === "limited"]
    );

    const cards = factions.map(f => `<div class="card"><h2>${escapeHtml(f)}</h2>
      <p><a href="/portal/faction/${encodeURIComponent(f)}">Open faction workspace →</a></p>${accessAtLeast(id, "view") ? `<p><a href="/portal/live/${encodeURIComponent(f)}">Current War dashboard →</a></p>` : ""}</div>`).join("");

    const recordRows = rows.rows.map(r => `<tr><td>${escapeHtml(r.faction_key)}</td><td><a href="/portal/record/${r.id}">${escapeHtml(r.enemy_name || r.public_title || "Archived payout")}</a></td>
      <td>${escapeHtml(r.war_id || r.legacy_key || "")}</td><td>${escapeHtml(r.outcome || "")}</td>
      <td>${id.access === "limited" ? "—" : escapeHtml(r.total_payout == null ? "" : "$"+Math.round(Number(r.total_payout)).toLocaleString("en-US"))}</td></tr>`).join("");

    const editNotice = id.access === "edit"
      ? '<div class="card"><b>Edit access confirmed.</b><p class="muted">Spreadsheet writes are intentionally disabled in this staging phase. The next step connects approved fields/actions to the Sheet bridge with audit logging.</p></div>'
      : "";

    res.send(portalShell(id, `${editNotice}<div class="grid">${cards}</div>
      <div class="card"><h2>Recent payout records</h2><div class="scroll"><table><thead><tr><th>Faction</th><th>Opponent</th><th>War / legacy ID</th><th>Outcome</th><th>Payout</th></tr></thead>
      <tbody>${recordRows || '<tr><td colspan="5">No records available for this access level.</td></tr>'}</tbody></table></div></div>`));
  });

  router.get("/portal/faction/:faction", async (req, res) => {
    const id = identity(req);
    const faction = String(req.params.faction || "").toLowerCase();
    if (!id || !factionAllowed(id, faction)) return res.status(403).send("Forbidden");
    const limited = id.access === "limited";
    const q = await pool.query(
      `SELECT id,faction_key,war_id,legacy_key,enemy_name,archived_at,published_at,outcome,total_payout
       FROM archive_wars WHERE faction_key=$1 AND ($2::boolean=FALSE OR published_at IS NOT NULL)
       ORDER BY COALESCE(archived_at::timestamptz, updated_at) DESC NULLS LAST LIMIT 60`,
      [faction, limited]
    );
    const body = q.rows.map(r => `<tr><td><a href="/portal/record/${r.id}">${escapeHtml(r.enemy_name || "Payout")}</a></td><td>${escapeHtml(r.war_id || r.legacy_key || "")}</td><td>${escapeHtml(r.outcome || "")}</td><td>${limited ? "—" : escapeHtml(r.total_payout == null ? "" : "$"+Math.round(Number(r.total_payout)).toLocaleString("en-US"))}</td></tr>`).join("");
    res.send(portalShell(id, `<div class="card"><p><a href="/portal">← Portal</a></p><h2>${escapeHtml(faction)} workspace</h2>
      <p class="muted">Access: ${escapeHtml(id.access)}. ${limited ? "Only published, server-redacted records are available." : "Faction records are available according to your signed role."}</p>
      <div class="scroll"><table><thead><tr><th>Opponent</th><th>War / legacy ID</th><th>Outcome</th><th>Payout</th></tr></thead><tbody>${body}</tbody></table></div></div>`));
  });

  router.get("/portal/record/:id", async (req, res) => {
    const ident = identity(req);
    if (!ident) return res.status(401).send("Unauthorized");
    const recordId = Number(req.params.id);
    if (!Number.isInteger(recordId) || recordId <= 0) return res.status(400).send("Invalid record");
    const q = await pool.query(
      `SELECT w.*, COALESCE(jsonb_agg(m.row_values ORDER BY m.row_order) FILTER (WHERE m.id IS NOT NULL), '[]'::jsonb) AS member_rows
       FROM archive_wars w LEFT JOIN archive_member_payouts m ON m.war_record_id=w.id
       WHERE w.id=$1 GROUP BY w.id`,
      [recordId]
    );
    if (!q.rowCount) return res.status(404).send("Not found");
    const raw = q.rows[0];
    if (!factionAllowed(ident, raw.faction_key)) return res.status(403).send("Forbidden");
    if (ident.access === "limited" && !raw.published_at) return res.status(404).send("Not found");
    const record = redactRecord(raw, ident.access);
    const headers = record.payout_headers || [];
    const tableHead = headers.map(h => `<th>${escapeHtml(h)}</th>`).join("");
    const tableRows = (record.member_rows || []).map(row => `<tr>${row.map(v => `<td>${escapeHtml(v)}</td>`).join("")}</tr>`).join("");
    const financials = ident.access === "limited" ? "" : `<div class="grid">
      <div class="card"><span class="muted">Revenue</span><h3>${record.total_revenue == null ? "—" : "$"+Math.round(Number(record.total_revenue)).toLocaleString("en-US")}</h3></div>
      <div class="card"><span class="muted">Total payout</span><h3>${record.total_payout == null ? "—" : "$"+Math.round(Number(record.total_payout)).toLocaleString("en-US")}</h3></div>
      <div class="card"><span class="muted">Faction profit</span><h3>${record.faction_profit == null ? "—" : "$"+Math.round(Number(record.faction_profit)).toLocaleString("en-US")}</h3></div></div>`;
    res.send(portalShell(ident, `<div class="card"><p><a href="/portal/faction/${encodeURIComponent(record.faction_key)}">← Back</a></p>
      <h2>${escapeHtml(record.enemy_name || record.public_title || "Payout")}</h2><p class="muted">${escapeHtml(record.faction_key)} · ${escapeHtml(record.war_id || record.legacy_key || "")}</p></div>
      ${financials}<div class="card"><h2>Member payout</h2><div class="scroll"><table><thead><tr>${tableHead}</tr></thead><tbody>${tableRows}</tbody></table></div></div>`));
  });

  router.get("/portal/live/:faction", async (req, res) => {
    const id = identity(req);
    const faction = String(req.params.faction || "").toLowerCase();
    if (!id || !accessAtLeast(id, "view") || !factionAllowed(id, faction)) {
      return res.status(403).send("Faction View access required.");
    }

    const liveApiUrl = "/portal/api/live/" + encodeURIComponent(faction);
    const body = \`
      <div class="card"><p><a href="/portal">← Portal</a></p>
        <div class="top"><div><h2 style="margin:0">Current War · \${escapeHtml(faction)}</h2>
        <p class="muted">Operational dashboard only. No payout or faction-financial data is loaded here.</p></div>
        <div class="badge" id="refreshStatus">Loading…</div></div>
      </div>
      <div id="liveRoot"><div class="card">Loading current-war data…</div></div>
      <script>
      const root = document.getElementById("liveRoot");
      const statusEl = document.getElementById("refreshStatus");
      const liveApiUrl = \${JSON.stringify(liveApiUrl)};
      const esc = v => String(v ?? "").replace(/[&<>"']/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[ch]));
      const num = v => Number(v || 0).toLocaleString();
      const money = v => "$" + Math.round(Number(v || 0)).toLocaleString();
      const when = v => {
        if (!v) return "—";
        const d = new Date(v);
        return Number.isNaN(d.getTime()) ? esc(v) : d.toLocaleString();
      };
      function stat(label, value) {
        return '<div class="card"><span class="muted">'+esc(label)+'</span><h2 style="margin:.3rem 0 0">'+esc(value)+'</h2></div>';
      }
      function render(data) {
        const w=data.war||{}, t=data.totals||{}, cw=data.chain_watch||{};
        const members=(data.members||[]).slice(0,50);
        const attacks=(data.recent_attacks||[]).slice(0,30);
        const bounties=(data.bounties||[]).slice(0,50);
        const memberRows=members.map(m=>'<tr><td>'+esc(m.name)+'</td><td>'+num(m.war_hits)+'</td><td>'+num(m.chain_hits)+'</td><td>'+num(m.assists)+'</td><td>'+Number(m.base_respect||0).toFixed(2)+'</td><td>'+num(m.retaliations)+'</td><td>'+num(m.interruptions)+'</td></tr>').join("");
        const attackRows=attacks.map(a=>'<tr><td>'+when(a.time)+'</td><td>'+esc(a.attacker_name)+'</td><td>'+esc(a.defender_name||a.defender_id)+'</td><td>'+esc(a.result)+'</td><td>'+Number(a.base_respect||0).toFixed(2)+'</td><td>'+(Number(a.chain_bonus||1)>1 ? esc(a.chain_bonus)+'x' : '—')+'</td></tr>').join("");
        const bountyRows=bounties.map(b=>'<tr><td>'+when(b.date_logged)+'</td><td>'+esc(b.placed_by)+'</td><td>'+esc(b.target)+'</td><td>'+num(b.quantity||1)+'</td><td>'+money(b.bounty_amount)+'</td><td>'+esc(b.status)+'</td></tr>').join("");
        let chainCard='<div class="card"><h2>Chain Watch</h2><p class="muted">No active Chain Watch information is available.</p></div>';
        if (cw && cw.active) {
          chainCard='<div class="card"><h2>Chain Watch</h2><div class="grid">'+
            stat("Risk", (cw.risk_status||"No Data")+" · "+num(cw.risk_score))+
            stat("Current watcher", cw.current_watcher_name||"Unassigned")+
            stat("Active watchers", num(cw.active_watcher_count))+
            stat("Next watcher", cw.next_watcher_name||"Unassigned")+
            stat("Scheduled next hour", num(cw.scheduled_next_hour))+
            stat("Uncovered next hour", num(cw.uncovered_slots_next_hour))+
          '</div></div>';
        }
        root.innerHTML =
          '<div class="card"><h2>'+esc(w.enemy_name||"Current War")+'</h2><p class="muted">War '+esc(w.war_id||"—")+' · score '+num(w.war_score)+' · '+esc(w.outcome||"Ongoing")+'</p></div>'+
          '<div class="grid">'+
            stat("War hits",num(t.war_hits))+stat("Chain hits",num(t.chain_hits))+stat("Assists",num(t.assists))+
            stat("Base respect",Number(t.base_respect||0).toFixed(2))+stat("Retaliations",num(t.retaliations))+
            stat("Interruptions",num(t.interruptions))+stat("Overseas hits",num(t.overseas_hits))+stat("Chain bonus hits",num(t.chain_bonus_hits))+
          '</div>'+
          chainCard+
          '<div class="card"><h2>Member Activity</h2><div class="scroll"><table><thead><tr><th>Member</th><th>War Hits</th><th>Chain Hits</th><th>Assists</th><th>Base Respect</th><th>Retals</th><th>Interrupts</th></tr></thead><tbody>'+memberRows+'</tbody></table></div></div>'+
          '<div class="card"><h2>Recent Attacks</h2><div class="scroll"><table><thead><tr><th>Time</th><th>Attacker</th><th>Defender</th><th>Result</th><th>Base Respect</th><th>Chain Bonus</th></tr></thead><tbody>'+attackRows+'</tbody></table></div></div>'+
          '<div class="card"><h2>Current Bounties</h2><div class="scroll"><table><thead><tr><th>Logged</th><th>Placed By</th><th>Target</th><th>Qty</th><th>Amount</th><th>Status</th></tr></thead><tbody>'+bountyRows+'</tbody></table></div></div>';
      }
      async function refreshLive() {
        statusEl.textContent="Updating…";
        try {
          const r=await fetch(liveApiUrl,{cache:"no-store"});
          const j=await r.json();
          if(!r.ok) throw new Error(j.error||"Unable to load");
          render(j.live);
          statusEl.textContent="Updated "+new Date().toLocaleTimeString();
        } catch(e) {
          statusEl.textContent="Update failed";
          root.innerHTML='<div class="card"><b>Live dashboard unavailable.</b><p class="muted">'+esc(e.message)+'</p></div>';
        }
      }
      refreshLive();
      setInterval(refreshLive,30000);
      </script>\`;
    res.send(portalShell(id, body));
  });

  router.get("/portal/api/live/:faction", async (req, res) => {
    const id = identity(req);
    const faction = String(req.params.faction || "").toLowerCase();
    if (!id || !accessAtLeast(id, "view") || !factionAllowed(id, faction)) {
      return res.status(403).json({ error: "Faction View access required" });
    }
    const now = Date.now();
    const cached = liveCache.get(faction);
    if (cached && now - cached.at < 15000) return res.json({ ok: true, live: cached.data, cached: true });
    try {
      const sheet = await callSheetBridge(sheetBridges[faction], faction, "war_live_snapshot", {});
      let chainWatch = null;
      try {
        chainWatch = await fetchChainWatch(chainWatchUpstream, factionIdForKey(faction));
      } catch (_e) {
        chainWatch = { active: false, status: "Chain Watch unavailable" };
      }
      const live = Object.assign({}, sheet, { chain_watch: chainWatch });
      liveCache.set(faction, { at: now, data: live });
      return res.json({ ok: true, live, cached: false });
    } catch (err) {
      return res.status(err.status || 502).json({ error: err.message || "Live war dashboard unavailable" });
    }
  });

  router.get("/portal/api/snapshot/:faction", async (req, res) => {
    const id = identity(req);
    const faction = String(req.params.faction || "").toLowerCase();
    if (!id || !factionAllowed(id, faction)) return res.status(403).json({ error: "Forbidden" });
    try {
      const result = await callSheetBridge(sheetBridges[faction], faction, "snapshot", {});
      return res.json({ ok: true, snapshot: redactSnapshot(result, id.access) });
    } catch (err) {
      return res.status(err.status || 502).json({ error: err.message || "Sheet bridge failed" });
    }
  });

  router.post("/portal/api/action/:faction", async (req, res) => {
    const id = identity(req);
    const faction = String(req.params.faction || "").toLowerCase();
    if (!id || !accessAtLeast(id, "edit") || !factionAllowed(id, faction)) {
      return res.status(403).json({ error: "Edit access required" });
    }

    const action = String((req.body && req.body.action) || "");
    const allowed = new Set([
      "set_dashboard_value", "set_dashboard_values", "add_push",
      "refresh_dashboard", "refresh_awards", "fetch_reports_and_rd",
      "build_payout", "calculate_payout", "build_final_payout"
    ]);
    if (!allowed.has(action)) return res.status(400).json({ error: "Unsupported portal action" });

    try {
      const result = await callSheetBridge(sheetBridges[faction], faction, action, (req.body && req.body.payload) || {});
      return res.json({ ok: true, result });
    } catch (err) {
      return res.status(err.status || 502).json({ error: err.message || "Sheet bridge failed" });
    }
  });

  router.get("/portal/api/session", (req, res) => {
    const id = identity(req);
    if (!id) return res.status(401).json({ error: "Unauthorized" });
    res.json({
      access: id.access,
      factions: id.factions,
      discord_user_id: id.discord_user_id,
      discord_name: id.discord_name || null,
      can_edit: accessAtLeast(id, "edit"),
      can_view_private: accessAtLeast(id, "view")
    });
  });

  return router;
}

module.exports = { createPortalRouter, verifyPortalToken, redactRecord };
