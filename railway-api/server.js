const express = require("express");
const fs = require("fs");
const path = require("path");
const { Pool } = require("pg");

const app = express();
app.use(express.json({ limit: "8mb" }));

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.PGSSL === "disable" ? false : { rejectUnauthorized: false }
});

const API_TOKEN = process.env.ARCHIVE_API_TOKEN || "";
const TORN_PROXY_TOKEN = process.env.TORN_PROXY_TOKEN || "";
const TORN_PROXY_UPSTREAM_URL = (process.env.TORN_PROXY_UPSTREAM_URL || "").replace(/\/+$/, "");
const PORT = Number(process.env.PORT || 3000);

function requireAuth(req, res, next) {
  if (!API_TOKEN) {
    return res.status(503).json({ error: "ARCHIVE_API_TOKEN is not configured" });
  }
  const bearer = (req.get("authorization") || "").replace(/^Bearer\s+/i, "");
  const header = req.get("x-archive-key") || "";
  if (bearer !== API_TOKEN && header !== API_TOKEN) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
}

function asText(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  return s === "" ? null : s;
}

function asNumber(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : null;
}

function asInt(v) {
  const n = asNumber(v);
  return n === null ? null : Math.trunc(n);
}

function asBool(v) {
  if (typeof v === "boolean") return v;
  const s = String(v ?? "").trim().toLowerCase();
  if (["yes", "true", "1", "y"].includes(s)) return true;
  if (["no", "false", "0", "n"].includes(s)) return false;
  return null;
}

function tornApiKeyForFaction(factionKey) {
  const key = String(factionKey || "").trim().toLowerCase();
  if (key === "ironsides") return process.env.TORN_API_KEY_IRONSIDES || "";
  if (key === "resolute") return process.env.TORN_API_KEY_RESOLUTE || "";
  return "";
}

const TORN_ALLOWED_SCOPES = new Set(["torn", "faction", "user"]);
const TORN_ALLOWED_SELECTIONS = new Set([
  "items", "rankedwarreport", "chainreport", "news", "basic",
  "fundsnews", "profile", "rankedwars"
]);

function validateTornRequest(scope, selectionsRaw) {
  if (!scope || !selectionsRaw) return { error: "scope and selections are required" };
  if (!TORN_ALLOWED_SCOPES.has(scope)) return { error: "Unsupported Torn API scope" };
  const selections = selectionsRaw.split(",").map(s => s.trim()).filter(Boolean);
  if (!selections.length || selections.some(s => !TORN_ALLOWED_SELECTIONS.has(s))) {
    return { error: "Unsupported Torn API selection" };
  }
  return { selections };
}

async function fetchTornWithKey(apiKey, scope, id, selections) {
  let url = "https://api.torn.com/" + encodeURIComponent(scope) + "/";
  if (id) url += encodeURIComponent(id);
  url += "?selections=" + encodeURIComponent(selections.join(",")) + "&key=" + encodeURIComponent(apiKey);

  const response = await fetch(url, { headers: { "User-Agent": "Torn-Payout-Railway/1.0" } });
  const bodyText = await response.text();
  let body;
  try { body = JSON.parse(bodyText); } catch (_e) { body = { error: { error: bodyText || "Invalid Torn API response" } }; }
  return { status: response.status, ok: response.ok, body };
}

async function fetchTornFromUpstream(factionKey, scope, id, selections) {
  if (!TORN_PROXY_UPSTREAM_URL || !TORN_PROXY_TOKEN) return null;
  const qs = new URLSearchParams({
    faction: factionKey,
    scope,
    selections: selections.join(",")
  });
  if (id) qs.set("id", id);

  const response = await fetch(TORN_PROXY_UPSTREAM_URL + "/internal/torn?" + qs.toString(), {
    headers: { Authorization: "Bearer " + TORN_PROXY_TOKEN }
  });
  const bodyText = await response.text();
  let body;
  try { body = JSON.parse(bodyText); } catch (_e) { body = { error: bodyText || "Invalid upstream response" }; }
  return { status: response.status, ok: response.ok, body };
}

async function initializeSchema() {
  const schemaPath = path.join(__dirname, "schema.sql");
  const sql = fs.readFileSync(schemaPath, "utf8");
  await pool.query(sql);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function fmtMoney(value) {
  const n = Number(value || 0);
  return Number.isFinite(n) ? "$" + Math.round(n).toLocaleString("en-US") : "";
}

function fmtDate(value) {
  if (!value) return "";
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? escapeHtml(value) : d.toISOString().slice(0, 10);
}

function publicPageShell(title, body) {
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<style>
:root{color-scheme:dark;--bg:#101214;--panel:#181b1f;--muted:#aeb6bf;--line:#30363d;--accent:#d97706;--green:#1f7a4d}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:#f3f4f6;font:15px/1.45 system-ui,-apple-system,Segoe UI,Roboto,sans-serif}
a{color:#fbbf24;text-decoration:none}a:hover{text-decoration:underline}.wrap{max-width:1280px;margin:auto;padding:28px 20px 48px}
h1{margin:0 0 6px;font-size:30px}h2{margin-top:28px}.sub{color:var(--muted);margin-bottom:22px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:18px;margin:14px 0}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px}.stat{background:#121519;border:1px solid var(--line);border-radius:10px;padding:12px}
.stat b{display:block;font-size:18px;margin-top:4px}.label{color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.04em}
table{width:100%;border-collapse:collapse;background:var(--panel);border:1px solid var(--line);border-radius:10px;overflow:hidden}
th,td{padding:9px 10px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}th{background:#20242a;position:sticky;top:0}
th:nth-child(1),td:nth-child(1){text-align:left}.scroll{overflow:auto;border-radius:10px}.pill{display:inline-block;padding:3px 8px;border-radius:999px;background:#2a2f36;color:#d8dee6;font-size:12px}
.search{display:flex;gap:8px;margin:18px 0}.search input{flex:1;padding:10px 12px;border-radius:8px;border:1px solid var(--line);background:#121519;color:white}
.search button{padding:10px 14px;border:0;border-radius:8px;background:var(--accent);color:white;font-weight:700}
footer{color:var(--muted);margin-top:30px;font-size:12px}
</style>
</head><body><div class="wrap">${body}<footer>Railway-backed payout archive</footer></div></body></html>`;
}

async function upsertWar(payload) {
  const factionKey = asText(payload.faction_key);
  if (!factionKey) throw new Error("faction_key is required");

  const warId = asInt(payload.war_id);
  const legacyKey = asText(payload.legacy_key);
  if (warId === null && !legacyKey) throw new Error("war_id or legacy_key is required");

  const summary = payload.summary || {};
  const payout = payload.payout || {};
  const topRow = Array.isArray(payout.top_row) ? payout.top_row : [];
  const headers = Array.isArray(payout.headers) ? payout.headers : [];
  const rows = Array.isArray(payout.rows) ? payout.rows : [];

  const client = await pool.connect();
  try {
    await client.query("BEGIN");

    let existing;
    if (warId !== null) {
      existing = await client.query(
        "SELECT id FROM archive_wars WHERE faction_key=$1 AND war_id=$2",
        [factionKey, warId]
      );
    } else {
      existing = await client.query(
        "SELECT id FROM archive_wars WHERE faction_key=$1 AND legacy_key=$2",
        [factionKey, legacyKey]
      );
    }

    const values = [
      factionKey,
      warId,
      legacyKey,
      asText(payload.enemy_name) || "",
      asInt(payload.enemy_faction_id),
      asText(payload.archived_at),
      asText(summary.outcome),
      asBool(summary.termed),
      asInt(summary.total_hits),
      asInt(summary.total_war_hits),
      asNumber(summary.war_score),
      asNumber(summary.total_revenue),
      asNumber(summary.total_payout),
      asNumber(summary.faction_profit),
      asText(summary.caches),
      asText(summary.official_war_start),
      asText(summary.official_war_end),
      JSON.stringify(summary),
      JSON.stringify(topRow),
      JSON.stringify(headers),
      asText(payload.source) || "apps-script"
    ];

    let warRecordId;
    if (existing.rowCount) {
      warRecordId = existing.rows[0].id;
      await client.query(
        `UPDATE archive_wars SET
          war_id=$2, legacy_key=$3, enemy_name=$4, enemy_faction_id=$5,
          archived_at=COALESCE($6::timestamptz, archived_at),
          outcome=$7, termed=$8, total_hits=$9, total_war_hits=$10,
          war_score=$11, total_revenue=$12, total_payout=$13,
          faction_profit=$14, caches=$15, official_war_start=$16,
          official_war_end=$17, summary=$18::jsonb, payout_top_row=$19::jsonb,
          payout_headers=$20::jsonb, source=$21, updated_at=NOW()
        WHERE id=$22`,
        [...values, warRecordId]
      );
    } else {
      const inserted = await client.query(
        `INSERT INTO archive_wars (
          faction_key, war_id, legacy_key, enemy_name, enemy_faction_id,
          archived_at, outcome, termed, total_hits, total_war_hits,
          war_score, total_revenue, total_payout, faction_profit, caches,
          official_war_start, official_war_end, summary, payout_top_row,
          payout_headers, source
        ) VALUES (
          $1,$2,$3,$4,$5,$6::timestamptz,$7,$8,$9,$10,$11,$12,$13,$14,$15,
          $16,$17,$18::jsonb,$19::jsonb,$20::jsonb,$21
        ) RETURNING id`,
        values
      );
      warRecordId = inserted.rows[0].id;
    }

    await client.query("DELETE FROM archive_member_payouts WHERE war_record_id=$1", [warRecordId]);

    for (let i = 0; i < rows.length; i++) {
      const row = Array.isArray(rows[i]) ? rows[i] : [];
      const memberId = asInt(row[0]);
      const memberName = asText(row[1]);
      await client.query(
        `INSERT INTO archive_member_payouts
          (war_record_id, row_order, member_id, member_name, row_values)
         VALUES ($1,$2,$3,$4,$5::jsonb)`,
        [warRecordId, i + 1, memberId, memberName, JSON.stringify(row)]
      );
    }

    await client.query("COMMIT");
    return { id: warRecordId, members: rows.length };
  } catch (err) {
    await client.query("ROLLBACK");
    throw err;
  } finally {
    client.release();
  }
}

app.get("/health", async (_req, res) => {
  try {
    await pool.query("SELECT 1");
    res.json({ ok: true, service: "torn-payout-archive-api" });
  } catch (err) {
    res.status(503).json({ ok: false, error: err.message });
  }
});


app.get("/", (_req, res) => {
  res.type("html").send(publicPageShell("Torn Payout Archive",
    "<h1>Torn Payout Archive</h1><p class=\"sub\">Use a faction-specific public payout link.</p>"
  ));
});

app.get("/public/:faction", async (req, res) => {
  try {
    const factionKey = asText(req.params.faction);
    const q = asText(req.query.q);
    const params = [factionKey];
    let where = "faction_key=$1 AND published_at IS NOT NULL";
    if (q) {
      params.push(`%${q}%`);
      where += " AND (enemy_name ILIKE $2 OR CAST(war_id AS TEXT) ILIKE $2)";
    }

    const result = await pool.query(
      `SELECT id, war_id, legacy_key, enemy_name, archived_at, published_at, outcome,
              total_hits, war_score, total_payout
         FROM archive_wars
        WHERE ${where}
        ORDER BY COALESCE(archived_at, published_at) DESC NULLS LAST, id DESC
        LIMIT 500`,
      params
    );

    const rows = result.rows.map(w => {
      const label = w.war_id || w.legacy_key || w.id;
      return `<tr>
        <td><a href="/public/${encodeURIComponent(factionKey)}/${w.id}">${escapeHtml(w.enemy_name || "Unknown")}</a></td>
        <td>${escapeHtml(label)}</td>
        <td>${fmtDate(w.archived_at || w.published_at)}</td>
        <td>${escapeHtml(w.outcome || "")}</td>
        <td>${Number(w.total_hits || 0).toLocaleString("en-US")}</td>
        <td>${Number(w.war_score || 0).toLocaleString("en-US")}</td>
        <td>${fmtMoney(w.total_payout)}</td>
      </tr>`;
    }).join("");

    const body = `
      <h1>${escapeHtml(factionKey)} Payout History</h1>
      <p class="sub">Published war payouts and historical results.</p>
      <form class="search" method="get">
        <input name="q" value="${escapeHtml(q || "")}" placeholder="Search enemy or War ID">
        <button type="submit">Search</button>
      </form>
      <div class="scroll"><table>
        <thead><tr><th>Opponent</th><th>War</th><th>Date</th><th>Outcome</th><th>Hits</th><th>Score</th><th>Total Payout</th></tr></thead>
        <tbody>${rows || '<tr><td colspan="7">No published payouts yet.</td></tr>'}</tbody>
      </table></div>`;

    res.type("html").send(publicPageShell(factionKey + " Payout History", body));
  } catch (err) {
    console.error(err);
    res.status(500).type("html").send(publicPageShell("Error", "<h1>Unable to load payout history</h1>"));
  }
});

app.get("/public/:faction/:recordId", async (req, res) => {
  try {
    const factionKey = asText(req.params.faction);
    const recordId = asInt(req.params.recordId);
    const warResult = await pool.query(
      `SELECT * FROM archive_wars WHERE id=$1 AND faction_key=$2 AND published_at IS NOT NULL`,
      [recordId, factionKey]
    );
    if (!warResult.rowCount) return res.status(404).type("html").send(publicPageShell("Not found", "<h1>Payout not found</h1>"));

    const war = warResult.rows[0];
    const memberResult = await pool.query(
      `SELECT row_values FROM archive_member_payouts WHERE war_record_id=$1 ORDER BY row_order`,
      [recordId]
    );

    const headers = Array.isArray(war.payout_headers) ? war.payout_headers : [];
    const top = Array.isArray(war.payout_top_row) ? war.payout_top_row : [];
    const hiddenNames = new Set(["member id", "total points"]);
    const visible = headers.map((h, i) => ({h: String(h || ""), i}))
      .filter(x => x.h && !hiddenNames.has(x.h.toLowerCase().trim()));

    const th = visible.map(x => {
      const weight = top[x.i];
      const w = weight !== "" && weight !== null && weight !== undefined && !Number.isNaN(Number(weight))
        ? `<div class="label">Weight ${escapeHtml(weight)}</div>` : "";
      return `<th>${escapeHtml(x.h)}${w}</th>`;
    }).join("");

    const bodyRows = memberResult.rows.map(r => {
      const row = Array.isArray(r.row_values) ? r.row_values : [];
      return "<tr>" + visible.map(x => {
        const header = x.h.toLowerCase();
        let v = row[x.i];
        if (header.includes("payout") || header.includes("$")) v = fmtMoney(v);
        else if (header.includes("contribution") && Number.isFinite(Number(v))) v = (Number(v) * 100).toFixed(2) + "%";
        else if (typeof v === "number") v = Number.isInteger(v) ? v.toLocaleString("en-US") : v.toLocaleString("en-US", {maximumFractionDigits:2});
        return `<td>${escapeHtml(v ?? "")}</td>`;
      }).join("") + "</tr>";
    }).join("");

    const s = war.summary || {};
    const warLabel = war.war_id || war.legacy_key || war.id;
    const body = `
      <p><a href="/public/${encodeURIComponent(factionKey)}">← All payouts</a></p>
      <h1>${escapeHtml(war.enemy_name || "Unknown")}</h1>
      <p class="sub">War / archive record ${escapeHtml(warLabel)}</p>
      <div class="grid">
        <div class="stat"><span class="label">Outcome</span><b>${escapeHtml(s.outcome || war.outcome || "")}</b></div>
        <div class="stat"><span class="label">Total Hits</span><b>${Number(s.total_hits || war.total_hits || 0).toLocaleString("en-US")}</b></div>
        <div class="stat"><span class="label">War Score</span><b>${Number(s.war_score || war.war_score || 0).toLocaleString("en-US")}</b></div>
        <div class="stat"><span class="label">Total Payout</span><b>${fmtMoney(s.total_payout || war.total_payout)}</b></div>
      </div>
      <h2>Member Payouts</h2>
      <div class="scroll"><table><thead><tr>${th}</tr></thead><tbody>${bodyRows}</tbody></table></div>`;

    res.type("html").send(publicPageShell((war.enemy_name || "Payout") + " — " + factionKey, body));
  } catch (err) {
    console.error(err);
    res.status(500).type("html").send(publicPageShell("Error", "<h1>Unable to load payout</h1>"));
  }
});

app.get("/internal/torn", async (req, res) => {
  const bearer = (req.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!TORN_PROXY_TOKEN || bearer !== TORN_PROXY_TOKEN) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  try {
    const factionKey = asText(req.query.faction);
    const scope = asText(req.query.scope);
    const id = asText(req.query.id);
    const selectionsRaw = asText(req.query.selections);
    if (!factionKey) return res.status(400).json({ error: "faction is required" });

    const valid = validateTornRequest(scope, selectionsRaw);
    if (valid.error) return res.status(400).json({ error: valid.error });

    const apiKey = tornApiKeyForFaction(factionKey);
    if (!apiKey) {
      return res.status(503).json({ error: "No Railway Torn API key configured for faction " + factionKey });
    }

    const result = await fetchTornWithKey(apiKey, scope, id, valid.selections);
    return res.status(result.status).json(result.body);
  } catch (err) {
    console.error(err);
    return res.status(502).json({ error: "Internal Torn proxy failed" });
  }
});

app.use("/api", requireAuth);

app.get("/api/torn", async (req, res) => {
  try {
    const factionKey = asText(req.query.faction);
    const scope = asText(req.query.scope);
    const id = asText(req.query.id);
    const selectionsRaw = asText(req.query.selections);

    if (!factionKey) return res.status(400).json({ error: "faction is required" });

    const valid = validateTornRequest(scope, selectionsRaw);
    if (valid.error) return res.status(400).json({ error: valid.error });

    const apiKey = tornApiKeyForFaction(factionKey);
    let result;

    if (apiKey) {
      result = await fetchTornWithKey(apiKey, scope, id, valid.selections);
    } else {
      result = await fetchTornFromUpstream(factionKey, scope, id, valid.selections);
      if (!result) {
        return res.status(503).json({ error: "No Railway Torn API key or upstream proxy configured for faction " + factionKey });
      }
    }

    return res.status(result.status).json(result.body);
  } catch (err) {
    console.error(err);
    return res.status(502).json({ error: "Torn API proxy failed" });
  }
});

app.get("/api/wars", async (req, res) => {
  const factionKey = asText(req.query.faction);
  if (!factionKey) return res.status(400).json({ error: "faction query parameter is required" });

  const limit = Math.min(Math.max(asInt(req.query.limit) || 100, 1), 500);
  const q = asText(req.query.q);

  const params = [factionKey, limit];
  let where = "faction_key=$1";
  if (q) {
    params.push(`%${q}%`);
    where += " AND (enemy_name ILIKE $3 OR CAST(war_id AS TEXT) ILIKE $3)";
  }

  const result = await pool.query(
    `SELECT id AS record_id, war_id, legacy_key, enemy_name, archived_at, outcome, termed,
            total_hits, total_war_hits, war_score, total_revenue,
            total_payout, faction_profit, caches, source
       FROM archive_wars
      WHERE ${where}
      ORDER BY archived_at DESC NULLS LAST, war_id DESC NULLS LAST
      LIMIT $2`,
    params
  );
  res.json({ faction_key: factionKey, wars: result.rows });
});

async function warResponseByRecordId(recordId) {
  const warResult = await pool.query(
    `SELECT * FROM archive_wars WHERE id=$1`,
    [recordId]
  );
  if (!warResult.rowCount) return null;

  const war = warResult.rows[0];
  const memberResult = await pool.query(
    `SELECT row_values FROM archive_member_payouts
      WHERE war_record_id=$1 ORDER BY row_order`,
    [war.id]
  );

  return {
    record_id: war.id,
    faction_key: war.faction_key,
    war_id: war.war_id,
    legacy_key: war.legacy_key,
    enemy_name: war.enemy_name,
    enemy_faction_id: war.enemy_faction_id,
    archived_at: war.archived_at,
    summary: war.summary,
    source: war.source,
    payout: {
      top_row: war.payout_top_row,
      headers: war.payout_headers,
      rows: memberResult.rows.map(r => r.row_values)
    }
  };
}

app.get("/api/records/:recordId", async (req, res) => {
  const factionKey = asText(req.query.faction);
  const recordId = asInt(req.params.recordId);
  if (!factionKey || recordId === null) {
    return res.status(400).json({ error: "valid faction and record ID are required" });
  }

  const record = await warResponseByRecordId(recordId);
  if (!record || record.faction_key !== factionKey) {
    return res.status(404).json({ error: "Archive record not found" });
  }
  res.json(record);
});

app.get("/api/wars/:warId", async (req, res) => {
  const factionKey = asText(req.query.faction);
  const warId = asInt(req.params.warId);
  if (!factionKey || warId === null) {
    return res.status(400).json({ error: "valid faction and war ID are required" });
  }

  const warResult = await pool.query(
    `SELECT * FROM archive_wars WHERE faction_key=$1 AND war_id=$2`,
    [factionKey, warId]
  );
  if (!warResult.rowCount) return res.status(404).json({ error: "War not found" });

  const war = warResult.rows[0];
  const memberResult = await pool.query(
    `SELECT row_values FROM archive_member_payouts
      WHERE war_record_id=$1 ORDER BY row_order`,
    [war.id]
  );

  res.json({
    faction_key: war.faction_key,
    war_id: war.war_id,
    enemy_name: war.enemy_name,
    enemy_faction_id: war.enemy_faction_id,
    archived_at: war.archived_at,
    summary: war.summary,
    source: war.source,
    payout: {
      top_row: war.payout_top_row,
      headers: war.payout_headers,
      rows: memberResult.rows.map(r => r.row_values)
    }
  });
});

app.post("/api/wars/:warId/publish", async (req, res) => {
  const factionKey = asText(req.body && req.body.faction_key);
  const warId = asInt(req.params.warId);
  if (!factionKey || warId === null) {
    return res.status(400).json({ error: "faction_key and valid war ID are required" });
  }
  const result = await pool.query(
    `UPDATE archive_wars
        SET published_at=NOW(), public_title=COALESCE($3, public_title), updated_at=NOW()
      WHERE faction_key=$1 AND war_id=$2
      RETURNING id, published_at`,
    [factionKey, warId, asText(req.body && req.body.public_title)]
  );
  if (!result.rowCount) return res.status(404).json({ error: "War not found" });
  res.json({ ok: true, ...result.rows[0] });
});

app.post("/api/records/:recordId/publish", async (req, res) => {
  const factionKey = asText(req.body && req.body.faction_key);
  const recordId = asInt(req.params.recordId);
  if (!factionKey || recordId === null) {
    return res.status(400).json({ error: "faction_key and valid record ID are required" });
  }
  const result = await pool.query(
    `UPDATE archive_wars
        SET published_at=NOW(), public_title=COALESCE($3, public_title), updated_at=NOW()
      WHERE faction_key=$1 AND id=$2
      RETURNING id, published_at`,
    [factionKey, recordId, asText(req.body && req.body.public_title)]
  );
  if (!result.rowCount) return res.status(404).json({ error: "Archive record not found" });
  res.json({ ok: true, ...result.rows[0] });
});

app.post("/api/wars", async (req, res) => {
  try {
    const result = await upsertWar(req.body || {});
    res.status(201).json({ ok: true, ...result });
  } catch (err) {
    console.error(err);
    res.status(400).json({ ok: false, error: err.message });
  }
});

initializeSchema()
  .then(() => {
    app.listen(PORT, "0.0.0.0", () => {
      console.log(`Archive API listening on port ${PORT}`);
    });
  })
  .catch(err => {
    console.error("Failed to initialize archive schema:", err);
    process.exit(1);
  });
