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

async function initializeSchema() {
  const schemaPath = path.join(__dirname, "schema.sql");
  const sql = fs.readFileSync(schemaPath, "utf8");
  await pool.query(sql);
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

app.use("/api", requireAuth);

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
