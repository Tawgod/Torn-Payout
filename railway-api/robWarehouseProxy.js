function intParam(value, fallback = null) {
  if (value === undefined || value === null || value === "") return fallback;
  const n = Number(value);
  if (!Number.isInteger(n)) return null;
  return n;
}

function buildRobWarehouseRequest(baseUrl, factionId, query = {}) {
  const base = String(baseUrl || "").trim().replace(/\/+$/, "");
  if (!/^https:\/\//i.test(base)) {
    return { error: "ROB payout export URL is not configured", status: 503 };
  }
  if (!Number.isInteger(factionId) || factionId <= 0) {
    return { error: "Unsupported faction", status: 400 };
  }

  const from = intParam(query.from);
  const to = intParam(query.to);
  const limit = intParam(query.limit, 5000);
  const cursorId = intParam(query.cursor_id, 0);
  const cursorStarted = String(query.cursor_started || "").trim();

  if (!from || !to || to <= from) {
    return { error: "Invalid payout time window", status: 400 };
  }
  if (to - from > 90 * 86400) {
    return { error: "Payout time window exceeds 90 days", status: 400 };
  }
  if (!limit || limit < 1 || limit > 5000) {
    return { error: "Invalid payout export limit", status: 400 };
  }
  if (cursorId === null || cursorId < 0) {
    return { error: "Invalid payout export cursor", status: 400 };
  }

  const params = new URLSearchParams({
    faction_id: String(factionId),
    from: String(from),
    to: String(to),
    limit: String(limit),
    cursor_id: String(cursorId)
  });
  if (cursorStarted) params.set("cursor_started", cursorStarted);

  return {
    url: base + "/internal/payout-attacks?" + params.toString(),
    status: 200
  };
}

async function fetchRobWarehousePage({ baseUrl, secret, factionId, query, fetchImpl = fetch }) {
  const request = buildRobWarehouseRequest(baseUrl, factionId, query);
  if (request.error) return { status: request.status, body: { error: request.error } };
  if (!secret) return { status: 503, body: { error: "ROB payout export secret is not configured" } };

  let response;
  try {
    response = await fetchImpl(request.url, {
      method: "GET",
      headers: {
        Authorization: "Bearer " + secret,
        Accept: "application/json",
        "User-Agent": "Torn-Payout-Warehouse-Proxy/1.0"
      },
      redirect: "error"
    });
  } catch (_err) {
    return { status: 502, body: { error: "ROB payout warehouse is unavailable" } };
  }

  const text = await response.text();
  let body;
  try {
    body = JSON.parse(text);
  } catch (_err) {
    return { status: 502, body: { error: "ROB payout warehouse returned invalid JSON" } };
  }

  if (!response.ok) {
    return {
      status: response.status >= 400 && response.status < 600 ? response.status : 502,
      body: { error: body && body.error ? String(body.error) : "ROB payout warehouse request failed" }
    };
  }
  if (!body || body.ok !== true || body.source !== "rob_shared_attack_warehouse" || !Array.isArray(body.attacks)) {
    return { status: 502, body: { error: "ROB payout warehouse returned an invalid contract" } };
  }

  return { status: 200, body };
}

module.exports = { buildRobWarehouseRequest, fetchRobWarehousePage };
