const assert = require("assert");
const { buildRobWarehouseRequest, fetchRobWarehousePage } = require("../railway-api/robWarehouseProxy");

(function testBuildRequestBoundsAndFaction() {
  let r = buildRobWarehouseRequest("", 40664, { from: 1000, to: 2000 });
  assert.strictEqual(r.status, 503);

  r = buildRobWarehouseRequest("https://rob.example", 0, { from: 1000, to: 2000 });
  assert.strictEqual(r.status, 400);

  r = buildRobWarehouseRequest("https://rob.example", 40664, { from: 1000, to: 2000, limit: 5001 });
  assert.strictEqual(r.status, 400);

  r = buildRobWarehouseRequest("https://rob.example", 40664, { from: 1000, to: 7777001 });
  assert.strictEqual(r.status, 400);

  r = buildRobWarehouseRequest("https://rob.example/", 40664, {
    from: 1000,
    to: 2000,
    limit: 2500,
    cursor_started: "1970-01-01T00:18:20.123456+00:00",
    cursor_id: 77,
    ignored: "do-not-forward"
  });
  assert.strictEqual(r.status, 200);
  assert.ok(r.url.startsWith("https://rob.example/internal/payout-attacks?"));
  assert.ok(r.url.includes("faction_id=40664"));
  assert.ok(r.url.includes("limit=2500"));
  assert.ok(r.url.includes("cursor_id=77"));
  assert.ok(r.url.includes("cursor_started=1970-01-01T00%3A18%3A20.123456%2B00%3A00"));
  assert.ok(!r.url.includes("ignored"));
})();

(async function testAuthenticatedProxyContract() {
  const calls = [];
  const goodFetch = async (url, options) => {
    calls.push({ url, options });
    return {
      ok: true,
      status: 200,
      text: async () => JSON.stringify({
        ok: true,
        source: "rob_shared_attack_warehouse",
        attacks: [{ attack_id: 1, chain_value: null, chain_bonus: 100, modifiers_complete: true }],
        next_cursor: null
      })
    };
  };
  const good = await fetchRobWarehousePage({
    baseUrl: "https://rob.example",
    secret: "shared-secret",
    factionId: 46442,
    query: { from: "1000", to: "2000" },
    fetchImpl: goodFetch
  });
  assert.strictEqual(good.status, 200);
  assert.strictEqual(good.body.attacks[0].chain_bonus, 100);
  assert.strictEqual(good.body.attacks[0].chain_value, null);
  assert.strictEqual(calls.length, 1);
  assert.strictEqual(calls[0].options.headers.Authorization, "Bearer shared-secret");
  assert.strictEqual(calls[0].options.redirect, "error");

  const missingSecret = await fetchRobWarehousePage({
    baseUrl: "https://rob.example",
    secret: "",
    factionId: 46442,
    query: { from: "1000", to: "2000" },
    fetchImpl: goodFetch
  });
  assert.strictEqual(missingSecret.status, 503);

  const invalidContract = await fetchRobWarehousePage({
    baseUrl: "https://rob.example",
    secret: "secret",
    factionId: 46442,
    query: { from: "1000", to: "2000" },
    fetchImpl: async () => ({ ok: true, status: 200, text: async () => JSON.stringify({ ok: true, attacks: [] }) })
  });
  assert.strictEqual(invalidContract.status, 502);

  const unavailable = await fetchRobWarehousePage({
    baseUrl: "https://rob.example",
    secret: "secret",
    factionId: 46442,
    query: { from: "1000", to: "2000" },
    fetchImpl: async () => { throw new Error("offline"); }
  });
  assert.strictEqual(unavailable.status, 502);

  console.log("ROB warehouse Railway proxy tests passed.");
})().catch(err => {
  console.error(err);
  process.exitCode = 1;
});
