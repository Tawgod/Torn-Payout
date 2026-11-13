const assert = require("assert");
const fs = require("fs");
const vm = require("vm");

const source = fs.readFileSync("ROB Attack Warehouse.gs", "utf8");

function response(status, body) {
  return {
    getResponseCode: () => status,
    getContentText: () => JSON.stringify(body)
  };
}

function makeContext(props, responses) {
  const calls = [];
  const context = {
    console,
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (name) => Object.prototype.hasOwnProperty.call(props, name) ? props[name] : null
      })
    },
    UrlFetchApp: {
      fetch: (url, options) => {
        calls.push({ url, options });
        if (!responses.length) throw new Error("Unexpected fetch");
        return responses.shift();
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(source, context);
  return { context, calls };
}

(function testDisabledSourceDoesNotFetch() {
  const { context, calls } = makeContext(
    { ROB_ATTACK_WAREHOUSE_ENABLED: "false" },
    []
  );
  assert.strictEqual(context.loadWarDataFromRobWarehouse_(40664, 1000000, 2000000), null);
  assert.strictEqual(calls.length, 0);
})();

(function testPaginationAndRdMapping() {
  const first = {
    ok: true,
    source: "rob_shared_attack_warehouse",
    attacks: [{
      attack_id: 101,
      started_at: "2026-11-01T12:00:00+00:00",
      ended_at: "2026-11-01T12:00:30+00:00",
      attacker_id: 1,
      attacker_name: "Alpha",
      attacker_faction_id: 40664,
      defender_id: 2,
      defender_name: "Beta",
      defender_faction_id: 999,
      result: "Attacked",
      respect_gain: 5000,
      fair_fight: 2,
      war_bonus: 1,
      retaliation: 1,
      group_attack: 1,
      overseas: 1,
      chain_bonus: 100,
      warlord_bonus: 1,
      respect_loss: 0,
      chain_value: null,
      direction: "outgoing",
      modifiers_complete: true
    }],
    next_cursor: { started: "2026-11-01T12:00:00.123456+00:00", attack_id: 101 }
  };
  const second = {
    ok: true,
    source: "rob_shared_attack_warehouse",
    attacks: [{
      attack_id: 101,
      started_at: "2026-11-01T12:00:00+00:00",
      chain_bonus: 100,
      chain_value: null,
      modifiers_complete: true
    }, {
      attack_id: 102,
      started_at: "2026-11-01T12:01:00+00:00",
      ended_at: "2026-11-01T12:01:20+00:00",
      attacker_id: 3,
      attacker_name: "Gamma",
      attacker_faction_id: 40664,
      defender_id: 4,
      defender_name: "Delta",
      defender_faction_id: 998,
      result: "Attacked",
      respect_gain: 25,
      chain_bonus: 1,
      chain_value: 50000,
      direction: "outgoing",
      modifiers_complete: true
    }],
    next_cursor: null
  };
  const { context, calls } = makeContext({
    ROB_ATTACK_WAREHOUSE_ENABLED: "true",
    ROB_ATTACK_WAREHOUSE_URL: "https://chain-watch-web-production.up.railway.app/",
    ROB_PAYOUT_EXPORT_SECRET: "secret"
  }, [response(200, first), response(200, second)]);

  const rows = context.loadWarDataFromRobWarehouse_(40664, 1793530000000, 1793540000000);
  assert.strictEqual(rows.length, 3); // header + 2 unique attacks
  assert.strictEqual(rows[0].length, 19);
  assert.strictEqual(rows[1][0], 101);
  assert.strictEqual(rows[1][10], 5000);
  assert.strictEqual(rows[1][16], 100); // payout Chain Bonus from chain_bonus
  assert.strictEqual(rows[2][0], 102);
  assert.strictEqual(rows[2][16], 1);
  assert.strictEqual(calls.length, 2);
  assert.strictEqual(calls[0].options.headers.Authorization, "Bearer secret");
  assert.ok(calls[0].url.startsWith("https://chain-watch-web-production.up.railway.app/internal/payout-attacks?"));
  assert.ok(calls[1].url.includes("cursor_started=2026-11-01T12%3A00%3A00.123456%2B00%3A00"));
  assert.ok(calls[1].url.includes("cursor_id=101"));
})();


(function testIncompleteModifiersFailClosed() {
  const { context } = makeContext({
    ROB_ATTACK_WAREHOUSE_ENABLED: "true",
    ROB_ATTACK_WAREHOUSE_URL: "https://example.invalid",
    ROB_PAYOUT_EXPORT_SECRET: "secret"
  }, [response(200, {
    ok: true,
    source: "rob_shared_attack_warehouse",
    attacks: [{
      attack_id: 999,
      started_at: "2026-11-01T12:00:00+00:00",
      chain_value: null,
      chain_bonus: null,
      modifiers_complete: false
    }],
    next_cursor: null
  })]);
  assert.throws(
    () => context.loadWarDataFromRobWarehouse_(40664, 1793530000000, 1793540000000),
    /without complete payout modifier data/
  );
})();

(function testNonAdvancingCursorFailsClosed() {
  const bad = {
    ok: true,
    source: "rob_shared_attack_warehouse",
    attacks: [],
    next_cursor: { started: "1970-01-01T00:16:40.000Z", attack_id: 0 }
  };
  const { context } = makeContext({
    ROB_ATTACK_WAREHOUSE_ENABLED: "true",
    ROB_ATTACK_WAREHOUSE_URL: "https://example.invalid",
    ROB_PAYOUT_EXPORT_SECRET: "secret"
  }, [response(200, bad)]);
  assert.throws(
    () => context.loadWarDataFromRobWarehouse_(40664, 1000000, 2000000),
    /cursor did not advance/
  );
})();

(function testImporterHasExplicitFallback() {
  const importer = fs.readFileSync("Importer.gs", "utf8");
  assert.ok(importer.includes('if (!rawData || rawData.length === 0) return false;'));
  assert.ok(!importer.includes('if (!rawData || rawData.length <= 1) return;'));
  assert.ok(importer.includes('return true;'));
  assert.ok(importer.includes('robAttackWarehouseEnabled_()'));
  assert.ok(importer.includes('loadWarDataFromRobWarehouse_'));
  assert.ok(importer.includes('SpreadsheetApp.openById(backendId)'));
  assert.ok(importer.indexOf('robAttackWarehouseEnabled_()') < importer.indexOf('SpreadsheetApp.openById(backendId)'));
})();


(function testOfficialReportAbortsOnFailedImport() {
  const dashboard = fs.readFileSync("Dashboard.gs", "utf8");
  assert.ok(dashboard.includes("const importSucceeded = importWarData(true);"));
  assert.ok(dashboard.includes("if (importSucceeded === false)"));
  assert.ok(dashboard.includes("Existing RD data was not used."));
})();

console.log("ROB attack warehouse importer contract tests passed.");
