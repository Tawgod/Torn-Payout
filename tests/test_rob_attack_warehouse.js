const assert = require("assert");
const fs = require("fs");
const vm = require("vm");

const source = fs.readFileSync("ROB Attack Warehouse.gs", "utf8");

function makeContext(props, responses, factionKey = "ironsides", configToggle = "") {
  const calls = [];
  const queue = responses.slice();
  const context = {
    console,
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (name) => Object.prototype.hasOwnProperty.call(props, name) ? props[name] : null
      })
    },
    SETTINGS: { configSheet: "Config", robWarehouseEnabledCell: "B9" },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: (name) => name === "Config" ? {
          getRange: (a1) => ({
            getDisplayValue: () => a1 === "B9" ? configToggle : ""
          })
        } : null
      })
    },
    getArchiveDatabaseConfig_: () => ({
      url: "https://payout.example.invalid",
      token: "archive-token",
      factionKey
    }),
    archiveApiRequest_: (path) => {
      calls.push(path);
      if (!queue.length) throw new Error("Unexpected archiveApiRequest_");
      const next = queue.shift();
      if (next instanceof Error) throw next;
      return next;
    }
  };
  vm.createContext(context);
  vm.runInContext(source, context);
  return { context, calls };
}

(function testConfigToggleCanEnableWithoutSecretProperty() {
  const { context } = makeContext({}, [], "ironsides", "ON");
  assert.strictEqual(context.robAttackWarehouseEnabled_(), true);
})();

(function testScriptPropertyExplicitFalseOverridesConfig() {
  const { context } = makeContext({ ROB_ATTACK_WAREHOUSE_ENABLED: "false" }, [], "ironsides", "ON");
  assert.strictEqual(context.robAttackWarehouseEnabled_(), false);
})();

(function testBlankConfigDefaultsOff() {
  const { context } = makeContext({}, [], "ironsides", "");
  assert.strictEqual(context.robAttackWarehouseEnabled_(), false);
})();

(function testDisabledSourceDoesNotCallRailway() {
  const { context, calls } = makeContext({ ROB_ATTACK_WAREHOUSE_ENABLED: "false" }, []);
  assert.strictEqual(context.loadWarDataFromRobWarehouse_(40664, 1000000, 2000000), null);
  assert.strictEqual(calls.length, 0);
})();

(function testFactionNamespaceMismatchFailsClosed() {
  const { context } = makeContext({ ROB_ATTACK_WAREHOUSE_ENABLED: "true" }, [], "ironsides");
  assert.throws(
    () => context.loadWarDataFromRobWarehouse_(46442, 1793530000000, 1793540000000),
    /does not match the configured Railway archive namespace/
  );
})();

(function testPaginationAndRdMappingThroughArchiveApi() {
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
      fair_fight: 1,
      war_bonus: 1,
      retaliation: 1,
      group_attack: 1,
      overseas: 1,
      chain_bonus: 1,
      warlord_bonus: 0,
      respect_loss: 0,
      chain_value: 50000,
      direction: "outgoing",
      modifiers_complete: true
    }],
    next_cursor: null
  };
  const { context, calls } = makeContext(
    { ROB_ATTACK_WAREHOUSE_ENABLED: "true" },
    [first, second]
  );

  const rows = context.loadWarDataFromRobWarehouse_(40664, 1793530000000, 1793540000000);
  assert.strictEqual(rows.length, 3);
  assert.strictEqual(rows[0].length, 19);
  assert.strictEqual(rows[1][0], 101);
  assert.strictEqual(rows[1][10], 5000);
  assert.strictEqual(rows[1][16], 100);
  assert.strictEqual(rows[2][0], 102);
  assert.strictEqual(rows[2][16], 1);
  assert.strictEqual(calls.length, 2);
  assert.ok(calls[0].startsWith("/api/rob-attack-warehouse?faction=ironsides&"));
  assert.ok(calls[1].includes("cursor_started=2026-11-01T12%3A00%3A00.123456%2B00%3A00"));
  assert.ok(calls[1].includes("cursor_id=101"));
})();

(function testIncompleteModifiersFailClosed() {
  const { context } = makeContext(
    { ROB_ATTACK_WAREHOUSE_ENABLED: "true" },
    [{
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
    }]
  );
  assert.throws(
    () => context.loadWarDataFromRobWarehouse_(40664, 1793530000000, 1793540000000),
    /without complete payout modifier data/
  );
})();

(function testSubMillisecondCursorAdvanceDoesNotUseTruncatedDateOrdering() {
  const { context, calls } = makeContext(
    { ROB_ATTACK_WAREHOUSE_ENABLED: "true" },
    [
      { ok: true, source: "rob_shared_attack_warehouse", attacks: [], next_cursor: { started: "2026-11-01T12:00:00.123456+00:00", attack_id: 200 } },
      { ok: true, source: "rob_shared_attack_warehouse", attacks: [], next_cursor: { started: "2026-11-01T12:00:00.123457+00:00", attack_id: 100 } },
      { ok: true, source: "rob_shared_attack_warehouse", attacks: [], next_cursor: null }
    ]
  );
  const rows = context.loadWarDataFromRobWarehouse_(40664, 1793530000000, 1793540000000);
  assert.strictEqual(rows.length, 1);
  assert.strictEqual(calls.length, 3);
  assert.ok(calls[2].includes("cursor_started=2026-11-01T12%3A00%3A00.123457%2B00%3A00"));
  assert.ok(calls[2].includes("cursor_id=100"));
})();

(function testNonAdvancingCursorFailsClosed() {
  const { context } = makeContext(
    { ROB_ATTACK_WAREHOUSE_ENABLED: "true" },
    [{ ok: true, source: "rob_shared_attack_warehouse", attacks: [], next_cursor: { started: "1970-01-01T00:16:40.000Z", attack_id: 0 } }]
  );
  assert.throws(
    () => context.loadWarDataFromRobWarehouse_(40664, 1000000, 2000000),
    /cursor did not advance/
  );
})();

(function testImporterHasExplicitFallbackAndFailClosedStatus() {
  const importer = fs.readFileSync("Importer.gs", "utf8");
  assert.ok(importer.includes('if (!rawData || rawData.length === 0) return false;'));
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
