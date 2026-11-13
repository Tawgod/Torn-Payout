const fs = require("fs");
const vm = require("vm");

const source = fs.readFileSync("PayoutMath.gs", "utf8");
const context = { console };
vm.createContext(context);
vm.runInContext(source, context, { filename: "PayoutMath.gs" });

const {
  isChainMilestoneBonus,
  chainMilestoneBonusRespect,
  ordinaryChainRespect,
  normalizeAttackChainCount,
} = context;

function assertEqual(actual, expected, message) {
  if (!Object.is(actual, expected)) {
    throw new Error(`${message}: expected ${expected}, got ${actual}`);
  }
}

function assertClose(actual, expected, tolerance, message) {
  if (Math.abs(actual - expected) > tolerance) {
    throw new Error(`${message}: expected ${expected}, got ${actual}`);
  }
}

// Ordinary chain scaling must remain ordinary respect, including the 2x scale
// reached late in a long chain.
assertEqual(isChainMilestoneBonus(1), false, "1x is not a milestone");
assertEqual(isChainMilestoneBonus(1.75), false, "1.75x is not a milestone");
assertEqual(isChainMilestoneBonus(2), false, "2x is not a milestone");
assertEqual(ordinaryChainRespect(10.69, 2), 10.69, "2x ordinary respect is preserved");
assertEqual(ordinaryChainRespect(5.5, null), 5.5, "null modifier preserves ordinary respect");
assertEqual(ordinaryChainRespect(5.5, undefined), 5.5, "missing modifier preserves ordinary respect");

// Flat Torn milestone bonuses are completely excluded from ordinary payout
// respect and remain available to the separate Bonus Chain Hits display.
for (const bonus of [10, 20, 40, 80, 160, 320, 640, 1280, 2560, 5120, 10240, 20480, 40960]) {
  assertEqual(isChainMilestoneBonus(bonus), true, `${bonus} is a milestone`);
  assertEqual(chainMilestoneBonusRespect(bonus, bonus), bonus, `${bonus} bonus respect detected`);
  assertEqual(ordinaryChainRespect(bonus, bonus), 0, `${bonus} contributes no ordinary respect`);
}

// Live Resolute 100k-chain fixture (Chain Report 72069225).
const officialRespect = 334034.5807440209;
const liveMilestones = [10, 20, 40, 80, 160, 320, 640, 1280, 2560, 5120, 10240, 20480, 40960];
const milestoneRespect = liveMilestones.reduce(
  (sum, bonus) => sum + chainMilestoneBonusRespect(bonus, bonus),
  0
);
assertEqual(milestoneRespect, 81910, "live milestone total");
assertClose(
  officialRespect - milestoneRespect,
  252124.58074402087,
  1e-9,
  "live ordinary respect after bonus removal"
);

// January 2027 attack-chain compatibility: null/missing remains unknown while a
// real zero remains zero. Payout currently consumes the shared RD warehouse,
// but exports/fixtures must not collapse these states.
assertEqual(normalizeAttackChainCount(null), null, "null chain stays unknown");
assertEqual(normalizeAttackChainCount(undefined), null, "missing chain stays unknown");
assertEqual(normalizeAttackChainCount(""), null, "blank chain stays unknown");
assertEqual(normalizeAttackChainCount(0), 0, "zero chain stays zero");
assertEqual(normalizeAttackChainCount(5), 5, "positive chain preserved");

console.log("Payout math regression tests passed.");


// Verify production script consumers use the shared milestone helpers rather
// than the old "any multiplier > 1" correction.
const calculator = fs.readFileSync("Calculator.gs", "utf8");
const dashboard = fs.readFileSync("Dashboard.gs", "utf8");

if (!calculator.includes("chainMilestoneBonusRespect(e.respect, e.cBonus)")) {
  throw new Error("Calculator must use shared milestone bonus detection");
}
if (calculator.includes("e.respect / e.cBonus")) {
  throw new Error("Calculator still contains legacy chain-multiplier division");
}
if (!dashboard.includes("chainMilestoneBonusRespect(respect, cBonus)")) {
  throw new Error("Dashboard must use shared milestone bonus detection");
}
if (!dashboard.includes("ordinaryChainRespect(respect, cBonus)")) {
  throw new Error("Dashboard fallback must preserve ordinary chain scaling");
}
if (dashboard.includes("respect / (cBonus > 1")) {
  throw new Error("Dashboard still contains legacy chain-multiplier division");
}
if (!dashboard.includes("isChainMilestoneBonus(cBonus)")) {
  throw new Error("Bonus Chain Hits leaderboard must use milestone classification");
}

console.log("Payout script integration assertions passed.");
