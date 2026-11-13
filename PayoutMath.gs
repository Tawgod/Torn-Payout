// ==========================================
// SHARED PAYOUT / CHAIN RESPECT HELPERS
// ==========================================

// Torn flat chain-milestone respect awards. Ordinary chain scaling is a
// multiplier between 1x and 2x and must not be mistaken for a milestone.
const CHAIN_MILESTONE_BONUS_VALUES = Object.freeze([
  10, 20, 40, 80, 160, 320, 640,
  1280, 2560, 5120, 10240, 20480, 40960
]);

function numericOrZero(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function isChainMilestoneBonus(chainBonus) {
  const value = numericOrZero(chainBonus);
  return CHAIN_MILESTONE_BONUS_VALUES.indexOf(value) !== -1;
}

function chainMilestoneBonusRespect(respect, chainBonus) {
  const value = Math.max(0, numericOrZero(respect));
  return isChainMilestoneBonus(chainBonus) ? value : 0;
}

function ordinaryChainRespect(respect, chainBonus) {
  const value = Math.max(0, numericOrZero(respect));
  return Math.max(0, value - chainMilestoneBonusRespect(value, chainBonus));
}

// The payout sheet consumes the shared RD warehouse rather than the raw v2
// attack chain counter. Keep a normalization helper for fixtures/exports so
// null or missing future API chain values stay distinct from a real zero.
function normalizeAttackChainCount(value) {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}
