function payoutRespectBreakdown_(rawRespect, chainBonus) {
  let raw = Number(rawRespect);
  if (!Number.isFinite(raw)) raw = 0;

  let multiplier = Number(chainBonus);
  if (!Number.isFinite(multiplier) || multiplier <= 0) multiplier = 1;

  const base = multiplier > 1 ? (raw / multiplier) : raw;
  return {
    rawRespect: raw,
    chainBonus: multiplier,
    baseRespect: base,
    bonusRespect: Math.max(0, raw - base)
  };
}
