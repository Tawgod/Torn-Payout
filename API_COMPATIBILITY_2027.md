# Torn-Payout — API deprecations effective January 1, 2027

**Status: BUG / VERIFY — tracked only; no code fix or deployment claimed.**  
**Primary source:** https://www.torn.com/api.html#patchNotes  
**Shared ROB compatibility tracker:** https://github.com/Tawgod/ironsides-faction-tools/blob/main/API_COMPATIBILITY_2027.md  
**Required development policies:** https://github.com/Tawgod/ironsides-faction-tools/blob/main/TORN_API_POLICY.md and https://github.com/Tawgod/ironsides-faction-tools/blob/main/API_USAGE.md

## Required before 2027

1. **Attack `chain: 0` becomes `chain: null`** in API v2 `user/attacks` and `faction/attacks`. Audit all Apps Script calculator/dashboard code, raw attack import, Railway attack warehouse consumer, legacy attack sheet, and spreadsheet formulas. Accept `null`, missing, `0`, positive and historical values. Only substitute zero for display or math when the field semantics justify it; do not lose a meaningful distinction or reject an attack.
2. **Check `faction/warfare` usage.** Search all Scripts, collectors, report pickers and older source branches. If present, migrate to official replacement selections and verify keys/scopes. Do not assume that the presence of `rankedwars` alone proves no legacy warfare dependence.
3. **`torn/itemdetails` becomes an array response.** Search item valuation, unique equipment and any report import. Normalize array and old object shapes without dropping UID mappings. Do not add global unique weapon/armor polling.
4. **Data-quality correction:** some interrupted attacks previously appeared as assists. Compare newly imported attack and assist classification before any historical adjustment. Preserve the existing payout rule: chain-bonus respect is recorded in **Bonus chain hits** only, not in ordinary payout respect.
5. **Other deprecations:** check whether `torn/stats` communication fields `friends`, `coworkers`, `spouses` are consumed; replacements must be taken from current Torn documentation before migration.

## Tests and deployment gate

- [ ] Locate and inventory affected Apps Script files and spreadsheet formulas; do not assume archived scripts match the live Google Sheet
- [ ] Use fixtures `chain: null`, `0`, `5`, missing `chain` and multi-page attack data
- [ ] Cross-check modern collector versus legacy attack sheet for identical time windows
- [ ] Verify bonus hits (including 50k/100k), respect totals, dashboards and no double-counting
- [ ] Verify null safety in script exports, CSV, report views and historical reimports
- [ ] Check replacement selections, Torn key access and API impact against canonical policies
- [ ] Keep Torn API calls and GitHub Actions runs minimal; use existing shared attack data where practical
- [ ] Run sandbox/test validations; obtain deployment approval where needed; verify live report after rollout

**Target readiness:** December 1, 2026, with a final January 1 deployment compatibility check. Do not mark LIVE merely because these notes exist.
