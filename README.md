# Torn Payout Sheet Script Archive

## Shared ROB attack warehouse

The payout attack importer now reads the Resolute Family ROB attack warehouse instead of opening a separate backend attack-data spreadsheet.

Configure these **Apps Script Script Properties** before deploying the importer:

- `ROB_PAYOUT_EXPORT_SECRET` — dedicated bearer secret shared with the ROB web service.
- `ROB_PAYOUT_EXPORT_URL` — optional override. If omitted, the importer uses the permanent direct Railway endpoint:
  `https://chain-watch-web-production.up.railway.app/internal/payout-attacks`

The secret must not be stored in spreadsheet cells, source code, or URLs.

The payout sheet's own faction ID remains in Config cell **B7**. Config B2 is retained only for backwards compatibility with older sheets and is no longer read by the attack importer.

### Data safety

- The importer makes no direct Torn API call for attack data.
- ROB's stored attack warehouse remains the source of truth.
- Import windows continue to use attack **end time**, matching the legacy payout importer.
- Payout `Chain Bonus` comes from the stored payout modifier, not the top-level Torn chain counter.
- If ROB reports incomplete payout modifiers for any attack, import stops rather than guessing multiplier values.
- The endpoint is intended for direct Railway server-to-server use and is not routed through the Timcorp Projects gateway.
