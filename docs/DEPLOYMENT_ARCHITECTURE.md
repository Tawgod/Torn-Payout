# Deployment Architecture

## Goal

Maintain one source of truth for Torn payout code while keeping:
- a dedicated test sheet for in-progress changes,
- Ironsides production,
- Resolute production,
- separate Railway staging/production backends,
- the existing Google sheets as backups.

## Recommended structure

GitHub is the only editable source of code.

Targets:
1. Test Google Sheet / Apps Script project
2. Ironsides production Google Sheet / Apps Script project
3. Resolute production Google Sheet / Apps Script project

Railway:
- staging: test database + payout API
- production: production database + payout API

The bound Apps Script projects remain because the spreadsheet custom menu is container-bound. They are deployment targets, not independent source copies.

## Release flow

1. Work happens on a feature branch.
2. CI deploys that branch to the Test Sheet Apps Script project and Railway staging.
3. Validate calculations, archive/recall, public payout output, and API behavior in Test.
4. Merge/approve the release.
5. CI deploys the exact tested revision to both Ironsides and Resolute production Apps Script projects and Railway production.
6. No manual copy/paste between production sheets.

## Environment configuration

Do not store Torn API keys in Google Sheets.

Production Railway holds faction API keys and the payout API selects the key from the sheet's faction namespace:
- ironsides -> TORN_API_KEY_IRONSIDES
- resolute -> TORN_API_KEY_RESOLUTE

Apps Script stores only non-Torn credentials/config needed to call the payout API:
- ARCHIVE_API_URL
- ARCHIVE_API_TOKEN
- ARCHIVE_FACTION_KEY
- PAYOUT_ENVIRONMENT

The test sheet should use:
- PAYOUT_ENVIRONMENT=test
- staging API URL/token
- an isolated staging database

Production sheets should use:
- PAYOUT_ENVIRONMENT=production
- production API URL/token
- the production database

## Future automated deployment

Use clasp or the Apps Script API from CI with three Apps Script project IDs stored as repository secrets:
- APPS_SCRIPT_TEST_ID
- APPS_SCRIPT_IRONSIDES_ID
- APPS_SCRIPT_RESOLUTE_ID

Suggested rules:
- feature/development branch -> Test target only
- release/main promotion -> Ironsides + Resolute targets together

This means production scripts are never edited separately by hand.

## Backup policy

The original production Google workbooks and legacy archive/public workbooks remain untouched unless a specific migration/promotion action explicitly targets them. Database migration reads legacy archives but does not delete or rewrite them.
