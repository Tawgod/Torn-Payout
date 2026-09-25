# Torn Payout Archive Database

This branch moves historical war/payout storage from one-Google-tab-per-war into the existing Railway Postgres database.

## Architecture

Main payout Google Sheet (Apps Script) → HTTPS Archive API on Railway → existing Postgres service.

The database stores the war summary, complete Payouts header/weight rows, every member payout row in original column order, and legacy payout-only records without a known numeric War ID.

## Railway service

Service: payout-archive-api

Required environment variables:

- DATABASE_URL=${{Postgres.DATABASE_URL}}
- ARCHIVE_API_TOKEN=<secret>
- PGSSL=disable for the current internal Railway connection

The service runs railway-api/schema.sql on startup. The schema statements are idempotent.

## Apps Script setup

After the Apps Script branch is installed in the bound spreadsheet:

1. Reload the spreadsheet.
2. Open Faction Tools → Database Archive → Configure Railway Database.
3. Enter the Railway service URL.
4. Enter the Railway ARCHIVE_API_TOKEN.
5. Enter a faction key such as ironsides or resolute.

The URL/token/faction key are stored in Google Apps Script Document Properties, not visible spreadsheet cells.

## Legacy migration

1. Run Migrate Legacy Google Archives. It reads Config!B3 and imports each numeric war tab. If Config!B8 is configured, the matching public payout tab supplies the saved weight row. Existing faction/War ID records are updated instead of duplicated.
2. Run Migrate Public-Only Payout Tabs. It reads Config!B8 and imports old public payout tabs that have no matching archived war. These are stored with a legacy key and can still be restored from the archive index.
3. Run Refresh War Archive Index and compare it with the old archive summary.
4. Repeat from each faction's main payout workbook using a distinct faction key.

## Restoring a prior war

Refresh War Archive Index creates/updates the local War Archive tab. Select a row and choose Load Selected Prior War. The script restores the archived Payouts top/weight row, headers, every member payout row, and the archived Dashboard summary values that are available.

Historical RD/raw attacks are not reconstructed because the legacy archive workbooks never stored those rows.

## New archive workflow

ARCHIVE TO DATABASE & RESET snapshots the current war to Postgres, waits for a successful API response, and only then resets the local workbook. If the database write fails, the workbook is not reset.

The old Google Sheet archive/publish commands remain available as Legacy commands during the transition.
