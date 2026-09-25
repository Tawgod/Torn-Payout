CREATE TABLE IF NOT EXISTS archive_wars (
  id BIGSERIAL PRIMARY KEY,
  faction_key TEXT NOT NULL,
  war_id BIGINT,
  legacy_key TEXT,
  enemy_name TEXT NOT NULL DEFAULT '',
  enemy_faction_id BIGINT,
  archived_at TIMESTAMPTZ,
  outcome TEXT,
  termed BOOLEAN,
  total_hits INTEGER,
  total_war_hits INTEGER,
  war_score NUMERIC,
  total_revenue NUMERIC(20,2),
  total_payout NUMERIC(20,2),
  faction_profit NUMERIC(20,2),
  caches TEXT,
  official_war_start TEXT,
  official_war_end TEXT,
  summary JSONB NOT NULL DEFAULT '{}'::jsonb,
  payout_top_row JSONB NOT NULL DEFAULT '[]'::jsonb,
  payout_headers JSONB NOT NULL DEFAULT '[]'::jsonb,
  source TEXT NOT NULL DEFAULT 'apps-script',
  published_at TIMESTAMPTZ,
  public_title TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE UNIQUE INDEX IF NOT EXISTS archive_wars_faction_war_id_uq
  ON archive_wars (faction_key, war_id)
  WHERE war_id IS NOT NULL;

CREATE UNIQUE INDEX IF NOT EXISTS archive_wars_faction_legacy_key_uq
  ON archive_wars (faction_key, legacy_key)
  WHERE war_id IS NULL AND legacy_key IS NOT NULL;

CREATE INDEX IF NOT EXISTS archive_wars_faction_archived_idx
  ON archive_wars (faction_key, archived_at DESC NULLS LAST);

CREATE TABLE IF NOT EXISTS archive_member_payouts (
  id BIGSERIAL PRIMARY KEY,
  war_record_id BIGINT NOT NULL REFERENCES archive_wars(id) ON DELETE CASCADE,
  row_order INTEGER NOT NULL,
  member_id BIGINT,
  member_name TEXT,
  row_values JSONB NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (war_record_id, row_order)
);

CREATE INDEX IF NOT EXISTS archive_member_member_id_idx
  ON archive_member_payouts (member_id);

CREATE INDEX IF NOT EXISTS archive_member_war_idx
  ON archive_member_payouts (war_record_id, row_order);

ALTER TABLE archive_wars ADD COLUMN IF NOT EXISTS published_at TIMESTAMPTZ;
ALTER TABLE archive_wars ADD COLUMN IF NOT EXISTS public_title TEXT;
CREATE INDEX IF NOT EXISTS archive_wars_public_idx ON archive_wars (faction_key, published_at DESC) WHERE published_at IS NOT NULL;
