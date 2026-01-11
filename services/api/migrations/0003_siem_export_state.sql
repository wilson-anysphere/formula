-- Track per-org SIEM export cursors + failure backoff state.

CREATE TABLE IF NOT EXISTS org_siem_export_state (
  org_id uuid PRIMARY KEY REFERENCES organizations(id) ON DELETE CASCADE,
  last_created_at timestamptz,
  last_event_id uuid,
  updated_at timestamptz NOT NULL DEFAULT now(),
  last_error text,
  consecutive_failures integer NOT NULL DEFAULT 0,
  disabled_until timestamptz
);

CREATE INDEX IF NOT EXISTS org_siem_export_state_disabled_until_idx
  ON org_siem_export_state(disabled_until);

