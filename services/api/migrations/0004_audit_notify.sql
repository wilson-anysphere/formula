-- Emit NOTIFY events for new audit_log rows so SSE streams can fan out across
-- multiple API instances.

-- NOTE: This migration is intentionally idempotent so it can be applied to
-- existing deployments without manual cleanup.

CREATE OR REPLACE FUNCTION formula_notify_audit_log_insert() RETURNS trigger AS $$
BEGIN
  PERFORM pg_notify(
    'formula_audit_events',
    json_build_object('orgId', NEW.org_id, 'id', NEW.id)::text
  );
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS audit_log_notify_insert ON audit_log;
CREATE TRIGGER audit_log_notify_insert
AFTER INSERT ON audit_log
FOR EACH ROW
EXECUTE FUNCTION formula_notify_audit_log_insert();

