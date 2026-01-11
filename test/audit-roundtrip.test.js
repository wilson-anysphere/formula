import assert from "node:assert/strict";
import test from "node:test";

import {
  assertAuditEvent,
  auditLogRowToAuditEvent,
  buildPostgresAuditLogInsert,
  createAuditEvent,
  serializeBatch,
  toCef,
  toLeef,
  validateAuditEvent
} from "../packages/audit-core/index.js";
import { SqliteAuditLogStore } from "../packages/security/src/audit/SqliteAuditLogStore.js";

test("audit event round-trip: schema -> sqlite -> query -> export", () => {
  const store = new SqliteAuditLogStore({ path: ":memory:" });

  const event = createAuditEvent({
    id: "evt_roundtrip",
    timestamp: "2026-01-01T00:00:00.000Z",
    eventType: "document.created",
    actor: { type: "user", id: "user_1" },
    context: {
      orgId: "org_1",
      userId: "user_1",
      userEmail: "user@example.com",
      ipAddress: "203.0.113.5",
      userAgent: "UnitTest/1.0",
      sessionId: "sess_1"
    },
    resource: { type: "document", id: "doc_1" },
    correlation: { requestId: "req_1", traceId: "trace_1" },
    success: true,
    details: {
      token: "supersecret",
      nested: { password: "p@ssw0rd" }
    }
  });

  assertAuditEvent(event);
  store.append(event);

  const queried = store.query({ eventType: "document.created" });
  assert.equal(queried.length, 1);
  assert.equal(queried[0].id, "evt_roundtrip");
  assert.equal(queried[0].context.orgId, "org_1");
  assert.equal(queried[0].resource.type, "document");
  assert.equal(queried[0].resource.id, "doc_1");

  const json = serializeBatch(queried, { format: "json" });
  const payload = JSON.parse(json.body.toString("utf8"));
  assert.equal(payload[0].details.token, "[REDACTED]");
  assert.equal(payload[0].details.nested.password, "[REDACTED]");

  const cef = toCef(queried[0]);
  assert.ok(cef.includes('"token":"[REDACTED]"'));

  const leef = toLeef(queried[0]);
  assert.ok(leef.includes('"token":"[REDACTED]"'));
});

test("SqliteAuditLogStore retention sweep deletes events older than retention window", () => {
  const store = new SqliteAuditLogStore({ path: ":memory:" });
  const now = Date.parse("2026-01-10T00:00:00.000Z");

  store.append(
    createAuditEvent({
      id: "evt_old",
      timestamp: "2026-01-01T00:00:00.000Z",
      eventType: "document.opened",
      actor: { type: "user", id: "user_1" },
      success: true,
      details: {}
    })
  );

  store.append(
    createAuditEvent({
      id: "evt_new",
      timestamp: "2026-01-09T00:00:00.000Z",
      eventType: "document.opened",
      actor: { type: "user", id: "user_1" },
      success: true,
      details: {}
    })
  );

  const deleted = store.sweepRetention({ retentionDays: 5, now });
  assert.equal(deleted, 1);

  const remaining = store.query({ eventType: "document.opened" });
  assert.deepEqual(
    remaining.map((e) => e.id).sort(),
    ["evt_new"]
  );
});

test("audit-core schema validator rejects legacy audit event shapes", () => {
  const legacy = {
    id: "evt_legacy",
    ts: Date.now(),
    eventType: "document.created",
    actor: { type: "user", id: "user_1" },
    success: true,
    metadata: {}
  };

  const result = validateAuditEvent(legacy);
  assert.equal(result.valid, false);
  assert.ok(result.errors.some((e) => e.includes("Legacy fields")));
});

test("audit-core postgres adapter persists actor/correlation via details meta and can reconstruct canonical events", () => {
  const event = createAuditEvent({
    id: "11111111-1111-4111-8111-111111111111",
    timestamp: "2026-01-01T00:00:00.000Z",
    eventType: "auth.login_failed",
    actor: { type: "anonymous", id: "user@example.com" },
    context: { orgId: "org_1", userEmail: "user@example.com", ipAddress: "203.0.113.5" },
    resource: { type: "user", id: null, name: "User record" },
    correlation: { requestId: "req_1", traceId: "trace_1" },
    success: false,
    error: { code: "invalid_credentials" },
    details: { token: "supersecret", nested: { password: "p@ssw0rd" } }
  });

  const { values } = buildPostgresAuditLogInsert(event);
  const storedDetails = JSON.parse(values[13]);
  assert.equal(storedDetails.__audit.actor.type, "anonymous");
  assert.equal(storedDetails.__audit.correlation.requestId, "req_1");
  assert.equal(storedDetails.__audit.resourceName, "User record");

  const reconstructed = auditLogRowToAuditEvent({
    id: event.id,
    org_id: event.context.orgId,
    user_id: event.context.userId ?? null,
    user_email: event.context.userEmail ?? null,
    event_type: event.eventType,
    resource_type: event.resource.type,
    resource_id: event.resource.id,
    ip_address: event.context.ipAddress ?? null,
    user_agent: event.context.userAgent ?? null,
    session_id: event.context.sessionId ?? null,
    success: event.success,
    error_code: event.error.code ?? null,
    error_message: event.error.message ?? null,
    details: storedDetails,
    created_at: event.timestamp
  });

  assert.equal(reconstructed.actor.type, "anonymous");
  assert.equal(reconstructed.actor.id, "user@example.com");
  assert.equal(reconstructed.correlation.requestId, "req_1");
  assert.equal(reconstructed.correlation.traceId, "trace_1");
  assert.equal(reconstructed.resource.name, "User record");
  assert.ok(!("__audit" in reconstructed.details));

  // Regression: some Postgres drivers/parsers may return jsonb columns as strings.
  const reconstructedFromString = auditLogRowToAuditEvent({
    id: event.id,
    org_id: event.context.orgId,
    user_id: event.context.userId ?? null,
    user_email: event.context.userEmail ?? null,
    event_type: event.eventType,
    resource_type: event.resource.type,
    resource_id: event.resource.id,
    ip_address: event.context.ipAddress ?? null,
    user_agent: event.context.userAgent ?? null,
    session_id: event.context.sessionId ?? null,
    success: event.success,
    error_code: event.error.code ?? null,
    error_message: event.error.message ?? null,
    details: JSON.stringify(storedDetails),
    created_at: event.timestamp
  });
  assert.equal(reconstructedFromString.actor.type, "anonymous");
  assert.equal(reconstructedFromString.correlation.requestId, "req_1");
  assert.ok(!("__audit" in reconstructedFromString.details));

  const exported = JSON.parse(serializeBatch([reconstructed]).body.toString("utf8"));
  assert.equal(exported[0].details.token, "[REDACTED]");
  assert.equal(exported[0].details.nested.password, "[REDACTED]");
});
