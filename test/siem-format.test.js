import assert from "node:assert/strict";
import test from "node:test";

import { toCef, toLeef } from "../packages/security/siem/format.js";

test("toCef formats an audit event and redacts sensitive fields", () => {
  const event = {
    id: "evt_123",
    timestamp: "2025-01-01T00:00:00.000Z",
    orgId: "org_1",
    eventType: "document.created",
    userId: "user_1",
    userEmail: "user@example.com",
    ipAddress: "203.0.113.5",
    userAgent: "UnitTest/1.0",
    sessionId: "sess_1",
    resourceType: "document",
    resourceId: "doc_1",
    details: {
      token: "supersecret",
      nested: { password: "p@ssw0rd" }
    },
    success: true
  };

  const formatted = toCef(event);
  assert.match(formatted, /^CEF:0\|Formula\|Spreadsheet\|1\.0\|document\.created\|document\.created\|/);
  assert.ok(formatted.includes("src=203.0.113.5"));
  assert.ok(formatted.includes("suser=user@example.com"));
  assert.ok(formatted.includes('"token":"[REDACTED]"'));
  assert.ok(!formatted.includes("supersecret"));
  assert.ok(!formatted.includes("p@ssw0rd"));
});

test("toCef escapes header fields", () => {
  const formatted = toCef({ eventType: "admin.settings|changed", timestamp: "2025-01-01T00:00:00.000Z" });
  assert.ok(formatted.includes("admin.settings\\|changed"));
});

test("toLeef formats an audit event with tab-delimited attributes and redaction", () => {
  const event = {
    id: "evt_456",
    timestamp: "2025-01-02T03:04:05.000Z",
    orgId: "org_2",
    eventType: "auth.login_failed",
    userEmail: "user@example.com",
    ipAddress: "198.51.100.10",
    details: {
      apiKey: "dd_api_key",
      token: "secret"
    },
    success: false
  };

  const formatted = toLeef(event);
  assert.match(formatted, /^LEEF:2\.0\|Formula\|Spreadsheet\|1\.0\|auth\.login_failed\|\t/);

  const segments = formatted.split("\t");
  assert.ok(segments.some((segment) => segment.startsWith("orgId=org_2")));
  assert.ok(segments.some((segment) => segment.startsWith("eventType=auth.login_failed")));
  assert.ok(segments.some((segment) => segment.includes('"apiKey":"[REDACTED]"')));
  assert.ok(!formatted.includes("dd_api_key"));
  assert.ok(!formatted.includes("secret"));
});
