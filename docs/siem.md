# SIEM Integration (Audit Logs)

Formula records audit events in Postgres and can export them to enterprise SIEM tools (Splunk, Elastic, Datadog, Azure Sentinel, or a custom HTTP receiver). Exports are **batched**, **retried with exponential backoff**, and **redacted** to avoid leaking secrets into downstream logging systems.

The current implementation lives in:

- `packages/audit-core/**` – canonical `AuditEvent` schema + redaction + JSON/CEF/LEEF serialization (`serializeBatch`).
- `services/api/src/routes/audit.ts` – audit query + export endpoints (Fastify).
- `services/api/src/routes/siem.ts` – SIEM config CRUD endpoints (Fastify).
- `services/api/src/siem/*` – background SIEM export worker + sender implementation.

## Supported export formats

### JSON (default)

- Content-Type: `application/json`
- Body: a single JSON array containing redacted audit event objects.

This is the best default for HTTP-based collectors (Splunk HEC, Datadog Logs HTTP intake, custom gateways).

### CEF

- Content-Type: `text/plain`
- Body: newline-delimited CEF records.
- Header example:

```
CEF:0|Formula|Spreadsheet|1.0|document.created|document.created|5|...
```

### LEEF

- Content-Type: `text/plain`
- Body: newline-delimited LEEF records (tab-delimited variant).
- Example:

```
LEEF:2.0|Formula|Spreadsheet|1.0|auth.login_failed|	...
```

The first delimiter after the header is a **literal tab character** (`\t`). Key/value pairs are separated by the same delimiter.

## Canonical audit event schema

Formula uses a canonical `AuditEvent` shape defined in `@formula/audit-core` (`createAuditEvent`). The API stores a flattened subset in Postgres (`audit_log` + `audit_log_archive`) and reconstructs canonical events on export via `auditLogRowToAuditEvent`.

Example:

```json
{
  "schemaVersion": 1,
  "id": "11111111-1111-4111-8111-111111111111",
  "timestamp": "2026-01-01T00:00:00.000Z",
  "eventType": "document.created",
  "actor": { "type": "user", "id": "user_1" },
  "context": {
    "orgId": "org_1",
    "userId": "user_1",
    "userEmail": "user@example.com",
    "ipAddress": "203.0.113.5",
    "userAgent": "UnitTest/1.0",
    "sessionId": "sess_1"
  },
  "resource": { "type": "document", "id": "doc_1" },
  "success": true,
  "details": { "title": "Q1 Plan" },
  "correlation": { "requestId": "req_123", "traceId": "trace_abc" }
}
```

## Per-organization SIEM configuration (Fastify API)

SIEM delivery is configured per organization and persisted in Postgres (`org_siem_configs`; see `services/api/migrations/0004_siem_configs.sql`). Only org admins can manage this configuration.

Endpoints:

- `PUT /orgs/:orgId/siem` – upsert SIEM configuration for an org and set `enabled` (defaults to `true` on first create).
- `GET /orgs/:orgId/siem` – fetch sanitized config (auth secrets are masked as `"***"`).
- `DELETE /orgs/:orgId/siem` – remove SIEM configuration (disables exports).

Request/response shape:

- `GET` returns `{ enabled, config }`.
- `PUT` accepts either `{ enabled, config }` (preferred) or the `config` object itself (backwards compatible).
- When updating an existing config, you can keep previously stored secret values by sending `"***"` for secret fields (the same masked value returned by `GET`).
- Setting `enabled: false` disables exports and deletes any stored secrets; re-enabling requires supplying secret values again.

Auth secrets are stored encrypted in the database-backed secret store (`secrets` table; key = `SECRET_STORE_KEY`) and referenced from `org_siem_configs.config` via `{ "secretRef": "siem:<orgId>:..." }` entries (never plaintext).

Example payload:

```json
{
  "enabled": true,
  "config": {
    "endpointUrl": "https://example.invalid/services/collector/event",
    "format": "json",
    "batchSize": 250,
    "timeoutMs": 10000,
    "idempotencyKeyHeader": "Idempotency-Key",
    "auth": {
      "type": "header",
      "name": "Authorization",
      "value": "Splunk <hec-token>"
    },
    "retry": {
      "maxAttempts": 5,
      "baseDelayMs": 500,
      "maxDelayMs": 30000,
      "jitter": true
    }
  }
}
```

Authentication options:

- `none`
- `bearer` (sets `Authorization: Bearer …`)
- `basic` (sets `Authorization: Basic …`)
- `header` (custom header name/value; useful for Splunk HEC and vendor-specific headers)

## Admin audit query + export API (Fastify)

These endpoints are for human/admin audit review and ad-hoc export (not the continuous SIEM feed).

- `GET /orgs/:orgId/audit` – query audit events (redacted).
- `GET /orgs/:orgId/audit/export` – export audit events (supports `format=json|cef|leef`).

Both endpoints read from `audit_log` and `audit_log_archive`. The export endpoint uses `@formula/audit-core.serializeBatch`.

## Background SIEM export worker

The API process starts a background worker (`SiemExportWorker` in `services/api/src/siem/worker.ts`) from `services/api/src/index.ts`.

High-level flow:

1. Load enabled org configs from Postgres (`DbSiemConfigProvider` in `services/api/src/siem/configProvider.ts`).
2. Fetch audit events (including archived) in ascending order via `services/api/src/siem/auditSource.ts`.
3. Send a batch to the org’s configured `endpointUrl` using `services/api/src/siem/sender.ts`.
4. Persist cursor + backoff state in `org_siem_export_state` (`services/api/migrations/0003_siem_export_state.sql`).

## Not implemented (design notes)

A legacy Node HTTP server previously documented endpoints for audit ingestion and SSE streaming. Those endpoints are **not** part of the current Fastify API.

- `POST /orgs/:orgId/audit` ingestion is not implemented; audit events are produced internally in `services/api/src/*` via `createAuditEvent` + `writeAuditEvent`.
- `GET /orgs/:orgId/audit/stream` SSE streaming is not implemented.

If external ingestion and/or streaming is desired, add it to `services/api/src/routes/audit.ts` and ensure it writes to Postgres (`audit_log`) using `writeAuditEvent`.
