# SIEM Integration (Audit Logs)

Formula records audit events in Postgres and can export them to enterprise SIEM tools (Splunk, Elastic, Datadog, Azure Sentinel, or a custom HTTP receiver). Exports are **batched**, **retried with exponential backoff**, and **redacted** to avoid leaking secrets into downstream logging systems.

Outbound SIEM delivery enforces a **minimum TLS version of TLS 1.3** and can optionally enforce **certificate pinning** (SHA-256 fingerprint allowlist) via org policy fields in `org_settings`. See `docs/tls-pinning.md` for how to compute and configure pins.

The current implementation lives in:

- `packages/audit-core/**` – canonical `AuditEvent` schema + redaction + JSON/CEF/LEEF serialization (`serializeBatch`).
- `services/api/src/routes/audit.ts` – audit query + export endpoints (Fastify).
- `services/api/src/routes/siem.ts` – SIEM config CRUD endpoints (Fastify).
- `services/api/src/siem/*` – background SIEM export worker + sender implementation.
- `packages/security/siem/**` – client-side/offline queue + delivery helpers (separate from the server worker).

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

Note: `GET /orgs/:orgId/audit/export` and the background SIEM export worker both emit this canonical `AuditEvent` shape. Any internal `details.__audit` metadata stored in Postgres is stripped during reconstruction.

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
- `GET /orgs/:orgId/siem` – fetch sanitized config, plus `secretConfigured` (auth secrets are never returned).
- `DELETE /orgs/:orgId/siem` – remove SIEM configuration (disables exports).

Request/response shape:

- `GET` returns `{ enabled, config, secretConfigured }`. If no config exists yet, it returns `{ enabled: false, config: null, secretConfigured: false }`.
- `PUT` accepts either `{ enabled, config }` (preferred) or the `config` object itself (backwards compatible).
- When updating an existing config, you can keep previously stored secret values by **omitting secret fields**. (Backwards compatible: `"***"` is also accepted.)
- Setting `enabled: false` disables exports and deletes any stored secrets; re-enabling requires supplying secret values again.

Auth secrets are stored encrypted in the database-backed secret store (`secrets` table; configured via `SECRET_STORE_KEYS_JSON` or legacy `SECRET_STORE_KEY`) and referenced internally from `org_siem_configs.config` via `{ "secretRef": "siem:<orgId>:..." }` entries (never plaintext). Secret names use stable keys such as:

- `siem:<orgId>:bearer_token`
- `siem:<orgId>:basic_username`
- `siem:<orgId>:basic_password`
- `siem:<orgId>:header_value`

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

Notes:

- `endpointUrl` must be `https:` in `NODE_ENV=production` (the API rejects plaintext HTTP in production).

Authentication options:

- `none`
- `bearer` (sets `Authorization: Bearer …`)
- `basic` (sets `Authorization: Basic …`)
- `header` (custom header name/value; useful for Splunk HEC and vendor-specific headers)

## Audit ingestion + query + export API (Fastify)

These endpoints are for:

- **Ingesting** client-side events into the server audit log
- **Admin review** and ad-hoc export
- **Near-real-time streaming** for admin/SOC consumption

Endpoints:

- `POST /orgs/:orgId/audit` – ingest an audit event (**authenticated**; any org member may write). The server derives:
  - `actor` (from the authenticated user or API key)
  - `context` (`orgId`, `userId`, `userEmail`, `sessionId`, `ipAddress`, `userAgent`)
  - `id` + `timestamp` (server-assigned via `createAuditEvent`)
  Client payload is limited to `eventType` + optional `resource`/`success`/`error`/`details`/`correlation`. Requests are rate-limited per org + IP.
- `GET /orgs/:orgId/audit` – query audit events (admin-only; returned events are redacted).
- `GET /orgs/:orgId/audit/export` – export audit events (admin-only; supports `format=json|cef|leef`).
- `GET /orgs/:orgId/audit/stream` – SSE stream of audit events (admin-only; returned events are redacted). Supports resume via `?after=<base64 cursor>` or `Last-Event-ID`.

All endpoints read from (or write to) Postgres audit tables: `audit_log` and `audit_log_archive`.

## Background SIEM export worker

The API process starts a background worker (`SiemExportWorker` in `services/api/src/siem/worker.ts`) from `services/api/src/index.ts`.

High-level flow:

1. Load enabled org configs from Postgres (`DbSiemConfigProvider` in `services/api/src/siem/configProvider.ts`).
2. Fetch audit events (including archived) in ascending order via `services/api/src/siem/auditSource.ts` and convert rows to canonical `AuditEvent` via `auditLogRowToAuditEvent`.
3. Serialize + redact batches using `@formula/audit-core` (`serializeBatch`) and POST to the org’s configured `endpointUrl` using `services/api/src/siem/sender.ts`.
4. Persist cursor + backoff state in `org_siem_export_state` (`services/api/migrations/0003_siem_export_state.sql`).

### Exported event shape (background worker)

The worker exports batches of canonical `AuditEvent` objects (same schema as above). Any internal `details.__audit` metadata stored in Postgres is removed before export.

## Client-side delivery helpers (desktop / offline-first)

This repo also includes a standalone JavaScript SIEM delivery library intended for **clients** (desktop / offline-first), not the cloud API runtime:

- `packages/security/siem/exporter.js` – `SiemExporter` (batching + retry + HTTP delivery).
- `packages/security/siem/offlineQueue.js` – `OfflineAuditQueue` (persist redacted events to Node FS or IndexedDB and flush later via `flushToExporter(exporter)`).

These helpers operate on canonical `AuditEvent` objects from `@formula/audit-core` and are separate from the server-side SIEM worker in `services/api/src/siem/*`.
