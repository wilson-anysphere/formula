# SIEM Integration (Audit Logs)

Formula can export audit events to enterprise SIEM tools (Splunk, Elastic, Datadog, Azure Sentinel, or a custom receiver) using a per-organization configuration. Audit events are **batched**, **retried with exponential backoff**, and **redacted** to avoid leaking secrets into downstream logging systems.

This repository includes:

- `packages/security/siem/**` – formatting (JSON/CEF/LEEF), redaction, batching + HTTP delivery, offline queue.
- `services/api/**` – minimal API wiring for SIEM config management and audit stream delivery.

## Supported export formats

### JSON (default)

- Content-Type: `application/json`
- Payload: a single JSON array containing redacted audit event objects.

This is the best default for HTTP-based collectors (Splunk HEC, Datadog Logs HTTP intake, custom gateways).

### CEF

- Content-Type: `text/plain`
- Payload: newline-delimited CEF records.
- Header example:

```
CEF:0|Formula|Spreadsheet|1.0|document.created|document.created|5|...
```

Formula maps common fields into CEF extension keys (`src`, `suser`, `rt`, etc.) and includes a JSON-encoded `details` payload in `cs6`.

### LEEF

- Content-Type: `text/plain`
- Payload: newline-delimited LEEF records.
- Formula uses the common tab-delimited variant:

```
LEEF:2.0|Formula|Spreadsheet|1.0|auth.login_failed|	...
```

The first delimiter after the header is a **literal tab character** (`\t`). Key/value pairs are separated by the same delimiter.

## Per-organization SIEM configuration

SIEM delivery is configured per organization:

```json
{
  "provider": "splunk",
  "endpointUrl": "https://example.invalid/services/collector/event",
  "format": "json",
  "batchSize": 250,
  "flushIntervalMs": 5000,
  "timeoutMs": 10000,
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
```

Authentication options supported by the exporter:

- `none`
- `bearer` (sets `Authorization: Bearer …`)
- `basic` (sets `Authorization: Basic …`)
- `header` (custom header name/value; useful for Splunk HEC and vendor-specific headers)

## Backend API (services/api)

`services/api/server.js` exposes a minimal set of endpoints for enterprise administration and audit consumption:

- `PUT /orgs/:orgId/siem` – upsert SIEM configuration for an org.
- `GET /orgs/:orgId/siem` – fetch sanitized config (secrets are masked).
- `DELETE /orgs/:orgId/siem` – remove SIEM configuration.
- `POST /orgs/:orgId/audit` – ingest an audit event (server assigns `id` + `timestamp`).
- `GET /orgs/:orgId/audit?limit=100` – pull recent audit events.
- `GET /orgs/:orgId/audit/stream` – SSE stream of audit events.

When SIEM is configured for an organization, `POST /orgs/:orgId/audit` also queues the event for webhook-style delivery via `SiemExporter`.

## Desktop offline-first delivery

For offline-first environments, use `OfflineAuditQueue`:

- Events are appended to a local JSONL file.
- When connectivity is restored, call `flushToExporter(exporter)` to forward queued events to the SIEM endpoint using the same batching, redaction, and retry logic.

## Security and redaction

Before formatting and delivery, audit events are scrubbed to avoid leaking credentials into SIEM systems.

Redaction rules are key-based and recursive. Any key matching patterns like:

- `password`, `secret`, `token`
- `apiKey`, `clientSecret`, `privateKey`
- `authorization`, `cookie`

is replaced with `[REDACTED]`. The redaction utility also detects common token string shapes (e.g., JWTs) and redacts them.

## Testing

- Unit tests cover CEF/LEEF formatting and escaping.
- Integration test uses a mock HTTP server to validate batching + retry behavior.

Run:

```bash
npm test
```

