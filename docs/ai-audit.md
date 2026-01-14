# AI audit logging (`@formula/ai-audit`)

This repo includes a small, host-agnostic audit trail library in `packages/ai-audit` (`@formula/ai-audit`).
It is used to record **what the AI was asked to do**, **which tools it invoked**, and **what verification/feedback happened**.

The design goals are:

- **Understandable**: a single `AIAuditEntry` captures one AI “run” (chat turn / agent step / cell function evaluation).
- **Bounded**: avoid unbounded payloads (entire `read_range` matrices, full external fetch bodies, etc).
- **Safe by default**: audit logs are *not* a DLP system. Sensitive data must be blocked/redacted **before** it is ever sent to a model *or* persisted to an audit store.

---

## 1) What is logged

### `AIAuditEntry`

The core record is `AIAuditEntry` (`packages/ai-audit/src/types.ts`):

| Field | Type | Meaning |
|---|---|---|
| `id` | `string` | Unique audit entry id (UUID when available). |
| `timestamp_ms` | `number` | Wall-clock time in milliseconds since epoch. |
| `session_id` | `string` | Conversation/session identifier (typically per chat/agent run). |
| `workbook_id?` | `string` | Optional workbook identifier for filtering across sessions. |
| `user_id?` | `string` | Optional user identifier (host-controlled). |
| `mode` | `AIMode` | One of: `"tab_completion" \| "inline_edit" \| "chat" \| "agent" \| "cell_function"`. |
| `input` | `unknown` | Host-provided input payload (often `{ prompt, attachments, ... }`). Treat as sensitive. May be compacted by `BoundedAIAuditStore` when enforcing per-entry size/serialization limits. |
| `model` | `string` | Model identifier used for the run (provider-agnostic). |
| `token_usage?` | `{ prompt_tokens, completion_tokens, total_tokens? }` | Optional token usage totals. `total_tokens` may be omitted by providers; `AIAuditRecorder` computes a running total when possible. |
| `latency_ms?` | `number` | Total measured latency (either accumulated via `recordModelLatency()` or set at finalize-time from start → end). |
| `tool_calls` | `ToolCallLog[]` | Ordered list of tool invocations (see below). |
| `verification?` | `AIVerificationResult` | Optional structured verification summary (see below). |
| `user_feedback?` | `"accepted" \| "rejected" \| "modified"` | Optional outcome/label from the UI surface. |

#### Tool calls: `ToolCallLog`

Each entry includes a `tool_calls: ToolCallLog[]` list with fields:

- `name: string` – tool name (e.g. `read_range`, `write_cell`).
- `parameters: unknown` – tool call parameters as logged (often size-capped; see below).
- `requires_approval?: boolean` / `approved?: boolean` – approval gating metadata.
- `ok?: boolean` / `error?: string` – success/failure signal and error string (if any).
- `duration_ms?: number` – tool execution duration (when available).
- `result?: unknown` – **optional** full tool result payload.
- `audit_result_summary?: unknown` – a **bounded** summary of the tool result intended for audit storage.
- `result_truncated?: boolean` – indicates the full `result` was omitted/truncated in the audit entry.

##### Why full tool results are usually omitted

Many spreadsheet tools can return large or sensitive payloads:

- `read_range` can return thousands of cells (values and/or formulas).
- `fetch_external_data` can return remote content.
- Some tools include user-provided secrets (URLs, headers) in parameters/results.

Storing full results has two practical downsides:

1. **Size / reliability**: LocalStorage-backed audit logs can easily exceed quota if full payloads are stored.
2. **Privacy**: full payloads frequently contain raw user data (including data that would be blocked/redacted for cloud processing).

For this reason, the audited tool-calling integration (`packages/ai-tools/src/llm/audited-run.ts`) defaults to:

- storing a **bounded** `audit_result_summary` (typically produced by `serializeToolResultForModel(...)`), and
- setting `result_truncated: true`.

Full results can be enabled explicitly (e.g. for a secure server-side store) via higher-level options such as
`store_full_tool_results`, but this should be considered a privileged/debug mode and reviewed for compliance.

#### Verification: `AIVerificationResult`

When used, `verification` summarizes whether the model used tools when needed and whether its claims were checked:

- `needs_tools: boolean`
- `used_tools: boolean`
- `verified: boolean`
- `confidence: number`
- `warnings: string[]`
- `claims?: Array<{ claim, verified, expected?, actual?, toolEvidence? }>` (optional claim-level details)

---

## 2) Storage backends and retention

### `AIAuditStore` interface

All backends implement:

```ts
export interface AIAuditStore {
  logEntry(entry: AIAuditEntry): Promise<void>;
  listEntries(filters?: AuditListFilters): Promise<AIAuditEntry[]>;
}
```

### `AuditListFilters` (`listEntries(...)`)

`listEntries(...)` supports:

- `session_id?: string`
- `workbook_id?: string`
- `mode?: AIMode | AIMode[]` (empty array is treated as no mode filter)
- `after_timestamp_ms?: number` – inclusive lower bound on `timestamp_ms` (≥)
- `before_timestamp_ms?: number` – exclusive upper bound on `timestamp_ms` (<)
- `cursor?: { before_timestamp_ms: number; before_id?: string }` – **stable pagination cursor**
  - Results are ordered newest-first.
  - When provided, stores return entries strictly **older** than the cursor:
    - older timestamp, or
    - same timestamp + `id < before_id` (when `before_id` is provided)
- `limit?: number`

Stores return entries ordered newest-first (`timestamp_ms` desc, then `id` desc as a tiebreaker when applicable).

### Defense-in-depth: `BoundedAIAuditStore` (per-entry size cap)

Even when upstream integrations try to keep audit payloads small (e.g. bounded tool result summaries),
it is still possible for a single entry to grow unexpectedly large (prompt attachments, large tool
parameters, etc).

`BoundedAIAuditStore` is a lightweight wrapper that enforces a hard cap on the serialized size of
each stored entry (default: **200k characters**). If an entry exceeds the cap **or is not JSON-serializable**
(e.g. contains `bigint` values or cycles), it stores a compacted copy that:

- preserves filter-critical fields (`id`, `timestamp_ms`, `session_id`, `workbook_id`, `mode`, `model`)
- replaces `input` with a truncated JSON string summary (with `audit_truncated` metadata)
- truncates tool `parameters` and `audit_result_summary` similarly
- drops full tool `result` payloads
- when possible, preserves/backfills `workbook_id` (from legacy `input.workbook_id` / `input.workbookId` or `session_id` patterns like `"workbook-123:<uuid>"`)

When compaction happens, `input`/tool payloads are replaced with a JSON-friendly summary object:

```ts
{
  audit_truncated: true,
  audit_original_chars: number,
  // Truncated stable JSON string prefix (may not be parseable JSON).
  // Note: stable JSON serialization coerces `bigint` values to decimal strings and replaces
  // cycles / throwing getters with placeholders like "[Circular]" / "[Unserializable]".
  audit_json: string
}
```

In extreme cases (very large inputs / tool logs), the compaction step may also drop some tool calls
and optional fields to fit under the cap.

If some tool calls are dropped, `BoundedAIAuditStore` appends a sentinel tool call with
`name: "audit_truncated_tool_calls"` describing how many were omitted.

Example:

```ts
import { BoundedAIAuditStore, LocalStorageAIAuditStore } from "@formula/ai-audit/browser";

const store = new BoundedAIAuditStore(new LocalStorageAIAuditStore(), {
  max_entry_chars: 200_000,
});
```

In the desktop app, the default audit store factory (`apps/desktop/src/ai/audit/auditStore.ts`)
wraps the persisted store with `BoundedAIAuditStore` so oversized entries cannot break local
persistence.

### Memory store (ephemeral)

- Class: `MemoryAIAuditStore`
- Persistence: none (in-process only)
- Retention knobs:
  - `max_entries?: number` (newest retained)
  - `max_age_ms?: number` (entries older than `now - max_age_ms` dropped at write-time)
- Intended for: tests, short-lived runs, debugging.

### JSON LocalStorage store (simple browser persistence)

- Class: `LocalStorageAIAuditStore`
- Persistence: `window.localStorage` as a JSON array under a key (default: `formula_ai_audit_log_entries`)
- Retention knobs:
  - `max_entries?: number` (default: `1000`); oldest entries are dropped.
  - `max_age_ms?: number`; entries older than `now - max_age_ms` are dropped at write-time and opportunistically on reads.
- Notes:
  - `max_entries` caps the *count* of entries, but does not enforce a strict per-entry size limit. For defense-in-depth against quota write failures from oversized entries, wrap the store with `BoundedAIAuditStore`.
  - LocalStorage can be unavailable or throw (private mode, quota exceeded, some Node “webstorage” environments); this store falls back to an in-memory buffer in those cases.
  - LocalStorage is **not** encrypted and is readable by any JS running in the origin. Treat it as user-accessible storage.

### IndexedDB store (browser persistence + indexing)

- Class: `IndexedDbAIAuditStore`
- Persistence: `indexedDB` (one entry per record)
  - Database name default: `formula_ai_audit`
  - Object store default: `ai_audit_log`
- Filtering: supports `timestamp_ms` (before/after/cursor pagination), `session_id`, `workbook_id`, and `mode`.
- Workbook metadata: on write, `workbook_id` is normalized for efficient filtering when possible (from `entry.workbook_id`,
  legacy `input.workbook_id` / `input.workbookId`, or `session_id` patterns like `"workbook-123:<uuid>"`).
- Retention knobs:
  - `max_entries?: number` (newest retained)
  - `max_age_ms?: number` (entries older than `now - max_age_ms` deleted)
- Notes:
  - Requires IndexedDB support; some environments (private mode, locked-down webviews) may block or clear IndexedDB.
  - Retention is **best-effort**: browsers can still evict/clear data, and quota behavior varies by platform.
  - IndexedDB also has storage quotas; wrapping with `BoundedAIAuditStore` can help ensure single oversized entries don't break persistence.

### SQLite store (sql.js)

- Class: `SqliteAIAuditStore` (WASM sqlite via `sql.js`)
- Persistence: depends on the configured `SqliteBinaryStorage`:
  - `InMemoryBinaryStorage` (default; no persistence)
  - `LocalStorageBinaryStorage` (base64-encodes the DB into LocalStorage)
  - `NodeFileBinaryStorage` (writes the DB to a file on disk in Node)
- Persistence knobs:
  - `auto_persist?: boolean` (default: `true`) – automatically persist after `logEntry()`
  - `auto_persist_interval_ms?: number` – debounce interval for persistence when `auto_persist` is enabled
  - When `auto_persist=false`, callers must call `flush()` (or `close()`) to durably save pending writes.
- Retention knobs (write-time enforcement):
  - `retention.max_entries?: number`
  - `retention.max_age_ms?: number`
- Workbook metadata: stored in a dedicated `workbook_id` column for efficient filtering. When missing, the store will attempt
  to infer/backfill it from legacy `input.workbook_id` / `input.workbookId` or `session_id` patterns like `"workbook-123:<uuid>"`.

SQLite is the recommended backend when you want:

- indexing/filtering by `session_id`, `workbook_id`, `mode`, and timestamp
- better scaling than JSON LocalStorage
- time-based retention (`max_age_ms`)

### Node file storage (SQLite persistence)

Node persistence is provided via `NodeFileBinaryStorage`, which implements `SqliteBinaryStorage` for `SqliteAIAuditStore`.
It writes the SQLite database bytes to a path using `fs.readFile` / `fs.writeFile`.

Operationally, this means:

- access control inherits from filesystem permissions (choose the path and permissions carefully)
- backups/retention are your responsibility (in addition to in-DB retention limits)

### Store wrappers (optional)

These are useful integrations when composing more complex setups:

- `BoundedAIAuditStore` – per-entry size cap wrapper (see above).
- `CompositeAIAuditStore` – fans out `logEntry(...)` writes to multiple stores (mode `"best_effort"` by default, or `"all"`),
  while delegating reads to the first configured store.

### Utility stores (optional)

- `NoopAIAuditStore` – a store that intentionally does nothing (useful to explicitly disable audit logging without null checks).
- `FailingAIAuditStore` – a store that always throws (useful for tests validating best-effort audit logging behavior).

---

## 3) Security posture (privacy/compliance)

### DLP happens **before** tool results reach the model

Tool results are typically fed back into the model as `role:"tool"` messages in the tool-calling loop.
If a cloud model is used, **this is a data egress path**.

Accordingly:

- DLP policy enforcement must happen at tool execution time (before results are added to the model context).
  - See: `packages/ai-tools/src/executor/tool-executor.ts` (`ToolExecutorOptions.dlp`) which can block or redact tool outputs.
- Audit logging should record the **same sanitized payloads** (post-DLP) that were safe to send to the model.

### Avoid logging full sensitive ranges

Even in local-only audit stores, avoid persisting raw cell matrices and other large/sensitive payloads.
Prefer one or more of:

- A1 references (`Sheet1!A1:D10`) + shape (`rows`, `cols`)
- bounded previews / samples (with `[REDACTED]` placeholders as needed)
- bounded summaries (`audit_result_summary`)
- hashes of large payloads (so you can correlate “same input” without storing the input)

### Treat audit logs as sensitive data

Audit logs can include:

- user prompts (`input`)
- tool parameters (may include URLs, identifiers, cell references)
- tool result summaries (may include data-derived values)

Therefore:

- store audit logs in locations appropriate for the workspace’s compliance requirements
- restrict access to exports and persisted stores
- configure retention (`max_entries`, `max_age_ms`) to the minimum compatible with debugging/compliance needs

---

## 4) How to use

### Choosing a store (`createDefaultAIAuditStore`)

Most hosts should use the default store factory instead of constructing a backend directly:

```ts
import { createDefaultAIAuditStore } from "@formula/ai-audit";

const store = await createDefaultAIAuditStore({
  max_entries: 10_000,
  max_age_ms: 30 * 24 * 60 * 60 * 1000,
  // `bounded` is enabled by default (per-entry size cap defense-in-depth).
  // bounded: false,
  // bounded: { max_entry_chars: 100_000 },
  // prefer: "indexeddb" | "localstorage" | "memory",
});
```

Default selection behavior:

- **Browser-like runtimes** (`window` exists): `IndexedDbAIAuditStore` → `LocalStorageAIAuditStore` → `MemoryAIAuditStore`.
- **Node runtimes** (no `window`): defaults to `MemoryAIAuditStore` (to avoid pulling in `sql.js` unless explicitly requested).
  - Use the Node entrypoint (`@formula/ai-audit/node`) with `prefer: "sqlite"` to opt into persistence.
  - Note: `prefer: "sqlite"` is not supported in the default/browser entrypoint.

### Recording a run with `AIAuditRecorder`

`AIAuditRecorder` builds an `AIAuditEntry` incrementally and writes it once at the end.

Note: `finalize()` is intentionally **best-effort** (it does not throw) because it is often called from `finally` blocks.
If you need to surface persistence failures, check `recorder.finalize_error` / `recorder.finalizeError` after calling `finalize()`.

```ts
import { AIAuditRecorder, MemoryAIAuditStore } from "@formula/ai-audit";

const store = new MemoryAIAuditStore();

const recorder = new AIAuditRecorder({
  store,
  session_id: "workbook-123:550e8400-e29b-41d4-a716-446655440000",
  workbook_id: "workbook-123",
  user_id: "user-42",
  mode: "chat",
  input: { prompt: "Summarize sales by region" },
  model: "cursor-default",
});

// Record model usage/latency as you call the provider:
recorder.recordTokenUsage({ prompt_tokens: 120, completion_tokens: 55 });
recorder.recordModelLatency(842);

// Record tool calls (optionally with approval metadata):
const callIndex = recorder.recordToolCall({
  id: "call-1",
  name: "read_range",
  parameters: { range: "Sheet1!A1:D100", include_formulas: false },
  requires_approval: false,
});

// Record tool outcome:
recorder.recordToolResult(callIndex, {
  ok: true,
  duration_ms: 12,
  // Typically the same bounded content that would be appended as a `role:"tool"` message.
  audit_result_summary: "read_range Sheet1!A1:D100 → (100x4) [values omitted]",
  result_truncated: true,
});

recorder.setVerification({
  needs_tools: true,
  used_tools: true,
  verified: true,
  confidence: 0.9,
  warnings: [],
});

recorder.setUserFeedback("accepted");
await recorder.finalize();
```

### `finalize()` is best-effort

`AIAuditRecorder.finalize()` never throws, even if the underlying store fails to persist the entry.

If you need to detect persistence failures (e.g. for telemetry), inspect:

```ts
await recorder.finalize();
if (recorder.finalizeError) {
  console.warn("Audit persistence failed:", recorder.finalizeError);
}
```

### Choosing a store (browser vs Node)

#### Convenience: `createDefaultAIAuditStore`

`createDefaultAIAuditStore(...)` picks a sensible backend for the current runtime:

- Browser-like runtimes (where `window` exists): prefer `IndexedDbAIAuditStore` (when available), then fall back to `LocalStorageAIAuditStore`, then `MemoryAIAuditStore`.
- Node runtimes (no `window`): default to `MemoryAIAuditStore` (opt into sqlite explicitly).

It also wraps the chosen store in `BoundedAIAuditStore` by default (pass `bounded: false` to disable).

Retention options can be provided either via the legacy `retention: { max_entries, max_age_ms }` object
or via top-level `max_entries` / `max_age_ms` (preferred).

Note: the browser entrypoint intentionally does **not** support `prefer: "sqlite"` (to avoid pulling `sql.js` into default web bundles).
If you want sqlite-backed persistence in browser/webview contexts, import `SqliteAIAuditStore` from `@formula/ai-audit/sqlite` and
construct it directly.

Browser example:

```ts
import { createDefaultAIAuditStore } from "@formula/ai-audit/browser";

const store = await createDefaultAIAuditStore({
  max_entries: 10_000,
  max_age_ms: 30 * 24 * 60 * 60 * 1000, // 30 days
  // bounded: { max_entry_chars: 200_000 }, // optional override
});
```

Node example (opt into sqlite):

```ts
import { createDefaultAIAuditStore, NodeFileBinaryStorage } from "@formula/ai-audit/node";

const store = await createDefaultAIAuditStore({
  prefer: "sqlite",
  sqlite_storage: new NodeFileBinaryStorage("./ai-audit.sqlite"),
  max_entries: 50_000,
  max_age_ms: 90 * 24 * 60 * 60 * 1000, // 90 days
});
```

#### Browser / webview

Small/simple persistence:

```ts
import { LocalStorageAIAuditStore } from "@formula/ai-audit/browser";

const store = new LocalStorageAIAuditStore({
  key: "formula_ai_audit_log_entries",
  max_entries: 1000,
  max_age_ms: 30 * 24 * 60 * 60 * 1000, // 30 days
});
```

IndexedDB-backed persistence (larger quota + incremental writes):

```ts
import { IndexedDbAIAuditStore } from "@formula/ai-audit/browser";

const store = new IndexedDbAIAuditStore({
  db_name: "formula_ai_audit",
  store_name: "ai_audit_log",
  max_entries: 10_000,
  max_age_ms: 30 * 24 * 60 * 60 * 1000, // 30 days
});
```

SQLite-backed persistence (better filtering + time-based retention):

```ts
import { SqliteAIAuditStore } from "@formula/ai-audit/sqlite";
import { BoundedAIAuditStore, LocalStorageBinaryStorage } from "@formula/ai-audit/browser";
import sqlWasmUrl from "sql.js/dist/sql-wasm.wasm?url";

const sqliteStore = await SqliteAIAuditStore.create({
  storage: new LocalStorageBinaryStorage("formula:ai_audit_db:v1"),
  locateFile: (file, prefix = "") => (file.endsWith(".wasm") ? sqlWasmUrl : prefix ? `${prefix}${file}` : file),
  retention: {
    max_entries: 10_000,
    max_age_ms: 30 * 24 * 60 * 60 * 1000, // 30 days
  },
});

// Optional but recommended: cap per-entry size for defense-in-depth against quota overruns.
const store = new BoundedAIAuditStore(sqliteStore, { max_entry_chars: 200_000 });
```

#### Node

SQLite persisted to a file:

```ts
import { NodeFileBinaryStorage } from "@formula/ai-audit/node";
import { createSqliteAIAuditStoreNode } from "@formula/ai-audit/sqlite";

const store = await createSqliteAIAuditStoreNode({
  storage: new NodeFileBinaryStorage("./ai-audit.sqlite"),
  retention: {
    max_entries: 50_000,
    max_age_ms: 90 * 24 * 60 * 60 * 1000, // 90 days
  },
});
```

#### Migrating LocalStorage → SQLite (optional)

If you have legacy entries stored via `LocalStorageAIAuditStore` (JSON array) and want to move to a
sqlite-backed store, `@formula/ai-audit` provides a helper:

```ts
import { migrateLocalStorageAuditEntriesToSqlite } from "@formula/ai-audit";
```

It is idempotent (skips entries that already exist in the destination store by primary key `id`) and
can optionally delete the source localStorage key after a successful migration.

Example:

```ts
import { migrateLocalStorageAuditEntriesToSqlite } from "@formula/ai-audit";
import { LocalStorageBinaryStorage } from "@formula/ai-audit/browser";
import { SqliteAIAuditStore } from "@formula/ai-audit/sqlite";
import sqlWasmUrl from "sql.js/dist/sql-wasm.wasm?url";

const destination = await SqliteAIAuditStore.create({
  storage: new LocalStorageBinaryStorage("formula:ai_audit_db:v1"),
  locateFile: (file, prefix = "") => (file.endsWith(".wasm") ? sqlWasmUrl : prefix ? `${prefix}${file}` : file),
});

await migrateLocalStorageAuditEntriesToSqlite({
  source: { key: "formula_ai_audit_log_entries" },
  destination,
  delete_source: true,
  // max_entries: 5_000, // optional safety cap (newest-first, matching listEntries semantics)
});
```

---

## 5) Exporting audit logs

`@formula/ai-audit` includes a deterministic export helper in `packages/ai-audit/src/export.ts`.

For convenience it is re-exported from `@formula/ai-audit`, but if you only need export functionality (and want to avoid importing store implementations) you can use the lightweight entrypoint:

```ts
import { serializeAuditEntries } from "@formula/ai-audit/export";
```

It takes an array of `AIAuditEntry` and returns a string:

- `serializeAuditEntries(entries, { format })`
- `format` is one of:
  - `"ndjson"` (default) – one JSON object per line
  - `"json"` – a single JSON array

Additional serialization notes (via the internal stable JSON helper):

- Object key ordering is deterministic (stable sorting).
- `bigint` values are serialized as decimal strings.
- Circular references are replaced with the placeholder string `"[Circular]"`.
- Values that throw during serialization (e.g. getters) are replaced with `"[Unserializable]"`.

### Recommended defaults (NDJSON + tool result redaction)

For large logs, prefer **NDJSON** (newline-delimited JSON). It can be streamed/processed incrementally and avoids holding a huge JSON array in memory.

By default, `serializeAuditEntries(...)` also uses `redactToolResults: true` (recommended). This:

- removes `tool_calls[].result` from the export (often large and/or sensitive),
- and truncates oversized `tool_calls[].audit_result_summary` payloads to `maxToolResultChars` (default: `10_000`), adding `export_truncated: true` on the affected tool call.

`maxToolResultChars` is a useful safety knob when you want exports to stay bounded even if a tool summary accidentally becomes large.

### Example: export as NDJSON (recommended for large logs)

```ts
import { AIAuditRecorder, MemoryAIAuditStore } from "@formula/ai-audit";
import { serializeAuditEntries } from "@formula/ai-audit/export";
import { writeFile } from "node:fs/promises";

// `store` can be any AIAuditStore implementation (Memory, LocalStorage, SQLite, etc).
const store = new MemoryAIAuditStore();

// (Optional) create an entry so the example produces output immediately.
const recorder = new AIAuditRecorder({
  store,
  session_id: "workbook-123:session-abc",
  mode: "chat",
  input: { prompt: "Summarize sales by region" },
  model: "cursor-default",
});
await recorder.finalize();

const entries = await store.listEntries({ session_id: "workbook-123:session-abc" });

// Stores return entries newest-first; reverse if you prefer chronological order.
entries.reverse();

const ndjson = serializeAuditEntries(entries, {
  format: "ndjson", // default
  redactToolResults: true, // default (recommended)
  maxToolResultChars: 10_000, // default
});

await writeFile("./ai-audit.ndjson", ndjson + "\n", "utf8");
```

### Example: export as a JSON array

```ts
import { AIAuditRecorder, MemoryAIAuditStore } from "@formula/ai-audit";
import { serializeAuditEntries } from "@formula/ai-audit/export";
import { writeFile } from "node:fs/promises";

const store = new MemoryAIAuditStore();

const recorder = new AIAuditRecorder({
  store,
  session_id: "workbook-123:session-abc",
  mode: "chat",
  input: { prompt: "Summarize sales by region" },
  model: "cursor-default",
});
await recorder.finalize();

const entries = await store.listEntries({ session_id: "workbook-123:session-abc" });

const jsonArray = serializeAuditEntries(entries, {
  format: "json",
  redactToolResults: true, // recommended even for JSON exports
});

await writeFile("./ai-audit.json", jsonArray, "utf8");
```
