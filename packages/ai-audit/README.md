# `@formula/ai-audit`

Browser-safe AI audit logging primitives (stores + recorder).

## Entry points

This package intentionally keeps the default/browser entry free of Node-only imports and **does not** re-export the sql.js-backed SQLite store (which pulls in `sql.js`).

### Browser / webview (safe default)

```ts
import { AIAuditRecorder, createDefaultAIAuditStore } from "@formula/ai-audit";

const store = await createDefaultAIAuditStore({
  // Browser-like runtimes prefer IndexedDB when available, falling back to
  // LocalStorage and then in-memory storage.
  retention: { max_entries: 10_000, max_age_ms: 30 * 24 * 60 * 60 * 1000 },
  // `bounded` is enabled by default (defense-in-depth against oversized entries).
  // bounded: false,
  // prefer: "indexeddb" | "localstorage" | "memory",
});
```

To hard-cap per-entry size (defense-in-depth for LocalStorage/IndexedDB quota limits), wrap any store:

```ts
import { BoundedAIAuditStore, LocalStorageAIAuditStore } from "@formula/ai-audit";

const store = new BoundedAIAuditStore(new LocalStorageAIAuditStore(), {
  max_entry_chars: 200_000,
});
```

When an entry is compacted, oversized fields like `input` and tool payloads are replaced with:

```ts
{
  audit_truncated: true,
  audit_original_chars: number,
  audit_json: string // truncated JSON prefix (may not be parseable JSON)
}
```

Or explicitly:

```ts
import { LocalStorageAIAuditStore } from "@formula/ai-audit/browser";
```

### SQLite store (explicit opt-in)

```ts
import { SqliteAIAuditStore } from "@formula/ai-audit/sqlite";
```

### Node-only helpers

```ts
import { NodeFileBinaryStorage } from "@formula/ai-audit/node";
```

In Node, `@formula/ai-audit/sqlite` also exposes helpers for resolving the `sql.js` wasm assets:

```ts
import { createSqliteAIAuditStoreNode } from "@formula/ai-audit/sqlite";
```

## Exporting audit entries (NDJSON / JSON)

For troubleshooting and compliance workflows you can serialize stored entries in a deterministic,
bounded way:

```ts
import type { AIAuditEntry } from "@formula/ai-audit";
import { serializeAuditEntries } from "@formula/ai-audit";

const ndjson = serializeAuditEntries(entries, {
  format: "ndjson",
  // Default: true. Removes `tool_calls[].result` (often large / sensitive).
  redactToolResults: true,
  // Truncates oversized `tool_calls[].audit_result_summary` payloads and sets
  // `export_truncated: true` when truncation occurs.
  maxToolResultChars: 10_000,
});
```

If you only need the export utility (and want to avoid importing store implementations), you can
use the dedicated entrypoint:

```ts
import { serializeAuditEntries } from "@formula/ai-audit/export";
```

Notes:
- Output ordering is deterministic (stable key sorting).
- `bigint` values are exported as decimal strings.
- Circular references are replaced with the placeholder string `"[Circular]"`.
- Values that throw during serialization (e.g. getters) are replaced with `"[Unserializable]"`.

## Querying audit entries (time ranges + pagination)

All stores implement `store.listEntries(filters?: AuditListFilters)` and always return entries in **newest-first** order.

Useful filters:
- `after_timestamp_ms` (inclusive lower bound)
- `before_timestamp_ms` (exclusive upper bound)
- `cursor: { before_timestamp_ms, before_id? }` for stable pagination
- `limit` applies after filtering

### Example: last 24h (first page)

```ts
const page1 = await store.listEntries({
  workbook_id,
  after_timestamp_ms: Date.now() - 24 * 60 * 60 * 1000,
  limit: 50,
});
```

### Example: fetch the next page (stable even with identical timestamps)

```ts
const last = page1.at(-1);
const page2 = last
  ? await store.listEntries({
      workbook_id,
      limit: 50,
      cursor: { before_timestamp_ms: last.timestamp_ms, before_id: last.id },
    })
  : [];
```
