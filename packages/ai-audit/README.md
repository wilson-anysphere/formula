# `@formula/ai-audit`

Browser-safe AI audit logging primitives (stores + recorder).

## Entry points

This package intentionally keeps the default/browser entry free of Node-only imports and **does not** re-export the sql.js-backed SQLite store (which pulls in `sql.js`).

### Browser / webview (safe default)

```ts
import { AIAuditRecorder, LocalStorageAIAuditStore } from "@formula/ai-audit";
```

To hard-cap per-entry size (defense-in-depth for LocalStorage/IndexedDB quota limits), wrap any store:

```ts
import { BoundedAIAuditStore, LocalStorageAIAuditStore } from "@formula/ai-audit";

const store = new BoundedAIAuditStore(new LocalStorageAIAuditStore(), {
  max_entry_chars: 200_000,
});
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
