# `@formula/ai-audit`

Browser-safe AI audit logging primitives (stores + recorder).

## Entry points

This package intentionally keeps the default/browser entry free of Node-only imports and **does not** re-export the sql.js-backed SQLite store (which pulls in `sql.js`).

### Browser / webview (safe default)

```ts
import { AIAuditRecorder, LocalStorageAIAuditStore } from "@formula/ai-audit";
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

