# `@formula/collab-offline` (deprecated)

This package provides a legacy “offline persistence” helper for Yjs documents:

```ts
import { attachOfflinePersistence } from "@formula/collab-offline";
```

It is retained for backwards compatibility, but **new code should use**
`@formula/collab-persistence` instead.

## Recommended replacement

Prefer the unified persistence interface:

```ts
import { createCollabSession } from "@formula/collab-session";
import { IndexedDbCollabPersistence } from "@formula/collab-persistence/indexeddb";

const session = createCollabSession({
  docId,
  persistence: new IndexedDbCollabPersistence(),
  connection: { wsUrl, docId },
});

await session.whenLocalPersistenceLoaded();
```

For Node/desktop environments:

```ts
import { FileCollabPersistence } from "@formula/collab-persistence/file";
```

## CollabSession legacy compatibility

`@formula/collab-session` still accepts a deprecated `options.offline` field for
backwards compatibility. Internally, it maps that option onto an appropriate
`CollabPersistence` implementation.

