import type * as Y from "yjs";

import { attachIndexeddbPersistence } from "./indexeddb.ts";
import type { OfflinePersistenceHandle, OfflinePersistenceOptions } from "./types.ts";

export type { OfflinePersistenceHandle, OfflinePersistenceMode, OfflinePersistenceOptions } from "./types.ts";

/**
 * Legacy offline persistence helper.
 *
 * @deprecated Prefer `@formula/collab-persistence` and `CollabSessionOptions.persistence`.
 */
export function attachOfflinePersistence(doc: Y.Doc, opts: OfflinePersistenceOptions): OfflinePersistenceHandle {
  const autoLoad = opts.autoLoad ?? true;
  const key = opts.key ?? doc.guid;

  let started = false;
  let handle: OfflinePersistenceHandle | null = null;

  const start = (): OfflinePersistenceHandle => {
    if (started) return handle!;
    started = true;

    if (opts.mode !== "indexeddb") {
      throw new Error('Offline persistence mode "file" is only supported in Node environments');
    }

    handle = attachIndexeddbPersistence(doc, { key });
    return handle;
  };

  let loaded: Promise<void> | null = null;
  const whenLoaded = () => {
    if (!loaded) {
      loaded = start().whenLoaded();
    }
    return loaded;
  };

  if (autoLoad) {
    void whenLoaded().catch(() => {
      // Best-effort: avoid unhandled rejections from auto-load in environments where persistence
      // is unavailable (private browsing, blocked IndexedDB, etc).
    });
  }

  return {
    whenLoaded,
    destroy: () => {
      if (!started) return;
      start().destroy();
    },
    clear: async () => {
      await start().clear();
    },
  };
}
