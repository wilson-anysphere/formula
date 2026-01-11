import type * as Y from "yjs";

import { attachFilePersistence } from "./file.ts";
import { attachIndexeddbPersistence } from "./indexeddb.ts";
import type { OfflinePersistenceHandle, OfflinePersistenceOptions } from "./types.ts";

export type { OfflinePersistenceHandle, OfflinePersistenceMode, OfflinePersistenceOptions } from "./types.ts";

export function attachOfflinePersistence(doc: Y.Doc, opts: OfflinePersistenceOptions): OfflinePersistenceHandle {
  const autoLoad = opts.autoLoad ?? true;
  const key = opts.key ?? doc.guid;

  let started = false;
  let handle: OfflinePersistenceHandle | null = null;

  const start = (): OfflinePersistenceHandle => {
    if (started) return handle!;
    started = true;

    if (opts.mode === "indexeddb") {
      handle = attachIndexeddbPersistence(doc, { key });
      return handle;
    }

    if (opts.mode === "file") {
      if (!opts.filePath) {
        throw new Error('Offline persistence mode "file" requires opts.filePath');
      }
      handle = attachFilePersistence(doc, { filePath: opts.filePath });
      return handle;
    }

    throw new Error(`Unsupported offline persistence mode: ${String((opts as any).mode)}`);
  };

  let loaded: Promise<void> | null = null;
  const whenLoaded = () => {
    if (!loaded) {
      loaded = start().whenLoaded();
    }
    return loaded;
  };

  if (autoLoad) {
    void whenLoaded();
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
