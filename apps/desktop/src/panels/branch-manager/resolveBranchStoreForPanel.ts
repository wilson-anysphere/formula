import type { CollabSession } from "@formula/collab-session";

// Import branching helpers from the browser-safe entrypoint so bundlers don't
// accidentally pull Node-only stores (e.g. SQLite) into the WebView bundle.
import { YjsBranchStore } from "../../../../../packages/versioning/branches/src/browser.js";

import type { BranchStore, CreateBranchStore } from "./branchStoreTypes.js";

export async function resolveBranchStoreForPanel({
  session,
  store,
  createBranchStore,
  compressionFallbackWarning,
}: {
  session: CollabSession;
  store?: BranchStore;
  createBranchStore?: CreateBranchStore;
  compressionFallbackWarning?: string;
}): Promise<{ store: BranchStore; storeWarning: string | null }> {
  if (store) return { store, storeWarning: null };

  if (createBranchStore) {
    const resolved = await createBranchStore(session);
    return { store: resolved, storeWarning: null };
  }

  // Conservative defaults so large branching commits don't exceed common sync-server
  // websocket message limits (close code 1009). Smaller chunks mean more Yjs updates,
  // but keeps the feature usable even when `SYNC_SERVER_MAX_MESSAGE_BYTES` is tuned low.
  const chunkSize = 8 * 1024;
  const maxChunksPerTransaction = 2;

  const proc = (globalThis as any).process;
  const isNodeRuntime = Boolean(proc?.versions?.node) && proc?.release?.name === "node";
  const CompressionStreamCtor = (globalThis as any).CompressionStream as any;
  const DecompressionStreamCtor = (globalThis as any).DecompressionStream as any;

  try {
    if (!isNodeRuntime) {
      if (typeof CompressionStreamCtor === "undefined" || typeof DecompressionStreamCtor === "undefined") {
        throw new Error("CompressionStream is unavailable");
      }
      // Some runtimes expose the constructor but don't support gzip.
      // (If either throws, fall back to JSON payloads.)
      // eslint-disable-next-line @typescript-eslint/no-unsafe-call
      void new CompressionStreamCtor("gzip");
      // eslint-disable-next-line @typescript-eslint/no-unsafe-call
      void new DecompressionStreamCtor("gzip");
    }

    return {
      store: new YjsBranchStore({
        ydoc: session.doc,
        payloadEncoding: "gzip-chunks",
        chunkSize,
        maxChunksPerTransaction,
      }) as any,
      storeWarning: null,
    };
  } catch {
    return {
      store: new YjsBranchStore({ ydoc: session.doc, payloadEncoding: "json" }) as any,
      storeWarning: compressionFallbackWarning ?? null,
    };
  }
}

