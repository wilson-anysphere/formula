import * as Y from "yjs";
import fs from "node:fs";
import path from "node:path";

import type { OfflinePersistenceHandle } from "./types.ts";

const persistenceOrigin = Symbol("formula.collab-offline.file");
const MAX_RECORD_BYTES = 128 * 1024 * 1024; // hard cap to guard against corrupt length prefixes

function encodeRecord(update: Uint8Array): Buffer {
  const header = Buffer.allocUnsafe(4);
  header.writeUInt32BE(update.byteLength, 0);
  return Buffer.concat([header, Buffer.from(update)]);
}

function writeAllSync(fd: number, buf: Buffer): void {
  let offset = 0;
  while (offset < buf.length) {
    const written = fs.writeSync(fd, buf, offset, buf.length - offset);
    if (written <= 0) throw new Error("Failed to write to persistence file");
    offset += written;
  }
}

function bestEffortFsync(fd: number): void {
  try {
    fs.fsyncSync(fd);
  } catch {
    // Best-effort: some environments may not support fsync (e.g. certain virtual filesystems).
  }
}

function loadAndRepairFile(opts: { doc: Y.Doc; fd: number }): void {
  let size = 0;
  try {
    size = fs.fstatSync(opts.fd).size;
  } catch {
    return;
  }

  const header = Buffer.allocUnsafe(4);
  let offset = 0;
  let lastGoodOffset = 0;

  while (offset + 4 <= size) {
    const recordStart = offset;
    const bytesRead = fs.readSync(opts.fd, header, 0, 4, offset);
    if (bytesRead !== 4) {
      offset = recordStart;
      break;
    }

    const len = header.readUInt32BE(0);
    offset += 4;

    if (len > MAX_RECORD_BYTES) {
      offset = recordStart;
      break;
    }

    if (offset + len > size) {
      // Partial tail record; drop it.
      offset = recordStart;
      break;
    }

    const record = Buffer.allocUnsafe(len);
    const recordRead = fs.readSync(opts.fd, record, 0, len, offset);
    if (recordRead !== len) {
      offset = recordStart;
      break;
    }
    offset += len;

    try {
      Y.applyUpdate(opts.doc, new Uint8Array(record), persistenceOrigin);
    } catch {
      // Corrupt tail record; drop it.
      offset = recordStart;
      break;
    }

    lastGoodOffset = offset;
  }

  if (lastGoodOffset !== size) {
    fs.ftruncateSync(opts.fd, lastGoodOffset);
    bestEffortFsync(opts.fd);
  }
}

export function attachFilePersistence(doc: Y.Doc, opts: { filePath: string }): OfflinePersistenceHandle {
  const dir = path.dirname(opts.filePath);
  fs.mkdirSync(dir, { recursive: true });

  const fd = fs.openSync(opts.filePath, "a+");

  let destroyed = false;
  let isLoading = false;
  /** @type {Uint8Array[]} */
  const bufferedUpdates: Uint8Array[] = [];

  const updateHandler = (update: Uint8Array, origin: unknown) => {
    if (destroyed) return;
    if (origin === persistenceOrigin) return;
    if (isLoading) {
      bufferedUpdates.push(update);
      return;
    }
    const record = encodeRecord(update);
    writeAllSync(fd, record);
    bestEffortFsync(fd);
  };

  // Attach listener first so we don't miss early client updates.
  doc.on("update", updateHandler);

  let loadPromise: Promise<void> | null = null;
  const whenLoaded = async () => {
    if (loadPromise) return loadPromise;
    loadPromise = (async () => {
      if (destroyed) return;

      isLoading = true;
      try {
        loadAndRepairFile({ doc, fd });

        // If the log is empty (brand new file, cleared file, or fully truncated due
        // to corruption), persist a baseline snapshot so future incremental updates
        // are replayable from scratch. This matches y-indexeddb's behavior of
        // persisting a state snapshot during initialization.
        try {
          if (fs.fstatSync(fd).size === 0) {
            const snapshot = Y.encodeStateAsUpdate(doc);
            writeAllSync(fd, encodeRecord(snapshot));
            bestEffortFsync(fd);
            bufferedUpdates.length = 0;
          }
        } catch {
          // Best-effort: failure to write the baseline should not prevent the doc
          // from loading. Subsequent edits may still be persisted incrementally.
        }
      } finally {
        isLoading = false;
      }

      if (destroyed) return;
      if (bufferedUpdates.length > 0) {
        for (const update of bufferedUpdates.splice(0, bufferedUpdates.length)) {
          const record = encodeRecord(update);
          writeAllSync(fd, record);
          bestEffortFsync(fd);
        }
      }
    })();

    return loadPromise;
  };

  const destroy = () => {
    if (destroyed) return;
    destroyed = true;
    doc.off("update", updateHandler);
    doc.off("destroy", destroy);
    try {
      fs.closeSync(fd);
    } catch {
      // ignore
    }
  };

  doc.on("destroy", destroy);

  return {
    whenLoaded,
    destroy,
    clear: async () => {
      if (destroyed) return;

      // Clearing the persisted state makes the on-disk log no longer a valid
      // replay source for the current in-memory document state. To avoid writing
      // a partial log after clearing, we fully detach persistence (callers can
      // re-attach if they want to resume persistence).
      destroy();
      try {
        fs.rmSync(opts.filePath, { force: true });
      } catch {
        // ignore
      }
    },
  };
}
