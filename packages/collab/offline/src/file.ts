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

function loadAndRepairFile(opts: { doc: Y.Doc; fd: number; filePath: string }): void {
  let data: Buffer;
  try {
    data = fs.readFileSync(opts.filePath);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "ENOENT") return;
    throw err;
  }

  let offset = 0;
  let lastGoodOffset = 0;

  while (offset + 4 <= data.length) {
    const recordStart = offset;
    const len = data.readUInt32BE(offset);
    offset += 4;

    if (len > MAX_RECORD_BYTES) {
      offset = recordStart;
      break;
    }

    if (offset + len > data.length) {
      // Partial tail record; drop it.
      offset = recordStart;
      break;
    }

    const record = data.subarray(offset, offset + len);
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

  if (lastGoodOffset !== data.length) {
    fs.ftruncateSync(opts.fd, lastGoodOffset);
    bestEffortFsync(opts.fd);
  }
}

export function attachFilePersistence(doc: Y.Doc, opts: { filePath: string }): OfflinePersistenceHandle {
  const dir = path.dirname(opts.filePath);
  fs.mkdirSync(dir, { recursive: true });

  const fd = fs.openSync(opts.filePath, "a+");

  let destroyed = false;

  const updateHandler = (update: Uint8Array, origin: unknown) => {
    if (destroyed) return;
    if (origin === persistenceOrigin) return;
    const record = encodeRecord(update);
    writeAllSync(fd, record);
    bestEffortFsync(fd);
  };

  // Attach listener first so we don't miss early client updates.
  doc.on("update", updateHandler);

  const whenLoaded = async () => {
    if (destroyed) return;
    loadAndRepairFile({ doc, fd, filePath: opts.filePath });
  };

  const destroy = () => {
    if (destroyed) return;
    destroyed = true;
    doc.off("update", updateHandler);
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
      fs.ftruncateSync(fd, 0);
      bestEffortFsync(fd);
    },
  };
}

