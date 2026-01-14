import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";

import * as Y from "yjs";

import type { KeyRing } from "../../../security/crypto/keyring.js";

import type { CollabPersistence, CollabPersistenceBinding, CollabPersistenceFlushOptions } from "./index.js";
import {
  FILE_FLAG_ENCRYPTED,
  FILE_HEADER_BYTES,
  atomicWriteFile,
  encodeEncryptedRecord,
  encodeFileHeader,
  encodeLegacyRecord,
  hasFileHeader,
  parseFileHeader,
  scanEncryptedRecords,
  scanLegacyRecords,
} from "./file-format.js";

type PendingQueue = Promise<void>;

type EncryptionConfig = { mode: "off" } | { mode: "keyring"; keyRing: KeyRing };

type PersistenceAadContext = {
  scope: "formula.collab.file-persistence";
  schemaVersion: 1;
  doc: string;
};

function persistenceAadContextForDocHash(docHash: string): PersistenceAadContext {
  return {
    scope: "formula.collab.file-persistence",
    schemaVersion: 1,
    doc: docHash,
  };
}

const persistenceOrigin = Symbol("formula.collab.file-persistence.origin");

export class FileCollabPersistence implements CollabPersistence {
  private readonly queues = new Map<string, PendingQueue>();
  private readonly updateCounts = new Map<string, number>();
  private readonly compactTimers = new Map<string, NodeJS.Timeout>();
  private readonly docsWithFailedEncryptedLoad = new Set<string>();

  private readonly dir: string;
  private readonly compactAfterUpdates: number;
  private readonly encryption: EncryptionConfig;

  constructor(dir: string, opts: { compactAfterUpdates?: number; keyRing?: KeyRing } = {}) {
    this.dir = dir;
    this.compactAfterUpdates = opts.compactAfterUpdates ?? 50;
    this.encryption = opts.keyRing ? { mode: "keyring", keyRing: opts.keyRing } : { mode: "off" };
  }

  private docHashForDocId(docId: string): string {
    return createHash("sha256").update(docId).digest("hex");
  }

  private filePathForDocId(docId: string): string {
    const docHash = this.docHashForDocId(docId);
    return path.join(this.dir, `${docHash}.yjs`);
  }

  private enqueue(docId: string, task: () => Promise<void>): Promise<void> {
    const prev = this.queues.get(docId) ?? Promise.resolve();
    const next = prev
      .catch(() => {
        // Keep the queue alive even if a previous write failed.
      })
      .then(task);
    this.queues.set(docId, next);
    // Prevent unbounded growth of `this.queues` and ensure queued tasks don't
    // trigger unhandled rejections when invoked in a fire-and-forget manner
    // (e.g. from Yjs `doc.on("update")` handlers).
    const cleanup = () => {
      if (this.queues.get(docId) === next) {
        this.queues.delete(docId);
      }
    };
    void next.then(cleanup, cleanup).catch(() => {
      // Best-effort: avoid unhandled rejections from internal queue bookkeeping.
    });
    return next;
  }

  private scheduleCompaction(docId: string, doc: Y.Doc): void {
    if (this.compactTimers.has(docId)) return;

    const timer = setTimeout(() => {
      this.compactTimers.delete(docId);
      void this.compactNow(docId, doc).catch(() => {
        // Best-effort: compaction should not crash consumers of persistence.
      });
    }, 250);
    this.compactTimers.set(docId, timer);
  }

  private async compactSnapshot(docId: string, snapshot: Uint8Array): Promise<void> {
    await this.enqueue(docId, async () => {
      await fs.mkdir(this.dir, { recursive: true });

      const docHash = this.docHashForDocId(docId);
      const filePath = this.filePathForDocId(docId);
      const aadContext = persistenceAadContextForDocHash(docHash);

      if (this.encryption.mode === "keyring") {
        const header = encodeFileHeader(FILE_FLAG_ENCRYPTED);
        const record = encodeEncryptedRecord(snapshot, {
          keyRing: this.encryption.keyRing,
          aadContext,
        });
        await atomicWriteFile(filePath, Buffer.concat([header, record]));
      } else {
        await atomicWriteFile(filePath, encodeLegacyRecord(snapshot));
      }

      this.updateCounts.set(docId, 0);
    });
  }

  private async compactNow(docId: string, doc: Y.Doc): Promise<void> {
    const snapshot = Y.encodeStateAsUpdate(doc);
    await this.compactSnapshot(docId, snapshot);
  }

  async load(docId: string, doc: Y.Doc): Promise<void> {
    await this.enqueue(docId, async () => {
      await fs.mkdir(this.dir, { recursive: true });

      const docHash = this.docHashForDocId(docId);
      const filePath = this.filePathForDocId(docId);
      const aadContext = persistenceAadContextForDocHash(docHash);

      let data: Buffer | null = null;
      try {
        data = await fs.readFile(filePath);
      } catch (err) {
        const code = (err as NodeJS.ErrnoException).code;
        if (code !== "ENOENT") throw err;
        data = null;
      }

        if (!data) {
          if (this.encryption.mode === "keyring") {
            await fs.writeFile(filePath, encodeFileHeader(FILE_FLAG_ENCRYPTED), {
              mode: 0o600,
            });
          }
          return;
        }

      if (hasFileHeader(data)) {
        const { flags } = parseFileHeader(data);
        if ((flags & FILE_FLAG_ENCRYPTED) !== FILE_FLAG_ENCRYPTED) {
          throw new Error("Unsupported persistence file flags; expected encrypted format");
        }
        if (this.encryption.mode !== "keyring") {
          this.docsWithFailedEncryptedLoad.add(docId);
          throw new Error("Encrypted persistence requires a KeyRing");
        }

        const { keyRing } = this.encryption;
        this.docsWithFailedEncryptedLoad.delete(docId);
        try {
          const { updates, lastGoodOffset } = scanEncryptedRecords(data, { keyRing, aadContext });
          if (lastGoodOffset < data.length) {
            await fs.truncate(filePath, lastGoodOffset);
          }
          for (const update of updates) {
            Y.applyUpdate(doc, update, persistenceOrigin);
          }
          return;
        } catch (err) {
          // If encrypted persistence cannot be decoded, disable writes for this doc
          // (otherwise we'd append records that are irrecoverable with the real keyring).
          this.docsWithFailedEncryptedLoad.add(docId);
          throw err;
        }
      }

      // Legacy plaintext file (no header). If encryption is enabled, upgrade it
      // in-place (atomically) before applying updates.
      const { updates: legacyUpdates, lastGoodOffset } = scanLegacyRecords(data);
      if (this.encryption.mode === "keyring") {
        const { keyRing } = this.encryption;
        const header = encodeFileHeader(FILE_FLAG_ENCRYPTED);
        const records = legacyUpdates.map((update) =>
          encodeEncryptedRecord(update, {
            keyRing,
            aadContext,
          })
        );
        await atomicWriteFile(filePath, Buffer.concat([header, ...records]));
      } else if (lastGoodOffset < data.length) {
        // Trim partial/corrupt tail records so future appends remain replayable.
        await fs.truncate(filePath, lastGoodOffset);
      }

      for (const update of legacyUpdates) {
        Y.applyUpdate(doc, update, persistenceOrigin);
      }
    });
  }

  bind(docId: string, doc: Y.Doc): CollabPersistenceBinding {
    const filePath = this.filePathForDocId(docId);
    const docHash = this.docHashForDocId(docId);
    const aadContext = persistenceAadContextForDocHash(docHash);

    // If file encryption settings don't match what's on disk, we must not
    // append new records (or compact) since that can corrupt the log or
    // accidentally downgrade encrypted persistence to plaintext.
    let persistenceEnabled = !this.docsWithFailedEncryptedLoad.has(docId);
    let sawUpdate = false;

    // Initialize (and validate) the persistence file before any update events are
    // queued. Update writes are serialized behind this task via `enqueue()`.
    void this.enqueue(docId, async () => {
      try {
        await fs.mkdir(this.dir, { recursive: true });

        let st: Awaited<ReturnType<typeof fs.stat>> | null = null;
        try {
          st = await fs.stat(filePath);
        } catch (err) {
          const code = (err as NodeJS.ErrnoException).code;
          if (code !== "ENOENT") throw err;
          st = null;
        }

        // New file: create an encrypted header when encryption is enabled.
        if (!st) {
          if (this.encryption.mode === "keyring") {
            await fs.writeFile(filePath, encodeFileHeader(FILE_FLAG_ENCRYPTED), {
              mode: 0o600,
            });
          }
          return;
        }

        if (st.size === 0) {
          if (this.encryption.mode === "keyring") {
            await fs.writeFile(filePath, encodeFileHeader(FILE_FLAG_ENCRYPTED), {
              mode: 0o600,
            });
          }
          return;
        }

        // Validate format by peeking the header bytes.
        const fd = await fs.open(filePath, "r");
        try {
          const header = Buffer.alloc(FILE_HEADER_BYTES);
          const { bytesRead } = await fd.read(header, 0, FILE_HEADER_BYTES, 0);
          const hasHeader = bytesRead === FILE_HEADER_BYTES && hasFileHeader(header);
          if (hasHeader) {
            const { flags } = parseFileHeader(header);
            const isEncrypted = (flags & FILE_FLAG_ENCRYPTED) === FILE_FLAG_ENCRYPTED;
            if (!isEncrypted) {
              persistenceEnabled = false;
              return;
            }
            if (this.encryption.mode !== "keyring") {
              persistenceEnabled = false;
              return;
            }
            return;
          }

          // No header -> legacy plaintext format. When encryption is enabled, we
          // require `load()` to upgrade the file before binding to avoid mixing
          // encrypted + plaintext records.
          if (this.encryption.mode === "keyring") {
            persistenceEnabled = false;
          }
        } finally {
          await fd.close();
        }
      } catch {
        // Keep the doc usable even if persistence setup fails (e.g. permissions).
        // We disable persistence to avoid corrupting on-disk state.
        persistenceEnabled = false;
      }
    });

    const updateHandler = (update: Uint8Array, origin: unknown) => {
      if (origin === persistenceOrigin) return;
      sawUpdate = true;

      void this.enqueue(docId, async () => {
          if (!persistenceEnabled) return;
          await fs.mkdir(this.dir, { recursive: true });

          if (this.encryption.mode === "keyring") {
            const record = encodeEncryptedRecord(update, {
              keyRing: this.encryption.keyRing,
              aadContext,
            });
            await fs.appendFile(filePath, record, { mode: 0o600 });
          } else {
            await fs.appendFile(filePath, encodeLegacyRecord(update), { mode: 0o600 });
          }

        const count = (this.updateCounts.get(docId) ?? 0) + 1;
        this.updateCounts.set(docId, count);

        if (count >= this.compactAfterUpdates) {
          this.scheduleCompaction(docId, doc);
        }
      });
    };

    doc.on("update", updateHandler);

    const destroy = async () => {
      doc.off("update", updateHandler);

      const timer = this.compactTimers.get(docId);
      if (timer) clearTimeout(timer);
      this.compactTimers.delete(docId);

      if (!sawUpdate) {
        // If the document never changed after binding, we don't need to enqueue a
        // final snapshot compaction. This avoids creating empty persistence files
        // when the caller simply opened and closed a document without edits (and
        // prevents teardown races in tests where the session is destroyed without
        // awaiting `binding.destroy()` / `persistence.flush()`).
        //
        // However, if the doc has content but there's no persistence file yet, we
        // still need to write a snapshot so state isn't lost. This can happen if
        // the caller made edits before `bind()` was attached (e.g. during initial
        // async session setup).
        const hasContent = doc.store.clients.size > 0;
        if (!hasContent) return;

        let shouldWriteSnapshot = false;
        try {
          const st = await fs.stat(filePath);
          const isEmptyPlaintext = st.size === 0;
          const isHeaderOnlyEncrypted = this.encryption.mode === "keyring" && st.size === FILE_HEADER_BYTES;
          shouldWriteSnapshot = isEmptyPlaintext || isHeaderOnlyEncrypted;
        } catch (err) {
          const code = (err as NodeJS.ErrnoException).code;
          if (code === "ENOENT") {
            shouldWriteSnapshot = true;
          } else {
            throw err;
          }
        }

        if (!shouldWriteSnapshot) return;
      }

      // Capture the final document state immediately so callers can destroy the
      // Y.Doc without racing the async persistence queue.
      const snapshot = Y.encodeStateAsUpdate(doc);

      // Enqueue the final compaction immediately so callers who *don't* await
      // `binding.destroy()` but do `await persistence.flush(docId)` still wait
      // for the snapshot rewrite.
      await this.enqueue(docId, async () => {
        if (!persistenceEnabled) return;
        await fs.mkdir(this.dir, { recursive: true });

        const docHash = this.docHashForDocId(docId);
        const filePath = this.filePathForDocId(docId);
        const aadContext = persistenceAadContextForDocHash(docHash);

        if (this.encryption.mode === "keyring") {
          const header = encodeFileHeader(FILE_FLAG_ENCRYPTED);
          const record = encodeEncryptedRecord(snapshot, {
            keyRing: this.encryption.keyRing,
            aadContext,
          });
          await atomicWriteFile(filePath, Buffer.concat([header, record]));
        } else {
          await atomicWriteFile(filePath, encodeLegacyRecord(snapshot));
        }

        this.updateCounts.set(docId, 0);
      });
    };

    doc.on("destroy", () => {
      doc.off("update", updateHandler);
      const timer = this.compactTimers.get(docId);
      if (timer) clearTimeout(timer);
      this.compactTimers.delete(docId);
    });

    return { destroy };
  }

  async flush(docId: string, _opts?: CollabPersistenceFlushOptions): Promise<void> {
    await (this.queues.get(docId) ?? Promise.resolve());
  }

  async clear(docId: string): Promise<void> {
    const timer = this.compactTimers.get(docId);
    if (timer) clearTimeout(timer);
    this.compactTimers.delete(docId);

    const filePath = this.filePathForDocId(docId);
    await this.enqueue(docId, async () => {
      await fs.rm(filePath, { force: true });
    });

    this.queues.delete(docId);
    this.updateCounts.delete(docId);
    this.compactTimers.delete(docId);
  }
}
