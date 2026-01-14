import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, truncateSync, promises as fs } from "node:fs";
import path from "node:path";

import type { Logger } from "pino";
import type * as YTypes from "yjs";

import type { KeyRing } from "../../../packages/security/crypto/keyring.js";
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
} from "../../../packages/collab/persistence/src/file-format.js";
import { Y } from "./yjs.js";

type PendingQueue = Promise<void>;

const persistenceOrigin = Symbol("formula.sync-server.persistence");

export type PersistenceOverloadScope = "doc" | "total";

export type FilePersistenceOptions = {
  /**
   * Maximum number of pending persistence tasks per document.
   *
   * `0` disables the limit (unbounded).
   */
  maxQueueDepthPerDoc?: number;
  /**
   * Maximum number of pending persistence tasks across all documents.
   *
   * `0` disables the limit (unbounded).
   */
  maxQueueDepthTotal?: number;
  /**
   * Callback invoked when persistence backpressure disables a document.
   */
  onOverload?: (docName: string, scope: PersistenceOverloadScope) => void;
};

type EncryptionConfig =
  | { mode: "off" }
  | {
      mode: "keyring";
      keyRing: KeyRing;
    };

type PersistenceAadContext = {
  scope: "formula.sync-server.persistence";
  schemaVersion: 1;
  doc: string;
};

function persistenceAadContextForDocHash(docHash: string): PersistenceAadContext {
  return {
    scope: "formula.sync-server.persistence",
    schemaVersion: 1,
    doc: docHash,
  };
}

export async function migrateLegacyPlaintextFilesToEncryptedFormat(opts: {
  dir: string;
  logger: Logger;
  keyRing: KeyRing;
}): Promise<void> {
  await fs.mkdir(opts.dir, { recursive: true });
  const entries = await fs.readdir(opts.dir, { withFileTypes: true });

  for (const entry of entries) {
    if (!entry.isFile()) continue;
    if (!entry.name.endsWith(".yjs")) continue;

    const docHash = entry.name.slice(0, -".yjs".length);
    const filePath = path.join(opts.dir, entry.name);

    const data = await fs.readFile(filePath);
    if (hasFileHeader(data)) {
      const { flags } = parseFileHeader(data);
      if ((flags & FILE_FLAG_ENCRYPTED) === FILE_FLAG_ENCRYPTED) continue;
    }

    const aadContext = persistenceAadContextForDocHash(docHash);
    const legacyUpdates = hasFileHeader(data)
      ? scanLegacyRecords(data, FILE_HEADER_BYTES).updates
      : scanLegacyRecords(data).updates;

    const header = encodeFileHeader(FILE_FLAG_ENCRYPTED);
    const records = legacyUpdates.map((update) =>
      encodeEncryptedRecord(update, { keyRing: opts.keyRing, aadContext })
    );

    await atomicWriteFile(filePath, Buffer.concat([header, ...records]));
    opts.logger.info({ filePath }, "persistence_migrated_to_encrypted_format");
  }
}

export class FilePersistence {
  private queues = new Map<string, PendingQueue>();
  private pendingCountsByDoc = new Map<string, number>();
  private pendingTotal = 0;
  private updateCounts = new Map<string, number>();
  private compactTimers = new Map<string, NodeJS.Timeout>();
  private loadPromises = new Map<string, Promise<void>>();
  private readonly shouldPersist: (docName: string) => boolean;
  private readonly disabledDocs = new Set<string>();
  private readonly maxQueueDepthPerDoc: number;
  private readonly maxQueueDepthTotal: number;
  private readonly onOverload?: FilePersistenceOptions["onOverload"];

  constructor(
    private readonly dir: string,
    private readonly logger: Logger,
    private readonly compactAfterUpdates: number,
    private readonly encryption: EncryptionConfig = { mode: "off" },
    shouldPersist?: (docName: string) => boolean,
    opts: FilePersistenceOptions = {}
  ) {
    this.shouldPersist = shouldPersist ?? (() => true);
    this.maxQueueDepthPerDoc = Math.max(0, opts.maxQueueDepthPerDoc ?? 0);
    this.maxQueueDepthTotal = Math.max(0, opts.maxQueueDepthTotal ?? 0);
    this.onOverload = opts.onOverload;
  }

  isDocDisabled(docName: string): boolean {
    return this.disabledDocs.has(docName);
  }

  private docHashForDoc(docName: string): string {
    return createHash("sha256").update(docName).digest("hex");
  }

  private filePathForDocHash(docHash: string): string {
    return path.join(this.dir, `${docHash}.yjs`);
  }

  private triggerOverload(docName: string, scope: PersistenceOverloadScope) {
    if (this.disabledDocs.has(docName)) return;
    this.disabledDocs.add(docName);
    const timer = this.compactTimers.get(docName);
    if (timer) clearTimeout(timer);
    this.compactTimers.delete(docName);
    this.logger.warn({ docName, scope }, "persistence_queue_overloaded");
    try {
      this.onOverload?.(docName, scope);
    } catch (err) {
      this.logger.warn({ docName, err }, "persistence_overload_callback_failed");
    }
  }

  private enqueue(
    docName: string,
    task: () => Promise<void>,
    opts: { bypassLimits?: boolean } = {}
  ): Promise<void> {
    if (!opts.bypassLimits) {
      const pendingForDoc = this.pendingCountsByDoc.get(docName) ?? 0;
      if (this.maxQueueDepthPerDoc > 0 && pendingForDoc >= this.maxQueueDepthPerDoc) {
        this.triggerOverload(docName, "doc");
        return Promise.resolve();
      }
      if (this.maxQueueDepthTotal > 0 && this.pendingTotal >= this.maxQueueDepthTotal) {
        this.triggerOverload(docName, "total");
        return Promise.resolve();
      }

      this.pendingCountsByDoc.set(docName, pendingForDoc + 1);
      this.pendingTotal += 1;
    } else {
      // Still track depth for observability, but never reject bypassed tasks.
      const pendingForDoc = this.pendingCountsByDoc.get(docName) ?? 0;
      this.pendingCountsByDoc.set(docName, pendingForDoc + 1);
      this.pendingTotal += 1;
    }

    const prev = this.queues.get(docName) ?? Promise.resolve();
    const next = prev
      .catch(() => {
        // Keep the queue alive even if a previous write failed.
      })
      .then(task);
    this.queues.set(docName, next);
    void next
      .finally(() => {
        this.pendingTotal = Math.max(0, this.pendingTotal - 1);
        const remaining = (this.pendingCountsByDoc.get(docName) ?? 1) - 1;
        if (remaining <= 0) {
          this.pendingCountsByDoc.delete(docName);
        } else {
          this.pendingCountsByDoc.set(docName, remaining);
        }
        if (this.queues.get(docName) === next) {
          this.queues.delete(docName);
        }
      })
      .catch(() => {
        // Best-effort: the returned `next` promise is awaited/handled by callers; avoid an
        // unhandled rejection from the internal `.finally` bookkeeping chain.
      });
    return next;
  }

  /**
   * Promise that resolves once persistence state has been loaded for `docName`.
   *
   * `y-websocket` does not await `bindState()`. sync-server calls this method
   * during websocket upgrade to ensure clients never sync against a transient
   * empty Y.Doc.
   */
  waitForLoaded(docName: string): Promise<void> {
    return this.loadPromises.get(docName) ?? Promise.resolve();
  }

  private scheduleCompaction(docName: string, doc: YTypes.Doc) {
    if (this.compactTimers.has(docName)) return;

    const timer = setTimeout(() => {
      this.compactTimers.delete(docName);
      if (!this.shouldPersist(docName)) return;
      if (this.disabledDocs.has(docName)) return;
      void this.compactNow(docName, doc);
    }, 250);

    this.compactTimers.set(docName, timer);
  }

  private async compactNow(docName: string, doc: YTypes.Doc): Promise<void> {
    if (!this.shouldPersist(docName)) return;
    if (this.disabledDocs.has(docName)) return;
    await this.enqueue(docName, async () => {
      if (this.disabledDocs.has(docName)) return;
      try {
        await fs.mkdir(this.dir, { recursive: true });

        const docHash = this.docHashForDoc(docName);
        const filePath = this.filePathForDocHash(docHash);
        const aadContext = persistenceAadContextForDocHash(docHash);
        const snapshot = Y.encodeStateAsUpdate(doc);

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

        this.updateCounts.set(docName, 0);
        this.logger.info({ docName }, "persistence_compacted");
      } catch (err) {
        this.disabledDocs.add(docName);
        this.logger.error({ docName, err }, "persistence_compaction_failed");
      }
    });
  }

  bindState(docName: string, doc: YTypes.Doc): void {
    const docHash = this.docHashForDoc(docName);
    const filePath = this.filePathForDocHash(docHash);
    const aadContext = persistenceAadContextForDocHash(docHash);

    const updateHandler = (update: Uint8Array, origin: unknown) => {
      if (origin === persistenceOrigin) return;
      if (!this.shouldPersist(docName)) return;
      if (this.disabledDocs.has(docName)) return;
      void this.enqueue(docName, async () => {
        if (this.disabledDocs.has(docName)) return;
        try {
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

          const count = (this.updateCounts.get(docName) ?? 0) + 1;
          this.updateCounts.set(docName, count);

          if (count >= this.compactAfterUpdates) {
            this.scheduleCompaction(docName, doc);
          }
        } catch (err) {
          this.disabledDocs.add(docName);
          this.logger.error({ docName, err }, "persistence_write_failed");
          const timer = this.compactTimers.get(docName);
          if (timer) clearTimeout(timer);
          this.compactTimers.delete(docName);
        }
      });
    };

    // Important: `y-websocket` does not await `bindState()`. Attach the update
    // listener first so we don't miss early client updates.
    doc.on("update", updateHandler);
    doc.on("destroy", () => {
      doc.off("update", updateHandler);
      const timer = this.compactTimers.get(docName);
      if (timer) clearTimeout(timer);
      this.compactTimers.delete(docName);
      this.disabledDocs.delete(docName);
      this.updateCounts.delete(docName);
      this.loadPromises.delete(docName);
    });

    // `y-websocket` server utilities do not await `bindState()`. To ensure
    // clients sync with the fully-hydrated document (and avoid racing client
    // schema initialization against late-arriving persisted updates), load the
    // plaintext persistence file synchronously when encryption is disabled.
    if (this.encryption.mode === "off") {
      // If persistence is disabled for this doc, treat it as immediately "loaded".
      if (!this.shouldPersist(docName)) {
        this.loadPromises.set(docName, Promise.resolve());
        return;
      }

      // Default to ready. Any error below flips the promise to rejected.
      this.loadPromises.set(docName, Promise.resolve());
      try {
        mkdirSync(this.dir, { recursive: true });

        let data: Buffer | null = null;
        try {
          data = readFileSync(filePath);
        } catch (err) {
          const code = (err as NodeJS.ErrnoException).code;
          if (code !== "ENOENT") throw err;
          data = null;
        }

        if (!data) return;

        if (hasFileHeader(data)) {
          const { flags } = parseFileHeader(data);
          if ((flags & FILE_FLAG_ENCRYPTED) === FILE_FLAG_ENCRYPTED) {
            throw new Error("Encrypted persistence file requires keyring encryption");
          }
        }

        const legacyScan = hasFileHeader(data)
          ? scanLegacyRecords(data, FILE_HEADER_BYTES)
          : scanLegacyRecords(data);
        for (const update of legacyScan.updates) {
          Y.applyUpdate(doc, update, persistenceOrigin);
        }
        if (legacyScan.lastGoodOffset < data.length) {
          truncateSync(filePath, legacyScan.lastGoodOffset);
          this.logger.warn({ docName }, "persistence_truncated_corrupt_tail");
        }
        this.logger.info({ docName }, "persistence_loaded");
      } catch (err) {
        this.disabledDocs.add(docName);
        this.logger.error({ docName, err }, "persistence_load_failed");
        const rejected = Promise.reject(err);
        // Avoid unhandled rejection warnings if a caller doesn't await this (e.g.
        // future internal endpoints).
        rejected.catch(() => {});
        this.loadPromises.set(docName, rejected);
      }
      return;
    }

    const loadPromise = this.enqueue(docName, async () => {
      if (!this.shouldPersist(docName)) return;
      if (this.disabledDocs.has(docName)) return;
      await fs.mkdir(this.dir, { recursive: true });

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
          throw new Error(
            "Encrypted persistence requires SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring"
          );
        }

        const { keyRing } = this.encryption;
        const { updates, lastGoodOffset } = scanEncryptedRecords(data, { keyRing, aadContext });
        for (const update of updates) {
          Y.applyUpdate(doc, update, persistenceOrigin);
        }
        if (lastGoodOffset < data.length) {
          await fs.truncate(filePath, lastGoodOffset);
          this.logger.warn({ docName }, "persistence_truncated_corrupt_tail");
        }
        this.logger.info({ docName }, "persistence_loaded");
        return;
      }

      // Legacy plaintext file (no header). If encryption is enabled, upgrade it
      // in-place (atomically) before applying updates.
      const legacyUpdates = scanLegacyRecords(data).updates;
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
      }

      for (const update of legacyUpdates) {
        Y.applyUpdate(doc, update, persistenceOrigin);
      }
      this.logger.info({ docName }, "persistence_loaded");
    }).catch((err) => {
      this.disabledDocs.add(docName);
      this.logger.error({ docName, err }, "persistence_load_failed");
      const timer = this.compactTimers.get(docName);
      if (timer) clearTimeout(timer);
      this.compactTimers.delete(docName);
      throw err;
    });

    // Avoid unhandled rejection warnings in case load fails before any websocket
    // upgrade awaits `waitForLoaded()`.
    loadPromise.catch(() => {});
    this.loadPromises.set(docName, loadPromise);
  }

  async writeState(docName: string, doc: YTypes.Doc): Promise<void> {
    if (!this.shouldPersist(docName)) return;
    if (this.disabledDocs.has(docName)) return;
    try {
      await this.compactNow(docName, doc);
    } catch (err) {
      this.disabledDocs.add(docName);
      this.logger.error({ docName, err }, "persistence_write_failed");
    }
  }

  async flush(): Promise<void> {
    for (const timer of this.compactTimers.values()) {
      clearTimeout(timer);
    }
    this.compactTimers.clear();

    while (this.queues.size > 0) {
      const pending = Array.from(this.queues.values());
      await Promise.allSettled(pending);
    }
  }

  async clearDocument(docName: string): Promise<void> {
    const timer = this.compactTimers.get(docName);
    if (timer) clearTimeout(timer);
    this.compactTimers.delete(docName);

    const docHash = this.docHashForDoc(docName);
    const filePath = this.filePathForDocHash(docHash);
    await this.enqueue(docName, async () => {
      await fs.rm(filePath, { force: true });
    }, { bypassLimits: true });

    // Reset per-document bookkeeping so future docs start from a clean slate.
    this.queues.delete(docName);
    this.updateCounts.delete(docName);
    this.disabledDocs.delete(docName);
    this.loadPromises.delete(docName);

    const timerAfter = this.compactTimers.get(docName);
    if (timerAfter) clearTimeout(timerAfter);
    this.compactTimers.delete(docName);
  }
}
