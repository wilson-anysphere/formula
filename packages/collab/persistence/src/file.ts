import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";

import * as Y from "yjs";

import type { KeyRing } from "../../../security/crypto/keyring.js";

import type { CollabPersistence, CollabPersistenceBinding } from "./index.js";
import {
  FILE_FLAG_ENCRYPTED,
  atomicWriteFile,
  decodeEncryptedRecords,
  decodeLegacyRecords,
  encodeEncryptedRecord,
  encodeFileHeader,
  encodeLegacyRecord,
  hasFileHeader,
  parseFileHeader,
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
    return next;
  }

  private scheduleCompaction(docId: string, doc: Y.Doc): void {
    if (this.compactTimers.has(docId)) return;

    const timer = setTimeout(() => {
      this.compactTimers.delete(docId);
      void this.compactNow(docId, doc);
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
          await fs.writeFile(filePath, encodeFileHeader(FILE_FLAG_ENCRYPTED));
        }
        return;
      }

      if (hasFileHeader(data)) {
        const { flags } = parseFileHeader(data);
        if ((flags & FILE_FLAG_ENCRYPTED) !== FILE_FLAG_ENCRYPTED) {
          throw new Error("Unsupported persistence file flags; expected encrypted format");
        }
        if (this.encryption.mode !== "keyring") {
          throw new Error("Encrypted persistence requires a KeyRing");
        }

        const { keyRing } = this.encryption;
        for (const update of decodeEncryptedRecords(data, { keyRing, aadContext })) {
          Y.applyUpdate(doc, update, persistenceOrigin);
        }
        return;
      }

      // Legacy plaintext file (no header). If encryption is enabled, upgrade it
      // in-place (atomically) before applying updates.
      const legacyUpdates = decodeLegacyRecords(data);
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
    });
  }

  bind(docId: string, doc: Y.Doc): CollabPersistenceBinding {
    // Initialize the persistence file (header when encryption is enabled) before
    // any update events are queued.
    void this.enqueue(docId, async () => {
      await fs.mkdir(this.dir, { recursive: true });
      if (this.encryption.mode !== "keyring") return;

      const filePath = this.filePathForDocId(docId);
      try {
        await fs.access(filePath);
      } catch (err) {
        const code = (err as NodeJS.ErrnoException).code;
        if (code !== "ENOENT") throw err;
        await fs.writeFile(filePath, encodeFileHeader(FILE_FLAG_ENCRYPTED));
      }
    });

    const docHash = this.docHashForDocId(docId);
    const filePath = this.filePathForDocId(docId);
    const aadContext = persistenceAadContextForDocHash(docHash);

    const updateHandler = (update: Uint8Array, origin: unknown) => {
      if (origin === persistenceOrigin) return;

      void this.enqueue(docId, async () => {
        await fs.mkdir(this.dir, { recursive: true });

        if (this.encryption.mode === "keyring") {
          const record = encodeEncryptedRecord(update, {
            keyRing: this.encryption.keyRing,
            aadContext,
          });
          await fs.appendFile(filePath, record);
        } else {
          await fs.appendFile(filePath, encodeLegacyRecord(update));
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

      const snapshot = Y.encodeStateAsUpdate(doc);
      await this.compactSnapshot(docId, snapshot);
      await this.flush(docId);
    };

    doc.on("destroy", () => {
      doc.off("update", updateHandler);
      const timer = this.compactTimers.get(docId);
      if (timer) clearTimeout(timer);
      this.compactTimers.delete(docId);
    });

    return { destroy };
  }

  async flush(docId: string): Promise<void> {
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
