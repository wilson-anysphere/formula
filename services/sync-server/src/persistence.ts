import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";

import type { Logger } from "pino";
import type * as YTypes from "yjs";

import type { KeyRing } from "../../../packages/security/crypto/keyring.js";
import { Y } from "./yjs.js";

type PendingQueue = Promise<void>;

const FILE_MAGIC = Buffer.from("FMLYJS01", "ascii");
const FILE_HEADER_BYTES = FILE_MAGIC.length + 1 + 3; // magic + flags + reserved
const FILE_FLAG_ENCRYPTED = 0b0000_0001;

const AES_GCM_IV_BYTES = 12;
const AES_GCM_TAG_BYTES = 16;
const AES_256_GCM_ALGORITHM = "aes-256-gcm";
const ENCRYPTED_RECORD_HEADER_BYTES = 4 + AES_GCM_IV_BYTES + AES_GCM_TAG_BYTES; // keyVersion + iv + tag

const persistenceOrigin = Symbol("formula.sync-server.persistence");

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

function encodeLegacyRecord(update: Uint8Array): Buffer {
  const header = Buffer.allocUnsafe(4);
  header.writeUInt32BE(update.byteLength, 0);
  return Buffer.concat([header, Buffer.from(update)]);
}

function decodeLegacyRecords(data: Buffer, offset = 0): Uint8Array[] {
  const out: Uint8Array[] = [];
  while (offset + 4 <= data.length) {
    const len = data.readUInt32BE(offset);
    offset += 4;
    if (offset + len > data.length) break;
    out.push(new Uint8Array(data.subarray(offset, offset + len)));
    offset += len;
  }
  return out;
}

function hasFileHeader(data: Buffer): boolean {
  if (data.length < FILE_HEADER_BYTES) return false;
  return data.subarray(0, FILE_MAGIC.length).equals(FILE_MAGIC);
}

function parseFileHeader(data: Buffer): { flags: number } {
  if (!hasFileHeader(data)) {
    throw new Error("not a sync-server persistence file");
  }
  return { flags: data.readUInt8(FILE_MAGIC.length) };
}

function encodeFileHeader(flags: number): Buffer {
  const header = Buffer.alloc(FILE_HEADER_BYTES);
  FILE_MAGIC.copy(header, 0);
  header.writeUInt8(flags, FILE_MAGIC.length);
  // remaining 3 bytes are reserved (0).
  return header;
}

function encodeEncryptedRecord(
  update: Uint8Array,
  opts: { keyRing: KeyRing; aadContext: unknown }
): Buffer {
  const encrypted = opts.keyRing.encryptBytes(Buffer.from(update), {
    aadContext: opts.aadContext,
  });

  if (encrypted.algorithm !== AES_256_GCM_ALGORITHM) {
    throw new Error(
      `Unsupported encryption algorithm for sync-server persistence: ${encrypted.algorithm}`
    );
  }

  if (encrypted.iv.byteLength !== AES_GCM_IV_BYTES) {
    throw new RangeError(
      `Invalid KeyRing iv length (expected ${AES_GCM_IV_BYTES}, got ${encrypted.iv.byteLength})`
    );
  }
  if (encrypted.tag.byteLength !== AES_GCM_TAG_BYTES) {
    throw new RangeError(
      `Invalid KeyRing tag length (expected ${AES_GCM_TAG_BYTES}, got ${encrypted.tag.byteLength})`
    );
  }

  const recordBytes = Buffer.allocUnsafe(
    ENCRYPTED_RECORD_HEADER_BYTES + encrypted.ciphertext.byteLength
  );
  let offset = 0;
  recordBytes.writeUInt32BE(encrypted.keyVersion, offset);
  offset += 4;
  encrypted.iv.copy(recordBytes, offset);
  offset += AES_GCM_IV_BYTES;
  encrypted.tag.copy(recordBytes, offset);
  offset += AES_GCM_TAG_BYTES;
  encrypted.ciphertext.copy(recordBytes, offset);

  const lenPrefix = Buffer.allocUnsafe(4);
  lenPrefix.writeUInt32BE(recordBytes.byteLength, 0);
  return Buffer.concat([lenPrefix, recordBytes]);
}

function decodeEncryptedRecords(
  data: Buffer,
  opts: { keyRing: KeyRing; aadContext: unknown },
  offset = FILE_HEADER_BYTES
): Uint8Array[] {
  const out: Uint8Array[] = [];

  while (offset + 4 <= data.length) {
    const recordLen = data.readUInt32BE(offset);
    offset += 4;
    if (offset + recordLen > data.length) break;

    if (recordLen < ENCRYPTED_RECORD_HEADER_BYTES) {
      throw new Error(
        `Invalid encrypted record length: ${recordLen} (< ${ENCRYPTED_RECORD_HEADER_BYTES})`
      );
    }

    const record = data.subarray(offset, offset + recordLen);
    offset += recordLen;

    const keyVersion = record.readUInt32BE(0);
    const ivOffset = 4;
    const tagOffset = ivOffset + AES_GCM_IV_BYTES;
    const ciphertextOffset = tagOffset + AES_GCM_TAG_BYTES;

    const plaintext = opts.keyRing.decryptBytes(
      {
        keyVersion,
        algorithm: AES_256_GCM_ALGORITHM,
        iv: record.subarray(ivOffset, ivOffset + AES_GCM_IV_BYTES),
        tag: record.subarray(tagOffset, tagOffset + AES_GCM_TAG_BYTES),
        ciphertext: record.subarray(ciphertextOffset),
      },
      { aadContext: opts.aadContext }
    );
    out.push(new Uint8Array(plaintext));
  }

  return out;
}

async function atomicWriteFile(filePath: string, contents: Buffer): Promise<void> {
  const tmpPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tmpPath, contents);
  try {
    await fs.rename(tmpPath, filePath);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "EEXIST" || code === "EPERM") {
      await fs.rm(filePath, { force: true });
      await fs.rename(tmpPath, filePath);
      return;
    }
    throw err;
  }
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
      ? decodeLegacyRecords(data, FILE_HEADER_BYTES)
      : decodeLegacyRecords(data);

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
  private updateCounts = new Map<string, number>();
  private compactTimers = new Map<string, NodeJS.Timeout>();
  private readonly shouldPersist: (docName: string) => boolean;

  constructor(
    private readonly dir: string,
    private readonly logger: Logger,
    private readonly compactAfterUpdates: number,
    private readonly encryption: EncryptionConfig = { mode: "off" },
    shouldPersist?: (docName: string) => boolean
  ) {
    this.shouldPersist = shouldPersist ?? (() => true);
  }

  private docHashForDoc(docName: string): string {
    return createHash("sha256").update(docName).digest("hex");
  }

  private filePathForDocHash(docHash: string): string {
    return path.join(this.dir, `${docHash}.yjs`);
  }

  private enqueue(docName: string, task: () => Promise<void>): Promise<void> {
    const prev = this.queues.get(docName) ?? Promise.resolve();
    const next = prev
      .catch(() => {
        // Keep the queue alive even if a previous write failed.
      })
      .then(task);
    this.queues.set(docName, next);
    return next;
  }

  private scheduleCompaction(docName: string, doc: YTypes.Doc) {
    if (this.compactTimers.has(docName)) return;

    const timer = setTimeout(() => {
      this.compactTimers.delete(docName);
      if (!this.shouldPersist(docName)) return;
      void this.compactNow(docName, doc);
    }, 250);

    this.compactTimers.set(docName, timer);
  }

  private async compactNow(docName: string, doc: YTypes.Doc): Promise<void> {
    if (!this.shouldPersist(docName)) return;
    await this.enqueue(docName, async () => {
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
    });
  }

  async bindState(docName: string, doc: YTypes.Doc): Promise<void> {
    const docHash = this.docHashForDoc(docName);
    const filePath = this.filePathForDocHash(docHash);
    const aadContext = persistenceAadContextForDocHash(docHash);

    const updateHandler = (update: Uint8Array, origin: unknown) => {
      if (origin === persistenceOrigin) return;
      if (!this.shouldPersist(docName)) return;
      void this.enqueue(docName, async () => {
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

        const count = (this.updateCounts.get(docName) ?? 0) + 1;
        this.updateCounts.set(docName, count);

        if (count >= this.compactAfterUpdates) {
          this.scheduleCompaction(docName, doc);
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
    });

    await this.enqueue(docName, async () => {
      if (!this.shouldPersist(docName)) return;
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
          await fs.writeFile(filePath, encodeFileHeader(FILE_FLAG_ENCRYPTED));
        }
        return;
      }

      if (hasFileHeader(data)) {
        const { flags } = parseFileHeader(data);
        if ((flags & FILE_FLAG_ENCRYPTED) !== FILE_FLAG_ENCRYPTED) {
          throw new Error(
            "Unsupported persistence file flags; expected encrypted format"
          );
        }
        if (this.encryption.mode !== "keyring") {
          throw new Error(
            "Encrypted persistence requires SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring"
          );
        }

        const { keyRing } = this.encryption;
        for (const update of decodeEncryptedRecords(data, {
          keyRing,
          aadContext,
        })) {
          Y.applyUpdate(doc, update, persistenceOrigin);
        }
        this.logger.info({ docName }, "persistence_loaded");
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
      this.logger.info({ docName }, "persistence_loaded");
    });
  }

  async writeState(docName: string, doc: YTypes.Doc): Promise<void> {
    if (!this.shouldPersist(docName)) return;
    await this.compactNow(docName, doc);
  }

  async clearDocument(docName: string): Promise<void> {
    const timer = this.compactTimers.get(docName);
    if (timer) clearTimeout(timer);
    this.compactTimers.delete(docName);

    const docHash = this.docHashForDoc(docName);
    const filePath = this.filePathForDocHash(docHash);
    await this.enqueue(docName, async () => {
      await fs.rm(filePath, { force: true });
    });

    // Reset per-document bookkeeping so future docs start from a clean slate.
    this.queues.delete(docName);
    this.updateCounts.delete(docName);

    const timerAfter = this.compactTimers.get(docName);
    if (timerAfter) clearTimeout(timerAfter);
    this.compactTimers.delete(docName);
  }
}
