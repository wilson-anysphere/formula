import { promises as fs } from "node:fs";

/**
 * This module contains the sync-server/file persistence record framing.
 *
 * It is intentionally JavaScript (with a .d.ts next to it) so it can be reused
 * by `services/sync-server` without pulling files outside that package's
 * TypeScript `rootDir` into its build output.
 */

export const FILE_MAGIC = Buffer.from("FMLYJS01", "ascii");
export const FILE_HEADER_BYTES = FILE_MAGIC.length + 1 + 3; // magic + flags + reserved
export const FILE_FLAG_ENCRYPTED = 0b0000_0001;

const AES_GCM_IV_BYTES = 12;
const AES_GCM_TAG_BYTES = 16;
const AES_256_GCM_ALGORITHM = "aes-256-gcm";
const ENCRYPTED_RECORD_HEADER_BYTES = 4 + AES_GCM_IV_BYTES + AES_GCM_TAG_BYTES; // keyVersion + iv + tag

export function encodeLegacyRecord(update) {
  const header = Buffer.allocUnsafe(4);
  header.writeUInt32BE(update.byteLength, 0);
  return Buffer.concat([header, Buffer.from(update)]);
}

export function scanLegacyRecords(data, offset = 0) {
  const updates = [];
  let lastGoodOffset = offset;
  while (offset + 4 <= data.length) {
    const len = data.readUInt32BE(offset);
    offset += 4;
    if (offset + len > data.length) break;
    updates.push(new Uint8Array(data.subarray(offset, offset + len)));
    offset += len;
    lastGoodOffset = offset;
  }
  return { updates, lastGoodOffset };
}

export function decodeLegacyRecords(data, offset = 0) {
  return scanLegacyRecords(data, offset).updates;
}

export function hasFileHeader(data) {
  if (data.length < FILE_HEADER_BYTES) return false;
  return data.subarray(0, FILE_MAGIC.length).equals(FILE_MAGIC);
}

export function parseFileHeader(data) {
  if (!hasFileHeader(data)) {
    throw new Error("not a yjs persistence file");
  }
  return { flags: data.readUInt8(FILE_MAGIC.length) };
}

export function encodeFileHeader(flags) {
  const header = Buffer.alloc(FILE_HEADER_BYTES);
  FILE_MAGIC.copy(header, 0);
  header.writeUInt8(flags, FILE_MAGIC.length);
  // remaining 3 bytes are reserved (0).
  return header;
}

export function encodeEncryptedRecord(update, opts) {
  const encrypted = opts.keyRing.encryptBytes(Buffer.from(update), {
    aadContext: opts.aadContext,
  });

  if (encrypted.algorithm !== AES_256_GCM_ALGORITHM) {
    throw new Error(`Unsupported encryption algorithm for yjs persistence: ${encrypted.algorithm}`);
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

export function scanEncryptedRecords(data, opts, offset = FILE_HEADER_BYTES) {
  const updates = [];
  let lastGoodOffset = offset;

  while (offset + 4 <= data.length) {
    const recordLen = data.readUInt32BE(offset);
    offset += 4;
    if (offset + recordLen > data.length) break;

    // Corrupt tail record; stop at the previous good boundary.
    if (recordLen < ENCRYPTED_RECORD_HEADER_BYTES) break;

    const record = data.subarray(offset, offset + recordLen);
    offset += recordLen;

    const keyVersion = record.readUInt32BE(0);
    const ivOffset = 4;
    const tagOffset = ivOffset + AES_GCM_IV_BYTES;
    const ciphertextOffset = tagOffset + AES_GCM_TAG_BYTES;

    try {
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
      updates.push(new Uint8Array(plaintext));
      lastGoodOffset = offset;
    } catch {
      // Decryption failed; treat as a corrupt tail record.
      break;
    }
  }

  return { updates, lastGoodOffset };
}

export function decodeEncryptedRecords(data, opts, offset = FILE_HEADER_BYTES) {
  return scanEncryptedRecords(data, opts, offset).updates;
}

export async function atomicWriteFile(filePath, contents) {
  const tmpPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tmpPath, contents);
  try {
    await fs.rename(tmpPath, filePath);
  } catch (err) {
    const code = err?.code;
    if (code === "EEXIST" || code === "EPERM") {
      await fs.rm(filePath, { force: true });
      await fs.rename(tmpPath, filePath);
      return;
    }
    throw err;
  }
}
