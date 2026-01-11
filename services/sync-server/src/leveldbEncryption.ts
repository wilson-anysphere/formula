import crypto from "node:crypto";

const AES_256_GCM = "aes-256-gcm";
const AES_GCM_IV_BYTES = 12;
const AES_GCM_TAG_BYTES = 16;

export const LEVELDB_ENCRYPTION_KEY_BYTES = 32;
export const LEVELDB_ENCRYPTION_KEY_VERSION = 1;
export const LEVELDB_ENCRYPTION_MAGIC = Buffer.from("FMLLDB01"); // 8 bytes
export const LEVELDB_ENCRYPTION_AAD = Buffer.from(
  "formula-sync-server-leveldb:v1",
  "utf8"
);

const LEVELDB_ENCRYPTION_HEADER_BYTES =
  LEVELDB_ENCRYPTION_MAGIC.length +
  4 + // keyVersion (u32 BE)
  AES_GCM_IV_BYTES +
  AES_GCM_TAG_BYTES;

function assertKey(key: Buffer): void {
  if (key.byteLength !== LEVELDB_ENCRYPTION_KEY_BYTES) {
    throw new Error(
      `Encryption key must be ${LEVELDB_ENCRYPTION_KEY_BYTES} bytes (got ${key.byteLength}).`
    );
  }
}

export function encryptLeveldbValue(plaintext: Buffer, key: Buffer): Buffer {
  if (!Buffer.isBuffer(plaintext)) {
    throw new TypeError("plaintext must be a Buffer");
  }
  assertKey(key);

  const iv = crypto.randomBytes(AES_GCM_IV_BYTES);
  const cipher = crypto.createCipheriv(AES_256_GCM, key, iv, {
    authTagLength: AES_GCM_TAG_BYTES,
  });
  cipher.setAAD(LEVELDB_ENCRYPTION_AAD);

  const ciphertext = Buffer.concat([cipher.update(plaintext), cipher.final()]);
  const tag = cipher.getAuthTag();

  const out = Buffer.allocUnsafe(
    LEVELDB_ENCRYPTION_HEADER_BYTES + ciphertext.byteLength
  );
  let offset = 0;
  LEVELDB_ENCRYPTION_MAGIC.copy(out, offset);
  offset += LEVELDB_ENCRYPTION_MAGIC.byteLength;

  out.writeUInt32BE(LEVELDB_ENCRYPTION_KEY_VERSION, offset);
  offset += 4;

  iv.copy(out, offset);
  offset += AES_GCM_IV_BYTES;

  tag.copy(out, offset);
  offset += AES_GCM_TAG_BYTES;

  ciphertext.copy(out, offset);
  return out;
}

export function decryptLeveldbValue(
  data: Buffer,
  key: Buffer,
  strict: boolean
): Buffer {
  if (!Buffer.isBuffer(data)) {
    throw new TypeError("ciphertext must be a Buffer");
  }
  assertKey(key);

  if (data.byteLength < LEVELDB_ENCRYPTION_MAGIC.byteLength) {
    if (strict) {
      throw new Error(
        "LevelDB value is not encrypted (missing magic header). If you're migrating an existing DB, set SYNC_SERVER_PERSISTENCE_ENCRYPTION_STRICT=false temporarily."
      );
    }
    return data;
  }

  const magic = data.subarray(0, LEVELDB_ENCRYPTION_MAGIC.byteLength);
  if (!magic.equals(LEVELDB_ENCRYPTION_MAGIC)) {
    if (strict) {
      throw new Error(
        "LevelDB value is not encrypted (missing magic header). If you're migrating an existing DB, set SYNC_SERVER_PERSISTENCE_ENCRYPTION_STRICT=false temporarily."
      );
    }
    return data;
  }

  if (data.byteLength < LEVELDB_ENCRYPTION_HEADER_BYTES) {
    throw new Error(
      "Encrypted LevelDB value is truncated (missing header bytes)."
    );
  }

  let offset = LEVELDB_ENCRYPTION_MAGIC.byteLength;
  const version = data.readUInt32BE(offset);
  offset += 4;
  if (version !== LEVELDB_ENCRYPTION_KEY_VERSION) {
    throw new Error(
      `Unsupported LevelDB encryption key version ${version} (expected ${LEVELDB_ENCRYPTION_KEY_VERSION}).`
    );
  }

  const iv = data.subarray(offset, offset + AES_GCM_IV_BYTES);
  offset += AES_GCM_IV_BYTES;

  const tag = data.subarray(offset, offset + AES_GCM_TAG_BYTES);
  offset += AES_GCM_TAG_BYTES;

  const ciphertext = data.subarray(offset);

  try {
    const decipher = crypto.createDecipheriv(AES_256_GCM, key, iv, {
      authTagLength: AES_GCM_TAG_BYTES,
    });
    decipher.setAAD(LEVELDB_ENCRYPTION_AAD);
    decipher.setAuthTag(tag);
    return Buffer.concat([
      decipher.update(ciphertext),
      decipher.final(),
    ]);
  } catch (err) {
    const reason = err instanceof Error ? err.message : String(err);
    throw new Error(
      `Failed to decrypt LevelDB value (${reason}). Check SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64.`
    );
  }
}

type LevelEncoding = {
  encode: (data: any) => any;
  decode: (data: any) => any;
  buffer?: boolean;
  type?: string;
  [key: string]: any;
};

function isLevelEncoding(value: unknown): value is LevelEncoding {
  return (
    typeof value === "object" &&
    value !== null &&
    "encode" in value &&
    "decode" in value &&
    typeof (value as LevelEncoding).encode === "function" &&
    typeof (value as LevelEncoding).decode === "function"
  );
}

function toBuffer(data: unknown): Buffer {
  if (Buffer.isBuffer(data)) return data;
  if (data instanceof Uint8Array) return Buffer.from(data);
  if (typeof data === "string") return Buffer.from(data, "utf8");
  throw new TypeError("LevelDB valueEncoding expects Buffer-like data");
}

export function createEncryptedLevelAdapter({
  baseLevel,
  key,
  strict,
}: {
  baseLevel: (location: string, opts: any) => any;
  key: Buffer | Uint8Array;
  strict: boolean;
}): (location: string, opts: any) => any {
  const keyBuf = Buffer.isBuffer(key) ? key : Buffer.from(key);
  assertKey(keyBuf);

  return (location: string, opts: any) => {
    const original = isLevelEncoding(opts?.valueEncoding)
      ? (opts.valueEncoding as LevelEncoding)
      : {
          buffer: true,
          type: "raw",
          encode: (v: any) => v,
          decode: (v: any) => v,
        };

    const valueEncoding: LevelEncoding = {
      ...original,
      buffer: true,
      type: `${original.type ?? "value"}+${AES_256_GCM}`,
      encode: (data: any) => {
        const encoded = original.encode(data);
        return encryptLeveldbValue(toBuffer(encoded), keyBuf);
      },
      decode: (data: any) => {
        const decrypted = decryptLeveldbValue(toBuffer(data), keyBuf, strict);
        return original.decode(decrypted);
      },
    };

    return baseLevel(location, { ...opts, valueEncoding });
  };
}

