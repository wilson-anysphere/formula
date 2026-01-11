import type { KeyRing } from "../../../packages/security/crypto/keyring.js";

const AES_GCM_IV_BYTES = 12;
const AES_GCM_TAG_BYTES = 16;

const DEFAULT_MAGIC = Buffer.from("FMLLDB01", "ascii");

export const DEFAULT_LEVELDB_VALUE_MAGIC = DEFAULT_MAGIC;

type LevelValueEncoding = {
  encode: (value: unknown) => unknown;
  decode: (value: unknown) => unknown;
  buffer?: boolean;
  type?: string;
  [key: string]: unknown;
};

export type LevelAdapter = (location: string, opts?: Record<string, unknown>) => any;

export type CreateEncryptedLevelAdapterOptions = {
  keyRing: KeyRing;
  strict: boolean;
  /**
   * 8-byte magic header to distinguish encrypted blobs from plaintext values.
   *
   * Defaults to `FMLLDB01`.
   */
  magic?: Buffer;
};

const LEVELDB_AAD_CONTEXT = {
  scope: "formula-sync-server-leveldb",
  schemaVersion: 1,
} as const;

/**
 * Identity encoding that can be used to bypass the encrypted `valueEncoding` in tests.
 */
export const RAW_VALUE_ENCODING: LevelValueEncoding = {
  buffer: true,
  type: "raw",
  encode: (value: unknown) => value,
  decode: (value: unknown) => value,
};

function assertU32(value: number, name: string) {
  if (!Number.isInteger(value) || value < 0 || value > 0xffffffff) {
    throw new RangeError(`${name} must be a uint32`);
  }
}

function asBuffer(value: unknown, name: string): Buffer {
  if (Buffer.isBuffer(value)) return value;
  if (value instanceof Uint8Array) return Buffer.from(value);
  throw new TypeError(`${name} must be a Buffer or Uint8Array`);
}

function isEncryptedBytes(bytes: Buffer, magic: Buffer): boolean {
  if (bytes.length < magic.length + 4 + AES_GCM_IV_BYTES + AES_GCM_TAG_BYTES) {
    return false;
  }
  return bytes.subarray(0, magic.length).equals(magic);
}

function encodeEncryptedBytes(
  encrypted: { keyVersion: number; iv: Buffer; tag: Buffer; ciphertext: Buffer },
  magic: Buffer
): Buffer {
  assertU32(encrypted.keyVersion, "keyVersion");
  if (!Buffer.isBuffer(encrypted.iv) || encrypted.iv.length !== AES_GCM_IV_BYTES) {
    throw new RangeError(`iv must be ${AES_GCM_IV_BYTES} bytes`);
  }
  if (!Buffer.isBuffer(encrypted.tag) || encrypted.tag.length !== AES_GCM_TAG_BYTES) {
    throw new RangeError(`tag must be ${AES_GCM_TAG_BYTES} bytes`);
  }
  if (!Buffer.isBuffer(encrypted.ciphertext)) {
    throw new TypeError("ciphertext must be a Buffer");
  }

  const headerBytes = magic.length + 4 + AES_GCM_IV_BYTES + AES_GCM_TAG_BYTES;
  const out = Buffer.allocUnsafe(headerBytes + encrypted.ciphertext.length);

  let offset = 0;
  magic.copy(out, offset);
  offset += magic.length;
  out.writeUInt32BE(encrypted.keyVersion, offset);
  offset += 4;
  encrypted.iv.copy(out, offset);
  offset += AES_GCM_IV_BYTES;
  encrypted.tag.copy(out, offset);
  offset += AES_GCM_TAG_BYTES;
  encrypted.ciphertext.copy(out, offset);

  return out;
}

function decodeEncryptedBytes(
  bytes: Buffer,
  magic: Buffer
): { keyVersion: number; iv: Buffer; tag: Buffer; ciphertext: Buffer } {
  if (!isEncryptedBytes(bytes, magic)) {
    throw new Error("bytes are not in encrypted value format");
  }

  const keyVersion = bytes.readUInt32BE(magic.length);
  const ivOffset = magic.length + 4;
  const tagOffset = ivOffset + AES_GCM_IV_BYTES;
  const ciphertextOffset = tagOffset + AES_GCM_TAG_BYTES;

  return {
    keyVersion,
    iv: bytes.subarray(ivOffset, ivOffset + AES_GCM_IV_BYTES),
    tag: bytes.subarray(tagOffset, tagOffset + AES_GCM_TAG_BYTES),
    ciphertext: bytes.subarray(ciphertextOffset),
  };
}

function wrapValueEncoding(params: {
  upstream: LevelValueEncoding;
  keyRing: KeyRing;
  strict: boolean;
  magic: Buffer;
}): LevelValueEncoding {
  const { upstream, keyRing, strict, magic } = params;

  const upstreamEncode = upstream.encode.bind(upstream);
  const upstreamDecode = upstream.decode.bind(upstream);

  return {
    ...upstream,
    buffer: true,
    type: upstream.type ? `encrypted(${String(upstream.type)})` : "encrypted",
    encode(value: unknown): Buffer {
      const encoded = upstreamEncode(value);
      const plaintext = asBuffer(encoded, "valueEncoding.encode()");

      const encrypted = keyRing.encryptBytes(plaintext, { aadContext: LEVELDB_AAD_CONTEXT });
      return encodeEncryptedBytes(encrypted, magic);
    },
    decode(value: unknown): unknown {
      const bytes = asBuffer(value, "valueEncoding.decode()");
      const hasMagic =
        bytes.length >= magic.length && bytes.subarray(0, magic.length).equals(magic);
      if (!hasMagic) {
        if (strict) {
          throw new Error(
            `Encountered unencrypted LevelDB value (missing ${magic.toString("ascii")} header). ` +
              "If this is legacy data, restart with SYNC_SERVER_PERSISTENCE_ENCRYPTION_STRICT=0 to allow migration."
          );
        }
        return upstreamDecode(bytes);
      }

      // If the magic header is present, treat truncation as corruption even in
      // non-strict mode (non-strict is only for legacy plaintext values).
      if (!isEncryptedBytes(bytes, magic)) {
        throw new Error("Encrypted LevelDB value is truncated (missing header bytes).");
      }

      const payload = decodeEncryptedBytes(bytes, magic);
      const plaintext = keyRing.decryptBytes(payload, { aadContext: LEVELDB_AAD_CONTEXT });
      return upstreamDecode(plaintext);
    },
  };
}

export function createEncryptedLevelAdapter(
  opts: CreateEncryptedLevelAdapterOptions
): (baseLevel: LevelAdapter) => LevelAdapter {
  const magic = opts.magic ?? DEFAULT_MAGIC;
  if (!Buffer.isBuffer(magic) || magic.length !== 8) {
    throw new RangeError("magic must be an 8-byte Buffer");
  }

  return (baseLevel: LevelAdapter) => {
    return (location: string, levelOptions: Record<string, unknown> = {}) => {
      const upstreamEncoding = levelOptions.valueEncoding;
      if (!upstreamEncoding || typeof upstreamEncoding !== "object") {
        throw new TypeError(
          "Encrypted LevelDB adapter requires valueEncoding to be an object with encode/decode functions"
        );
      }

      const valueEncoding = upstreamEncoding as LevelValueEncoding;
      if (typeof valueEncoding.encode !== "function" || typeof valueEncoding.decode !== "function") {
        throw new TypeError(
          "Encrypted LevelDB adapter requires valueEncoding.encode and valueEncoding.decode"
        );
      }

      return baseLevel(location, {
        ...levelOptions,
        valueEncoding: wrapValueEncoding({
          upstream: valueEncoding,
          keyRing: opts.keyRing,
          strict: opts.strict,
          magic,
        }),
      });
    };
  };
}
