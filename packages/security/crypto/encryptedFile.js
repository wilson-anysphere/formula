import { AES_GCM_IV_BYTES, AES_GCM_TAG_BYTES } from "./aes256gcm.js";

// 8-byte magic header to distinguish encrypted blobs from plaintext (e.g. SQLite files).
const MAGIC = Buffer.from("FMLENC01", "ascii");
const HEADER_BYTES = MAGIC.length + 4 + AES_GCM_IV_BYTES + AES_GCM_TAG_BYTES;

function assertU32(value, name) {
  if (!Number.isInteger(value) || value < 0 || value > 0xffffffff) {
    throw new RangeError(`${name} must be a uint32`);
  }
}

export function isEncryptedFileBytes(bytes) {
  if (!Buffer.isBuffer(bytes)) return false;
  if (bytes.length < HEADER_BYTES) return false;
  return bytes.subarray(0, MAGIC.length).equals(MAGIC);
}

export function encodeEncryptedFileBytes({ keyVersion, iv, tag, ciphertext }) {
  assertU32(keyVersion, "keyVersion");
  if (!Buffer.isBuffer(iv) || iv.length !== AES_GCM_IV_BYTES) {
    throw new RangeError(`iv must be ${AES_GCM_IV_BYTES} bytes`);
  }
  if (!Buffer.isBuffer(tag) || tag.length !== AES_GCM_TAG_BYTES) {
    throw new RangeError(`tag must be ${AES_GCM_TAG_BYTES} bytes`);
  }
  if (!Buffer.isBuffer(ciphertext)) {
    throw new TypeError("ciphertext must be a Buffer");
  }

  const out = Buffer.allocUnsafe(HEADER_BYTES + ciphertext.length);
  let offset = 0;
  MAGIC.copy(out, offset);
  offset += MAGIC.length;
  out.writeUInt32BE(keyVersion, offset);
  offset += 4;
  iv.copy(out, offset);
  offset += AES_GCM_IV_BYTES;
  tag.copy(out, offset);
  offset += AES_GCM_TAG_BYTES;
  ciphertext.copy(out, offset);
  return out;
}

export function decodeEncryptedFileBytes(bytes) {
  if (!isEncryptedFileBytes(bytes)) {
    throw new Error("bytes are not in encrypted file format");
  }
  const keyVersion = bytes.readUInt32BE(MAGIC.length);
  const ivOffset = MAGIC.length + 4;
  const tagOffset = ivOffset + AES_GCM_IV_BYTES;
  const ciphertextOffset = tagOffset + AES_GCM_TAG_BYTES;

  return {
    keyVersion,
    iv: bytes.subarray(ivOffset, ivOffset + AES_GCM_IV_BYTES),
    tag: bytes.subarray(tagOffset, tagOffset + AES_GCM_TAG_BYTES),
    ciphertext: bytes.subarray(ciphertextOffset)
  };
}

