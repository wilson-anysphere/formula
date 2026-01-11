import crypto from "node:crypto";
import { AES_256_KEY_BYTES, AES_GCM_IV_BYTES, AES_GCM_TAG_BYTES, decryptAes256Gcm, encryptAes256Gcm } from "../aes256gcm";
import { aadFromContext, assertBufferLength } from "../utils";
import type { DecryptKeyParams, EncryptKeyParams, EncryptKeyResult, KmsProvider } from "./types";

const WRAPPED_DEK_VERSION = 1;

function encodeWrappedDekV1({
  iv,
  tag,
  ciphertext
}: {
  iv: Buffer;
  tag: Buffer;
  ciphertext: Buffer;
}): Buffer {
  assertBufferLength(iv, AES_GCM_IV_BYTES, "iv");
  assertBufferLength(tag, AES_GCM_TAG_BYTES, "tag");
  return Buffer.concat([Buffer.from([WRAPPED_DEK_VERSION]), iv, tag, ciphertext]);
}

function decodeWrappedDek(blob: Buffer): { version: number; iv: Buffer; tag: Buffer; ciphertext: Buffer } {
  if (!Buffer.isBuffer(blob)) {
    throw new TypeError("encryptedDek must be a Buffer");
  }
  const minLength = 1 + AES_GCM_IV_BYTES + AES_GCM_TAG_BYTES + 1;
  if (blob.length < minLength) {
    throw new RangeError(`encryptedDek too short (got ${blob.length} bytes)`);
  }
  const version = blob.readUInt8(0);
  if (version !== WRAPPED_DEK_VERSION) {
    throw new Error(`Unsupported wrapped DEK version: ${version}`);
  }
  const ivStart = 1;
  const tagStart = ivStart + AES_GCM_IV_BYTES;
  const ciphertextStart = tagStart + AES_GCM_TAG_BYTES;
  return {
    version,
    iv: blob.subarray(ivStart, tagStart),
    tag: blob.subarray(tagStart, ciphertextStart),
    ciphertext: blob.subarray(ciphertextStart)
  };
}

/**
 * Local KMS provider for tests and single-node dev deployments.
 *
 * This provides envelope-style key wrapping via a master secret read from config/env.
 * A per-org + per-kmsKeyId Key Encryption Key (KEK) is derived with HKDF-SHA256,
 * and used to wrap Data Encryption Keys (DEKs) using AES-256-GCM.
 *
 * In production, replace with a real KMS provider (AWS/GCP/Azure).
 */
export class LocalKmsProvider implements KmsProvider {
  readonly provider = "local";
  private readonly masterKey: Buffer;

  constructor({ masterKey }: { masterKey: Buffer }) {
    assertBufferLength(masterKey, AES_256_KEY_BYTES, "masterKey");
    this.masterKey = masterKey;
  }

  private deriveKek(orgId: string, kmsKeyId: string): Buffer {
    if (!orgId) throw new Error("orgId is required");
    if (!kmsKeyId) throw new Error("kmsKeyId is required");

    const salt = Buffer.from(`formula:local-kms:org:${orgId}`, "utf8");
    const info = Buffer.from(`formula:local-kms:kmsKeyId:${kmsKeyId}`, "utf8");
    const derived = crypto.hkdfSync("sha256", this.masterKey, salt, info, AES_256_KEY_BYTES);
    const buf = Buffer.isBuffer(derived) ? derived : Buffer.from(derived);
    assertBufferLength(buf, AES_256_KEY_BYTES, "kek");
    return buf;
  }

  private wrapAad(orgId: string, kmsKeyId: string): Buffer {
    const aad = aadFromContext({
      v: WRAPPED_DEK_VERSION,
      purpose: "dek-wrap",
      orgId,
      kmsKeyId
    });
    // aadFromContext only returns null when context is null/undefined; we always pass an object.
    return aad!;
  }

  async encryptKey({ plaintextDek, orgId, keyId }: EncryptKeyParams): Promise<EncryptKeyResult> {
    assertBufferLength(plaintextDek, AES_256_KEY_BYTES, "plaintextDek");
    if (!keyId) throw new Error("keyId is required");

    const kek = this.deriveKek(orgId, keyId);
    const encrypted = encryptAes256Gcm({
      plaintext: plaintextDek,
      key: kek,
      aad: this.wrapAad(orgId, keyId)
    });

    return {
      encryptedDek: encodeWrappedDekV1(encrypted),
      kmsKeyId: keyId
    };
  }

  async decryptKey({ encryptedDek, orgId, kmsKeyId }: DecryptKeyParams): Promise<Buffer> {
    if (!kmsKeyId) throw new Error("kmsKeyId is required");
    const parsed = decodeWrappedDek(encryptedDek);
    if (parsed.version !== WRAPPED_DEK_VERSION) {
      throw new Error(`Unsupported wrapped DEK version: ${parsed.version}`);
    }

    const kek = this.deriveKek(orgId, kmsKeyId);
    const plaintextDek = decryptAes256Gcm({
      ciphertext: parsed.ciphertext,
      key: kek,
      iv: parsed.iv,
      tag: parsed.tag,
      aad: this.wrapAad(orgId, kmsKeyId)
    });

    assertBufferLength(plaintextDek, AES_256_KEY_BYTES, "plaintextDek");
    return plaintextDek;
  }
}
