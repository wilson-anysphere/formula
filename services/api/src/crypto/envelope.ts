import { AES_256_GCM, decryptAes256Gcm, encryptAes256Gcm, generateAes256Key } from "./aes256gcm";
import type { KmsProvider } from "./kms";
import { aadFromContext, canonicalJson } from "./utils";

export const ENVELOPE_VERSION = 1 as const;

export type EnvelopeAadContext = Record<string, unknown>;

export type EncryptedEnvelopeV1 = {
  envelopeVersion: typeof ENVELOPE_VERSION;
  algorithm: typeof AES_256_GCM;
  ciphertext: Buffer;
  iv: Buffer;
  tag: Buffer;
  encryptedDek: Buffer;
  kmsProvider: string;
  kmsKeyId: string;
  aad: EnvelopeAadContext;
};

export type EncryptedEnvelope = EncryptedEnvelopeV1;

function ensureAadContext(value: EnvelopeAadContext): void {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new TypeError("aadContext must be a plain object");
  }
  // Ensure JSON-serializable determinism early (throws on circular refs).
  canonicalJson(value);
}

/**
 * Encrypt data using a per-object DEK (data encryption key), wrapped via the
 * supplied KMS provider (envelope encryption).
 */
export async function encryptEnvelope({
  plaintext,
  kmsProvider,
  orgId,
  keyId,
  aadContext
}: {
  plaintext: Buffer;
  kmsProvider: KmsProvider;
  orgId: string;
  keyId: string;
  aadContext: EnvelopeAadContext;
}): Promise<EncryptedEnvelopeV1> {
  if (!Buffer.isBuffer(plaintext)) {
    throw new TypeError("plaintext must be a Buffer");
  }
  ensureAadContext(aadContext);

  const dek = generateAes256Key();
  const aad = aadFromContext(aadContext);

  const encrypted = encryptAes256Gcm({ plaintext, key: dek, aad });
  const wrapped = await kmsProvider.encryptKey({ plaintextDek: dek, orgId, keyId });

  return {
    envelopeVersion: ENVELOPE_VERSION,
    algorithm: AES_256_GCM,
    ciphertext: encrypted.ciphertext,
    iv: encrypted.iv,
    tag: encrypted.tag,
    encryptedDek: wrapped.encryptedDek,
    kmsProvider: kmsProvider.provider,
    kmsKeyId: wrapped.kmsKeyId,
    aad: structuredClone(aadContext)
  };
}

export async function decryptEnvelope({
  envelope,
  kmsProvider,
  orgId,
  aadContext
}: {
  envelope: EncryptedEnvelope;
  kmsProvider: KmsProvider;
  orgId: string;
  aadContext: EnvelopeAadContext;
}): Promise<Buffer> {
  if (envelope.envelopeVersion !== ENVELOPE_VERSION) {
    throw new Error(`Unsupported envelopeVersion: ${envelope.envelopeVersion}`);
  }
  if (envelope.algorithm !== AES_256_GCM) {
    throw new Error(`Unsupported algorithm: ${envelope.algorithm}`);
  }
  ensureAadContext(aadContext);
  if (kmsProvider.provider !== envelope.kmsProvider) {
    throw new Error(`KMS provider mismatch (expected ${envelope.kmsProvider}, got ${kmsProvider.provider})`);
  }

  const dek = await kmsProvider.decryptKey({
    encryptedDek: envelope.encryptedDek,
    orgId,
    kmsKeyId: envelope.kmsKeyId
  });

  const aad = aadFromContext(aadContext);
  return decryptAes256Gcm({
    ciphertext: envelope.ciphertext,
    key: dek,
    iv: envelope.iv,
    tag: envelope.tag,
    aad
  });
}

