/// <reference types="node" />

export type EncryptionContext = Record<string, unknown> | null;

export interface EnvelopeKmsProvider<WrappedKey = unknown> {
  /**
   * Stable identifier for storage/debugging (e.g. "local", "aws").
   */
  readonly provider: string;
  wrapKey(args: {
    plaintextKey: Buffer;
    encryptionContext?: EncryptionContext;
  }): WrappedKey | Promise<WrappedKey>;
  unwrapKey(args: { wrappedKey: WrappedKey; encryptionContext?: EncryptionContext }): Buffer | Promise<Buffer>;
}

export type EncryptedEnvelope<WrappedKey = unknown> = {
  schemaVersion: 1;
  wrappedDek: WrappedKey;
  algorithm: string;
  iv: string;
  ciphertext: string;
  tag: string;
};

export function encryptEnvelope<WrappedKey = unknown>(args: {
  plaintext: Buffer;
  kmsProvider: EnvelopeKmsProvider<WrappedKey>;
  encryptionContext?: EncryptionContext;
}): Promise<EncryptedEnvelope<WrappedKey>>;

export function decryptEnvelope<WrappedKey = unknown>(args: {
  encryptedEnvelope: EncryptedEnvelope<WrappedKey>;
  kmsProvider: EnvelopeKmsProvider<WrappedKey>;
  encryptionContext?: EncryptionContext;
}): Promise<Buffer>;

