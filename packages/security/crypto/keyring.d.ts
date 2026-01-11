/// <reference types="node" />

export type KeyRingEncryptedBytes = {
  keyVersion: number;
  iv: Buffer;
  ciphertext: Buffer;
  tag: Buffer;
};

export class KeyRing {
  static create(): KeyRing;
  static fromJSON(value: unknown): KeyRing;

  currentVersion: number;

  rotate(): number;

  encryptBytes(
    plaintext: Buffer,
    opts?: { aadContext?: unknown | null }
  ): KeyRingEncryptedBytes;

  decryptBytes(
    encrypted: KeyRingEncryptedBytes,
    opts?: { aadContext?: unknown | null }
  ): Buffer;

  toJSON(): unknown;
}

