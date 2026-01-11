/// <reference types="node" />

export type KeyRingJSON = {
  currentVersion: number;
  keys: Record<string, string>;
};

export type KeyRingEncrypted = {
  keyVersion: number;
  algorithm: string;
  iv: string;
  ciphertext: string;
  tag: string;
};

export type KeyRingEncryptedBytes = {
  keyVersion: number;
  algorithm: string;
  iv: Buffer;
  ciphertext: Buffer;
  tag: Buffer;
};

export class KeyRing {
  currentVersion: number;
  keysByVersion: Map<number, Buffer>;

  constructor(opts: { currentVersion: number; keysByVersion: Map<number, Buffer> });

  static create(): KeyRing;
  static fromJSON(value: unknown): KeyRing;

  rotate(): number;
  getKey(version: number): Buffer;

  encrypt(plaintext: Buffer, opts?: { aadContext?: unknown | null }): KeyRingEncrypted;
  encryptBytes(
    plaintext: Buffer,
    opts?: { aadContext?: unknown | null }
  ): KeyRingEncryptedBytes;

  decrypt(encrypted: KeyRingEncrypted, opts?: { aadContext?: unknown | null }): Buffer;
  decryptBytes(
    encrypted: KeyRingEncryptedBytes,
    opts?: { aadContext?: unknown | null }
  ): Buffer;

  toJSON(): KeyRingJSON;
}
