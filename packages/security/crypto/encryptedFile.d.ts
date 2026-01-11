/// <reference types="node" />

export function isEncryptedFileBytes(bytes: Buffer): boolean;
export function encodeEncryptedFileBytes(args: {
  keyVersion: number;
  iv: Buffer;
  tag: Buffer;
  ciphertext: Buffer;
}): Buffer;
export function decodeEncryptedFileBytes(bytes: Buffer): {
  keyVersion: number;
  iv: Buffer;
  tag: Buffer;
  ciphertext: Buffer;
};
