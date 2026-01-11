/// <reference types="node" />

export const AES_256_GCM: "aes-256-gcm";
export const AES_GCM_IV_BYTES: 12;
export const AES_GCM_TAG_BYTES: 16;
export const AES_256_KEY_BYTES: 32;

export type Aes256GcmEncryptedPayload = {
  algorithm: string;
  iv: Buffer;
  ciphertext: Buffer;
  tag: Buffer;
};

export type SerializedEncryptedPayload = {
  algorithm: string;
  iv: string;
  ciphertext: string;
  tag: string;
};

export function generateAes256Key(): Buffer;

export function encryptAes256Gcm(args: {
  plaintext: Buffer;
  key: Buffer;
  aad?: Buffer | null;
  iv?: Buffer | null;
}): Aes256GcmEncryptedPayload;

export function decryptAes256Gcm(args: {
  ciphertext: Buffer;
  key: Buffer;
  iv: Buffer;
  tag: Buffer;
  aad?: Buffer | null;
}): Buffer;

export function serializeEncryptedPayload(payload: Aes256GcmEncryptedPayload): SerializedEncryptedPayload;
export function deserializeEncryptedPayload(payload: SerializedEncryptedPayload): Aes256GcmEncryptedPayload;
