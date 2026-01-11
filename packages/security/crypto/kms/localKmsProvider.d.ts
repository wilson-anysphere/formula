/// <reference types="node" />

import type { EncryptionContext, EnvelopeKmsProvider } from "../envelope.js";

export type LocalKmsWrappedKey = {
  kmsProvider: "local";
  kmsKeyVersion: number;
  algorithm: string;
  iv: string;
  ciphertext: string;
  tag: string;
};

export type LocalKmsProviderJSON = {
  provider: "local";
  currentVersion: number;
  keks: Record<string, string>;
};

export class LocalKmsProvider implements EnvelopeKmsProvider<LocalKmsWrappedKey> {
  readonly provider: "local";
  currentVersion: number;

  constructor(opts?: { currentVersion?: number; keksByVersion?: Map<number, Buffer> });

  rotateKey(): number;

  wrapKey(args: { plaintextKey: Buffer; encryptionContext?: EncryptionContext }): Promise<LocalKmsWrappedKey>;
  unwrapKey(args: { wrappedKey: LocalKmsWrappedKey; encryptionContext?: EncryptionContext }): Promise<Buffer>;

  toJSON(): LocalKmsProviderJSON;
  static fromJSON(value: unknown): LocalKmsProvider;
}

