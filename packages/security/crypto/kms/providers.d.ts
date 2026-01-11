/// <reference types="node" />

import type { EncryptionContext, EnvelopeKmsProvider } from "../envelope.js";

export type AwsKmsWrappedKey = {
  kmsProvider: "aws";
  kmsKeyId: string;
  ciphertext: string;
};

export class AwsKmsProvider implements EnvelopeKmsProvider<AwsKmsWrappedKey> {
  readonly provider: "aws";

  constructor(opts: { region: string; keyId?: string | null });

  wrapKey(args: { plaintextKey: Buffer; encryptionContext?: EncryptionContext }): Promise<AwsKmsWrappedKey>;
  unwrapKey(args: { wrappedKey: AwsKmsWrappedKey; encryptionContext?: EncryptionContext }): Promise<Buffer>;
}

export class GcpKmsProvider implements EnvelopeKmsProvider {
  readonly provider: "gcp";
  wrapKey(args: { plaintextKey: Buffer; encryptionContext?: EncryptionContext }): Promise<unknown>;
  unwrapKey(args: { wrappedKey: unknown; encryptionContext?: EncryptionContext }): Promise<Buffer>;
}

export class AzureKeyVaultProvider implements EnvelopeKmsProvider {
  readonly provider: "azure";
  wrapKey(args: { plaintextKey: Buffer; encryptionContext?: EncryptionContext }): Promise<unknown>;
  unwrapKey(args: { wrappedKey: unknown; encryptionContext?: EncryptionContext }): Promise<Buffer>;
}

