export type EncryptKeyParams = {
  plaintextDek: Buffer;
  orgId: string;
  /**
   * The KMS key identifier (provider-specific). For AWS this is a KeyId/ARN/alias.
   */
  keyId: string;
};

export type EncryptKeyResult = {
  /**
   * Provider-specific wrapped DEK. Must be treated as opaque bytes.
   */
  encryptedDek: Buffer;
  /**
   * The KMS key id used to produce `encryptedDek` (may differ from the requested `keyId`).
   */
  kmsKeyId: string;
};

export type DecryptKeyParams = {
  encryptedDek: Buffer;
  orgId: string;
  kmsKeyId: string;
};

export interface KmsProvider {
  /**
   * Stable identifier for storage and debugging (e.g. "local", "aws").
   */
  readonly provider: string;

  encryptKey(params: EncryptKeyParams): Promise<EncryptKeyResult>;
  decryptKey(params: DecryptKeyParams): Promise<Buffer>;
}

