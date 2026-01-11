import type { DecryptKeyParams, EncryptKeyParams, EncryptKeyResult, KmsProvider } from "./types";

type AwsKmsClient = {
  send(command: unknown): Promise<any>;
};

function loadAwsSdk(): any {
  try {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    return require("@aws-sdk/client-kms");
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    throw new Error(
      [
        "AwsKmsProvider is not available in this build.",
        "Install @aws-sdk/client-kms and set AWS_KMS_ENABLED=true to enable it.",
        `Underlying error: ${message}`
      ].join(" ")
    );
  }
}

/**
 * Optional AWS KMS implementation.
 *
 * This is intentionally dependency-light: the AWS SDK is loaded lazily so local
 * dev/test environments do not require AWS packages/credentials.
 */
export class AwsKmsProvider implements KmsProvider {
  readonly provider = "aws";
  private readonly region: string;
  private client: AwsKmsClient | null = null;

  constructor({ region }: { region: string }) {
    if (!region) throw new Error("AwsKmsProvider requires region");
    this.region = region;
  }

  private getClient(): AwsKmsClient {
    const existing = this.client;
    if (existing) return existing;
    const sdk = loadAwsSdk();
    // eslint-disable-next-line new-cap
    const client: AwsKmsClient = new sdk.KMSClient({ region: this.region });
    this.client = client;
    return client;
  }

  async encryptKey({ plaintextDek, orgId, keyId }: EncryptKeyParams): Promise<EncryptKeyResult> {
    const sdk = loadAwsSdk();
    const client = this.getClient();
    const res = await client.send(
      new sdk.EncryptCommand({
        KeyId: keyId,
        Plaintext: plaintextDek,
        EncryptionContext: { orgId }
      })
    );

    const ciphertext = res.CiphertextBlob ? Buffer.from(res.CiphertextBlob) : null;
    if (!ciphertext) {
      throw new Error("AWS KMS EncryptCommand returned empty CiphertextBlob");
    }

    return {
      encryptedDek: ciphertext,
      kmsKeyId: (res.KeyId as string | undefined) ?? keyId
    };
  }

  async decryptKey({ encryptedDek, orgId, kmsKeyId }: DecryptKeyParams): Promise<Buffer> {
    const sdk = loadAwsSdk();
    const client = this.getClient();
    const res = await client.send(
      new sdk.DecryptCommand({
        CiphertextBlob: encryptedDek,
        KeyId: kmsKeyId,
        EncryptionContext: { orgId }
      })
    );

    const plaintext = res.Plaintext ? Buffer.from(res.Plaintext) : null;
    if (!plaintext) {
      throw new Error("AWS KMS DecryptCommand returned empty Plaintext");
    }
    return plaintext;
  }
}
