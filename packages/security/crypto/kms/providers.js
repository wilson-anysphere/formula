let cachedAwsSdk = null;

async function loadAwsSdk() {
  if (cachedAwsSdk) return cachedAwsSdk;
  try {
    cachedAwsSdk = await import("@aws-sdk/client-kms");
    return cachedAwsSdk;
  } catch (err) {
    cachedAwsSdk = null;
    const message = err instanceof Error ? err.message : String(err);
    throw new Error(
      [
        "AwsKmsProvider is not available in this build.",
        "Install @aws-sdk/client-kms to enable it.",
        `Underlying error: ${message}`
      ].join(" ")
    );
  }
}

function awsEncryptionContext(encryptionContext) {
  if (!encryptionContext || typeof encryptionContext !== "object") return undefined;
  const orgId = encryptionContext.orgId;
  if (typeof orgId === "string" && orgId) return { orgId };
  return undefined;
}

export class AwsKmsProvider {
  constructor({ region, keyId = null } = {}) {
    this.provider = "aws";
    if (!region) {
      throw new Error("AwsKmsProvider requires region");
    }
    this.region = region;
    this.keyId = keyId;
    this.client = null;
  }

  async _getClient() {
    if (this.client) return this.client;
    const sdk = await loadAwsSdk();
    this.client = new sdk.KMSClient({ region: this.region });
    return this.client;
  }

  async wrapKey({ plaintextKey, encryptionContext = null }) {
    if (!this.keyId) {
      throw new Error("AwsKmsProvider.wrapKey requires keyId");
    }
    if (!Buffer.isBuffer(plaintextKey)) {
      throw new TypeError("plaintextKey must be a Buffer");
    }

    const sdk = await loadAwsSdk();
    const client = await this._getClient();
    const res = await client.send(
      new sdk.EncryptCommand({
        KeyId: this.keyId,
        Plaintext: plaintextKey,
        EncryptionContext: awsEncryptionContext(encryptionContext)
      })
    );

    const ciphertext = res.CiphertextBlob ? Buffer.from(res.CiphertextBlob) : null;
    if (!ciphertext) {
      throw new Error("AWS KMS EncryptCommand returned empty CiphertextBlob");
    }

    return {
      kmsProvider: this.provider,
      kmsKeyId: (res.KeyId ?? this.keyId) || this.keyId,
      ciphertext: ciphertext.toString("base64")
    };
  }

  async unwrapKey({ wrappedKey, encryptionContext = null }) {
    if (!wrappedKey || typeof wrappedKey !== "object") {
      throw new TypeError("wrappedKey must be an object");
    }
    if (wrappedKey.kmsProvider !== this.provider) {
      throw new Error(`Unsupported kmsProvider: ${wrappedKey.kmsProvider}`);
    }
    if (typeof wrappedKey.ciphertext !== "string") {
      throw new TypeError("wrappedKey.ciphertext must be a base64 string");
    }
    if (typeof wrappedKey.kmsKeyId !== "string" || !wrappedKey.kmsKeyId) {
      throw new TypeError("wrappedKey.kmsKeyId must be a string");
    }

    const sdk = await loadAwsSdk();
    const client = await this._getClient();
    const res = await client.send(
      new sdk.DecryptCommand({
        CiphertextBlob: Buffer.from(wrappedKey.ciphertext, "base64"),
        KeyId: wrappedKey.kmsKeyId,
        EncryptionContext: awsEncryptionContext(encryptionContext)
      })
    );

    const plaintext = res.Plaintext ? Buffer.from(res.Plaintext) : null;
    if (!plaintext) {
      throw new Error("AWS KMS DecryptCommand returned empty Plaintext");
    }
    return plaintext;
  }
}

export class GcpKmsProvider {
  constructor() {
    this.provider = "gcp";
  }

  async wrapKey() {
    throw new Error("GcpKmsProvider.wrapKey is not implemented in this reference repo");
  }

  async unwrapKey() {
    throw new Error("GcpKmsProvider.unwrapKey is not implemented in this reference repo");
  }
}

export class AzureKeyVaultProvider {
  constructor() {
    this.provider = "azure";
  }

  async wrapKey() {
    throw new Error("AzureKeyVaultProvider.wrapKey is not implemented in this reference repo");
  }

  async unwrapKey() {
    throw new Error("AzureKeyVaultProvider.unwrapKey is not implemented in this reference repo");
  }
}
