export class AwsKmsProvider {
  constructor() {
    this.provider = "aws";
  }

  wrapKey() {
    throw new Error("AwsKmsProvider.wrapKey is not implemented in this reference repo");
  }

  unwrapKey() {
    throw new Error("AwsKmsProvider.unwrapKey is not implemented in this reference repo");
  }
}

export class GcpKmsProvider {
  constructor() {
    this.provider = "gcp";
  }

  wrapKey() {
    throw new Error("GcpKmsProvider.wrapKey is not implemented in this reference repo");
  }

  unwrapKey() {
    throw new Error("GcpKmsProvider.unwrapKey is not implemented in this reference repo");
  }
}

export class AzureKeyVaultProvider {
  constructor() {
    this.provider = "azure";
  }

  wrapKey() {
    throw new Error("AzureKeyVaultProvider.wrapKey is not implemented in this reference repo");
  }

  unwrapKey() {
    throw new Error("AzureKeyVaultProvider.unwrapKey is not implemented in this reference repo");
  }
}

