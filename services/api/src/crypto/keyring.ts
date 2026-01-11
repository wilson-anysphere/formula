import crypto from "node:crypto";
import type { AppConfig } from "../config";
import type { KmsProvider } from "./kms";
import { AwsKmsProvider, LocalKmsProvider } from "./kms";

function deriveLocalMasterKey(secret: string): Buffer {
  if (!secret) {
    throw new Error("LOCAL_KMS_MASTER_KEY must be set (or provided via AppConfig.localKmsMasterKey)");
  }
  // Hashing makes it easy to accept arbitrary-length secrets while ensuring we
  // always end up with a 32-byte AES-256 key.
  return crypto.createHash("sha256").update(secret, "utf8").digest();
}

export class Keyring {
  private readonly local: LocalKmsProvider;
  private aws: AwsKmsProvider | null = null;
  private readonly awsEnabled: boolean;
  private readonly awsRegion: string | null;

  constructor({
    localMasterKey,
    awsKmsEnabled,
    awsRegion
  }: {
    localMasterKey: string;
    awsKmsEnabled: boolean;
    awsRegion: string | null;
  }) {
    this.local = new LocalKmsProvider({ masterKey: deriveLocalMasterKey(localMasterKey) });
    this.awsEnabled = awsKmsEnabled;
    this.awsRegion = awsRegion;
  }

  get(provider: string): KmsProvider {
    switch (provider) {
      case "local":
        return this.local;
      case "aws":
        if (!this.awsEnabled) {
          throw new Error(
            "AWS KMS provider requested but disabled (set AWS_KMS_ENABLED=true and configure AWS_REGION)"
          );
        }
        if (!this.awsRegion) {
          throw new Error("AWS KMS provider requested but AWS_REGION is not set");
        }
        if (!this.aws) {
          this.aws = new AwsKmsProvider({ region: this.awsRegion });
        }
        return this.aws;
      default:
        throw new Error(`Unsupported kms provider: ${provider}`);
    }
  }
}

export function createKeyring(config: AppConfig): Keyring {
  return new Keyring({
    localMasterKey: config.localKmsMasterKey,
    awsKmsEnabled: config.awsKmsEnabled,
    awsRegion: config.awsRegion ?? null
  });
}

