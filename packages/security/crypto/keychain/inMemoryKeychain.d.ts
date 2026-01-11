/// <reference types="node" />

export class InMemoryKeychainProvider {
  constructor();
  getSecret(args: { service: string; account: string }): Promise<Buffer | null>;
  setSecret(args: { service: string; account: string; secret: Buffer }): Promise<void>;
  deleteSecret(args: { service: string; account: string }): Promise<void>;
}
