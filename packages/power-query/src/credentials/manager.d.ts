export class CredentialManager {
  constructor(options?: any);

  onCredentialRequest(connectorId: string, details: unknown): Promise<unknown>;
}
