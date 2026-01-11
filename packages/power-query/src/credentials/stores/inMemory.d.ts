export class InMemoryCredentialStore {
  constructor();

  get(scope: any): Promise<{ id: string; secret: unknown } | null>;

  set(scope: any, secret: unknown): Promise<{ id: string; secret: unknown }>;

  delete(scope: any): Promise<void>;
}
