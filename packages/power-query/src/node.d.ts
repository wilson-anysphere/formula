/**
 * Node-only TypeScript surface for `@formula/power-query/node`.
 *
 * This entrypoint re-exports the main `@formula/power-query` API plus a small
 * set of helpers that depend on Node built-ins (e.g. `node:crypto`).
 */

export * from "./index.d.ts";

export class EncryptedFileSystemCacheStore implements CacheStore {
  constructor(options: any);

  get(key: string): Promise<CacheEntry | null>;
  set(key: string, entry: CacheEntry): Promise<void>;
  delete(key: string): Promise<void>;
  clear(): Promise<void>;
  pruneExpired(nowMs?: number): Promise<void>;
  prune(options: { nowMs: number; maxEntries?: number; maxBytes?: number }): Promise<void>;
}

export function createNodeCryptoCacheProvider(options: { keyVersion: number; keyBytes: Uint8Array }): CacheCryptoProvider;

export function createNodeCredentialStore(opts: { filePath: string; keychainProvider?: any; service?: string }): CredentialStore;
