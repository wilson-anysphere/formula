import { CredentialManager } from "../../../../packages/power-query/src/credentials/manager.js";
import { InMemoryCredentialStore } from "../../../../packages/power-query/src/credentials/stores/inMemory.js";
import { hasTauriInvoke, TauriCredentialStore } from "./tauriCredentialStore.ts";

export type PowerQueryCredentialPrompt = (args: {
  connectorId: string;
  scope: unknown;
  request: unknown;
}) => Promise<unknown | null | undefined>;

/**
 * Minimal desktop integration helper for Power Query credentials.
 *
 * When running in Tauri, this defaults to a persistent, OS-keychain-backed
 * credential store. Tests can still override the store implementation.
 * For now this returns an `onCredentialRequest` implementation that can be
 * passed into `new QueryEngine({ onCredentialRequest })`.
 */
export function createPowerQueryCredentialManager(opts?: {
  store?: {
    get: (scope: any) => Promise<{ id: string; secret: unknown } | null>;
    set: (scope: any, secret: unknown) => Promise<{ id: string; secret: unknown }>;
    delete: (scope: any) => Promise<void>;
  };
  prompt?: PowerQueryCredentialPrompt;
}) {
  const store = opts?.store ?? (hasTauriInvoke() ? new TauriCredentialStore() : new InMemoryCredentialStore());
  const manager = new CredentialManager({ store, prompt: opts?.prompt });
  return {
    store,
    manager,
    onCredentialRequest: manager.onCredentialRequest.bind(manager),
  };
}
