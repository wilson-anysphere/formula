import { CredentialStoreOAuthTokenStore, InMemoryCredentialStore, OAuth2Manager } from "@formula/power-query";

import { oauthBroker } from "./oauthBroker.ts";
import { hasTauriInvoke, TauriCredentialStore } from "./tauriCredentialStore.ts";

export function createDesktopOAuth2Manager(opts?: {
  /**
   * Credential store override (useful for tests). Defaults to the persistent Tauri
   * store when available.
   */
  store?: {
    get: (scope: any) => Promise<{ id: string; secret: unknown } | null>;
    set: (scope: any, secret: unknown) => Promise<{ id: string; secret: unknown }>;
    delete: (scope: any) => Promise<void>;
  };
  fetch?: typeof fetch;
  persistAccessToken?: boolean;
}) {
  const store = opts?.store ?? (hasTauriInvoke() ? new TauriCredentialStore() : new InMemoryCredentialStore());
  const tokenStore = new CredentialStoreOAuthTokenStore(store);
  const oauth2 = new OAuth2Manager({
    tokenStore,
    fetch: opts?.fetch,
    persistAccessToken: opts?.persistAccessToken,
  });

  return { oauth2, broker: oauthBroker, store, tokenStore };
}
