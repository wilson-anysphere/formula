import { createNodeCredentialStore } from "../credentials/node.js";
import { CredentialStoreOAuthTokenStore } from "./credentialStoreTokenStore.js";
import { OAuth2Manager } from "./manager.js";

/**
 * Node convenience helper: create an OAuth token store backed by the credential
 * store framework.
 *
 * On Linux/Windows this typically uses an encrypted file store (since the
 * keychain CLI helpers are read-only today). On macOS it can use the OS keychain
 * directly.
 *
 * @param {{
 *   filePath: string;
 *   keychainProvider?: any;
 *   service?: string;
 * }} opts
 */
export function createNodeOAuthTokenStore(opts) {
  const credentialStore = createNodeCredentialStore(opts);
  const tokenStore = new CredentialStoreOAuthTokenStore(credentialStore);
  return { credentialStore, tokenStore };
}

/**
 * Node convenience helper: create an `OAuth2Manager` with a secure persistent token store.
 *
 * @param {{
 *   filePath: string;
 *   keychainProvider?: any;
 *   service?: string;
 *   fetch?: typeof fetch;
 *   now?: (() => number);
 *   clockSkewMs?: number;
 *   persistAccessToken?: boolean;
 * }} opts
 */
export function createNodeOAuth2Manager(opts) {
  const { credentialStore, tokenStore } = createNodeOAuthTokenStore({
    filePath: opts.filePath,
    keychainProvider: opts.keychainProvider,
    service: opts.service,
  });
  const manager = new OAuth2Manager({
    tokenStore,
    fetch: opts.fetch,
    now: opts.now,
    clockSkewMs: opts.clockSkewMs,
    persistAccessToken: opts.persistAccessToken,
  });
  return { manager, tokenStore, credentialStore };
}

