import { oauth2Scope } from "../credentials/scopes.js";

/**
 * @typedef {import("./tokenStore.js").OAuth2TokenStore} OAuth2TokenStore
 * @typedef {import("./tokenStore.js").OAuth2TokenStoreEntry} OAuth2TokenStoreEntry
 * @typedef {import("./tokenStore.js").OAuth2TokenStoreKey} OAuth2TokenStoreKey
 * @typedef {import("../credentials/store.js").CredentialStore} CredentialStore
 */

/**
 * OAuth2TokenStore implementation backed by the Power Query credential store
 * framework (Task 46).
 *
 * This allows host apps to persist OAuth refresh tokens using:
 * - `KeychainCredentialStore` (OS keychain)
 * - `EncryptedFileCredentialStore` (encrypted-at-rest on disk, Node)
 * - or any other `CredentialStore` implementation
 *
 * @implements {OAuth2TokenStore}
 */
export class CredentialStoreOAuthTokenStore {
  /**
   * @param {CredentialStore} store
   */
  constructor(store) {
    if (!store) throw new TypeError("store is required");
    this.store = store;
  }

  /**
   * @param {OAuth2TokenStoreKey} key
   */
  scope(key) {
    return oauth2Scope({ providerId: key.providerId, scopesHash: key.scopesHash });
  }

  /**
   * @param {OAuth2TokenStoreKey} key
   * @returns {Promise<OAuth2TokenStoreEntry | null>}
   */
  async get(key) {
    const entry = await this.store.get(this.scope(key));
    if (!entry) return null;
    const secret = entry.secret;
    if (!secret || typeof secret !== "object") return null;

    // @ts-ignore - runtime access
    const providerId = typeof secret.providerId === "string" ? secret.providerId : key.providerId;
    // @ts-ignore - runtime access
    const scopesHash = typeof secret.scopesHash === "string" ? secret.scopesHash : key.scopesHash;
    // @ts-ignore - runtime access
    const scopes = Array.isArray(secret.scopes) ? secret.scopes.filter((s) => typeof s === "string") : [];
    /** @type {string | null} */
    // @ts-ignore - runtime access
    const refreshToken = typeof secret.refreshToken === "string" ? secret.refreshToken : null;
    /** @type {string | null | undefined} */
    // @ts-ignore - runtime access
    const accessToken = typeof secret.accessToken === "string" ? secret.accessToken : secret.accessToken === null ? null : undefined;
    /** @type {number | null | undefined} */
    // @ts-ignore - runtime access
    const expiresAtMs = typeof secret.expiresAtMs === "number" ? secret.expiresAtMs : secret.expiresAtMs === null ? null : undefined;

    return { providerId, scopesHash, scopes, refreshToken, accessToken, expiresAtMs };
  }

  /**
   * @param {OAuth2TokenStoreKey} key
   * @param {OAuth2TokenStoreEntry} value
   * @returns {Promise<void>}
   */
  async set(key, value) {
    await this.store.set(this.scope(key), value);
  }

  /**
   * @param {OAuth2TokenStoreKey} key
   * @returns {Promise<void>}
   */
  async delete(key) {
    await this.store.delete(this.scope(key));
  }
}
