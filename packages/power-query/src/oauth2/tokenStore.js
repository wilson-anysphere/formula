import { hashValue } from "../cache/key.js";

/**
 * OAuth token storage.
 *
 * This is intentionally an interface (via JSDoc) so host applications can plug
 * in a secure implementation (keychain, encrypted database, etc).
 *
 * Security note:
 * - Refresh tokens SHOULD be stored encrypted at rest by production stores.
 * - Access tokens may be treated as ephemeral (in-memory) to reduce risk, but a
 *   store MAY persist them encrypted to improve startup performance.
 */

/**
 * @typedef {Object} OAuth2TokenStoreKey
 * @property {string} providerId Stable provider/config identifier.
 * @property {string} scopesHash Hash of normalized scope list.
 */

/**
 * @typedef {Object} OAuth2TokenStoreEntry
 * @property {string} providerId
 * @property {string} scopesHash
 * @property {string[]} scopes Normalized scope list.
 * @property {string | null} refreshToken
 * @property {string | null | undefined} [accessToken]
 * @property {number | null | undefined} [expiresAtMs]
 */

/**
 * @typedef {Object} OAuth2TokenStore
 * @property {(key: OAuth2TokenStoreKey) => Promise<OAuth2TokenStoreEntry | null>} get
 * @property {(key: OAuth2TokenStoreKey, entry: OAuth2TokenStoreEntry) => Promise<void>} set
 * @property {(key: OAuth2TokenStoreKey) => Promise<void>} delete
 */

/**
 * @param {string[] | string | undefined} scopes
 * @returns {{ scopes: string[], scopesHash: string }}
 */
export function normalizeScopes(scopes) {
  /** @type {unknown[]} */
  let raw = [];
  if (Array.isArray(scopes)) raw = scopes;
  else if (typeof scopes === "string") raw = scopes.split(/[\s,]+/).filter(Boolean);
  const normalized = raw
    .filter((s) => typeof s === "string")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  normalized.sort();
  const deduped = Array.from(new Set(normalized));
  return { scopes: deduped, scopesHash: hashValue(deduped.join(" ")) };
}

/**
 * Simple in-memory store used in tests and environments where a secure store is
 * not available.
 *
 * @implements {OAuth2TokenStore}
 */
export class InMemoryOAuthTokenStore {
  /**
   * @param {Record<string, OAuth2TokenStoreEntry> | undefined} [snapshot]
   */
  constructor(snapshot) {
    /** @type {Map<string, OAuth2TokenStoreEntry>} */
    this.entries = new Map();
    if (snapshot && typeof snapshot === "object") {
      for (const [key, value] of Object.entries(snapshot)) {
        if (!value || typeof value !== "object") continue;
        // @ts-ignore - trust snapshot
        this.entries.set(key, value);
      }
    }
  }

  /**
   * @param {OAuth2TokenStoreKey} key
   */
  static keyString(key) {
    return `${key.providerId}:${key.scopesHash}`;
  }

  /** @returns {Record<string, OAuth2TokenStoreEntry>} */
  snapshot() {
    /** @type {Record<string, OAuth2TokenStoreEntry>} */
    const out = {};
    for (const [k, v] of this.entries.entries()) {
      out[k] = { ...v, scopes: v.scopes.slice() };
    }
    return out;
  }

  /**
   * @param {OAuth2TokenStoreKey} key
   * @returns {Promise<OAuth2TokenStoreEntry | null>}
   */
  async get(key) {
    const existing = this.entries.get(InMemoryOAuthTokenStore.keyString(key));
    if (!existing) return null;
    return { ...existing, scopes: existing.scopes.slice() };
  }

  /**
   * @param {OAuth2TokenStoreKey} key
   * @param {OAuth2TokenStoreEntry} entry
   * @returns {Promise<void>}
   */
  async set(key, entry) {
    this.entries.set(InMemoryOAuthTokenStore.keyString(key), { ...entry, scopes: entry.scopes.slice() });
  }

  /**
   * @param {OAuth2TokenStoreKey} key
   * @returns {Promise<void>}
   */
  async delete(key) {
    this.entries.delete(InMemoryOAuthTokenStore.keyString(key));
  }
}
