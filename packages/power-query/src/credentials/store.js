import { stableStringify } from "../cache/key.js";

/**
 * @typedef {import("./scopes.js").CredentialScope} CredentialScope
 */

/**
 * A stored secret plus a stable identifier that is safe to embed in cache keys.
 *
 * @typedef {{
 *   id: string;
 *   secret: unknown;
 * }} CredentialEntry
 */

/**
 * @typedef {{
 *   get: (scope: CredentialScope) => Promise<CredentialEntry | null>;
 *   set: (scope: CredentialScope, secret: unknown) => Promise<CredentialEntry>;
 *   delete: (scope: CredentialScope) => Promise<void>;
 * }} CredentialStore
 */

/**
 * Stable, JSON-serializable key for a credential scope.
 *
 * Stores should use this when they need to index secrets by scope.
 *
 * @param {CredentialScope} scope
 * @returns {string}
 */
export function credentialScopeKey(scope) {
  // Using the engine's stable JSON canonicalization keeps the key deterministic
  // across environments (Node/browser) without introducing Node-only crypto
  // dependencies.
  return stableStringify(scope);
}

