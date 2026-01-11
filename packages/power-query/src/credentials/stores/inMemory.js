import { credentialScopeKey } from "../store.js";
import { randomId } from "../utils.js";

/**
 * @typedef {import("../store.js").CredentialEntry} CredentialEntry
 * @typedef {import("../scopes.js").CredentialScope} CredentialScope
 */

export class InMemoryCredentialStore {
  constructor() {
    /** @type {Map<string, CredentialEntry>} */
    this._entries = new Map();
  }

  /**
   * @param {CredentialScope} scope
   * @returns {Promise<CredentialEntry | null>}
   */
  async get(scope) {
    return this._entries.get(credentialScopeKey(scope)) ?? null;
  }

  /**
   * @param {CredentialScope} scope
   * @param {unknown} secret
   * @returns {Promise<CredentialEntry>}
   */
  async set(scope, secret) {
    const entry = { id: randomId(), secret };
    this._entries.set(credentialScopeKey(scope), entry);
    return entry;
  }

  /**
   * @param {CredentialScope} scope
   * @returns {Promise<void>}
   */
  async delete(scope) {
    this._entries.delete(credentialScopeKey(scope));
  }
}

