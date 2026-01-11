import { hashValue } from "../../cache/key.js";
import { randomId } from "../utils.js";

/**
 * @typedef {import("../store.js").CredentialEntry} CredentialEntry
 * @typedef {import("../scopes.js").CredentialScope} CredentialScope
 */

/**
 * @typedef {{
 *   getSecret: (opts: { service: string; account: string }) => Promise<Buffer | null>;
 *   setSecret: (opts: { service: string; account: string; secret: Buffer }) => Promise<void>;
 *   deleteSecret: (opts: { service: string; account: string }) => Promise<void>;
 * }} KeychainProvider
 */

function assertJsonSerializable(value, name) {
  try {
    JSON.stringify(value);
  } catch (err) {
    throw new TypeError(`${name} must be JSON-serializable`);
  }
}

/**
 * Credential store backed by a KeychainProvider.
 *
 * This is the most secure option when an OS keychain is available.
 */
export class KeychainCredentialStore {
  /**
   * @param {{
   *   keychainProvider: KeychainProvider;
   *   service?: string;
   *   accountPrefix?: string;
   * }} opts
   */
  constructor(opts) {
    if (!opts?.keychainProvider) {
      throw new TypeError("keychainProvider is required");
    }
    this.keychainProvider = opts.keychainProvider;
    this.service = opts.service ?? "formula.power-query";
    this.accountPrefix = opts.accountPrefix ?? "pq-cred:";
  }

  /**
   * @param {CredentialScope} scope
   */
  _account(scope) {
    // Keep the keychain account identifier short and opaque to avoid leaking
    // scope details into OS keychain UIs and to avoid platform length limits.
    return `${this.accountPrefix}${scope.type}:${hashValue(scope)}`;
  }

  /**
   * @param {CredentialScope} scope
   * @returns {Promise<CredentialEntry | null>}
   */
  async get(scope) {
    const account = this._account(scope);
    const bytes = await this.keychainProvider.getSecret({ service: this.service, account });
    if (!bytes) return null;
    const parsed = JSON.parse(bytes.toString("utf8"));
    if (!parsed || typeof parsed !== "object") return null;
    const id = /** @type {any} */ (parsed).id;
    const secret = /** @type {any} */ (parsed).secret;
    if (typeof id !== "string" || id.length === 0) return null;
    return { id, secret };
  }

  /**
   * @param {CredentialScope} scope
   * @param {unknown} secret
   * @returns {Promise<CredentialEntry>}
   */
  async set(scope, secret) {
    assertJsonSerializable(secret, "secret");
    const entry = { id: randomId(), secret };
    const account = this._account(scope);
    const json = JSON.stringify(entry);
    await this.keychainProvider.setSecret({
      service: this.service,
      account,
      secret: Buffer.from(json, "utf8"),
    });
    return entry;
  }

  /**
   * @param {CredentialScope} scope
   * @returns {Promise<void>}
   */
  async delete(scope) {
    const account = this._account(scope);
    await this.keychainProvider.deleteSecret({ service: this.service, account });
  }
}
