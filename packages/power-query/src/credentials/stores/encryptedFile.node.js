import crypto from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";

import { KeyRing } from "../../../../security/crypto/keyring.js";
import {
  decodeEncryptedFileBytes,
  encodeEncryptedFileBytes,
  isEncryptedFileBytes,
} from "../../../../security/crypto/encryptedFile.js";

import { credentialScopeKey } from "../store.js";
import { randomId } from "../utils.js";

/**
 * @typedef {import("../store.js").CredentialEntry} CredentialEntry
 * @typedef {import("../scopes.js").CredentialScope} CredentialScope
 */

function defaultAadContext(filePath) {
  return { scope: "formula.power-query.credentials", schemaVersion: 1, filePath };
}

/**
 * Encrypted credential store persisted to a single on-disk file.
 *
 * The file contents are encrypted with a `KeyRing`. The key material is stored
 * in an OS keychain when possible; otherwise, it falls back to a local file
 * with restrictive permissions. This fallback is intended for tests and local
 * tooling environments where an OS keychain may not be available.
 */
export class EncryptedFileCredentialStore {
  /**
   * @param {{
   *   filePath: string;
   *   keychainProvider?: {
   *     getSecret: (opts: { service: string; account: string }) => Promise<Buffer | null>;
   *     setSecret: (opts: { service: string; account: string; secret: Buffer }) => Promise<void>;
   *     deleteSecret: (opts: { service: string; account: string }) => Promise<void>;
   *   } | null;
   *   keychainService?: string;
   *   keychainAccount?: string;
   *   keyringFilePath?: string;
   *   aadContext?: any;
   * }} opts
   */
  constructor(opts) {
    if (!opts?.filePath) throw new TypeError("filePath is required");
    this.filePath = opts.filePath;
    this.keychainProvider = opts.keychainProvider ?? null;
    this.keychainService = opts.keychainService ?? "formula.power-query";
    this.keychainAccount = opts.keychainAccount ?? this._defaultKeychainAccount();
    this.keyringFilePath = opts.keyringFilePath ?? `${this.filePath}.keyring.json`;
    this.aadContext = opts.aadContext ?? defaultAadContext(this.filePath);

    /** @type {KeyRing | null} */
    this._keyRing = null;
    /** @type {Map<string, CredentialEntry> | null} */
    this._entries = null;
    /** @type {Promise<void> | null} */
    this._initPromise = null;
    /** @type {Promise<void>} */
    this._persistChain = Promise.resolve();
  }

  _defaultKeychainAccount() {
    const hash = crypto.createHash("sha256").update(this.filePath).digest("hex").slice(0, 16);
    return `power-query-credentials:${hash}`;
  }

  async _ensureDir() {
    await fs.mkdir(path.dirname(this.filePath), { recursive: true });
  }

  async _loadKeyRingFromKeychain() {
    if (!this.keychainProvider) return null;
    try {
      const secret = await this.keychainProvider.getSecret({
        service: this.keychainService,
        account: this.keychainAccount,
      });
      if (!secret) return null;
      const parsed = JSON.parse(secret.toString("utf8"));
      return KeyRing.fromJSON(parsed);
    } catch {
      return null;
    }
  }

  async _storeKeyRingToKeychain(keyRing) {
    if (!this.keychainProvider) return false;
    try {
      const json = JSON.stringify(keyRing.toJSON());
      await this.keychainProvider.setSecret({
        service: this.keychainService,
        account: this.keychainAccount,
        secret: Buffer.from(json, "utf8"),
      });
      return true;
    } catch {
      return false;
    }
  }

  async _loadKeyRingFromFile() {
    try {
      const raw = await fs.readFile(this.keyringFilePath, "utf8");
      return KeyRing.fromJSON(JSON.parse(raw));
    } catch {
      return null;
    }
  }

  async _storeKeyRingToFile(keyRing) {
    await this._ensureDir();
    const json = JSON.stringify(keyRing.toJSON());
    const tmp = `${this.keyringFilePath}.tmp`;
    await fs.writeFile(tmp, json, { mode: 0o600 });
    await fs.rename(tmp, this.keyringFilePath);
  }

  async _ensureKeyRing() {
    if (this._keyRing) return this._keyRing;
    const fromKeychain = await this._loadKeyRingFromKeychain();
    if (fromKeychain) {
      this._keyRing = fromKeychain;
      return fromKeychain;
    }
    const fromFile = await this._loadKeyRingFromFile();
    if (fromFile) {
      this._keyRing = fromFile;
      // Best-effort: migrate to keychain if configured.
      await this._storeKeyRingToKeychain(fromFile);
      return fromFile;
    }

    const created = KeyRing.create();
    // Prefer keychain when available; otherwise fall back to local file.
    const storedInKeychain = await this._storeKeyRingToKeychain(created);
    if (!storedInKeychain) {
      await this._storeKeyRingToFile(created);
    }
    this._keyRing = created;
    return created;
  }

  async _loadEntries() {
    await this._ensureDir();
    /** @type {Buffer | null} */
    let bytes = null;
    try {
      bytes = await fs.readFile(this.filePath);
    } catch {
      bytes = null;
    }

    if (!bytes) {
      this._entries = new Map();
      return;
    }

    let jsonBytes = bytes;
    if (isEncryptedFileBytes(bytes)) {
      const ring = await this._ensureKeyRing();
      const decoded = decodeEncryptedFileBytes(bytes);
      jsonBytes = ring.decryptBytes(decoded, { aadContext: this.aadContext });
    }

    const parsed = JSON.parse(Buffer.from(jsonBytes).toString("utf8"));
    const entriesObj = parsed && typeof parsed === "object" ? parsed.entries : null;
    const entries = new Map();
    if (entriesObj && typeof entriesObj === "object") {
      for (const [k, v] of Object.entries(entriesObj)) {
        if (!v || typeof v !== "object") continue;
        const id = /** @type {any} */ (v).id;
        const secret = /** @type {any} */ (v).secret;
        if (typeof id !== "string" || id.length === 0) continue;
        entries.set(k, { id, secret });
      }
    }

    this._entries = entries;
  }

  async _open() {
    if (this._entries) return;
    if (this._initPromise) return this._initPromise;
    this._initPromise = this._loadEntries();
    await this._initPromise;
  }

  async _persist() {
    await this._open();
    if (!this._entries) return;
    await this._ensureDir();
    const ring = await this._ensureKeyRing();

    const payload = {
      version: 1,
      entries: Object.fromEntries(this._entries.entries()),
    };
    const plaintext = Buffer.from(JSON.stringify(payload), "utf8");
    const encrypted = ring.encryptBytes(plaintext, { aadContext: this.aadContext });
    const out = encodeEncryptedFileBytes({
      keyVersion: encrypted.keyVersion,
      iv: encrypted.iv,
      tag: encrypted.tag,
      ciphertext: encrypted.ciphertext,
    });

    const tmp = `${this.filePath}.tmp`;
    await fs.writeFile(tmp, out, { mode: 0o600 });
    await fs.rename(tmp, this.filePath);
  }

  async _queuePersist() {
    this._persistChain = this._persistChain.then(() => this._persist());
    return this._persistChain;
  }

  /**
   * @param {CredentialScope} scope
   * @returns {Promise<CredentialEntry | null>}
   */
  async get(scope) {
    await this._open();
    return this._entries?.get(credentialScopeKey(scope)) ?? null;
  }

  /**
   * @param {CredentialScope} scope
   * @param {unknown} secret
   * @returns {Promise<CredentialEntry>}
   */
  async set(scope, secret) {
    await this._open();
    if (!this._entries) throw new Error("store not initialized");
    const entry = { id: randomId(), secret };
    this._entries.set(credentialScopeKey(scope), entry);
    await this._queuePersist();
    return entry;
  }

  /**
   * @param {CredentialScope} scope
   * @returns {Promise<void>}
   */
  async delete(scope) {
    await this._open();
    if (!this._entries) return;
    this._entries.delete(credentialScopeKey(scope));
    await this._queuePersist();
  }
}

