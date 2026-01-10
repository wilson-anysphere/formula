import { promises as fs } from "node:fs";
import path from "node:path";

import { KeyRing } from "../../../../../packages/security/crypto/keyring.js";

const DEFAULT_AAD_CONTEXT = { scope: "formula-desktop-store", schemaVersion: 1 };

export class DesktopEncryptedDocumentStore {
  constructor({
    filePath,
    keychainProvider,
    keychainService = "formula.desktop",
    keychainAccount = "storage-keyring"
  }) {
    if (typeof filePath !== "string" || filePath.length === 0) {
      throw new TypeError("filePath must be a non-empty string");
    }
    if (!keychainProvider) {
      throw new TypeError("keychainProvider is required");
    }

    this.filePath = filePath;
    this.keychainProvider = keychainProvider;
    this.keychainService = keychainService;
    this.keychainAccount = keychainAccount;
  }

  async _loadKeyRing() {
    const secret = await this.keychainProvider.getSecret({
      service: this.keychainService,
      account: this.keychainAccount
    });
    if (!secret) return null;
    const parsed = JSON.parse(secret.toString("utf8"));
    return KeyRing.fromJSON(parsed);
  }

  async _storeKeyRing(keyRing) {
    const json = JSON.stringify(keyRing.toJSON());
    await this.keychainProvider.setSecret({
      service: this.keychainService,
      account: this.keychainAccount,
      secret: Buffer.from(json, "utf8")
    });
  }

  async _deleteKeyRing() {
    await this.keychainProvider.deleteSecret({
      service: this.keychainService,
      account: this.keychainAccount
    });
  }

  async _readOnDisk() {
    try {
      const raw = await fs.readFile(this.filePath, "utf8");
      return JSON.parse(raw);
    } catch (error) {
      if (error && error.code === "ENOENT") {
        return {
          schemaVersion: 1,
          encrypted: false,
          documents: {}
        };
      }
      throw error;
    }
  }

  async _writeOnDisk(data) {
    await fs.mkdir(path.dirname(this.filePath), { recursive: true });
    await fs.writeFile(this.filePath, JSON.stringify(data, null, 2), "utf8");
  }

  async _loadPlaintextDocuments() {
    const onDisk = await this._readOnDisk();
    if (!onDisk.encrypted) {
      return onDisk.documents ?? {};
    }

    const keyRing = await this._loadKeyRing();
    if (!keyRing) {
      throw new Error("Encrypted store present but no keyring in keychain");
    }

    const plaintextBytes = keyRing.decrypt(onDisk, { aadContext: DEFAULT_AAD_CONTEXT });
    const parsed = JSON.parse(plaintextBytes.toString("utf8"));
    return parsed.documents ?? {};
  }

  async _writeDocuments({ documents, encrypted }) {
    if (!encrypted) {
      await this._writeOnDisk({
        schemaVersion: 1,
        encrypted: false,
        documents
      });
      return;
    }

    let keyRing = await this._loadKeyRing();
    if (!keyRing) {
      keyRing = KeyRing.create();
      await this._storeKeyRing(keyRing);
    }

    const plaintextBytes = Buffer.from(
      JSON.stringify({ schemaVersion: 1, encrypted: false, documents }),
      "utf8"
    );

    const encryptedPayload = keyRing.encrypt(plaintextBytes, {
      aadContext: DEFAULT_AAD_CONTEXT
    });

    await this._writeOnDisk({
      schemaVersion: 1,
      encrypted: true,
      ...encryptedPayload
    });
  }

  async enableEncryption() {
    const documents = await this._loadPlaintextDocuments();
    await this._writeDocuments({ documents, encrypted: true });
  }

  async disableEncryption({ deleteKey = true } = {}) {
    const documents = await this._loadPlaintextDocuments();
    await this._writeDocuments({ documents, encrypted: false });
    if (deleteKey) {
      await this._deleteKeyRing();
    }
  }

  async rotateKey() {
    const onDisk = await this._readOnDisk();
    if (!onDisk.encrypted) {
      throw new Error("Cannot rotate key: store is not encrypted");
    }

    const documents = await this._loadPlaintextDocuments();

    const keyRing = await this._loadKeyRing();
    if (!keyRing) {
      throw new Error("Cannot rotate key: missing keyring");
    }

    keyRing.rotate();
    await this._storeKeyRing(keyRing);
    await this._writeDocuments({ documents, encrypted: true });
  }

  async saveDocument(docId, document) {
    if (typeof docId !== "string" || docId.length === 0) {
      throw new TypeError("docId must be a non-empty string");
    }

    const onDisk = await this._readOnDisk();
    const documents = await this._loadPlaintextDocuments();
    documents[docId] = document;
    await this._writeDocuments({ documents, encrypted: Boolean(onDisk.encrypted) });
  }

  async loadDocument(docId) {
    const documents = await this._loadPlaintextDocuments();
    return documents[docId] ?? null;
  }
}

