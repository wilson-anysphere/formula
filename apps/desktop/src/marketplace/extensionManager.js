import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import semverPkg from "../../../../shared/semver.js";
import extensionPackagePkg from "../../../../shared/extension-package/index.js";

const { compareSemver } = semverPkg;
const {
  detectExtensionPackageFormatVersion,
  verifyAndExtractExtensionPackage,
  verifyExtractedExtensionDir,
} = extensionPackagePkg;

async function readJsonIfExists(filePath, fallback) {
  try {
    const raw = await fs.readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error && (error.code === "ENOENT" || error.code === "ENOTDIR")) return fallback;
    if (error instanceof SyntaxError) return fallback;
    throw error;
  }
}

async function atomicWriteJson(filePath, data) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const tmp = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tmp, JSON.stringify(data, null, 2));
  try {
    await fs.rename(tmp, filePath);
  } catch (error) {
    if (error?.code === "EEXIST" || error?.code === "EPERM") {
      try {
        await fs.rm(filePath, { force: true });
        await fs.rename(tmp, filePath);
        return;
      } catch (renameError) {
        await fs.rm(tmp, { force: true }).catch(() => {});
        throw renameError;
      }
    }
    await fs.rm(tmp, { force: true }).catch(() => {});
    throw error;
  }
}

function ensureSignaturePresent(signatureBase64) {
  if (!signatureBase64 || typeof signatureBase64 !== "string") {
    throw new Error("Marketplace download missing package signature (mandatory)");
  }
}

function sha256Hex(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

export class ExtensionManager {
  constructor({ marketplaceClient, extensionsDir, statePath }) {
    if (!marketplaceClient) throw new Error("marketplaceClient is required");
    if (!extensionsDir) throw new Error("extensionsDir is required");
    if (!statePath) throw new Error("statePath is required");

    this.marketplaceClient = marketplaceClient;
    this.extensionsDir = extensionsDir;
    this.statePath = statePath;

    /** @type {Set<(event: any) => void>} */
    this._listeners = new Set();
  }

  /**
   * Subscribe to install/update/uninstall events.
   *
   * @param {(event: { action: "install" | "update" | "uninstall", id: string, record?: any }) => void} listener
   * @returns {{ dispose: () => void }}
   */
  onDidChange(listener) {
    if (typeof listener !== "function") {
      throw new Error("onDidChange listener must be a function");
    }
    this._listeners.add(listener);
    return {
      dispose: () => {
        this._listeners.delete(listener);
      },
    };
  }

  _emit(event) {
    for (const listener of [...this._listeners]) {
      try {
        listener(event);
      } catch {
        // ignore
      }
    }
  }

  async _loadState() {
    const state = await readJsonIfExists(this.statePath, { installed: {} });
    if (!state || typeof state !== "object") return { installed: {} };
    if (!state.installed || typeof state.installed !== "object") state.installed = {};
    return state;
  }

  async _saveState(state) {
    await atomicWriteJson(this.statePath, state);
  }

  async listInstalled() {
    const state = await this._loadState();
    return Object.values(state.installed);
  }

  async getInstalled(id) {
    const state = await this._loadState();
    return state.installed[id] || null;
  }

  async _markCorrupted(state, id, reason) {
    const installed = state.installed[id];
    if (!installed) return;
    installed.corrupted = true;
    installed.corruptedAt = new Date().toISOString();
    installed.corruptedReason = reason;
    await this._saveState(state);
  }

  async _installInternal(id, version = null) {
    const ext = await this.marketplaceClient.getExtension(id);
    if (!ext) throw new Error(`Extension not found: ${id}`);

    const resolvedVersion = version || ext.latestVersion;
    if (!resolvedVersion) throw new Error("Marketplace did not provide latestVersion");

    const rawPublisherKeys = Array.isArray(ext.publisherKeys) ? ext.publisherKeys : [];
    const hasPublisherKeySet = rawPublisherKeys.length > 0;

    /** @type {{ id?: string | null, publicKeyPem: string }[]} */
    let candidateKeys = [];
    if (hasPublisherKeySet) {
      candidateKeys = rawPublisherKeys
        .filter((k) => k && typeof k.publicKeyPem === "string" && typeof k.id === "string")
        .filter((k) => !k.revoked)
        .map((k) => ({ id: k.id, publicKeyPem: k.publicKeyPem }));

      if (candidateKeys.length === 0) {
        throw new Error("All publisher signing keys are revoked (refusing to install)");
      }
    } else {
      const publicKeyPem = ext.publisherPublicKeyPem;
      if (!publicKeyPem) {
        throw new Error("Marketplace did not provide publisher public key (mandatory)");
      }
      candidateKeys = [{ id: null, publicKeyPem }];
    }

    const download = await this.marketplaceClient.downloadPackage(id, resolvedVersion);
    if (!download) throw new Error(`Package not found: ${id}@${resolvedVersion}`);

    const computedPackageSha256 = sha256Hex(download.bytes);
    if (download.sha256 && download.sha256 !== computedPackageSha256) {
      // Treat a sha256 mismatch as a signature verification failure: both indicate the downloaded
      // package bytes cannot be trusted (tampering/corruption).
      throw new Error("Extension signature verification failed (sha256 mismatch)");
    }

    if (hasPublisherKeySet && download.publisherKeyId) {
      const keyId = String(download.publisherKeyId);
      const preferred = candidateKeys.find((k) => k.id === keyId);
      if (preferred) {
        candidateKeys = [preferred, ...candidateKeys.filter((k) => k.id !== keyId)];
      }
    }

    const formatVersion =
      typeof download.formatVersion === "number" && Number.isFinite(download.formatVersion)
        ? download.formatVersion
        : detectExtensionPackageFormatVersion(download.bytes);

    const installDir = path.join(this.extensionsDir, id);
    if (formatVersion === 1) {
      ensureSignaturePresent(download.signatureBase64);
    }

    let verified = null;
    /** @type {any} */
    let lastSignatureError = null;
    for (const key of candidateKeys) {
      try {
        verified = await verifyAndExtractExtensionPackage(download.bytes, installDir, {
          publicKeyPem: key.publicKeyPem,
          signatureBase64: download.signatureBase64,
          formatVersion,
          expectedId: id,
          expectedVersion: resolvedVersion,
        });
        break;
      } catch (error) {
        const message = String(error?.message ?? error);
        if (message.toLowerCase().includes("signature verification failed")) {
          lastSignatureError = error;
          continue;
        }
        throw error;
      }
    }

    if (!verified) {
      throw lastSignatureError || new Error("Extension signature verification failed (mandatory)");
    }

    try {
      const expectedFiles = Array.isArray(verified.files) ? verified.files : [];
      const signatureBase64 = verified.signatureBase64 ?? null;
      const verifiedFormatVersion =
        typeof verified.formatVersion === "number" && Number.isFinite(verified.formatVersion)
          ? verified.formatVersion
          : formatVersion;

      let expectedFilesSha256 = null;
      if (expectedFiles.length > 0) {
        expectedFilesSha256 = sha256Hex(Buffer.from(JSON.stringify(expectedFiles), "utf8"));
        const headerFilesSha256 =
          download.filesSha256 && typeof download.filesSha256 === "string"
            ? download.filesSha256.trim().toLowerCase()
            : null;
        if (headerFilesSha256 && /^[0-9a-f]{64}$/.test(headerFilesSha256) && headerFilesSha256 !== expectedFilesSha256) {
          // Defense-in-depth: the package is already signature-verified, but a mismatch here indicates
          // the marketplace's recorded file inventory doesn't match the signed payload we extracted.
          throw new Error("Extension signature verification failed (files sha256 mismatch)");
        }
      }

      const scanStatus = download.scanStatus && typeof download.scanStatus === "string" ? download.scanStatus : null;

      const state = await this._loadState();
      state.installed[id] = {
        id,
        version: resolvedVersion,
        installedAt: new Date().toISOString(),
        formatVersion: verifiedFormatVersion,
        packageSha256: computedPackageSha256,
        signatureBase64,
        scanStatus,
        filesSha256: expectedFilesSha256,
        files: expectedFiles,
      };
      await this._saveState(state);

      return state.installed[id];
    } catch (error) {
      // verifyAndExtractExtensionPackage has already written the extracted extension to disk. If we fail
      // after that (e.g. provenance mismatch), ensure we don't leave behind an untracked directory.
      await fs.rm(installDir, { recursive: true, force: true }).catch(() => {});
      throw error;
    }
  }

  async install(id, version = null) {
    const record = await this._installInternal(id, version);
    this._emit({ action: "install", id, record });
    return record;
  }

  async uninstall(id) {
    const installDir = path.join(this.extensionsDir, id);
    await fs.rm(installDir, { recursive: true, force: true });

    const state = await this._loadState();
    delete state.installed[id];
    await this._saveState(state);

    this._emit({ action: "uninstall", id });
  }

  async checkForUpdates() {
    const installed = await this.listInstalled();
    const updates = [];

    for (const item of installed) {
      const ext = await this.marketplaceClient.getExtension(item.id);
      if (!ext || !ext.latestVersion) continue;
      if (compareSemver(ext.latestVersion, item.version) > 0) {
        updates.push({
          id: item.id,
          currentVersion: item.version,
          latestVersion: ext.latestVersion,
        });
      }
    }

    return updates;
  }

  async update(id) {
    const installed = await this.getInstalled(id);
    if (!installed) throw new Error(`Not installed: ${id}`);

    const ext = await this.marketplaceClient.getExtension(id);
    if (!ext || !ext.latestVersion) throw new Error(`Marketplace missing latestVersion for ${id}`);
    if (compareSemver(ext.latestVersion, installed.version) <= 0) {
      return installed;
    }

    const record = await this._installInternal(id, ext.latestVersion);
    this._emit({ action: "update", id, record });
    return record;
  }

  async verifyInstalled(id) {
    const state = await this._loadState();
    const installed = state.installed[id];
    if (!installed) throw new Error(`Not installed: ${id}`);

    if (installed.corrupted) {
      return { ok: false, reason: installed.corruptedReason || "Extension is quarantined" };
    }

    if (!Array.isArray(installed.files) || installed.files.length === 0) {
      const reason =
        "Missing integrity metadata (installed with an older version). Repair (reinstall) is required.";
      await this._markCorrupted(state, id, reason);
      return { ok: false, reason };
    }

    const installDir = path.join(this.extensionsDir, id);
    let result;
    try {
      result = await verifyExtractedExtensionDir(installDir, installed.files, {
        ignoreExtraPaths: [".DS_Store", "Thumbs.db", "desktop.ini"],
      });
    } catch (error) {
      result = { ok: false, reason: error?.message ?? String(error) };
    }

    if (!result.ok) {
      await this._markCorrupted(state, id, result.reason || "Extension integrity check failed");
    }

    return result;
  }

  async verifyAllInstalled() {
    const installed = await this.listInstalled();
    /** @type {Record<string, { ok: boolean, reason?: string }>} */
    const out = {};
    for (const item of installed) {
      try {
        out[item.id] = await this.verifyInstalled(item.id);
      } catch (error) {
        out[item.id] = { ok: false, reason: error?.message ?? String(error) };
      }
    }
    return out;
  }

  async loadIntoHost(extensionHost, id) {
    if (!extensionHost || typeof extensionHost.loadExtension !== "function") {
      throw new Error("extensionHost with loadExtension() is required");
    }

    const verification = await this.verifyInstalled(id);
    if (!verification.ok) {
      const reason = verification.reason || "unknown reason";
      throw new Error(
        `Extension integrity check failed for ${id}: ${reason}. Run repair() to reinstall the extension.`
      );
    }

    const installDir = path.join(this.extensionsDir, id);
    return extensionHost.loadExtension(installDir);
  }

  async repair(id) {
    const installed = await this.getInstalled(id);
    if (!installed) throw new Error(`Not installed: ${id}`);

    try {
      return await this.install(id, installed.version);
    } catch (error) {
      const msg = error?.message ?? String(error);
      if (!/Package not found/i.test(msg)) throw error;
      return this.install(id);
    }
  }
}
