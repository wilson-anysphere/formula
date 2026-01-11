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
    throw error;
  }
}

async function atomicWriteJson(filePath, data) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const tmp = `${filePath}.tmp`;
  await fs.writeFile(tmp, JSON.stringify(data, null, 2));
  await fs.rename(tmp, filePath);
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

    const publicKeyPem = ext.publisherPublicKeyPem;
    if (!publicKeyPem) {
      throw new Error("Marketplace did not provide publisher public key (mandatory)");
    }

    const download = await this.marketplaceClient.downloadPackage(id, resolvedVersion);
    if (!download) throw new Error(`Package not found: ${id}@${resolvedVersion}`);

    const computedPackageSha256 = sha256Hex(download.bytes);
    if (download.sha256 && download.sha256 !== computedPackageSha256) {
      throw new Error("Marketplace download sha256 does not match downloaded bytes");
    }

    const formatVersion =
      typeof download.formatVersion === "number" && Number.isFinite(download.formatVersion)
        ? download.formatVersion
        : detectExtensionPackageFormatVersion(download.bytes);

    const installDir = path.join(this.extensionsDir, id);
    if (formatVersion === 1) {
      ensureSignaturePresent(download.signatureBase64);
    }

    const verified = await verifyAndExtractExtensionPackage(download.bytes, installDir, {
      publicKeyPem,
      signatureBase64: download.signatureBase64,
      formatVersion,
      expectedId: id,
      expectedVersion: resolvedVersion,
    });
    const expectedFiles = Array.isArray(verified.files) ? verified.files : [];
    const signatureBase64 = verified.signatureBase64 ?? null;
    const verifiedFormatVersion =
      typeof verified.formatVersion === "number" && Number.isFinite(verified.formatVersion)
        ? verified.formatVersion
        : formatVersion;

    const state = await this._loadState();
    state.installed[id] = {
      id,
      version: resolvedVersion,
      installedAt: new Date().toISOString(),
      formatVersion: verifiedFormatVersion,
      packageSha256: computedPackageSha256,
      signatureBase64,
      files: expectedFiles,
    };
    await this._saveState(state);

    return state.installed[id];
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
