import fs from "node:fs/promises";
import path from "node:path";

import semverPkg from "../../../../shared/semver.js";
import extensionPackagePkg from "../../../../shared/extension-package/index.js";
import signingPkg from "../../../../shared/crypto/signing.js";

const { compareSemver } = semverPkg;
const { extractExtensionPackage, readExtensionPackage } = extensionPackagePkg;
const { verifyBytesSignature } = signingPkg;

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

export class ExtensionManager {
  constructor({ marketplaceClient, extensionsDir, statePath }) {
    if (!marketplaceClient) throw new Error("marketplaceClient is required");
    if (!extensionsDir) throw new Error("extensionsDir is required");
    if (!statePath) throw new Error("statePath is required");

    this.marketplaceClient = marketplaceClient;
    this.extensionsDir = extensionsDir;
    this.statePath = statePath;
  }

  async _loadState() {
    return readJsonIfExists(this.statePath, { installed: {} });
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

  async install(id, version = null) {
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
    ensureSignaturePresent(download.signatureBase64);

    const signatureOk = verifyBytesSignature(download.bytes, download.signatureBase64, publicKeyPem);
    if (!signatureOk) throw new Error("Extension signature verification failed (mandatory)");

    // Basic cross-check: the manifest inside the package must match the requested extension id.
    const bundle = readExtensionPackage(download.bytes);
    const manifest = bundle.manifest;
    const bundleId = `${manifest.publisher}.${manifest.name}`;
    if (bundleId !== id) {
      throw new Error(`Package id mismatch: expected ${id} but got ${bundleId}`);
    }
    if (manifest.version !== resolvedVersion) {
      throw new Error(`Package version mismatch: expected ${resolvedVersion} but got ${manifest.version}`);
    }

    const installDir = path.join(this.extensionsDir, id);
    await fs.rm(installDir, { recursive: true, force: true });
    await extractExtensionPackage(download.bytes, installDir);

    const state = await this._loadState();
    state.installed[id] = {
      id,
      version: resolvedVersion,
      installedAt: new Date().toISOString(),
    };
    await this._saveState(state);

    return state.installed[id];
  }

  async uninstall(id) {
    const installDir = path.join(this.extensionsDir, id);
    await fs.rm(installDir, { recursive: true, force: true });

    const state = await this._loadState();
    delete state.installed[id];
    await this._saveState(state);
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

    return this.install(id, ext.latestVersion);
  }
}
