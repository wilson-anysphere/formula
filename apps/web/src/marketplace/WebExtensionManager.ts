import extensionApiSource from "@formula/extension-api?raw";

import type {
  ReadExtensionPackageV2Result,
  VerifiedExtensionPackageV2
} from "../../../../shared/extension-package/v2-browser.mjs";
import {
  readExtensionPackageV2,
  verifyExtensionPackageV2Browser
} from "../../../../shared/extension-package/v2-browser.mjs";

import { MarketplaceClient } from "./MarketplaceClient";

export interface InstalledExtensionRecord {
  id: string;
  version: string;
  installedAt: string;
}

export interface BrowserExtensionHostLike {
  loadExtension(args: {
    extensionId: string;
    extensionPath: string;
    manifest: Record<string, any>;
    mainUrl: string;
  }): Promise<string>;
  unloadExtension?(extensionId: string): Promise<void> | void;
  listExtensions(): Array<{ id: string }>;
}

export interface InstalledExtensionWithManifest extends InstalledExtensionRecord {
  manifest: Record<string, any>;
  verified: VerifiedExtensionPackageV2;
}

interface StoredPackageRecord {
  key: string; // `${id}@${version}`
  id: string;
  version: string;
  bytes: ArrayBuffer;
  verified: VerifiedExtensionPackageV2;
}

const DB_NAME = "formula.webExtensions";
const DB_VERSION = 1;

const STORE_INSTALLED = "installed";
const STORE_PACKAGES = "packages";

function compareSemver(a: string, b: string): number {
  // Minimal semver compare (major.minor.patch[-prerelease]) to avoid pulling in a dependency.
  const semverRe =
    /^(?<major>0|[1-9]\d*)\.(?<minor>0|[1-9]\d*)\.(?<patch>0|[1-9]\d*)(?:-(?<prerelease>[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+(?<build>[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?$/;

  const ma = semverRe.exec(a);
  const mb = semverRe.exec(b);
  if (!ma?.groups || !mb?.groups) {
    throw new Error(`Invalid semver compare: "${a}" vs "${b}"`);
  }
  const pa = {
    major: Number(ma.groups.major),
    minor: Number(ma.groups.minor),
    patch: Number(ma.groups.patch),
    prerelease: ma.groups.prerelease ? ma.groups.prerelease.split(".") : null
  };
  const pb = {
    major: Number(mb.groups.major),
    minor: Number(mb.groups.minor),
    patch: Number(mb.groups.patch),
    prerelease: mb.groups.prerelease ? mb.groups.prerelease.split(".") : null
  };

  if (pa.major !== pb.major) return pa.major < pb.major ? -1 : 1;
  if (pa.minor !== pb.minor) return pa.minor < pb.minor ? -1 : 1;
  if (pa.patch !== pb.patch) return pa.patch < pb.patch ? -1 : 1;

  const aPre = pa.prerelease;
  const bPre = pb.prerelease;
  if (!aPre && !bPre) return 0;
  if (!aPre) return 1;
  if (!bPre) return -1;

  const idRe = /^[0-9]+$/;
  const max = Math.max(aPre.length, bPre.length);
  for (let i = 0; i < max; i++) {
    const ai = aPre[i];
    const bi = bPre[i];
    if (ai === undefined) return -1;
    if (bi === undefined) return 1;
    const aiNum = idRe.test(ai);
    const biNum = idRe.test(bi);
    if (aiNum && biNum) {
      const av = Number(ai);
      const bv = Number(bi);
      if (av !== bv) return av < bv ? -1 : 1;
      continue;
    }
    if (aiNum !== biNum) return aiNum ? -1 : 1;
    if (ai !== bi) return ai < bi ? -1 : 1;
  }

  return 0;
}

function detectExtensionPackageFormatVersion(packageBytes: Uint8Array): number {
  if (packageBytes.length >= 2 && packageBytes[0] === 0x1f && packageBytes[1] === 0x8b) return 1;
  return 2;
}

function requestToPromise<T = unknown>(req: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error("IndexedDB request failed"));
  });
}

function openDb(): Promise<IDBDatabase> {
  if (typeof indexedDB === "undefined") {
    throw new Error("IndexedDB is required to install extensions in the browser");
  }

  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onerror = () => reject(req.error ?? new Error("Failed to open IndexedDB"));
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_INSTALLED)) {
        db.createObjectStore(STORE_INSTALLED, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(STORE_PACKAGES)) {
        const store = db.createObjectStore(STORE_PACKAGES, { keyPath: "key" });
        store.createIndex("byId", "id", { unique: false });
      }
    };
    req.onsuccess = () => resolve(req.result);
  });
}

function txDone(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
    tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
  });
}

function bytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }
  if (typeof btoa === "function") {
    let bin = "";
    for (const b of bytes) bin += String.fromCharCode(b);
    return btoa(bin);
  }
  throw new Error("Base64 encoding is not available in this runtime");
}

function bytesToDataUrl(bytes: Uint8Array, mime: string): string {
  return `data:${mime};base64,${bytesToBase64(bytes)}`;
}

function createModuleUrl(bytes: Uint8Array, mime = "text/javascript"): { url: string; revoke: () => void } {
  const isNodeRuntime =
    typeof process !== "undefined" && typeof (process as any)?.versions?.node === "string";

  if (
    !isNodeRuntime &&
    typeof URL !== "undefined" &&
    typeof URL.createObjectURL === "function" &&
    typeof Blob !== "undefined"
  ) {
    const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
    return { url, revoke: () => URL.revokeObjectURL(url) };
  }
  const url = bytesToDataUrl(bytes, mime);
  return { url, revoke: () => {} };
}

function createModuleUrlFromText(source: string): { url: string; revoke: () => void } {
  const bytes = new TextEncoder().encode(source);
  return createModuleUrl(bytes);
}

function extractEntrypointPath(manifest: Record<string, any>): string {
  const entry = manifest.browser ?? manifest.module ?? manifest.main;
  if (typeof entry !== "string" || entry.trim().length === 0) {
    throw new Error("Extension manifest missing entrypoint (browser/module/main)");
  }

  let rel = entry.trim().replace(/\\/g, "/");
  while (rel.startsWith("./")) rel = rel.slice(2);
  const parts = rel.split("/");
  if (parts.some((p) => p === "" || p === "." || p === "..")) {
    throw new Error(`Invalid extension entrypoint path: ${entry}`);
  }
  return parts.join("/");
}

function rewriteEntrypointSource(source: string, { extensionApiUrl }: { extensionApiUrl: string }): string {
  const rewritten = source
    .replace(/from\s+["']@formula\/extension-api["']/g, `from "${extensionApiUrl}"`)
    .replace(/import\s+["']@formula\/extension-api["']/g, `import "${extensionApiUrl}"`);

  const specifiers = new Set<string>();
  const importRe = /\bimport\s+(?:[^"']*?\s+from\s+)?["']([^"']+)["']/g;
  const exportRe = /\bexport\s+(?:\*|\{[^}]*\})\s+from\s+["']([^"']+)["']/g;
  for (const re of [importRe, exportRe]) {
    for (;;) {
      const match = re.exec(rewritten);
      if (!match) break;
      specifiers.add(match[1]);
    }
  }

  for (const specifier of specifiers) {
    // The loader only supports importing verified code. The only allowed imports are other in-memory
    // modules (blob/data URLs). Anything else would require fetching unverified code.
    if (!specifier.startsWith("blob:") && !specifier.startsWith("data:")) {
      throw new Error(
        `Unsupported import specifier "${specifier}". Browser extensions must be bundled as a single-file entrypoint.`
      );
    }
  }

  return rewritten;
}

export class WebExtensionManager {
  readonly marketplaceClient: MarketplaceClient;
  readonly host: BrowserExtensionHostLike | null;

  private readonly _loadedMainUrls = new Map<string, { mainUrl: string; revoke: () => void }>();
  private _extensionApiModule: { url: string; revoke: () => void } | null = null;

  constructor(options: { marketplaceClient?: MarketplaceClient; host?: BrowserExtensionHostLike | null } = {}) {
    this.marketplaceClient = options.marketplaceClient ?? new MarketplaceClient();
    this.host = options.host ?? null;
  }

  search(params: Parameters<MarketplaceClient["search"]>[0]) {
    return this.marketplaceClient.search(params);
  }

  getExtension(id: string) {
    return this.marketplaceClient.getExtension(id);
  }

  async listInstalled(): Promise<InstalledExtensionRecord[]> {
    const db = await openDb();
    try {
      const tx = db.transaction([STORE_INSTALLED], "readonly");
      const store = tx.objectStore(STORE_INSTALLED);
      const records = (await requestToPromise(store.getAll())) as InstalledExtensionRecord[];
      await txDone(tx);
      return records;
    } finally {
      db.close();
    }
  }

  async getInstalled(id: string): Promise<InstalledExtensionRecord | null> {
    const db = await openDb();
    try {
      const tx = db.transaction([STORE_INSTALLED], "readonly");
      const store = tx.objectStore(STORE_INSTALLED);
      const record = (await requestToPromise(store.get(String(id)))) as InstalledExtensionRecord | undefined;
      await txDone(tx);
      return record ?? null;
    } finally {
      db.close();
    }
  }

  async getInstalledWithManifest(id: string): Promise<InstalledExtensionWithManifest | null> {
    const installed = await this.getInstalled(id);
    if (!installed) return null;
    const pkg = await this._getPackage(installed.id, installed.version);
    if (!pkg) return null;
    return {
      ...installed,
      manifest: pkg.verified.manifest,
      verified: pkg.verified
    };
  }

  async install(id: string, version: string | null = null): Promise<InstalledExtensionRecord> {
    const ext = await this.marketplaceClient.getExtension(id);
    if (!ext) throw new Error(`Extension not found: ${id}`);

    const resolvedVersion = version ?? ext.latestVersion;
    if (!resolvedVersion) {
      throw new Error("Marketplace did not provide latestVersion");
    }

    const rawPublisherKeys = Array.isArray(ext.publisherKeys) ? ext.publisherKeys : [];
    const hasPublisherKeySet = rawPublisherKeys.length > 0;

    let candidateKeys: Array<{ id: string | null; publicKeyPem: string }> = [];
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
    if (formatVersion !== 2) {
      throw new Error(`Unsupported extension package formatVersion: ${formatVersion}`);
    }

    let verified: VerifiedExtensionPackageV2 | null = null;
    let lastSignatureError: unknown = null;
    for (const key of candidateKeys) {
      try {
        // eslint-disable-next-line no-await-in-loop
        verified = await verifyExtensionPackageV2Browser(download.bytes, key.publicKeyPem);
        break;
      } catch (error) {
        const msg = String((error as Error)?.message ?? error);
        if (msg.toLowerCase().includes("signature verification failed")) {
          lastSignatureError = error;
          continue;
        }
        throw new Error(`Extension signature verification failed (mandatory): ${msg}`);
      }
    }

    if (!verified) {
      throw new Error(
        `Extension signature verification failed (mandatory): ${String(
          (lastSignatureError as Error)?.message ?? lastSignatureError ?? "unknown error"
        )}`
      );
    }

    const manifest = verified.manifest;
    const bundleId = `${manifest.publisher}.${manifest.name}`;
    if (bundleId !== id) {
      throw new Error(`Package id mismatch: expected ${id} but got ${bundleId}`);
    }
    if (manifest.version !== resolvedVersion) {
      throw new Error(`Package version mismatch: expected ${resolvedVersion} but got ${manifest.version}`);
    }

    if (download.signatureBase64 && download.signatureBase64 !== verified.signatureBase64) {
      throw new Error("Marketplace signature header does not match package signature");
    }

    const installedAt = new Date().toISOString();
    const db = await openDb();
    try {
      const tx = db.transaction([STORE_INSTALLED, STORE_PACKAGES], "readwrite");
      const installedStore = tx.objectStore(STORE_INSTALLED);
      const packagesStore = tx.objectStore(STORE_PACKAGES);

      const prev = (await requestToPromise(installedStore.get(String(id)))) as InstalledExtensionRecord | undefined;
      const key = `${id}@${resolvedVersion}`;

      const pkgRecord: StoredPackageRecord = {
        key,
        id,
        version: resolvedVersion,
        bytes: download.bytes.buffer.slice(download.bytes.byteOffset, download.bytes.byteOffset + download.bytes.byteLength),
        verified
      };

      packagesStore.put(pkgRecord);
      installedStore.put({ id, version: resolvedVersion, installedAt });

      if (prev && prev.version && prev.version !== resolvedVersion) {
        packagesStore.delete(`${id}@${prev.version}`);
      }

      await txDone(tx);
    } finally {
      db.close();
    }

    return { id, version: resolvedVersion, installedAt };
  }

  async uninstall(id: string): Promise<void> {
    const existing = await this.getInstalled(id);
    if (!existing) return;

    await this.unload(id).catch(() => {});

    const db = await openDb();
    try {
      const tx = db.transaction([STORE_INSTALLED, STORE_PACKAGES], "readwrite");
      tx.objectStore(STORE_INSTALLED).delete(String(id));
      tx.objectStore(STORE_PACKAGES).delete(`${id}@${existing.version}`);
      await txDone(tx);
    } finally {
      db.close();
    }
  }

  async checkForUpdates(): Promise<Array<{ id: string; currentVersion: string; latestVersion: string }>> {
    const installed = await this.listInstalled();
    const updates: Array<{ id: string; currentVersion: string; latestVersion: string }> = [];

    for (const item of installed) {
      // eslint-disable-next-line no-await-in-loop
      const ext = await this.marketplaceClient.getExtension(item.id);
      if (!ext?.latestVersion) continue;
      if (compareSemver(ext.latestVersion, item.version) > 0) {
        updates.push({ id: item.id, currentVersion: item.version, latestVersion: ext.latestVersion });
      }
    }

    return updates;
  }

  async update(id: string): Promise<InstalledExtensionRecord> {
    const installed = await this.getInstalled(id);
    if (!installed) throw new Error(`Not installed: ${id}`);

    const ext = await this.marketplaceClient.getExtension(id);
    if (!ext?.latestVersion) throw new Error(`Marketplace missing latestVersion for ${id}`);
    if (compareSemver(ext.latestVersion, installed.version) <= 0) {
      return installed;
    }

    const next = await this.install(id, ext.latestVersion);
    if (this.isLoaded(id)) {
      await this.unload(id);
      await this.loadInstalled(id);
    }
    return next;
  }

  isLoaded(id: string): boolean {
    return this._loadedMainUrls.has(String(id));
  }

  async loadInstalled(id: string): Promise<string> {
    if (!this.host) throw new Error("WebExtensionManager requires a BrowserExtensionHost to load extensions");
    const installed = await this.getInstalled(id);
    if (!installed) throw new Error(`Not installed: ${id}`);

    const pkg = await this._getPackage(installed.id, installed.version);
    if (!pkg) {
      throw new Error(`Missing stored package for ${installed.id}@${installed.version}`);
    }

    if (!this._extensionApiModule) {
      this._extensionApiModule = createModuleUrlFromText(extensionApiSource);
    }

    const { mainUrl, revoke } = await this._createMainModuleUrl(pkg, this._extensionApiModule.url);

    const extensionPath = `indexeddb://formula/extensions/${installed.id}/${installed.version}`;
    const manifest = pkg.verified.manifest;

    // If already loaded, the host will throw; ensure the manager keeps a single source of truth.
    if (this.host.listExtensions().some((e) => e.id === installed.id)) {
      throw new Error(`Extension already loaded: ${installed.id}`);
    }

    await this.host.loadExtension({
      extensionId: installed.id,
      extensionPath,
      manifest,
      mainUrl
    });

    this._loadedMainUrls.set(installed.id, { mainUrl, revoke });
    return installed.id;
  }

  async unload(id: string): Promise<void> {
    if (!this.host) return;
    const existing = this._loadedMainUrls.get(String(id));
    if (existing) {
      this._loadedMainUrls.delete(String(id));
      try {
        existing.revoke();
      } catch {
        // ignore
      }
    }

    await this.host.unloadExtension?.(String(id));
  }

  async dispose(): Promise<void> {
    for (const id of [...this._loadedMainUrls.keys()]) {
      // eslint-disable-next-line no-await-in-loop
      await this.unload(id);
    }
    if (this._extensionApiModule) {
      try {
        this._extensionApiModule.revoke();
      } catch {
        // ignore
      }
      this._extensionApiModule = null;
    }
  }

  private async _getPackage(id: string, version: string): Promise<StoredPackageRecord | null> {
    const db = await openDb();
    try {
      const tx = db.transaction([STORE_PACKAGES], "readonly");
      const store = tx.objectStore(STORE_PACKAGES);
      const key = `${id}@${version}`;
      const record = (await requestToPromise(store.get(key))) as StoredPackageRecord | undefined;
      await txDone(tx);
      return record ?? null;
    } finally {
      db.close();
    }
  }

  private async _createMainModuleUrl(
    pkg: StoredPackageRecord,
    extensionApiUrl: string
  ): Promise<{ mainUrl: string; revoke: () => void }> {
    const bytes = new Uint8Array(pkg.bytes);
    const parsed: ReadExtensionPackageV2Result = readExtensionPackageV2(bytes);

    const entryRel = extractEntrypointPath(pkg.verified.manifest);
    const entryBytes = parsed.files.get(entryRel);
    if (!entryBytes) {
      throw new Error(`Extension entrypoint missing from package: ${entryRel}`);
    }

    const source = decodeUtf8FromBytes(entryBytes);
    const rewritten = rewriteEntrypointSource(source, { extensionApiUrl });
    const { url, revoke } = createModuleUrlFromText(rewritten);
    return { mainUrl: url, revoke };
  }
}

function decodeUtf8FromBytes(bytes: Uint8Array): string {
  return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
}
