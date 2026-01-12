import type {
  ReadExtensionPackageV2Result,
  VerifiedExtensionPackageV2
} from "../../../shared/extension-package/v2-browser.mjs";
import {
  readExtensionPackageV2,
  verifyExtensionPackageV2Browser
} from "../../../shared/extension-package/v2-browser.mjs";
import { validateExtensionManifest } from "../../../shared/extension-manifest/index.mjs";

import { MarketplaceClient } from "./MarketplaceClient";

// The extension worker (`packages/extension-host/src/browser/extension-worker.mjs`) eagerly imports
// `@formula/extension-api`, which initializes the runtime and installs the API object on
// `globalThis[Symbol.for("formula.extensionApi.api")]`.
//
// Browser-loaded extensions cannot reliably import `@formula/extension-api` by bare specifier (no
// import maps in workers), and we also cannot import the package entrypoint via `data:`/`blob:`
// because it has relative imports. Instead we provide a tiny, self-contained ESM shim that
// re-exports the already-initialized API object.
const EXTENSION_API_SHIM_SOURCE = `const api = globalThis[Symbol.for(\"formula.extensionApi.api\")];\nif (!api) { throw new Error(\"@formula/extension-api runtime failed to initialize\"); }\nexport const workbook = api.workbook;\nexport const sheets = api.sheets;\nexport const cells = api.cells;\nexport const commands = api.commands;\nexport const functions = api.functions;\nexport const dataConnectors = api.dataConnectors;\nexport const network = api.network;\nexport const clipboard = api.clipboard;\nexport const ui = api.ui;\nexport const storage = api.storage;\nexport const config = api.config;\nexport const events = api.events;\nexport const context = api.context;\nexport const __setTransport = api.__setTransport;\nexport const __setContext = api.__setContext;\nexport const __handleMessage = api.__handleMessage;\nexport default api;\n`;

export interface InstalledExtensionRecord {
  id: string;
  version: string;
  installedAt: string;
  warnings?: ExtensionInstallWarning[];
  /**
   * When set, the installed record has been quarantined due to an integrity
   * verification failure (e.g. IndexedDB corruption / partial write).
   *
   * Call `repair(id)` to re-download the package and clear this flag.
   */
  corrupted?: boolean;
  corruptedAt?: string;
  corruptedReason?: string;
  /**
   * Marketplace-provided security scan status for the installed package/version.
   *
   * Typically sourced from the download header (`x-package-scan-status`) and/or
   * the per-version metadata field.
   */
  scanStatus?: string | null;
  /**
   * Marketplace key id for the publisher signing key that signed this package.
   *
   * Typically sourced from `x-publisher-key-id` and/or the per-version `signingKeyId`.
   */
  signingKeyId?: string | null;

  /**
   * When set, the installed record has been quarantined because its stored
   * manifest is invalid or incompatible with the current Formula engine version.
   *
   * Installing a compatible extension version (or changing engine version)
   * should clear this flag.
   */
  incompatible?: boolean;
  incompatibleAt?: string;
  incompatibleReason?: string;
}

export interface BrowserExtensionHostLike {
  readonly engineVersion?: string;
  getEngineVersion?(): string;
  loadExtension(args: {
    extensionId: string;
    extensionPath: string;
    manifest: Record<string, any>;
    mainUrl: string;
  }): Promise<string>;
  unloadExtension?(extensionId: string): Promise<void | boolean> | void | boolean;
  /**
   * Preferred single-call API for uninstall flows (clears permissions + storage).
   */
  resetExtensionState?(extensionId: string): Promise<void> | void;
  /**
   * Back-compat: clear persisted permission grants.
   */
  revokePermissions?(extensionId: string, permissions?: string[]): Promise<void> | void;
  /**
   * Back-compat: clear persisted extension storage/config state.
   */
  clearExtensionStorage?(extensionId: string): Promise<void> | void;
  listExtensions(): Array<{ id: string }>;
  /**
   * Starts the host and delivers startup activation events + initial workbook snapshot.
   *
   * Optional for backward compatibility with older BrowserExtensionHost versions.
   */
  startup?: () => Promise<void>;
  /**
   * Starts a single extension if the host is already running, ensuring it receives the
   * initial workbook snapshot. Optional for older hosts.
   */
  startupExtension?: (extensionId: string) => Promise<void>;
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
  /**
   * SHA-256 hex digest of the full package bytes as installed, used to detect
   * post-install IndexedDB corruption.
   */
  packageSha256?: string;
}

const DB_NAME = "formula.webExtensions";
const DB_VERSION = 1;

const STORE_INSTALLED = "installed";
const STORE_PACKAGES = "packages";

const PERMISSIONS_STORE_KEY = "formula.extensionHost.permissions";
const CONTRIBUTED_PANELS_SEED_STORE_KEY = "formula.extensions.contributedPanels.v1";

type ContributedPanelSeed = {
  extensionId: string;
  title: string;
  icon?: string | null;
  defaultDock?: "left" | "right" | "bottom";
};

function getLocalStorage(): Storage | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

function readContributedPanelSeedStore(storage: Storage): Record<string, ContributedPanelSeed> {
  try {
    const raw = storage.getItem(CONTRIBUTED_PANELS_SEED_STORE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return {};
    const out: Record<string, ContributedPanelSeed> = {};
    for (const [panelId, value] of Object.entries(parsed as Record<string, unknown>)) {
      if (typeof panelId !== "string" || panelId.trim().length === 0) continue;
      if (!value || typeof value !== "object" || Array.isArray(value)) continue;
      const record = value as any;
      const extensionId = typeof record.extensionId === "string" ? record.extensionId.trim() : "";
      const title = typeof record.title === "string" ? record.title.trim() : "";
      if (!extensionId || !title) continue;
      const icon = record.icon === undefined ? undefined : record.icon === null ? null : String(record.icon);
      const defaultDock =
        record.defaultDock === "left" || record.defaultDock === "right" || record.defaultDock === "bottom"
          ? record.defaultDock
          : undefined;
      out[panelId] = {
        extensionId,
        title,
        ...(icon !== undefined ? { icon } : {}),
        ...(defaultDock ? { defaultDock } : {})
      };
    }
    return out;
  } catch {
    return {};
  }
}

function writeContributedPanelSeedStore(storage: Storage, data: Record<string, ContributedPanelSeed>): void {
  storage.setItem(CONTRIBUTED_PANELS_SEED_STORE_KEY, JSON.stringify(data));
}

function removeContributedPanelSeedsForExtension(storage: Storage, extensionId: string): void {
  const owner = String(extensionId ?? "").trim();
  if (!owner) return;
  const current = readContributedPanelSeedStore(storage);
  const next: Record<string, ContributedPanelSeed> = {};
  let changed = false;
  for (const [panelId, seed] of Object.entries(current)) {
    if (seed.extensionId === owner) {
      changed = true;
      continue;
    }
    next[panelId] = seed;
  }
  if (!changed) return;
  writeContributedPanelSeedStore(storage, next);
}

function removePermissionGrantsForExtension(storage: Storage, extensionId: string): void {
  const owner = String(extensionId ?? "").trim();
  if (!owner) return;
  try {
    const raw = storage.getItem(PERMISSIONS_STORE_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return;
    if (!Object.prototype.hasOwnProperty.call(parsed, owner)) return;
    delete (parsed as Record<string, unknown>)[owner];
    if (Object.keys(parsed as Record<string, unknown>).length === 0) {
      storage.removeItem(PERMISSIONS_STORE_KEY);
      return;
    }
    storage.setItem(PERMISSIONS_STORE_KEY, JSON.stringify(parsed));
  } catch {
    // ignore
  }
}

function prepareContributedPanelSeedsUpdate(
  storage: Storage,
  extensionId: string,
  panels: Array<{ id?: unknown; title?: unknown; icon?: unknown; defaultDock?: unknown; position?: unknown }>
): Record<string, ContributedPanelSeed> {
  const owner = String(extensionId ?? "").trim();
  if (!owner) return readContributedPanelSeedStore(storage);

  const current = readContributedPanelSeedStore(storage);
  const next: Record<string, ContributedPanelSeed> = {};

  for (const [panelId, seed] of Object.entries(current)) {
    if (seed.extensionId === owner) continue;
    next[panelId] = seed;
  }

  const seenInExtension = new Set<string>();
  for (const panel of panels ?? []) {
    const panelId = typeof panel?.id === "string" ? panel.id.trim() : "";
    if (!panelId) continue;
    if (seenInExtension.has(panelId)) continue;
    seenInExtension.add(panelId);

    const existing = next[panelId];
    if (existing && existing.extensionId !== owner) {
      throw new Error(
        `Panel id already contributed by another extension: ${panelId} (existing: ${existing.extensionId}, new: ${owner})`
      );
    }

    const titleRaw = typeof panel?.title === "string" ? panel.title.trim() : "";
    const title = titleRaw || panelId;
    const icon = panel?.icon === undefined ? undefined : panel.icon === null ? null : String(panel.icon);
    const dockCandidate = (panel as any)?.defaultDock ?? (panel as any)?.position;
    const defaultDock =
      dockCandidate === "left" || dockCandidate === "right" || dockCandidate === "bottom" ? dockCandidate : undefined;
    next[panelId] = {
      extensionId: owner,
      title,
      ...(icon !== undefined ? { icon } : {}),
      ...(defaultDock ? { defaultDock } : {})
    };
  }

  return next;
}

export type ExtensionInstallWarningKind = "deprecated" | "scanStatus";

export interface ExtensionInstallWarning {
  kind: ExtensionInstallWarningKind;
  message: string;
  scanStatus?: string | null;
}

export type ExtensionScanPolicy = "enforce" | "allow" | "ignore";

export interface WebExtensionInstallOptions {
  /**
   * Policy for marketplace package scan statuses that are not "passed".
   *
   * - "enforce": refuse install
   * - "allow": allow install but return a warning (and optionally require confirmation)
   * - "ignore": ignore scan status completely
   *
   * Default: "enforce". Override via `FORMULA_EXTENSION_SCAN_POLICY` /
   * `FORMULA_WEB_EXTENSION_SCAN_POLICY` (node) or `VITE_FORMULA_EXTENSION_SCAN_POLICY` (web).
   */
  scanPolicy?: ExtensionScanPolicy;
  /**
   * Optional confirmation callback invoked for "allowed-but-warned" install states
   * (deprecated extensions, non-passed scan status when scanPolicy="allow").
   *
   * Return `true` to proceed, `false` to cancel the install.
   */
  confirm?: (warning: ExtensionInstallWarning) => Promise<boolean> | boolean;
}

function isNodeRuntime(): boolean {
  return typeof process !== "undefined" && typeof (process as any)?.versions?.node === "string";
}

function defaultScanPolicyFromEnv(): ExtensionScanPolicy {
  if (isNodeRuntime()) {
    const env = (process as any)?.env as Record<string, string | undefined> | undefined;
    const explicit = env?.FORMULA_EXTENSION_SCAN_POLICY ?? env?.FORMULA_WEB_EXTENSION_SCAN_POLICY;
    if (explicit) {
      const normalized = String(explicit).trim().toLowerCase();
      if (normalized === "enforce" || normalized === "allow" || normalized === "ignore") {
        return normalized as ExtensionScanPolicy;
      }
    }
    return "enforce";
  }

  const metaEnv = (import.meta as any)?.env as Record<string, unknown> | undefined;
  const explicit = metaEnv?.VITE_FORMULA_EXTENSION_SCAN_POLICY;
  if (explicit) {
    const normalized = String(explicit).trim().toLowerCase();
    if (normalized === "enforce" || normalized === "allow" || normalized === "ignore") {
      return normalized as ExtensionScanPolicy;
    }
  }
  return "enforce";
}

function normalizeOptionalString(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function normalizeScanStatus(value: unknown): string | null {
  const normalized = normalizeOptionalString(value);
  return normalized ? normalized.toLowerCase() : null;
}

function isSha256Hex(value: string): boolean {
  return typeof value === "string" && /^[0-9a-f]{64}$/i.test(value.trim());
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle?.digest) {
    throw new Error("WebExtensionManager requires crypto.subtle.digest() to verify downloads");
  }
  // `crypto.subtle.digest` expects a BufferSource backed by an `ArrayBuffer`. TypeScript models
  // `Uint8Array` as potentially backed by a `SharedArrayBuffer` (`ArrayBufferLike`), so normalize
  // to an `ArrayBuffer`-backed view for type safety.
  const normalized: Uint8Array<ArrayBuffer> =
    bytes.buffer instanceof ArrayBuffer ? (bytes as Uint8Array<ArrayBuffer>) : new Uint8Array(bytes);

  const hash = new Uint8Array(await subtle.digest("SHA-256", normalized));
  let out = "";
  for (const b of hash) out += b.toString(16).padStart(2, "0");
  return out;
}

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

  const normalized: Uint8Array<ArrayBuffer> =
    bytes.buffer instanceof ArrayBuffer ? (bytes as Uint8Array<ArrayBuffer>) : new Uint8Array(bytes);

  if (
    !isNodeRuntime &&
    typeof URL !== "undefined" &&
    typeof URL.createObjectURL === "function" &&
    typeof Blob !== "undefined"
  ) {
    const url = URL.createObjectURL(new Blob([normalized], { type: mime }));
    return { url, revoke: () => URL.revokeObjectURL(url) };
  }
  const url = bytesToDataUrl(normalized, mime);
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
    .replace(/import\s+["']@formula\/extension-api["']/g, `import "${extensionApiUrl}"`)
    .replace(/from\s+["']formula["']/g, `from "${extensionApiUrl}"`)
    .replace(/import\s+["']formula["']/g, `import "${extensionApiUrl}"`);

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
  readonly engineVersion: string;
  readonly scanPolicy: ExtensionScanPolicy;

  private readonly _loadedMainUrls = new Map<string, { mainUrl: string; revoke: () => void }>();
  private _extensionApiModule: { url: string; revoke: () => void } | null = null;
  private _didHostStartup = false;

  constructor(
    options: {
      marketplaceClient?: MarketplaceClient;
      host?: BrowserExtensionHostLike | null;
      engineVersion?: string;
      scanPolicy?: ExtensionScanPolicy;
    } = {}
  ) {
    this.marketplaceClient = options.marketplaceClient ?? new MarketplaceClient();
    this.host = options.host ?? null;
    const fromHost =
      this.host?.engineVersion ??
      (typeof this.host?.getEngineVersion === "function" ? this.host.getEngineVersion() : null);
    const rawEngine = options.engineVersion ?? fromHost ?? "1.0.0";
    this.engineVersion = typeof rawEngine === "string" && rawEngine.trim().length > 0 ? rawEngine.trim() : "1.0.0";
    this.scanPolicy = options.scanPolicy ?? defaultScanPolicyFromEnv();
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

  async install(
    id: string,
    version: string | null = null,
    options: WebExtensionInstallOptions = {}
  ): Promise<InstalledExtensionRecord> {
    const ext = await this.marketplaceClient.getExtension(id);
    if (!ext) throw new Error(`Extension not found: ${id}`);

    if (ext.blocked) {
      throw new Error(`Extension is blocked and cannot be installed: ${id}`);
    }
    if (ext.malicious) {
      throw new Error(`Extension is marked as malicious and cannot be installed: ${id}`);
    }
    if (ext.publisherRevoked) {
      throw new Error(`Extension publisher is revoked (refusing to install): ${id}`);
    }

    const warnings: ExtensionInstallWarning[] = [];
    const confirm = typeof options.confirm === "function" ? options.confirm : null;
    const addWarning = async (warning: ExtensionInstallWarning) => {
      warnings.push(warning);
      if (confirm) {
        const ok = await confirm(warning);
        if (!ok) {
          throw new Error("Extension install cancelled");
        }
      }
    };

    if (ext.deprecated) {
      await addWarning({
        kind: "deprecated",
        message: `Extension ${id} is deprecated. It may be unmaintained and could be removed from the marketplace.`,
      });
    }

    const resolvedVersion = version ?? ext.latestVersion;
    if (!resolvedVersion) {
      throw new Error("Marketplace did not provide latestVersion");
    }

    const versionMeta = Array.isArray(ext.versions)
      ? ext.versions.find((v) => v && typeof (v as any).version === "string" && (v as any).version === resolvedVersion) ??
        null
      : null;

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

    const signingKeyIdFromHeader = normalizeOptionalString(download.publisherKeyId);
    const signingKeyIdFromVersion = normalizeOptionalString((versionMeta as any)?.signingKeyId);
    if (signingKeyIdFromHeader && signingKeyIdFromVersion && signingKeyIdFromHeader !== signingKeyIdFromVersion) {
      throw new Error(
        `Marketplace signing key mismatch for ${id}@${resolvedVersion}: header=${signingKeyIdFromHeader} version=${signingKeyIdFromVersion}`
      );
    }
    const signingKeyId = signingKeyIdFromHeader ?? signingKeyIdFromVersion ?? null;

    if (hasPublisherKeySet && signingKeyId) {
      const keyId = String(signingKeyId);
      const preferred = candidateKeys.find((k) => k.id === keyId);
      if (preferred) {
        candidateKeys = [preferred, ...candidateKeys.filter((k) => k.id !== keyId)];
      }
    }

    const scanStatusFromHeader = normalizeScanStatus(download.scanStatus);
    const scanStatusFromVersion = normalizeScanStatus((versionMeta as any)?.scanStatus);
    if (scanStatusFromHeader && scanStatusFromVersion && scanStatusFromHeader !== scanStatusFromVersion) {
      throw new Error(
        `Marketplace provided conflicting package scan statuses for ${id}@${resolvedVersion}: header=${scanStatusFromHeader} version=${scanStatusFromVersion}`
      );
    }

    const scanStatus = scanStatusFromHeader ?? scanStatusFromVersion ?? null;
    const policy = options.scanPolicy ?? this.scanPolicy;
    if (policy !== "ignore") {
      if (!scanStatus) {
        if (policy === "enforce") {
          throw new Error(
            `Refusing to install ${id}@${resolvedVersion}: package scan status is missing (expected "passed")`
          );
        }
        if (policy === "allow") {
          await addWarning({
            kind: "scanStatus",
            scanStatus: null,
            message: `Extension ${id}@${resolvedVersion} is missing package scan status. Proceed with caution.`
          });
        }
      } else if (scanStatus !== "passed") {
        if (policy === "enforce") {
          throw new Error(
            `Refusing to install ${id}@${resolvedVersion}: package scan status is "${scanStatus}" (expected "passed")`
          );
        }
        if (policy === "allow") {
          await addWarning({
            kind: "scanStatus",
            scanStatus,
            message: `Extension ${id}@${resolvedVersion} has package scan status "${scanStatus}". Proceed with caution.`
          });
        }
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

    const manifest = this._validateManifest(verified.manifest, { id, version: resolvedVersion });
    verified = { ...verified, manifest };
    const bundleId = `${manifest.publisher}.${manifest.name}`;
    if (bundleId !== id) {
      throw new Error(`Package id mismatch: expected ${id} but got ${bundleId}`);
    }
    if (manifest.version !== resolvedVersion) {
      throw new Error(`Package version mismatch: expected ${resolvedVersion} but got ${manifest.version}`);
    }

    // Prepare a localStorage seed-store update before committing the install to IndexedDB.
    // This ensures we fail fast on duplicate panel ids (two installed extensions claiming the same
    // panel id), and avoids ending up with an installed extension whose panels cannot be seeded
    // synchronously at startup.
    const seedStorage = getLocalStorage();
    const seedUpdate =
      seedStorage && Array.isArray(manifest.contributes?.panels)
        ? prepareContributedPanelSeedsUpdate(seedStorage, id, manifest.contributes.panels)
        : null;

    if (download.signatureBase64 && download.signatureBase64 !== verified.signatureBase64) {
      throw new Error("Marketplace signature header does not match package signature");
    }

    const headerFilesSha256 = download.filesSha256 ? String(download.filesSha256).trim().toLowerCase() : null;
    if (headerFilesSha256 && isSha256Hex(headerFilesSha256)) {
      const filesJson = JSON.stringify(verified.files || []);
      const computedFilesSha = await sha256Hex(new TextEncoder().encode(filesJson));
      if (computedFilesSha !== headerFilesSha256) {
        throw new Error("Marketplace files sha256 header does not match verified package contents");
      }
    }

    // Persist the full package sha256 so we can detect IndexedDB corruption on subsequent loads.
    //
    // The marketplace client (when used) already validates `download.sha256` against the bytes,
    // but compute our own digest here so custom marketplace implementations cannot accidentally
    // store a bad hash.
    const computedPackageSha256 = await sha256Hex(download.bytes);
    const headerPackageSha256 = download.sha256 ? String(download.sha256).trim().toLowerCase() : null;
    if (headerPackageSha256 && isSha256Hex(headerPackageSha256) && headerPackageSha256 !== computedPackageSha256) {
      throw new Error("Marketplace package sha256 header does not match downloaded bytes");
    }

    const installedAt = new Date().toISOString();
    const installedRecord: InstalledExtensionRecord = {
      id,
      version: resolvedVersion,
      installedAt,
      scanStatus,
      signingKeyId,
      ...(warnings.length > 0 ? { warnings } : {})
    };
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
        verified,
        packageSha256: computedPackageSha256
      };

      packagesStore.put(pkgRecord);
      installedStore.put(installedRecord);

      if (prev && prev.version && prev.version !== resolvedVersion) {
        packagesStore.delete(`${id}@${prev.version}`);
      }

      await txDone(tx);
    } finally {
      db.close();
    }

    if (seedStorage && seedUpdate) {
      try {
        writeContributedPanelSeedStore(seedStorage, seedUpdate);
      } catch (error) {
        // Best-effort: localStorage may be unavailable or full. Failing to write the seed store
        // should not prevent extension installation; it only affects layout persistence across
        // restarts.
        // eslint-disable-next-line no-console
        console.warn(
          `Failed to persist contributed panel seed store for ${id}@${resolvedVersion}: ${String(
            (error as Error)?.message ?? error
          )}`
        );
      }
    }

    return installedRecord;
  }

  async uninstall(id: string): Promise<void> {
    const existing = await this.getInstalled(id);
    if (!existing) return;

    await this.unload(id).catch(() => {});
    try {
      const host = this.host;
      if (host && typeof host.resetExtensionState === "function") {
        await host.resetExtensionState(String(id));
      } else {
        await host?.revokePermissions?.(String(id));
        await host?.clearExtensionStorage?.(String(id));
      }
    } catch {
      // ignore (host/storage might be unavailable)
    }

    const db = await openDb();
    try {
      const tx = db.transaction([STORE_INSTALLED, STORE_PACKAGES], "readwrite");
      tx.objectStore(STORE_INSTALLED).delete(String(id));
      tx.objectStore(STORE_PACKAGES).delete(`${id}@${existing.version}`);
      await txDone(tx);
    } finally {
      db.close();
    }

    // Best-effort cleanup of persisted state owned by the uninstalled extension so a reinstall
    // behaves like a clean install.
    //
    // Note: We prefer clearing state via the host (resetExtensionState/revokePermissions/clearExtensionStorage),
    // but we also clear known default localStorage keys as a fallback when the host is not available.
    const localStorage = getLocalStorage();
    if (localStorage) {
      try {
        removePermissionGrantsForExtension(localStorage, String(id));
      } catch {
        // ignore
      }

      try {
        localStorage.removeItem(`formula.extensionHost.storage.${String(id)}`);
      } catch {
        // ignore
      }

      try {
        removeContributedPanelSeedsForExtension(localStorage, String(id));
      } catch {
        // ignore
      }
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

  async repair(id: string): Promise<InstalledExtensionRecord> {
    const installed = await this.getInstalled(id);
    if (!installed) throw new Error(`Not installed: ${id}`);
    if (this.isLoaded(id)) {
      await this.unload(id);
    }
    // Repair always re-downloads the currently-installed version, replacing the stored bytes and
    // clearing any corruption markers.
    try {
      return await this.install(id, installed.version);
    } catch (error) {
      const msg = String((error as Error)?.message ?? error);
      // If the exact version is no longer available (yanked/deleted), fall back to installing the
      // latest version so users still have a recovery path without manual DB clearing.
      if (!/Package not found/i.test(msg)) throw error;
      return this.install(id);
    }
  }

  isLoaded(id: string): boolean {
    return this._loadedMainUrls.has(String(id));
  }

  async loadInstalled(id: string): Promise<string> {
    return this._loadInstalledInternal(id, { start: true });
  }

  /**
   * Loads all extensions installed in IndexedDB, then triggers host startup once (if supported).
   *
   * This is the preferred entrypoint for desktop/web boot: it ensures extensions that rely on
   * `activationEvents: ["onStartupFinished"]` are activated and receive the initial
   * `workbookOpened` event.
   */
  async loadAllInstalled(): Promise<string[]> {
    if (!this.host) throw new Error("WebExtensionManager requires a BrowserExtensionHost to load extensions");
    const installed = await this.listInstalled();

    const newlyLoaded: string[] = [];
    for (const record of installed) {
      // Avoid throwing if the caller invokes `loadAllInstalled()` multiple times.
      if (this.isLoaded(record.id) || this.host.listExtensions().some((e) => e.id === record.id)) continue;
      // eslint-disable-next-line no-await-in-loop
      await this._loadInstalledInternal(record.id, { start: false });
      newlyLoaded.push(record.id);
    }

    // Preferred behavior: call host.startup() exactly once so existing extensions don't receive
    // duplicate startup events.
    if (this.host.startup && !this._didHostStartup) {
      await this.host.startup();
      this._didHostStartup = true;
      return newlyLoaded;
    }

    // If the host is already running, ensure any newly loaded extensions get startup semantics
    // without spamming already-active extensions.
    if (this.host.startupExtension && newlyLoaded.length > 0) {
      for (const id of newlyLoaded) {
        // eslint-disable-next-line no-await-in-loop
        await this.host.startupExtension(id);
      }
    }

    return newlyLoaded;
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

  private async _loadInstalledInternal(id: string, options: { start: boolean }): Promise<string> {
    if (!this.host) throw new Error("WebExtensionManager requires a BrowserExtensionHost to load extensions");
    const installed = await this.getInstalled(id);
    if (!installed) throw new Error(`Not installed: ${id}`);
    if (installed.corrupted) {
      const reason = installed.corruptedReason ? `: ${installed.corruptedReason}` : "";
      throw new Error(
        `Extension ${installed.id}@${installed.version} is corrupted${reason}. ` +
          "Call WebExtensionManager.repair(id) to re-download the package."
      );
    }

    const pkg = await this._getPackage(installed.id, installed.version);
    if (!pkg) {
      const reason = `missing stored package record for ${installed.id}@${installed.version}`;
      await this._quarantineCorruptedInstall(installed, reason);
      throw new Error(
        `Extension package integrity check failed for ${installed.id}@${installed.version}: ${reason}. ` +
          "Call WebExtensionManager.repair(id) to re-download the package."
      );
    }

    await this._verifyStoredPackageIntegrity(installed, pkg);

    let manifest: Record<string, any>;
    try {
      manifest = this._validateManifest(pkg.verified.manifest, { id: installed.id, version: installed.version });
    } catch (error) {
      const msg = String((error as Error)?.message ?? error);
      await this._quarantineIncompatibleInstall(installed, msg).catch(() => {});
      throw error instanceof Error ? error : new Error(msg);
    }

    if (installed.incompatible) {
      await this._clearIncompatibleInstall(installed).catch(() => {});
    }

    const bundleId = `${manifest.publisher}.${manifest.name}`;
    if (bundleId !== installed.id) {
      const reason = `package id mismatch (expected ${installed.id} but got ${bundleId})`;
      await this._quarantineCorruptedInstall(installed, reason);
      throw new Error(`Extension package integrity check failed for ${installed.id}@${installed.version}: ${reason}`);
    }
    if (manifest.version !== installed.version) {
      const reason = `package version mismatch (expected ${installed.version} but got ${manifest.version})`;
      await this._quarantineCorruptedInstall(installed, reason);
      throw new Error(`Extension package integrity check failed for ${installed.id}@${installed.version}: ${reason}`);
    }

    if (!this._extensionApiModule) {
      this._extensionApiModule = createModuleUrlFromText(EXTENSION_API_SHIM_SOURCE);
    }

    const { mainUrl, revoke } = await this._createMainModuleUrl(pkg, manifest, this._extensionApiModule.url);

    const extensionPath = `indexeddb://formula/extensions/${installed.id}/${installed.version}`;

    // If already loaded, the host will throw; ensure the manager keeps a single source of truth.
    const existing = this.host.listExtensions();
    if (existing.some((e) => e.id === installed.id)) {
      throw new Error(`Extension already loaded: ${installed.id}`);
    }

    const hadOtherExtensions = existing.length > 0;

    await this.host.loadExtension({
      extensionId: installed.id,
      extensionPath,
      manifest,
      mainUrl
    });

    this._loadedMainUrls.set(installed.id, { mainUrl, revoke });

    if (options.start) {
      if (this.host.startupExtension) {
        await this.host.startupExtension(installed.id);
      } else if (this.host.startup && !hadOtherExtensions && !this._didHostStartup) {
        // Fallback for older hosts: only call startup() when we're confident it won't
        // re-emit startup events to extensions that were already running.
        await this.host.startup();
        this._didHostStartup = true;
      }
    }

    return installed.id;
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

  private async _verifyStoredPackageIntegrity(
    installed: InstalledExtensionRecord,
    pkg: StoredPackageRecord
  ): Promise<void> {
    const bytes = new Uint8Array(pkg.bytes);
    const computedSha = await sha256Hex(bytes);
    const expectedSha =
      typeof pkg.packageSha256 === "string" && isSha256Hex(pkg.packageSha256) ? pkg.packageSha256.trim().toLowerCase() : null;

    if (expectedSha) {
      if (computedSha !== expectedSha) {
        const reason = `package sha256 mismatch (expected ${expectedSha} but got ${computedSha})`;
        await this._quarantineCorruptedInstall(installed, reason);
        throw new Error(`Extension package integrity check failed for ${installed.id}@${installed.version}: ${reason}`);
      }
      return;
    }

    // Legacy install: before we had `packageSha256`, we can still verify integrity by re-checking
    // the stored file digests we computed at install time. If that passes, persist the package
    // sha256 for faster checks going forward.
    try {
      const expectedFiles = new Map<string, { sha256: string; size: number }>();
      const storedFiles = Array.isArray(pkg.verified?.files) ? pkg.verified.files : [];
      for (const f of storedFiles) {
        if (!f || typeof (f as any).path !== "string") continue;
        const sha = typeof (f as any).sha256 === "string" ? String((f as any).sha256).trim().toLowerCase() : "";
        const size = (f as any).size;
        if (!isSha256Hex(sha) || typeof size !== "number" || !Number.isFinite(size) || size < 0) continue;
        expectedFiles.set(String((f as any).path), { sha256: sha, size });
      }

      if (expectedFiles.size === 0) {
        throw new Error("missing integrity metadata (no packageSha256 and no verified file list)");
      }

      const parsed: ReadExtensionPackageV2Result = readExtensionPackageV2(bytes);
      for (const [relPath, fileBytes] of parsed.files.entries()) {
        const expected = expectedFiles.get(relPath);
        if (!expected) {
          throw new Error(`unexpected file in archive: ${relPath}`);
        }
        // eslint-disable-next-line no-await-in-loop
        const fileSha = await sha256Hex(fileBytes);
        if (fileSha !== expected.sha256) {
          throw new Error(`checksum mismatch for ${relPath}`);
        }
        if (fileBytes.length !== expected.size) {
          throw new Error(`size mismatch for ${relPath}`);
        }
        expectedFiles.delete(relPath);
      }

      if (expectedFiles.size > 0) {
        throw new Error(`archive missing expected files: ${[...expectedFiles.keys()].join(", ")}`);
      }

      await this._persistPackageSha256(pkg, computedSha);
    } catch (error) {
      const reason = String((error as Error)?.message ?? error);
      await this._quarantineCorruptedInstall(installed, reason);
      throw new Error(`Extension package integrity check failed for ${installed.id}@${installed.version}: ${reason}`);
    }
  }

  private async _persistPackageSha256(pkg: StoredPackageRecord, packageSha256: string): Promise<void> {
    if (!isSha256Hex(packageSha256)) return;
    if (pkg.packageSha256 && pkg.packageSha256.trim().toLowerCase() === packageSha256.trim().toLowerCase()) return;
    const db = await openDb();
    try {
      const tx = db.transaction([STORE_PACKAGES], "readwrite");
      tx.objectStore(STORE_PACKAGES).put({ ...pkg, packageSha256: packageSha256.trim().toLowerCase() });
      await txDone(tx);
    } finally {
      db.close();
    }
  }

  private async _quarantineCorruptedInstall(installed: InstalledExtensionRecord, reason: string): Promise<void> {
    const db = await openDb();
    try {
      const tx = db.transaction([STORE_INSTALLED], "readwrite");
      const store = tx.objectStore(STORE_INSTALLED);
      store.put({
        ...installed,
        corrupted: true,
        corruptedAt: new Date().toISOString(),
        corruptedReason: String(reason || "unknown error")
      });
      await txDone(tx);
    } finally {
      db.close();
    }
  }

  private async _createMainModuleUrl(
    pkg: StoredPackageRecord,
    manifest: Record<string, any>,
    extensionApiUrl: string
  ): Promise<{ mainUrl: string; revoke: () => void }> {
    const bytes = new Uint8Array(pkg.bytes);
    const parsed: ReadExtensionPackageV2Result = readExtensionPackageV2(bytes);

    const entryRel = extractEntrypointPath(manifest);
    const entryBytes = parsed.files.get(entryRel);
    if (!entryBytes) {
      throw new Error(`Extension entrypoint missing from package: ${entryRel}`);
    }

    const source = decodeUtf8FromBytes(entryBytes);
    const rewritten = rewriteEntrypointSource(source, { extensionApiUrl });
    const { url, revoke } = createModuleUrlFromText(rewritten);
    return { mainUrl: url, revoke };
  }

  private _validateManifest(
    manifest: Record<string, any>,
    context: { id: string; version: string }
  ): Record<string, any> {
    try {
      return validateExtensionManifest(manifest, {
        engineVersion: this.engineVersion,
        enforceEngine: true
      }) as Record<string, any>;
    } catch (error) {
      const msg = String((error as Error)?.message ?? error);
      throw new Error(`Invalid extension manifest for ${context.id}@${context.version}: ${msg}`);
    }
  }

  private async _quarantineIncompatibleInstall(installed: InstalledExtensionRecord, reason: string): Promise<void> {
    const db = await openDb();
    try {
      const tx = db.transaction([STORE_INSTALLED], "readwrite");
      const store = tx.objectStore(STORE_INSTALLED);
      store.put({
        ...installed,
        incompatible: true,
        incompatibleAt: new Date().toISOString(),
        incompatibleReason: String(reason || "unknown error")
      });
      await txDone(tx);
    } finally {
      db.close();
    }
  }

  private async _clearIncompatibleInstall(installed: InstalledExtensionRecord): Promise<void> {
    if (!installed.incompatible && !installed.incompatibleAt && !installed.incompatibleReason) return;
    const db = await openDb();
    try {
      const tx = db.transaction([STORE_INSTALLED], "readwrite");
      const store = tx.objectStore(STORE_INSTALLED);
      const next: InstalledExtensionRecord & Record<string, any> = { ...installed };
      delete (next as any).incompatible;
      delete (next as any).incompatibleAt;
      delete (next as any).incompatibleReason;
      store.put(next);
      await txDone(tx);
    } finally {
      db.close();
    }
  }
}

function decodeUtf8FromBytes(bytes: Uint8Array): string {
  return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
}
