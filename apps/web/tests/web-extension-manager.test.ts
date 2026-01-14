import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { Worker as NodeWorker } from "node:worker_threads";

import { IDBKeyRange, indexedDB } from "fake-indexeddb";
import { afterAll, afterEach, beforeEach, expect, test } from "vitest";

import { WebExtensionManager } from "@formula/extension-marketplace";

// CJS helpers (shared/* is CommonJS).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const signingPkg: any = await import("../../../shared/crypto/signing.js");
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const extensionPackagePkg: any = await import("../../../shared/extension-package/index.js");

const { generateEd25519KeyPair } = signingPkg;
const { createExtensionPackageV2 } = extensionPackagePkg;

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const browserHostPkg: any = await import("../../../packages/extension-host/src/browser/index.mjs");
const { BrowserExtensionHost } = browserHostPkg;

const WORKER_WRAPPER_URL = new URL("./helpers/node-web-worker.mjs", import.meta.url);

const originalGlobals = {
  indexedDB: (globalThis as any).indexedDB,
  IDBKeyRange: (globalThis as any).IDBKeyRange,
  crypto: (globalThis as any).crypto,
  Worker: (globalThis as any).Worker,
  localStorage: (globalThis as any).localStorage
};

class MemoryStorage implements Storage {
  private readonly _map = new Map<string, string>();

  get length(): number {
    return this._map.size;
  }

  clear(): void {
    this._map.clear();
  }

  getItem(key: string): string | null {
    const value = this._map.get(String(key));
    return value === undefined ? null : value;
  }

  key(index: number): string | null {
    return [...this._map.keys()][index] ?? null;
  }

  removeItem(key: string): void {
    this._map.delete(String(key));
  }

  setItem(key: string, value: string): void {
    this._map.set(String(key), String(value));
  }
}

async function waitFor(condition: () => boolean, { timeoutMs = 3000, intervalMs = 25 } = {}) {
  const start = Date.now();
  for (;;) {
    if (condition()) return;
    if (Date.now() - start > timeoutMs) {
      throw new Error("Timed out waiting for condition");
    }
    // eslint-disable-next-line no-await-in-loop
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
}

class NodeWebWorkerShim {
  private readonly _worker: NodeWorker;
  private readonly _listeners = new Map<string, Set<(event: any) => void>>();

  constructor(url: URL | string) {
    const workerUrl = typeof url === "string" ? url : url.href;
    this._worker = new NodeWorker(WORKER_WRAPPER_URL, { type: "module", workerData: { url: workerUrl } });

    this._worker.on("message", (data) => this._emit("message", { data }));
    this._worker.on("error", (err) => this._emit("error", { message: err?.message ?? String(err), error: err }));
  }

  addEventListener(type: string, listener: (event: any) => void) {
    const key = String(type);
    if (!this._listeners.has(key)) this._listeners.set(key, new Set());
    this._listeners.get(key)!.add(listener);
  }

  removeEventListener(type: string, listener: (event: any) => void) {
    this._listeners.get(String(type))?.delete(listener);
  }

  postMessage(message: any) {
    this._worker.postMessage(message);
  }

  terminate() {
    void this._worker.terminate();
  }

  private _emit(type: string, event: any) {
    const set = this._listeners.get(String(type));
    if (!set) return;
    for (const listener of [...set]) {
      try {
        listener(event);
      } catch {
        // ignore
      }
    }
  }
}

class TestSpreadsheetApi {
  private readonly _cells = new Map<string, any>();
  private _selection = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };

  setSelection(range: { startRow: number; startCol: number; endRow: number; endCol: number }) {
    this._selection = { ...range };
  }

  getSelection() {
    const { startRow, startCol, endRow, endCol } = this._selection;
    const values: any[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      const row: any[] = [];
      for (let c = startCol; c <= endCol; c++) {
        row.push(this.getCell(r, c));
      }
      values.push(row);
    }
    return { startRow, startCol, endRow, endCol, values };
  }

  getCell(row: number, col: number) {
    return this._cells.get(`${row},${col}`) ?? null;
  }

  async setCell(row: number, col: number, value: any) {
    this._cells.set(`${row},${col}`, value);
  }
}

function createMockMarketplace({
  extensionId,
  latestVersion,
  publicKeyPem,
  packages,
  publisherKeys = null,
  publisherKeyIds = null,
  scanStatuses = null,
  extensionVersions = null,
  deprecated = false,
  blocked = false,
  malicious = false,
  publisherRevoked = false
}: any) {
  return {
    async getExtension(id: string) {
      if (id !== extensionId) return null;
      return {
        id,
        name: id.split(".")[1],
        displayName: id,
        publisher: id.split(".")[0],
        description: "",
        latestVersion,
        verified: true,
        featured: false,
        categories: [],
        tags: [],
        screenshots: [],
        downloadCount: 0,
        updatedAt: new Date().toISOString(),
        versions: Array.isArray(extensionVersions) ? extensionVersions : [],
        readme: "",
        publisherPublicKeyPem: publicKeyPem,
        publisherKeys: Array.isArray(publisherKeys) ? publisherKeys : undefined,
        createdAt: new Date().toISOString(),
        deprecated: Boolean(deprecated),
        blocked: Boolean(blocked),
        malicious: Boolean(malicious),
        publisherRevoked: Boolean(publisherRevoked)
      };
    },
    async downloadPackage(id: string, version: string) {
      if (id !== extensionId) return null;
      const pkg = packages[version];
      if (!pkg) return null;
      return {
        bytes: new Uint8Array(pkg),
        signatureBase64: null,
        sha256: null,
        formatVersion: 2,
        publisher: id.split(".")[0],
        publisherKeyId:
          publisherKeyIds && typeof publisherKeyIds === "object" && typeof publisherKeyIds[version] === "string"
            ? publisherKeyIds[version]
            : null,
        scanStatus:
          scanStatuses && typeof scanStatuses === "object" && Object.prototype.hasOwnProperty.call(scanStatuses, version)
            ? (scanStatuses as any)[version]
            : "passed",
        filesSha256: null
      };
    }
  };
}

async function deleteExtensionDb() {
  await new Promise<void>((resolve) => {
    const req = indexedDB.deleteDatabase("formula.webExtensions");
    req.onsuccess = () => resolve();
    req.onerror = () => resolve();
    req.onblocked = () => resolve();
  });
}

function requestToPromise<T = unknown>(req: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error("IndexedDB request failed"));
  });
}

function txDone(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
    tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
  });
}

beforeEach(async () => {
  // Provide browser primitives.
  // eslint-disable-next-line no-global-assign
  (globalThis as any).indexedDB = indexedDB;
  // eslint-disable-next-line no-global-assign
  (globalThis as any).IDBKeyRange = IDBKeyRange;
  // eslint-disable-next-line no-global-assign
  (globalThis as any).localStorage = new MemoryStorage();
  if (!globalThis.crypto?.subtle) {
    // eslint-disable-next-line no-global-assign
    (globalThis as any).crypto = crypto.webcrypto;
  }
  // eslint-disable-next-line no-global-assign
  (globalThis as any).Worker = NodeWebWorkerShim;

  await deleteExtensionDb();
});

afterEach(async () => {
  await deleteExtensionDb();
});

afterAll(() => {
  // Restore any globals we replaced so other test suites are unaffected.
  try {
    // eslint-disable-next-line no-global-assign
    (globalThis as any).indexedDB = originalGlobals.indexedDB;
  } catch {
    // ignore
  }
  try {
    // eslint-disable-next-line no-global-assign
    (globalThis as any).IDBKeyRange = originalGlobals.IDBKeyRange;
  } catch {
    // ignore
  }
  try {
    // eslint-disable-next-line no-global-assign
    (globalThis as any).crypto = originalGlobals.crypto;
  } catch {
    // Some Node versions expose `crypto` as a read-only getter; we only assign it in
    // the test setup when it's missing, so ignore restoration failures.
  }
  try {
    // eslint-disable-next-line no-global-assign
    (globalThis as any).Worker = originalGlobals.Worker;
  } catch {
    // ignore
  }
  try {
    // eslint-disable-next-line no-global-assign
    (globalThis as any).localStorage = originalGlobals.localStorage;
  } catch {
    // ignore
  }
});

test("install → verify → load → execute a command (sample-hello)", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const spreadsheet = new TestSpreadsheetApi();
  spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
  await spreadsheet.setCell(0, 0, 1);
  await spreadsheet.setCell(0, 1, 2);
  await spreadsheet.setCell(1, 0, 3);
  await spreadsheet.setCell(1, 1, 4);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: spreadsheet,
    permissionPrompt: async () => true
  });

  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  const installed = await manager.install("formula.sample-hello");
  expect(installed.scanStatus).toBe("passed");
  expect(installed.signingKeyId ?? null).toBeNull();
  expect((await manager.getInstalled("formula.sample-hello"))?.scanStatus).toBe("passed");

  // Installing an extension should synchronously persist its contributed panels so the desktop
  // layout system can seed the panel registry before deserializing persisted layouts.
  const seedRaw = globalThis.localStorage.getItem("formula.extensions.contributedPanels.v1");
  expect(seedRaw).not.toBeNull();
  const seed = JSON.parse(String(seedRaw));
  expect(seed["sampleHello.panel"]).toMatchObject({
    extensionId: "formula.sample-hello",
    title: "Sample Hello Panel",
  });
  await manager.loadInstalled("formula.sample-hello");

  const result = await host.executeCommand("sampleHello.sumSelection");
  expect(result).toBe(10);
  expect(spreadsheet.getCell(2, 0)).toBe(10);

  await manager.dispose();
  await host.dispose();
});

test("install persists contributed panel metadata (icon + defaultDock) into the seed store", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-panel-seed-metadata-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "panel-seed-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.mjs",
    engines: { formula: "^1.0.0" },
    contributes: {
      panels: [
        {
          id: "test.panelSeed",
          title: "Seed Panel",
          icon: "seed-icon",
          // Not (yet) validated by the shared manifest schema, but allowed as an extra field.
          // Desktop uses this as the default dock side when seeding the panel registry.
          position: "left"
        }
      ]
    }
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  const source = `export async function activate() {}\n`;
  await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "test.panel-seed-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
    // Activation involves worker spin-up + module loading; bump timeout in this integration test
    // to reduce flake under heavy CI load.
    activationTimeoutMs: 60000
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await manager.install("test.panel-seed-ext");

  const seedRaw = globalThis.localStorage.getItem("formula.extensions.contributedPanels.v1");
  expect(seedRaw).not.toBeNull();
  const seed = JSON.parse(String(seedRaw));
  expect(seed["test.panelSeed"]).toMatchObject({
    extensionId: "test.panel-seed-ext",
    title: "Seed Panel",
    icon: "seed-icon",
    defaultDock: "left"
  });

  await manager.dispose();
  await host.dispose();
  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("install overwrites a corrupted contributed panel seed store", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const manager = new WebExtensionManager({ marketplaceClient, host: null, engineVersion: "1.0.0" });

  // Simulate a corrupted seed store from an older/broken build.
  globalThis.localStorage.setItem("formula.extensions.contributedPanels.v1", "{not-json");

  await manager.install("formula.sample-hello");

  const seedRaw = globalThis.localStorage.getItem("formula.extensions.contributedPanels.v1");
  expect(seedRaw).not.toBeNull();
  const seed = JSON.parse(String(seedRaw));
  expect(seed["sampleHello.panel"]).toMatchObject({
    extensionId: "formula.sample-hello",
    title: "Sample Hello Panel"
  });

  await manager.dispose();
});

test("install clears contributed panel seed store entries when updating to a version without panels", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-panel-seed-update-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const extensionId = "test.panel-seed-update-ext";

  const source = `export async function activate() {}\n`;
  await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);

  const manifestV1 = {
    name: "panel-seed-update-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.mjs",
    engines: { formula: "^1.0.0" },
    contributes: {
      panels: [{ id: "test.panelSeedUpdate", title: "Seed Update Panel" }],
    },
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifestV1, null, 2));
  const pkgV1 = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });

  // Update removes the panels contribution entirely (no `contributes.panels` array).
  // WebExtensionManager should treat this as "no panels" and remove any previously persisted
  // contributed panel seed entries for the extension.
  const manifestV2 = {
    ...manifestV1,
    version: "1.0.1",
    contributes: {},
  };
  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifestV2, null, 2));
  const pkgV2 = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId,
    latestVersion: "1.0.1",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgV1, "1.0.1": pkgV2 },
  });

  const manager = new WebExtensionManager({ marketplaceClient, host: null, engineVersion: "1.0.0" });

  try {
    await manager.install(extensionId, "1.0.0");

    const seedKey = "formula.extensions.contributedPanels.v1";
    const seedRaw = globalThis.localStorage.getItem(seedKey);
    expect(seedRaw).not.toBeNull();
    const seed = JSON.parse(String(seedRaw));
    expect(seed["test.panelSeedUpdate"]).toMatchObject({
      extensionId,
      title: "Seed Update Panel",
    });

    await manager.install(extensionId, "1.0.1");

    // When the last contributed panel seed is removed, the key should be deleted entirely
    // (avoids leaving behind an empty "{}" record).
    expect(globalThis.localStorage.getItem(seedKey)).toBeNull();
  } finally {
    await manager.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("loadInstalled triggers onStartupFinished activation + initial workbookOpened (storage key)", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-startup-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "startup-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onStartupFinished"],
    permissions: ["storage"],
    contributes: {}
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(
    path.join(extDir, "dist", "extension.js"),
    `import * as formula from "@formula/extension-api";\nexport async function activate() {\n  formula.events.onWorkbookOpened(() => {\n    // Event handlers are not awaited by the runtime.\n    formula.storage.set("startup.workbookOpened", "ok").catch(() => {});\n  });\n}\n`
  );

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.startup-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host });

  try {
    await manager.install("test.startup-ext");
    await manager.loadInstalled("test.startup-ext");

    const store = (host as any)._storageApi.getExtensionStore("test.startup-ext");
    await waitFor(() => store["startup.workbookOpened"] === "ok");
    expect(store["startup.workbookOpened"]).toBe("ok");
    expect(host.listExtensions().find((e: any) => e.id === "test.startup-ext")?.active).toBe(true);
  } finally {
    await manager.dispose();
    await host.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("loadInstalled falls back to host.startup() when startupExtension is unavailable (first extension)", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-startup-fallback-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "startup-fallback-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onStartupFinished"],
    contributes: {}
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), `export async function activate() {}\n`);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.startup-fallback-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const loaded: Array<{ id: string }> = [];
  let startupCalls = 0;
  const host = {
    loadExtension: async ({ extensionId }: { extensionId: string }) => {
      loaded.push({ id: extensionId });
      return extensionId;
    },
    listExtensions: () => loaded,
    startup: async () => {
      startupCalls++;
    }
  };

  const manager = new WebExtensionManager({ marketplaceClient, host: host as any, engineVersion: "1.0.0" });
  try {
    await manager.install("test.startup-fallback-ext");
    await manager.loadInstalled("test.startup-fallback-ext");
    expect(startupCalls).toBe(1);
  } finally {
    await manager.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("loadInstalled can call host.startup() even if the manager already considers startup complete (no other extensions)", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-startup-fallback-repeat-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "startup-fallback-repeat-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onStartupFinished"],
    contributes: {}
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), `export async function activate() {}\n`);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.startup-fallback-repeat-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const loaded: Array<{ id: string }> = [];
  let startupCalls = 0;
  const host = {
    loadExtension: async ({ extensionId }: { extensionId: string }) => {
      loaded.push({ id: extensionId });
      return extensionId;
    },
    listExtensions: () => loaded,
    startup: async () => {
      startupCalls++;
    }
  };

  const manager = new WebExtensionManager({ marketplaceClient, host: host as any, engineVersion: "1.0.0" });
  // Simulate a host that ran `startup()` before any extensions were loaded. For older hosts (no
  // `startupExtension`), we still need to re-run startup when loading the first onStartupFinished
  // extension.
  (manager as any)._didHostStartup = true;

  try {
    await manager.install("test.startup-fallback-repeat-ext");
    await manager.loadInstalled("test.startup-fallback-repeat-ext");
    expect(startupCalls).toBe(1);
  } finally {
    await manager.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("loadAllInstalled can call host.startup() even if startup already ran with zero extensions (no startupExtension)", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-startup-all-fallback-repeat-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "startup-all-fallback-repeat-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onStartupFinished"],
    contributes: {}
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), `export async function activate() {}\n`);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.startup-all-fallback-repeat-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const loaded: Array<{ id: string }> = [];
  let startupCalls = 0;
  const host = {
    loadExtension: async ({ extensionId }: { extensionId: string }) => {
      loaded.push({ id: extensionId });
      return extensionId;
    },
    listExtensions: () => loaded,
    startup: async () => {
      startupCalls++;
    }
  };

  const manager = new WebExtensionManager({ marketplaceClient, host: host as any, engineVersion: "1.0.0" });

  try {
    // Simulate boot calling the helper before any extensions are installed.
    await manager.loadAllInstalled();
    expect(startupCalls).toBe(1);

    // After installing an extension, loadAllInstalled should still run startup again so the new
    // extension receives onStartupFinished + workbookOpened (even if the manager considers startup
    // complete).
    await manager.install("test.startup-all-fallback-repeat-ext");
    await manager.loadAllInstalled();
    expect(startupCalls).toBe(2);
  } finally {
    await manager.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("loadAllInstalled can re-run host.startup() when the host has preloaded inactive extensions (no startupExtension)", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-startup-all-fallback-preloaded-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "startup-all-fallback-preloaded-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onStartupFinished"],
    contributes: {}
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), `export async function activate() {}\n`);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.startup-all-fallback-preloaded-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  // Pretend the host already has a (built-in) extension loaded, but not started yet (active=false).
  const loaded: Array<{ id: string; active?: boolean }> = [{ id: "test.preloaded-ext", active: false }];
  let startupCalls = 0;
  const host = {
    loadExtension: async ({ extensionId }: { extensionId: string }) => {
      loaded.push({ id: extensionId, active: false });
      return extensionId;
    },
    listExtensions: () => loaded,
    startup: async () => {
      startupCalls++;
    }
  };

  const manager = new WebExtensionManager({ marketplaceClient, host: host as any, engineVersion: "1.0.0" });

  try {
    // Initial boot: no installed extensions, but host has preloaded inactive ones. We should still
    // call startup once so the host is initialized.
    await manager.loadAllInstalled();
    expect(startupCalls).toBe(1);

    // Installing a new onStartupFinished extension later should cause a second safe startup() call
    // since no pre-existing extensions are active and there is no startupExtension() API.
    await manager.install("test.startup-all-fallback-preloaded-ext");
    await manager.loadAllInstalled();
    expect(startupCalls).toBe(2);
  } finally {
    await manager.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("loadAllInstalled does not call host.startup() when the host is already running (avoids workbookOpened spam)", async () => {
  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });

  const manager = new WebExtensionManager({ host, engineVersion: "1.0.0" });

  try {
    const extensionId = "test.started-host";
    const source = `const api = globalThis[Symbol.for(\"formula.extensionApi.api\")];\nif (!api) { throw new Error(\"missing formula api\"); }\nexport async function activate() {\n  api.events.onWorkbookOpened(() => {\n    void (async () => {\n      const prev = await api.storage.get(\"started.count\");\n      const next = (typeof prev === \"number\" ? prev : 0) + 1;\n      await api.storage.set(\"started.count\", next);\n    })().catch(() => {});\n  });\n}\n`;
    const mainUrl = `data:text/javascript,${encodeURIComponent(source)}`;

    await host.loadExtension({
      extensionId,
      extensionPath: "memory://started-host/",
      manifest: {
        name: "started-host",
        publisher: "test",
        version: "1.0.0",
        main: "./dist/extension.mjs",
        browser: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: ["onStartupFinished"],
        permissions: ["storage"],
        contributes: {}
      },
      mainUrl
    });

    // Start the host before calling into WebExtensionManager.
    await host.startup();

    const store = (host as any)._storageApi.getExtensionStore(extensionId);
    await waitFor(() => store["started.count"] === 1);
    expect(store["started.count"]).toBe(1);
    expect(host.listExtensions().find((e: any) => e.id === extensionId)?.active).toBe(true);

    // Should *not* call host.startup() again (would re-broadcast workbookOpened).
    await manager.loadAllInstalled();
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(store["started.count"]).toBe(1);
  } finally {
    await manager.dispose();
    await host.dispose();
  }
});

test("loadAllInstalled starts newly installed onStartupFinished extensions when the host is already running (no workbookOpened spam)", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-startup-all-running-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const installedId = "test.startup-all-running-ext";
  const installedVersion = "1.0.0";

  const manifest = {
    name: "startup-all-running-ext",
    publisher: "test",
    version: installedVersion,
    main: "./dist/extension.js",
    browser: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onStartupFinished"],
    permissions: ["storage"],
    contributes: {}
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(
    path.join(extDir, "dist", "extension.js"),
    `import * as formula from "@formula/extension-api";\nexport async function activate() {\n  formula.events.onWorkbookOpened(() => {\n    void (async () => {\n      const prev = await formula.storage.get(\"installed.count\");\n      const next = (typeof prev === \"number\" ? prev : 0) + 1;\n      await formula.storage.set(\"installed.count\", next);\n    })().catch(() => {});\n  });\n}\n`
  );

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: installedId,
    latestVersion: installedVersion,
    publicKeyPem: keys.publicKeyPem,
    packages: { [installedVersion]: pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });

  const runningId = "test.running-host";
  const runningSource = `const api = globalThis[Symbol.for(\"formula.extensionApi.api\")];\nif (!api) { throw new Error(\"missing formula api\"); }\nexport async function activate() {\n  api.events.onWorkbookOpened(() => {\n    void (async () => {\n      const prev = await api.storage.get(\"running.count\");\n      const next = (typeof prev === \"number\" ? prev : 0) + 1;\n      await api.storage.set(\"running.count\", next);\n    })().catch(() => {});\n  });\n}\n`;
  const runningUrl = `data:text/javascript,${encodeURIComponent(runningSource)}`;
  await host.loadExtension({
    extensionId: runningId,
    extensionPath: "memory://running-host/",
    manifest: {
      name: "running-host",
      publisher: "test",
      version: "1.0.0",
      main: "./dist/extension.mjs",
      browser: "./dist/extension.mjs",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onStartupFinished"],
      permissions: ["storage"],
      contributes: {}
    },
    mainUrl: runningUrl
  });

  // Start the host (activates running-host + emits workbookOpened).
  await host.startup();
  const runningStore = (host as any)._storageApi.getExtensionStore(runningId);
  await waitFor(() => runningStore["running.count"] === 1);
  expect(runningStore["running.count"]).toBe(1);

  const manager = new WebExtensionManager({ marketplaceClient, host });
  try {
    await manager.install(installedId);
    await manager.loadAllInstalled();

    // Existing active extensions should not get an extra startup workbookOpened event.
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(runningStore["running.count"]).toBe(1);

    const installedStore = (host as any)._storageApi.getExtensionStore(installedId);
    await waitFor(() => installedStore["installed.count"] === 1);
    expect(installedStore["installed.count"]).toBe(1);
    expect(host.listExtensions().find((e: any) => e.id === installedId)?.active).toBe(true);
  } finally {
    await manager.dispose();
    await host.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("loadInstalled does not call host.startup() when startupExtension is unavailable and other extensions exist", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-startup-fallback-existing-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "startup-fallback-existing-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onStartupFinished"],
    contributes: {}
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), `export async function activate() {}\n`);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.startup-fallback-existing-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const loaded: Array<{ id: string }> = [{ id: "test.preloaded-ext" }];
  let startupCalls = 0;
  const host = {
    loadExtension: async ({ extensionId }: { extensionId: string }) => {
      loaded.push({ id: extensionId });
      return extensionId;
    },
    listExtensions: () => loaded,
    startup: async () => {
      startupCalls++;
    }
  };

  const manager = new WebExtensionManager({
    marketplaceClient,
    host: host as any,
    engineVersion: "1.0.0"
  });
  try {
    await manager.install("test.startup-fallback-existing-ext");
    await manager.loadInstalled("test.startup-fallback-existing-ext");
    expect(startupCalls).toBe(0);
  } finally {
    await manager.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("loadInstalled calls host.startup() when existing extensions are loaded but none are active (startupExtension unavailable)", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-startup-fallback-inactive-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "startup-fallback-inactive-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onStartupFinished"],
    contributes: {}
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), `export async function activate() {}\n`);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.startup-fallback-inactive-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  // Pretend the host already has a (built-in) extension loaded but not started.
  const loaded: Array<{ id: string; active?: boolean }> = [{ id: "test.preloaded-ext", active: false }];
  let startupCalls = 0;
  const host = {
    loadExtension: async ({ extensionId }: { extensionId: string }) => {
      loaded.push({ id: extensionId, active: false });
      return extensionId;
    },
    listExtensions: () => loaded,
    startup: async () => {
      startupCalls++;
    }
  };

  const manager = new WebExtensionManager({
    marketplaceClient,
    host: host as any,
    engineVersion: "1.0.0"
  });
  try {
    await manager.install("test.startup-fallback-inactive-ext");
    await manager.loadInstalled("test.startup-fallback-inactive-ext");
    expect(startupCalls).toBe(1);
  } finally {
    await manager.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("loadAllInstalled triggers onStartupFinished + workbookOpened and is idempotent", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-startup-all-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "startup-all-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onStartupFinished"],
    permissions: ["storage"],
    contributes: {}
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(
    path.join(extDir, "dist", "extension.js"),
    `import * as formula from "@formula/extension-api";\nexport async function activate() {\n  formula.events.onWorkbookOpened(() => {\n    void (async () => {\n      const prev = await formula.storage.get("startup.count");\n      const next = (typeof prev === "number" ? prev : 0) + 1;\n      await formula.storage.set("startup.count", next);\n    })().catch(() => {});\n  });\n}\n`
  );

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.startup-all-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host });

  try {
    await manager.install("test.startup-all-ext");
    await manager.loadAllInstalled();

    const store = (host as any)._storageApi.getExtensionStore("test.startup-all-ext");
    await waitFor(() => store["startup.count"] === 1);
    expect(store["startup.count"]).toBe(1);

    // Calling loadAllInstalled again should be a no-op and should not re-broadcast workbookOpened.
    await manager.loadAllInstalled();
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(store["startup.count"]).toBe(1);
  } finally {
    await manager.dispose();
    await host.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("loadInstalled then loadAllInstalled does not duplicate the startup workbookOpened event", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-startup-installed-then-all-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "startup-installed-then-all",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onStartupFinished"],
    permissions: ["storage"],
    contributes: {}
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(
    path.join(extDir, "dist", "extension.js"),
    `import * as formula from "@formula/extension-api";\nexport async function activate() {\n  formula.events.onWorkbookOpened(() => {\n    void (async () => {\n      const prev = await formula.storage.get("startup.count");\n      const next = (typeof prev === "number" ? prev : 0) + 1;\n      await formula.storage.set("startup.count", next);\n    })().catch(() => {});\n  });\n}\n`
  );

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.startup-installed-then-all",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host });

  try {
    await manager.install("test.startup-installed-then-all");

    // First, load a single extension.
    await manager.loadInstalled("test.startup-installed-then-all");
    const store = (host as any)._storageApi.getExtensionStore("test.startup-installed-then-all");
    await waitFor(() => store["startup.count"] === 1);
    expect(store["startup.count"]).toBe(1);

    // Then, call the boot helper. It should not re-broadcast the startup workbookOpened event
    // to extensions that were already started via startupExtension().
    await manager.loadAllInstalled();
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(store["startup.count"]).toBe(1);
  } finally {
    await manager.dispose();
    await host.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("loadAllInstalled starts onStartupFinished extensions already loaded in the host even if startup was skipped", async () => {
  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ host, engineVersion: "1.0.0" });

  try {
    const extensionId = "test.manual-startup";
    const source = `const api = globalThis[Symbol.for(\"formula.extensionApi.api\")];\nif (!api) { throw new Error(\"missing formula api\"); }\nexport async function activate() {\n  api.events.onWorkbookOpened(() => {\n    api.storage.set(\"manual.startup\", \"ok\").catch(() => {});\n  });\n}\n`;
    const mainUrl = `data:text/javascript,${encodeURIComponent(source)}`;
    await host.loadExtension({
      extensionId,
      extensionPath: "memory://manual-startup/",
      manifest: {
        name: "manual-startup",
        publisher: "test",
        version: "1.0.0",
        main: "./dist/extension.mjs",
        browser: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: ["onStartupFinished"],
        permissions: ["storage"],
        contributes: {}
      },
      mainUrl
    });

    // Simulate an environment where something else already started the host (or we want to avoid
    // calling host.startup() because it would re-broadcast workbookOpened).
    (manager as any)._didHostStartup = true;

    await manager.loadAllInstalled();

    const store = (host as any)._storageApi.getExtensionStore(extensionId);
    await waitFor(() => store["manual.startup"] === "ok");
    expect(store["manual.startup"]).toBe("ok");
    expect(host.listExtensions().find((e: any) => e.id === extensionId)?.active).toBe(true);
  } finally {
    await manager.dispose();
    await host.dispose();
  }
});

test("update reloads onStartupFinished extension and re-delivers workbookOpened", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-startup-update-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const writeVersion = async (version: string, marker: string) => {
    const manifest = {
      name: "startup-ext",
      publisher: "test",
      version,
      main: "./dist/extension.js",
      browser: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onStartupFinished"],
      permissions: ["storage"],
      contributes: {}
    };
    await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
    await fs.writeFile(
      path.join(extDir, "dist", "extension.js"),
      `import * as formula from "@formula/extension-api";\nexport async function activate() {\n  formula.events.onWorkbookOpened(() => {\n    formula.storage.set("startup.marker", ${JSON.stringify(marker)}).catch(() => {});\n  });\n}\n`
    );
    return createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  };

  const pkgV1 = await writeVersion("1.0.0", "v1");
  const pkgV2 = await writeVersion("1.0.1", "v2");

  const marketplaceClient = createMockMarketplace({
    extensionId: "test.startup-ext",
    latestVersion: "1.0.1",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgV1, "1.0.1": pkgV2 }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host });

  try {
    await manager.install("test.startup-ext", "1.0.0");
    await manager.loadInstalled("test.startup-ext");

    const store = (host as any)._storageApi.getExtensionStore("test.startup-ext");
    await waitFor(() => store["startup.marker"] === "v1");
    expect(store["startup.marker"]).toBe("v1");

    await manager.update("test.startup-ext");
    expect(await manager.getInstalled("test.startup-ext")).toMatchObject({ version: "1.0.1" });

    await waitFor(() => store["startup.marker"] === "v2");
    expect(store["startup.marker"]).toBe("v2");
  } finally {
    await manager.dispose();
    await host.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("tampered package fails verification and is not installed", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const tampered = Buffer.from(pkgBytes);
  tampered[100] ^= 0x01;

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": tampered }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });

  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await expect(manager.install("formula.sample-hello")).rejects.toThrow(/verification failed/i);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await host.dispose();
});

test("install fails when engines.formula is incompatible with engineVersion", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-engine-mismatch-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "engine-mismatch-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.mjs",
    engines: { formula: "^2.0.0" }
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  const source = `export async function activate() {}\n`;
  await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.engine-mismatch-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const manager = new WebExtensionManager({ marketplaceClient, engineVersion: "1.0.0" });

  await expect(manager.install("test.engine-mismatch-ext")).rejects.toThrow(/engine mismatch/i);
  expect(await manager.listInstalled()).toEqual([]);

  // Ensure we didn't persist any package bytes.
  const db = await requestToPromise(indexedDB.open("formula.webExtensions"));
  try {
    const tx = db.transaction(["packages"], "readonly");
    const store = tx.objectStore("packages");
    const records = (await requestToPromise(store.getAll())) as any[];
    await txDone(tx);
    expect(records).toEqual([]);
  } finally {
    db.close();
  }

  await manager.dispose();
  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("install fails when publisher/name contain path separators (invalid extension id)", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-invalid-id-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "bad-ext",
    publisher: "evil/publisher",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.mjs",
    engines: { formula: "^1.0.0" }
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  const source = `export async function activate() {}\n`;
  await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "evil/publisher.bad-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const manager = new WebExtensionManager({ marketplaceClient, engineVersion: "1.0.0" });

  await expect(manager.install("evil/publisher.bad-ext")).rejects.toThrow(/invalid extension id/i);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("loadInstalled quarantines when a stored manifest becomes incompatible with engineVersion", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-quarantine-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "quarantine-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.mjs",
    engines: { formula: "^1.0.0" }
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  const source = `export async function activate() {}\n`;
  await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.quarantine-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const installer = new WebExtensionManager({ marketplaceClient, engineVersion: "1.0.0" });
  await installer.install("test.quarantine-ext");
  await installer.dispose();

  const host = new BrowserExtensionHost({
    engineVersion: "2.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "2.0.0" });

  await expect(manager.loadInstalled("test.quarantine-ext")).rejects.toThrow(/engine mismatch/i);
  const stored = await manager.getInstalled("test.quarantine-ext");
  expect(stored?.incompatible).toBe(true);
  expect(String(stored?.incompatibleReason ?? "")).toMatch(/engine mismatch/i);
  expect(manager.isLoaded("test.quarantine-ext")).toBe(false);
  expect(host.listExtensions().some((e: any) => e.id === "test.quarantine-ext")).toBe(false);

  await manager.dispose();
  await host.dispose();
  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("verifyInstalled clears incompatible flag when the manifest becomes compatible again", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-unquarantine-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "unquarantine-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.mjs",
    engines: { formula: "^1.0.0" }
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  const source = `export async function activate() {}\n`;
  await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.unquarantine-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const installer = new WebExtensionManager({ marketplaceClient, engineVersion: "1.0.0" });
  await installer.install("test.unquarantine-ext");
  await installer.dispose();

  // Mark the install incompatible under a different engine version.
  const managerV2 = new WebExtensionManager({ marketplaceClient, engineVersion: "2.0.0" });
  const firstCheck = await managerV2.verifyInstalled("test.unquarantine-ext");
  expect(firstCheck.ok).toBe(false);
  expect(String(firstCheck.reason ?? "")).toMatch(/engine mismatch/i);
  const storedAfterV2 = await managerV2.getInstalled("test.unquarantine-ext");
  expect(storedAfterV2?.incompatible).toBe(true);
  await managerV2.dispose();

  // Switching back to a compatible engine should clear the quarantine flag via verifyInstalled().
  const managerV1 = new WebExtensionManager({ marketplaceClient, engineVersion: "1.0.0" });
  const secondCheck = await managerV1.verifyInstalled("test.unquarantine-ext");
  expect(secondCheck.ok).toBe(true);
  const storedAfterV1 = await managerV1.getInstalled("test.unquarantine-ext");
  expect(storedAfterV1?.incompatible).not.toBe(true);
  expect(storedAfterV1?.incompatibleAt).toBeUndefined();
  expect(storedAfterV1?.incompatibleReason).toBeUndefined();
  await managerV1.dispose();

  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("loadInstalled clears incompatible markers even when the incompatible flag is missing", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-load-unquarantine-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "load-unquarantine-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.mjs",
    engines: { formula: "^1.0.0" }
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  const source = `export async function activate() {}\n`;
  await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.load-unquarantine-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const installer = new WebExtensionManager({ marketplaceClient, engineVersion: "1.0.0" });
  await installer.install("test.load-unquarantine-ext");
  await installer.dispose();

  // Mark incompatible under a different engine version.
  const managerV2 = new WebExtensionManager({ marketplaceClient, engineVersion: "2.0.0" });
  const verificationV2 = await managerV2.verifyInstalled("test.load-unquarantine-ext");
  expect(verificationV2.ok).toBe(false);
  expect(String(verificationV2.reason ?? "")).toMatch(/engine mismatch/i);
  const storedAfterV2 = await managerV2.getInstalled("test.load-unquarantine-ext");
  expect(storedAfterV2?.incompatible).toBe(true);
  expect(String(storedAfterV2?.incompatibleReason ?? "")).toMatch(/engine mismatch/i);
  await managerV2.dispose();

  // Simulate a legacy/partial record where `incompatibleAt`/`incompatibleReason` were persisted
  // without the boolean `incompatible` flag.
  const db = await requestToPromise(indexedDB.open("formula.webExtensions"));
  try {
    const tx = db.transaction(["installed"], "readwrite");
    const store = tx.objectStore("installed");
    const record: any = await requestToPromise(store.get("test.load-unquarantine-ext"));
    expect(record).toBeTruthy();
    expect(record.incompatibleAt).toBeTruthy();
    expect(record.incompatibleReason).toBeTruthy();
    delete record.incompatible;
    store.put(record);
    await txDone(tx);
  } finally {
    db.close();
  }

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const managerV1 = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await managerV1.loadInstalled("test.load-unquarantine-ext");
  const storedAfterLoad = await managerV1.getInstalled("test.load-unquarantine-ext");
  expect(storedAfterLoad?.incompatible).not.toBe(true);
  expect(storedAfterLoad?.incompatibleAt).toBeUndefined();
  expect(storedAfterLoad?.incompatibleReason).toBeUndefined();
  expect(managerV1.isLoaded("test.load-unquarantine-ext")).toBe(true);
  expect(host.listExtensions().some((e: any) => e.id === "test.load-unquarantine-ext")).toBe(true);

  await managerV1.dispose();
  await host.dispose();
  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("verifyInstalled quarantines and unloads an extension when its stored manifest becomes incompatible", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-incompatible-unload-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "incompatible-unload-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.mjs",
    engines: { formula: "^1.0.0" }
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  const source = `export async function activate() {}\n`;
  await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.incompatible-unload-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await manager.install("test.incompatible-unload-ext");
  await manager.loadInstalled("test.incompatible-unload-ext");
  expect(manager.isLoaded("test.incompatible-unload-ext")).toBe(true);
  expect(host.listExtensions().some((e: any) => e.id === "test.incompatible-unload-ext")).toBe(true);

  // Tamper the stored manifest to make it incompatible with the current engine.
  const db = await requestToPromise(indexedDB.open("formula.webExtensions"));
  try {
    const tx = db.transaction(["packages"], "readwrite");
    const store = tx.objectStore("packages");
    const record: any = await requestToPromise(store.get("test.incompatible-unload-ext@1.0.0"));
    expect(record).toBeTruthy();
    record.verified = {
      ...record.verified,
      manifest: {
        ...record.verified.manifest,
        engines: { formula: "^2.0.0" }
      }
    };
    store.put(record);
    await txDone(tx);
  } finally {
    db.close();
  }

  const verification = await manager.verifyInstalled("test.incompatible-unload-ext");
  expect(verification.ok).toBe(false);
  expect(String(verification.reason ?? "")).toMatch(/engine mismatch/i);

  const stored = await manager.getInstalled("test.incompatible-unload-ext");
  expect(stored?.incompatible).toBe(true);
  expect(String(stored?.incompatibleReason ?? "")).toMatch(/engine mismatch/i);

  // The incompatible quarantine should proactively unload the running extension.
  expect(manager.isLoaded("test.incompatible-unload-ext")).toBe(false);
  expect(host.listExtensions().some((e: any) => e.id === "test.incompatible-unload-ext")).toBe(false);

  await manager.dispose();
  await host.dispose();
  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("update flow replaces installed version and reloads when loaded", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-update-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const writeVersion = async (version: string, marker: string) => {
    const manifest = {
      name: "test-ext",
      publisher: "test",
      version,
      main: "./dist/extension.js",
      browser: "./dist/extension.mjs",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCommand:test.version"],
      permissions: ["ui.commands"],
      contributes: { commands: [{ command: "test.version", title: "Version" }] }
    };
    await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
    const source = `import * as formula from "@formula/extension-api";\nexport async function activate() {\n  await formula.commands.registerCommand("test.version", () => "${marker}");\n}\n`;
    await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
    await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);
    return createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  };

  const pkgV1 = await writeVersion("1.0.0", "v1");
  const pkgV2 = await writeVersion("1.0.1", "v2");

  const marketplaceClient = createMockMarketplace({
    extensionId: "test.test-ext",
    latestVersion: "1.0.1",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgV1, "1.0.1": pkgV2 }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await manager.install("test.test-ext", "1.0.0");
  await manager.loadInstalled("test.test-ext");
  expect(await host.executeCommand("test.version")).toBe("v1");

  await manager.update("test.test-ext");
  expect(await manager.getInstalled("test.test-ext")).toMatchObject({ version: "1.0.1" });
  expect(await host.executeCommand("test.version")).toBe("v2");

  await manager.dispose();
  await host.dispose();
  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("install supports publisher signing key rotation via publisherKeys + per-version key id", async () => {
  const keysA = generateEd25519KeyPair();
  const keysB = generateEd25519KeyPair();

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-key-rotation-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const writeVersion = async (version: string, marker: string, privateKeyPem: string) => {
    const manifest = {
      name: "test-ext",
      publisher: "test",
      version,
      main: "./dist/extension.js",
      browser: "./dist/extension.mjs",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCommand:test.version"],
      permissions: ["ui.commands"],
      contributes: { commands: [{ command: "test.version", title: "Version" }] }
    };
    await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
    const source = `import * as formula from "@formula/extension-api";\nexport async function activate() {\n  await formula.commands.registerCommand("test.version", () => "${marker}");\n}\n`;
    await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
    await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);
    return createExtensionPackageV2(extDir, { privateKeyPem });
  };

  const pkgV1 = await writeVersion("1.0.0", "v1", keysA.privateKeyPem);
  const pkgV2 = await writeVersion("1.0.1", "v2", keysB.privateKeyPem);

  // Simulate rotation: publisherPublicKeyPem points at the *new* key, but the publisherKeys
  // set contains both keys. The download response provides the per-version key id so clients
  // can prefer the correct key immediately.
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.test-ext",
    latestVersion: "1.0.1",
    publicKeyPem: keysB.publicKeyPem,
    publisherKeys: [
      { id: "keyB", publicKeyPem: keysB.publicKeyPem, revoked: false },
      { id: "keyA", publicKeyPem: keysA.publicKeyPem, revoked: false }
    ],
    publisherKeyIds: { "1.0.0": "keyA", "1.0.1": "keyB" },
    packages: { "1.0.0": pkgV1, "1.0.1": pkgV2 }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await manager.install("test.test-ext", "1.0.0");
  expect((await manager.getInstalled("test.test-ext"))?.signingKeyId).toBe("keyA");
  await manager.loadInstalled("test.test-ext");
  expect(await host.executeCommand("test.version")).toBe("v1");

  await manager.update("test.test-ext");
  expect(await manager.getInstalled("test.test-ext")).toMatchObject({ version: "1.0.1", signingKeyId: "keyB" });
  expect(await host.executeCommand("test.version")).toBe("v2");

  await manager.dispose();
  await host.dispose();
  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("install refuses when all publisher signing keys are revoked", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    publisherKeys: [{ id: "keyA", publicKeyPem: keys.publicKeyPem, revoked: true }],
    publisherKeyIds: { "1.0.0": "keyA" },
    packages: { "1.0.0": pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await expect(manager.install("formula.sample-hello", "1.0.0")).rejects.toThrow(/all publisher signing keys are revoked/i);

  await manager.dispose();
  await host.dispose();
});

test("install rejects conflicting contributed panel ids across installed extensions", async () => {
  const keys = generateEd25519KeyPair();

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-panel-conflict-"));
  const extDirA = path.join(tmpRoot, "ext-a");
  const extDirB = path.join(tmpRoot, "ext-b");

  const writeExtension = async (dir: string, { name, publisher, version }: { name: string; publisher: string; version: string }) => {
    await fs.mkdir(path.join(dir, "dist"), { recursive: true });
    const manifest = {
      name,
      publisher,
      version,
      main: "./dist/extension.js",
      browser: "./dist/extension.mjs",
      engines: { formula: "^1.0.0" },
      contributes: {
        panels: [{ id: "shared.panel", title: "Shared Panel" }]
      }
    };
    await fs.writeFile(path.join(dir, "package.json"), JSON.stringify(manifest, null, 2));
    const source = `export async function activate() {}\n`;
    await fs.writeFile(path.join(dir, "dist", "extension.mjs"), source);
    await fs.writeFile(path.join(dir, "dist", "extension.js"), source);
    return createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });
  };

  const pkgA = await writeExtension(extDirA, { publisher: "acme", name: "one", version: "1.0.0" });
  const pkgB = await writeExtension(extDirB, { publisher: "acme", name: "two", version: "1.0.0" });

  const packages: Record<string, Record<string, ArrayBuffer>> = {
    "acme.one": { "1.0.0": pkgA },
    "acme.two": { "1.0.0": pkgB },
  };

  const marketplaceClient = {
    async getExtension(id: string) {
      if (!packages[id]) return null;
      return {
        id,
        name: id.split(".")[1],
        displayName: id,
        publisher: id.split(".")[0],
        description: "",
        latestVersion: "1.0.0",
        verified: true,
        featured: false,
        categories: [],
        tags: [],
        screenshots: [],
        downloadCount: 0,
        updatedAt: new Date().toISOString(),
        versions: [],
        readme: "",
        publisherPublicKeyPem: keys.publicKeyPem,
        createdAt: new Date().toISOString(),
        deprecated: false,
        blocked: false,
        malicious: false
      };
    },
    async downloadPackage(id: string, version: string) {
      const bytes = packages[id]?.[version];
      if (!bytes) return null;
      return {
        bytes: new Uint8Array(bytes),
        signatureBase64: null,
        sha256: null,
        formatVersion: 2,
        publisher: id.split(".")[0],
        publisherKeyId: null,
        scanStatus: "passed",
        filesSha256: null
      };
    }
  };

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await manager.install("acme.one");
  expect(await manager.listInstalled()).toEqual([
    expect.objectContaining({ id: "acme.one", version: "1.0.0" })
  ]);

  // Installing a second extension that claims the same panel id should be rejected, and should
  // not corrupt the installed list or the seed store.
  await expect(manager.install("acme.two")).rejects.toThrow(/panel id already contributed/i);

  const installedAfter = await manager.listInstalled();
  expect(installedAfter.map((r) => r.id).sort()).toEqual(["acme.one"]);

  const seedRaw = globalThis.localStorage.getItem("formula.extensions.contributedPanels.v1");
  expect(seedRaw).not.toBeNull();
  const seed = JSON.parse(String(seedRaw));
  expect(seed["shared.panel"]).toMatchObject({ extensionId: "acme.one" });

  await manager.dispose();
  await host.dispose();
  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("install fails for packages signed by a revoked key even if other keys exist", async () => {
  const keysA = generateEd25519KeyPair();
  const keysB = generateEd25519KeyPair();

  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keysA.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    // Simulate rotation where publisherPublicKeyPem is the *new* key, but the old key is revoked.
    publicKeyPem: keysB.publicKeyPem,
    publisherKeys: [
      { id: "keyB", publicKeyPem: keysB.publicKeyPem, revoked: false },
      { id: "keyA", publicKeyPem: keysA.publicKeyPem, revoked: true }
    ],
    publisherKeyIds: { "1.0.0": "keyA" },
    packages: { "1.0.0": pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await expect(manager.install("formula.sample-hello", "1.0.0")).rejects.toThrow(/signature verification failed/i);

  await manager.dispose();
  await host.dispose();
});

test("supports importing extension API via the \"formula\" alias", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-alias-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "alias-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.mjs",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onCommand:test.alias"],
    permissions: ["ui.commands"],
    contributes: { commands: [{ command: "test.alias", title: "Alias" }] }
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  const source = `import * as formula from "formula";\nexport async function activate() {\n  await formula.commands.registerCommand("test.alias", () => "ok");\n}\n`;
  await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "test.alias-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });

  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await manager.install("test.alias-ext");
  await manager.loadInstalled("test.alias-ext");
  expect(await host.executeCommand("test.alias")).toBe("ok");

  await manager.dispose();
  await host.dispose();
  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("rejects browser entrypoints that contain dynamic import()", async () => {
  const keys = generateEd25519KeyPair();
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-dynamic-import-"));
  const extDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

  const manifest = {
    name: "dynamic-import-ext",
    publisher: "test",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.mjs",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onCommand:test.dynamic"],
    permissions: ["ui.commands"],
    contributes: { commands: [{ command: "test.dynamic", title: "Dynamic" }] }
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
  const source = `import * as formula from "@formula/extension-api";\nexport async function activate() {\n  await import("data:text/javascript,export default 123");\n  await formula.commands.registerCommand("test.dynamic", () => "ok");\n}\n`;
  await fs.writeFile(path.join(extDir, "dist", "extension.mjs"), source);
  await fs.writeFile(path.join(extDir, "dist", "extension.js"), source);

  const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
  const marketplaceClient = createMockMarketplace({
    extensionId: "test.dynamic-import-ext",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await manager.install("test.dynamic-import-ext");
  await manager.loadInstalled("test.dynamic-import-ext");
  await expect(host.executeCommand("test.dynamic")).rejects.toThrow(/dynamic import/i);
  expect(manager.isLoaded("test.dynamic-import-ext")).toBe(true);
  const loaded = host.listExtensions().find((e: any) => e.id === "test.dynamic-import-ext");
  expect(Boolean(loaded)).toBe(true);
  expect(loaded?.active).toBe(false);

  await manager.dispose();
  await host.dispose();
  await fs.rm(tmpRoot, { recursive: true, force: true });
});

test("uninstall clears persisted permission + extension storage state (localStorage)", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await manager.install("formula.sample-hello");

  // Simulate previously-granted permissions + persisted extension storage.
  globalThis.localStorage.setItem(
    "formula.extensionHost.permissions",
    JSON.stringify({
      "formula.sample-hello": { storage: true, network: { mode: "full" } },
      "other.extension": { "ui.commands": true },
    })
  );
  globalThis.localStorage.setItem(
    "formula.extensionHost.storage.formula.sample-hello",
    JSON.stringify({ foo: "bar" })
  );
  globalThis.localStorage.setItem(
    "formula.extensions.contributedPanels.v1",
    JSON.stringify({
      "sampleHello.panel": { extensionId: "formula.sample-hello", title: "Sample Hello Panel" },
      "other.panel": { extensionId: "other.extension", title: "Other Panel" }
    })
  );

  // Ensure we preserve other extensions' panel seeds while removing the uninstalled extension.
  const seedKey = "formula.extensions.contributedPanels.v1";
  const seedRaw = globalThis.localStorage.getItem(seedKey);
  const seed = seedRaw ? JSON.parse(String(seedRaw)) : {};
  seed["other.panel"] = { extensionId: "other.extension", title: "Other Panel" };
  globalThis.localStorage.setItem(seedKey, JSON.stringify(seed));

  // Ensure the host has the extension store cached in-memory; uninstall should clear it too.
  const storeBefore = (host as any)._storageApi.getExtensionStore("formula.sample-hello");
  expect(storeBefore.foo).toBe("bar");

  await manager.uninstall("formula.sample-hello");

  expect(globalThis.localStorage.getItem("formula.extensionHost.storage.formula.sample-hello")).toBe(null);
  const storeAfter = (host as any)._storageApi.getExtensionStore("formula.sample-hello");
  expect(storeAfter).not.toBe(storeBefore);
  expect(storeAfter.foo).toBeUndefined();

  const panelSeedsRaw = globalThis.localStorage.getItem("formula.extensions.contributedPanels.v1");
  expect(panelSeedsRaw).not.toBe(null);
  const panelSeeds = JSON.parse(String(panelSeedsRaw));
  expect(panelSeeds["sampleHello.panel"]).toBeUndefined();
  expect(panelSeeds["other.panel"]).toEqual({ extensionId: "other.extension", title: "Other Panel" });

  const permissionsRaw = globalThis.localStorage.getItem("formula.extensionHost.permissions");
  expect(permissionsRaw).not.toBe(null);
  const permissions = JSON.parse(String(permissionsRaw));
  expect(permissions["formula.sample-hello"]).toBeUndefined();
  expect(permissions["other.extension"]).toEqual({ "ui.commands": true });

  const seedAfterRaw = globalThis.localStorage.getItem(seedKey);
  expect(seedAfterRaw).not.toBeNull();
  const seedAfter = JSON.parse(String(seedAfterRaw));
  expect(seedAfter["sampleHello.panel"]).toBeUndefined();
  expect(seedAfter["other.panel"]).toMatchObject({ extensionId: "other.extension", title: "Other Panel" });

  await manager.dispose();
  await host.dispose();
});

test("uninstall clears persisted permissions even when host is absent", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const manager = new WebExtensionManager({ marketplaceClient, host: null, engineVersion: "1.0.0" });

  await manager.install("formula.sample-hello");

  globalThis.localStorage.setItem(
    "formula.extensionHost.permissions",
    JSON.stringify({
      "formula.sample-hello": { storage: true },
      "other.extension": { "ui.commands": true }
    })
  );
  globalThis.localStorage.setItem(
    "formula.extensionHost.storage.formula.sample-hello",
    JSON.stringify({ foo: "bar" })
  );

  await manager.uninstall("formula.sample-hello");

  const permissionsRaw = globalThis.localStorage.getItem("formula.extensionHost.permissions");
  expect(permissionsRaw).not.toBe(null);
  const permissions = JSON.parse(String(permissionsRaw));
  expect(permissions["formula.sample-hello"]).toBeUndefined();
  expect(permissions["other.extension"]).toEqual({ "ui.commands": true });

  expect(globalThis.localStorage.getItem("formula.extensionHost.storage.formula.sample-hello")).toBe(null);

  await manager.dispose();
});

test("uninstall removes corrupted permissions store key even when host is absent", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
  });

  const manager = new WebExtensionManager({ marketplaceClient, host: null, engineVersion: "1.0.0" });

  await manager.install("formula.sample-hello");

  // Simulate a corrupted permissions store value (invalid JSON).
  globalThis.localStorage.setItem("formula.extensionHost.permissions", "{not-json");
  expect(globalThis.localStorage.getItem("formula.extensionHost.permissions")).toBe("{not-json");

  await manager.uninstall("formula.sample-hello");

  // Uninstall should clear the corrupted permissions key so future installs start clean.
  expect(globalThis.localStorage.getItem("formula.extensionHost.permissions")).toBe(null);

  await manager.dispose();
});

test("uninstall removes permissions store key when the last extension is removed", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
  });

  const manager = new WebExtensionManager({ marketplaceClient, host: null, engineVersion: "1.0.0" });
  await manager.install("formula.sample-hello");

  globalThis.localStorage.setItem(
    "formula.extensionHost.permissions",
    JSON.stringify({
      "formula.sample-hello": { storage: true },
    }),
  );

  await manager.uninstall("formula.sample-hello");

  expect(globalThis.localStorage.getItem("formula.extensionHost.permissions")).toBe(null);

  await manager.dispose();
});

test("uninstall removes permissions store key when the last extension is removed (host present)", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });
  await manager.install("formula.sample-hello");

  // Simulate the only extension having stored permission grants.
  globalThis.localStorage.setItem(
    "formula.extensionHost.permissions",
    JSON.stringify({
      "formula.sample-hello": { storage: true },
    }),
  );

  await manager.uninstall("formula.sample-hello");

  // The browser host's permission manager persists `{}` when the last extension is removed;
  // WebExtensionManager should clear the key entirely on uninstall so reinstall behaves like a
  // clean slate.
  expect(globalThis.localStorage.getItem("formula.extensionHost.permissions")).toBe(null);

  await manager.dispose();
  await host.dispose();
});

test("uninstall removes contributed panel seed store key when empty", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
  });

  const manager = new WebExtensionManager({ marketplaceClient, host: null, engineVersion: "1.0.0" });

  await manager.install("formula.sample-hello");

  expect(globalThis.localStorage.getItem("formula.extensions.contributedPanels.v1")).not.toBeNull();

  await manager.uninstall("formula.sample-hello");

  expect(globalThis.localStorage.getItem("formula.extensions.contributedPanels.v1")).toBeNull();

  await manager.dispose();
});

test("uninstall removes contributed panel seed store key when it is already empty", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
  });

  const manager = new WebExtensionManager({ marketplaceClient, host: null, engineVersion: "1.0.0" });

  await manager.install("formula.sample-hello");

  // Simulate an older client leaving behind an empty `"{}"` seed store record.
  const seedKey = "formula.extensions.contributedPanels.v1";
  globalThis.localStorage.setItem(seedKey, "{}");
  expect(globalThis.localStorage.getItem(seedKey)).toBe("{}");

  await manager.uninstall("formula.sample-hello");

  expect(globalThis.localStorage.getItem(seedKey)).toBeNull();

  await manager.dispose();
});

test("uninstall removes contributed panel seed store key when it is corrupted", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
  });

  const manager = new WebExtensionManager({ marketplaceClient, host: null, engineVersion: "1.0.0" });

  await manager.install("formula.sample-hello");

  const seedKey = "formula.extensions.contributedPanels.v1";
  globalThis.localStorage.setItem(seedKey, "{not-json");
  expect(globalThis.localStorage.getItem(seedKey)).toBe("{not-json");

  await manager.uninstall("formula.sample-hello");

  expect(globalThis.localStorage.getItem(seedKey)).toBeNull();

  await manager.dispose();
});

test("uninstall removes all IndexedDB package records for an extension id", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
  });

  const manager = new WebExtensionManager({ marketplaceClient, host: null, engineVersion: "1.0.0" });

  await manager.install("formula.sample-hello");

  // Inject an extra package record (simulating an older/orphaned version left behind after a crash).
  const db = await requestToPromise(indexedDB.open("formula.webExtensions"));
  try {
    const tx = db.transaction(["packages"], "readwrite");
    const store = tx.objectStore("packages");
    const primaryKey = "formula.sample-hello@1.0.0";
    const existing = await requestToPromise<any>(store.get(primaryKey));
    expect(existing).toBeTruthy();
    store.put({ ...existing, key: "formula.sample-hello@0.9.0", version: "0.9.0" });
    await txDone(tx);
  } finally {
    db.close();
  }

  const dbBefore = await requestToPromise(indexedDB.open("formula.webExtensions"));
  try {
    const tx = dbBefore.transaction(["packages"], "readonly");
    const index = tx.objectStore("packages").index("byId");
    const count = await requestToPromise(index.count("formula.sample-hello"));
    expect(count).toBe(2);
    await txDone(tx);
  } finally {
    dbBefore.close();
  }

  await manager.uninstall("formula.sample-hello");

  const dbAfter = await requestToPromise(indexedDB.open("formula.webExtensions"));
  try {
    const tx = dbAfter.transaction(["packages"], "readonly");
    const store = tx.objectStore("packages");
    const index = store.index("byId");
    const count = await requestToPromise(index.count("formula.sample-hello"));
    expect(count).toBe(0);
    expect(await requestToPromise(store.get("formula.sample-hello@1.0.0"))).toBeUndefined();
    expect(await requestToPromise(store.get("formula.sample-hello@0.9.0"))).toBeUndefined();
    await txDone(tx);
  } finally {
    dbAfter.close();
  }

  await manager.dispose();
});

test("detects IndexedDB corruption on load and supports repair()", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const spreadsheet = new TestSpreadsheetApi();
  spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
  await spreadsheet.setCell(0, 0, 1);
  await spreadsheet.setCell(0, 1, 2);
  await spreadsheet.setCell(1, 0, 3);
  await spreadsheet.setCell(1, 1, 4);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: spreadsheet,
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await manager.install("formula.sample-hello");

  // Tamper the stored bytes in IndexedDB.
  const db = await requestToPromise(indexedDB.open("formula.webExtensions"));
  try {
    const tx = db.transaction(["packages"], "readwrite");
    const store = tx.objectStore("packages");
    const record: any = await requestToPromise(store.get("formula.sample-hello@1.0.0"));
    expect(record).toBeTruthy();

    const bytes = new Uint8Array(record.bytes as ArrayBuffer);
    bytes[0] ^= 0xff;
    store.put(record);
    await txDone(tx);
  } finally {
    db.close();
  }

  await expect(manager.loadInstalled("formula.sample-hello")).rejects.toThrow(/integrity|sha256/i);

  const corrupted = await manager.getInstalled("formula.sample-hello");
  expect(corrupted).toMatchObject({ id: "formula.sample-hello", version: "1.0.0", corrupted: true });
  expect(corrupted?.corruptedAt).toMatch(/^\d{4}-\d{2}-\d{2}T/);
  expect(corrupted?.corruptedReason).toMatch(/sha256|checksum|integrity/i);

  await manager.repair("formula.sample-hello");
  const repaired = await manager.getInstalled("formula.sample-hello");
  expect(repaired?.corrupted).not.toBe(true);

  await manager.loadInstalled("formula.sample-hello");
  expect(await host.executeCommand("sampleHello.sumSelection")).toBe(10);
  await manager.dispose();
  await host.dispose();
});

test("install refuses blocked extensions returned by the marketplace", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    blocked: true,
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await expect(manager.install("formula.sample-hello", "1.0.0")).rejects.toThrow(/blocked/i);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await host.dispose();
});

test("install refuses malicious extensions returned by the marketplace", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    malicious: true,
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await expect(manager.install("formula.sample-hello", "1.0.0")).rejects.toThrow(/malicious/i);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await host.dispose();
});

test("install refuses extensions whose publisher is revoked", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    publisherRevoked: true,
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await expect(manager.install("formula.sample-hello", "1.0.0")).rejects.toThrow(/revoked/i);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await host.dispose();
});

test("install refuses extensions when all publisher signing keys are revoked", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    publisherKeys: [
      {
        id: "key-1",
        publicKeyPem: keys.publicKeyPem,
        revoked: true,
      },
    ],
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await expect(manager.install("formula.sample-hello", "1.0.0")).rejects.toThrow(/revoked/i);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await host.dispose();
});

test("install warns when installing a deprecated extension", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    deprecated: true,
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  const result = await manager.install("formula.sample-hello", "1.0.0");
  expect(result.warnings?.some((w) => w.kind === "deprecated")).toBe(true);
  expect(await manager.getInstalled("formula.sample-hello")).toMatchObject({ version: "1.0.0" });

  await manager.dispose();
  await host.dispose();
});

test("install can be cancelled via confirm() when installing a deprecated extension", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    deprecated: true,
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  let confirmCalled = false;
  await expect(
    manager.install("formula.sample-hello", "1.0.0", {
      confirm: async (warning) => {
        confirmCalled = true;
        expect(warning.kind).toBe("deprecated");
        return false;
      },
    })
  ).rejects.toThrow(/cancel/i);
  expect(confirmCalled).toBe(true);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await host.dispose();
});

test("install uses per-version scanStatus metadata when download scanStatus header is missing", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    scanStatuses: { "1.0.0": null },
    extensionVersions: [{ version: "1.0.0", scanStatus: "passed" }],
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  const installed = await manager.install("formula.sample-hello", "1.0.0");
  expect(installed.scanStatus).toBe("passed");
  expect((await manager.getInstalled("formula.sample-hello"))?.scanStatus).toBe("passed");

  await manager.dispose();
  await host.dispose();
});

test("install refuses when scanStatus is missing by default", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    scanStatuses: { "1.0.0": null },
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await expect(manager.install("formula.sample-hello", "1.0.0")).rejects.toThrow(/scan status.*missing/i);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await host.dispose();
});

test("install enforces scanStatus when configured", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    scanStatuses: { "1.0.0": "pending" },
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await expect(manager.install("formula.sample-hello", "1.0.0", { scanPolicy: "enforce" })).rejects.toThrow(/scan status/i);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await host.dispose();
});

test("install can allow non-passed scanStatus in dev mode when configured", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    scanStatuses: { "1.0.0": "pending" },
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  const result = await manager.install("formula.sample-hello", "1.0.0", { scanPolicy: "allow" });
  expect(result.warnings?.some((w) => w.kind === "scanStatus" && w.scanStatus === "pending")).toBe(true);
  expect(await manager.getInstalled("formula.sample-hello")).toMatchObject({ version: "1.0.0" });

  await manager.dispose();
  await host.dispose();
});

test("install can be cancelled via confirm() when scanStatus is non-passed and policy=allow", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    scanStatuses: { "1.0.0": "pending" },
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  let confirmCalled = false;
  await expect(
    manager.install("formula.sample-hello", "1.0.0", {
      scanPolicy: "allow",
      confirm: async (warning) => {
        confirmCalled = true;
        expect(warning.kind).toBe("scanStatus");
        expect(warning.scanStatus).toBe("pending");
        return false;
      },
    })
  ).rejects.toThrow(/cancel/i);
  expect(confirmCalled).toBe(true);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await host.dispose();
});

test("scanStatus is enforced by default", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes },
    scanStatuses: { "1.0.0": "pending" },
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true,
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await expect(manager.install("formula.sample-hello", "1.0.0")).rejects.toThrow(/scan status/i);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await host.dispose();
});

test("missing stored package record is quarantined and can be repaired", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await manager.install("formula.sample-hello");

  // Delete the stored package record to simulate a partial write / IndexedDB corruption.
  const db = await requestToPromise(indexedDB.open("formula.webExtensions"));
  try {
    const tx = db.transaction(["packages"], "readwrite");
    tx.objectStore("packages").delete("formula.sample-hello@1.0.0");
    await txDone(tx);
  } finally {
    db.close();
  }

  await expect(manager.loadInstalled("formula.sample-hello")).rejects.toThrow(/missing stored package|integrity/i);
  expect(await manager.getInstalled("formula.sample-hello")).toMatchObject({ corrupted: true });

  await manager.repair("formula.sample-hello");
  const repaired = await manager.getInstalled("formula.sample-hello");
  expect(repaired?.corrupted).not.toBe(true);

  await manager.loadInstalled("formula.sample-hello");

  await manager.dispose();
  await host.dispose();
});

test("loadInstalled quarantines when stored bytes are not parseable as a v2 package", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    packages: { "1.0.0": pkgBytes }
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: new TestSpreadsheetApi(),
    permissionPrompt: async () => true
  });
  const manager = new WebExtensionManager({ marketplaceClient, host, engineVersion: "1.0.0" });

  await manager.install("formula.sample-hello");

  // Replace bytes with an invalid buffer but keep sha256 metadata consistent so we exercise the
  // "not parseable" path (rather than sha mismatch).
  const invalid = Buffer.from([0x01, 0x02, 0x03, 0x04]);
  const invalidSha256 = crypto.createHash("sha256").update(invalid).digest("hex");

  const db = await requestToPromise(indexedDB.open("formula.webExtensions"));
  try {
    const tx = db.transaction(["packages"], "readwrite");
    const store = tx.objectStore("packages");
    const record: any = await requestToPromise(store.get("formula.sample-hello@1.0.0"));
    record.bytes = invalid.buffer.slice(invalid.byteOffset, invalid.byteOffset + invalid.byteLength);
    record.packageSha256 = invalidSha256;
    store.put(record);
    await txDone(tx);
  } finally {
    db.close();
  }

  await expect(manager.loadInstalled("formula.sample-hello")).rejects.toThrow(/valid v2 extension package/i);
  const installed = await manager.getInstalled("formula.sample-hello");
  expect(installed?.corrupted).toBe(true);
  expect(String(installed?.corruptedReason ?? "")).toMatch(/valid v2 extension package/i);

  const verification = await manager.verifyInstalled("formula.sample-hello");
  expect(verification.ok).toBe(false);
  expect(verification.reason).toMatch(/valid v2 extension package/i);

  expect(host.listExtensions().some((e: any) => e.id === "formula.sample-hello")).toBe(false);

  await manager.dispose();
  await host.dispose();
});
