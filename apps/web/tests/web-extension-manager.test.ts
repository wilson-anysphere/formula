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
  deprecated = false,
  blocked = false,
  malicious = false,
  publisherRevoked = false,
  scanStatuses = null
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
        versions: [],
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
          scanStatuses && typeof scanStatuses === "object" && typeof scanStatuses[version] === "string"
            ? scanStatuses[version]
            : null,
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

  await manager.install("formula.sample-hello");

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

  await manager.dispose();
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
  await manager.loadInstalled("test.test-ext");
  expect(await host.executeCommand("test.version")).toBe("v1");

  await manager.update("test.test-ext");
  expect(await manager.getInstalled("test.test-ext")).toMatchObject({ version: "1.0.1" });
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

test("scanStatus is allowed by default in dev/test builds (warn-only)", async () => {
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

  const result = await manager.install("formula.sample-hello", "1.0.0");
  expect(result.warnings?.some((w) => w.kind === "scanStatus")).toBe(true);

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
