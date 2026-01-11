import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { Worker as NodeWorker } from "node:worker_threads";

import { IDBKeyRange, indexedDB } from "fake-indexeddb";
import { afterAll, afterEach, beforeEach, expect, test } from "vitest";

import { WebExtensionManager } from "../src/marketplace/WebExtensionManager";

// CJS helpers (shared/* is CommonJS).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const signingPkg: any = await import("../../../shared/crypto/signing.js");
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const extensionPackagePkg: any = await import("../../../shared/extension-package/index.js");

const { generateEd25519KeyPair } = signingPkg;
const { createExtensionPackageV2 } = extensionPackagePkg;

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const browserHostPkg: any = await import("@formula/extension-host/browser");
const { BrowserExtensionHost } = browserHostPkg;

const WORKER_WRAPPER_URL = new URL("./helpers/node-web-worker.mjs", import.meta.url);

const originalGlobals = {
  indexedDB: (globalThis as any).indexedDB,
  IDBKeyRange: (globalThis as any).IDBKeyRange,
  crypto: (globalThis as any).crypto,
  Worker: (globalThis as any).Worker
};

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

function createMockMarketplace({ extensionId, latestVersion, publicKeyPem, packages, publisherKeys = null, publisherKeyIds = null }: any) {
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
        deprecated: false,
        blocked: false,
        malicious: false
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
            : null
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

beforeEach(async () => {
  // Provide browser primitives.
  // eslint-disable-next-line no-global-assign
  (globalThis as any).indexedDB = indexedDB;
  // eslint-disable-next-line no-global-assign
  (globalThis as any).IDBKeyRange = IDBKeyRange;
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

  const manager = new WebExtensionManager({ marketplaceClient, host });

  await manager.install("formula.sample-hello");
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

  const manager = new WebExtensionManager({ marketplaceClient, host });

  await expect(manager.install("formula.sample-hello")).rejects.toThrow(/verification failed/i);
  expect(await manager.listInstalled()).toEqual([]);

  await manager.dispose();
  await host.dispose();
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
      main: "./dist/extension.mjs",
      browser: "./dist/extension.mjs",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCommand:test.version"],
      permissions: ["ui.commands"],
      contributes: { commands: [{ command: "test.version", title: "Version" }] }
    };
    await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
    await fs.writeFile(
      path.join(extDir, "dist", "extension.mjs"),
      `import * as formula from "@formula/extension-api";\nexport async function activate() {\n  await formula.commands.registerCommand("test.version", () => "${marker}");\n}\n`
    );
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
  const manager = new WebExtensionManager({ marketplaceClient, host });

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
      main: "./dist/extension.mjs",
      browser: "./dist/extension.mjs",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCommand:test.version"],
      permissions: ["ui.commands"],
      contributes: { commands: [{ command: "test.version", title: "Version" }] }
    };
    await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
    await fs.writeFile(
      path.join(extDir, "dist", "extension.mjs"),
      `import * as formula from "@formula/extension-api";\nexport async function activate() {\n  await formula.commands.registerCommand("test.version", () => "${marker}");\n}\n`
    );
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
  const manager = new WebExtensionManager({ marketplaceClient, host });

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
