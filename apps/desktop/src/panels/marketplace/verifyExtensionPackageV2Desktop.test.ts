import crypto from "node:crypto";
import path from "node:path";

import { IDBKeyRange, indexedDB } from "fake-indexeddb";
import { afterAll, afterEach, beforeEach, expect, test, vi } from "vitest";

import { WebExtensionManager } from "../../../../../packages/extension-marketplace/src/WebExtensionManager";

import { verifyExtensionPackageV2Desktop } from "./verifyExtensionPackageV2Desktop";

// CJS helpers (shared/* is CommonJS).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const signingPkg: any = await import("../../../../../shared/crypto/signing.js");
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const extensionPackagePkg: any = await import("../../../../../shared/extension-package/index.js");

const { generateEd25519KeyPair } = signingPkg;
const { createExtensionPackageV2 } = extensionPackagePkg;

const originalGlobals = {
  indexedDB: (globalThis as any).indexedDB,
  IDBKeyRange: (globalThis as any).IDBKeyRange,
  __TAURI__: (globalThis as any).__TAURI__
};

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

  await deleteExtensionDb();
});

afterEach(async () => {
  await deleteExtensionDb();
  vi.restoreAllMocks();
  // eslint-disable-next-line no-global-assign
  (globalThis as any).__TAURI__ = originalGlobals.__TAURI__;
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
});

function createMockMarketplace({ extensionId, latestVersion, publicKeyPem, pkgBytes }: any) {
  const now = new Date().toISOString();
  return {
    // `WebExtensionManager` accepts a `MarketplaceClient` (class) and relies on structural typing
    // in tests. Provide a minimal compatible surface area.
    baseUrl: "https://example.invalid/api",
    async search(_params: any) {
      return { total: 0, results: [], nextCursor: null };
    },
    async getExtension(id: string) {
      if (id !== extensionId) return null;
      return {
        id,
        name: id.split(".").slice(1).join("."),
        displayName: id,
        publisher: id.split(".")[0] ?? "unknown",
        description: "",
        latestVersion,
        verified: true,
        featured: false,
        categories: [],
        tags: [],
        screenshots: [],
        downloadCount: 0,
        updatedAt: now,
        createdAt: now,
        versions: [
          {
            version: latestVersion,
            sha256: "0".repeat(64),
            uploadedAt: now,
            yanked: false,
            scanStatus: "passed",
            signingKeyId: null,
            formatVersion: 2,
          },
        ],
        readme: "",
        deprecated: false,
        blocked: false,
        malicious: false,
        publisherPublicKeyPem: publicKeyPem
      };
    },
    async downloadPackage(id: string, version: string) {
      if (id !== extensionId) return null;
      if (version !== latestVersion) return null;
      const rawBytes: Uint8Array = pkgBytes instanceof Uint8Array ? pkgBytes : new Uint8Array(pkgBytes);
      const bytes: Uint8Array<ArrayBuffer> =
        rawBytes.buffer instanceof ArrayBuffer ? (rawBytes as Uint8Array<ArrayBuffer>) : new Uint8Array(rawBytes);
      return {
        bytes,
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
}

test("falls back to Tauri Ed25519 verification when WebCrypto Ed25519 is unsupported", async () => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.resolve("extensions/sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const marketplaceClient = createMockMarketplace({
    extensionId: "formula.sample-hello",
    latestVersion: "1.0.0",
    publicKeyPem: keys.publicKeyPem,
    pkgBytes
  });

  const invoke = vi.fn().mockResolvedValue(true);
  // eslint-disable-next-line no-global-assign
  (globalThis as any).__TAURI__ = { core: { invoke } };

  const importKeySpy = vi
    .spyOn(globalThis.crypto.subtle, "importKey")
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    .mockImplementation(async () => {
      throw { name: "NotSupportedError" } as any;
    });

  const manager = new WebExtensionManager({
    marketplaceClient,
    host: null,
    verifyPackage: verifyExtensionPackageV2Desktop
  });

  await expect(manager.install("formula.sample-hello")).resolves.toMatchObject({
    id: "formula.sample-hello",
    version: "1.0.0"
  });

  expect(importKeySpy).toHaveBeenCalled();
  expect(invoke).toHaveBeenCalledWith("verify_ed25519_signature", expect.any(Object));
  expect(invoke.mock.calls[0][1]).toMatchObject({
    signature_base64: expect.any(String),
    public_key_pem: keys.publicKeyPem
  });

  await manager.dispose();
});
