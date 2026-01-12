import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

// CJS helpers (shared/* is CommonJS).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const signingPkg: any = await import("@formula/marketplace-shared/crypto/signing.js");
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const extensionPackagePkg: any = await import("@formula/marketplace-shared/extension-package/index.js");
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const browserVerifierPkg: any = await import("@formula/marketplace-shared/extension-package/v2-browser.mjs");

const { generateEd25519KeyPair } = signingPkg;
const { createExtensionPackageV2 } = extensionPackagePkg;
const { verifyExtensionPackageV2Browser } = browserVerifierPkg;

async function createTempExtensionDir({ version }: { version: string }) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-extpkg-webcrypto-fallback-"));
  await fs.mkdir(path.join(dir, "dist"), { recursive: true });

  const manifest = {
    name: "temp-ext",
    publisher: "temp-pub",
    version,
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" }
  };

  await fs.writeFile(path.join(dir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(path.join(dir, "dist", "extension.js"), `export default {};\n`);
  await fs.writeFile(path.join(dir, "README.md"), "hello\n");

  return { dir, manifest };
}

describe("verifyExtensionPackageV2Browser (desktop Ed25519 fallback)", () => {
  const originalTauri = (globalThis as any).__TAURI__;
  const subtle = globalThis.crypto?.subtle as SubtleCrypto | undefined;
  const originalImportKey = subtle?.importKey;

  let tmpDir: string | null = null;

  afterEach(async () => {
    (globalThis as any).__TAURI__ = originalTauri;
    if (subtle && originalImportKey) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (subtle as any).importKey = originalImportKey;
    }

    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
      tmpDir = null;
    }
  });

  it("falls back to Tauri invoke when WebCrypto doesn't support Ed25519", async () => {
    if (!subtle || typeof originalImportKey !== "function") {
      throw new Error("Test requires WebCrypto subtle.importKey to exist");
    }

    const keys = generateEd25519KeyPair();
    const { dir, manifest } = await createTempExtensionDir({ version: "1.0.0" });
    tmpDir = dir;

    const pkgBytes = await createExtensionPackageV2(dir, { privateKeyPem: keys.privateKeyPem });

    const invoke = vi.fn(async (cmd: string, args?: Record<string, unknown>) => {
      expect(cmd).toBe("verify_ed25519_signature");
      expect(args).toBeTruthy();
      expect(Array.isArray(args!.payload)).toBe(true);
      expect(args!.signature_base64).toBeTypeOf("string");
      expect(args!.public_key_pem).toBe(keys.publicKeyPem);
      return true;
    });

    (globalThis as any).__TAURI__ = { core: { invoke } };

    // Simulate a WebView whose WebCrypto implementation lacks Ed25519 (e.g. WKWebView/WebKitGTK).
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (subtle as any).importKey = vi.fn(async () => {
      throw { name: "NotSupportedError", message: "Ed25519 not supported" };
    });

    const result = await verifyExtensionPackageV2Browser(new Uint8Array(pkgBytes), keys.publicKeyPem);
    expect(result.manifest).toEqual(manifest);
    expect(invoke).toHaveBeenCalledTimes(1);
  });
});
