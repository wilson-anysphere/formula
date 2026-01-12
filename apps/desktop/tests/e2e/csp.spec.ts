import { expect, test } from "@playwright/test";
import crypto from "node:crypto";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

const requireFromHere = createRequire(import.meta.url);

// shared/* is CommonJS (shared/package.json sets type=commonjs).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const signingPkg: any = requireFromHere("../../../../shared/crypto/signing.js");
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const extensionPackagePkg: any = requireFromHere("../../../../shared/extension-package/index.js");

const { generateEd25519KeyPair } = signingPkg;
const { createExtensionPackageV2, readExtensionPackageV2 } = extensionPackagePkg;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "../../../..");

function viteFsUrl(absPath: string) {
  return `/@fs${absPath}`;
}

test.describe("Content Security Policy (Tauri parity)", () => {
  test("startup has no CSP violations and allows WASM-in-worker execution", async ({ page }) => {
    const cspViolations: string[] = [];

    page.on("console", (msg) => {
      if (msg.type() !== "error" && msg.type() !== "warning") {
        return;
      }
      const text = msg.text();
      if (/content security policy/i.test(text)) {
        cspViolations.push(text);
      }
    });

    const response = await page.goto("/");
    const cspHeader = response?.headers()["content-security-policy"];
    expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

    // The CSP smoke test doesn't need the full UI to render; it only needs the
    // document to load so we can validate WASM + Worker execution under the
    // configured policy.
    await expect(page.locator("#grid")).toHaveCount(1);

    const { mainThreadAnswer, workerAnswer } = await page.evaluate(async () => {
      const wasmBytes = new Uint8Array([
        0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00, 0x01, 0x05, 0x01, 0x60, 0x00, 0x01, 0x7f,
        0x03, 0x02, 0x01, 0x00, 0x07, 0x0a, 0x01, 0x06, 0x61, 0x6e, 0x73, 0x77, 0x65, 0x72, 0x00,
        0x00, 0x0a, 0x06, 0x01, 0x04, 0x00, 0x41, 0x2a, 0x0b
      ]);

      const { instance } = await WebAssembly.instantiate(wasmBytes);
      const mainThreadAnswer = (instance.exports as any).answer() as number;

      const workerAnswer = await new Promise<number>((resolve, reject) => {
        const code = `
          const wasmBytes = new Uint8Array(${JSON.stringify(Array.from(wasmBytes))});
          self.onmessage = async () => {
            try {
              const { instance } = await WebAssembly.instantiate(wasmBytes);
              self.postMessage(instance.exports.answer());
            } catch (err) {
              self.postMessage({ error: String(err) });
            }
          };
        `;

        const blob = new Blob([code], { type: "text/javascript" });
        const url = URL.createObjectURL(blob);
        const worker = new Worker(url, { type: "module" });

        const cleanup = () => {
          worker.terminate();
          URL.revokeObjectURL(url);
        };

        worker.onmessage = (event) => {
          const payload = event.data as any;
          if (payload && typeof payload === "object" && "error" in payload) {
            cleanup();
            reject(new Error(String(payload.error)));
            return;
          }
          cleanup();
          resolve(payload as number);
        };

        worker.onerror = (event) => {
          cleanup();
          reject(new Error((event as ErrorEvent).message || "worker error"));
        };

        worker.postMessage(null);
      });

      return { mainThreadAnswer, workerAnswer };
    });

    expect(mainThreadAnswer).toBe(42);
    expect(workerAnswer).toBe(42);

    expect(cspViolations, `Unexpected CSP violations:\\n${cspViolations.join("\n")}`).toEqual([]);
  });

  test("@formula/engine loads formula-wasm in a module Worker under CSP", async ({ page }) => {
    const cspViolations: string[] = [];

    page.on("console", (msg) => {
      if (msg.type() !== "error" && msg.type() !== "warning") {
        return;
      }
      const text = msg.text();
      if (/content security policy/i.test(text)) {
        cspViolations.push(text);
      }
    });

    const response = await page.goto("/");
    const cspHeader = response?.headers()["content-security-policy"];
    expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

    await expect(page.locator("#grid")).toHaveCount(1);

    const engineEntryUrl = viteFsUrl(path.join(repoRoot, "packages/engine/src/index.ts"));

    const result = await page.evaluate(async ({ engineEntryUrl }) => {
      const { createEngineClient } = await import(engineEntryUrl);

      const engine = createEngineClient();
      try {
        await engine.init();
        await engine.newWorkbook();
        await engine.setCell("A1", 1);
        await engine.setCell("A2", 2);
        await engine.setCell("B1", "=A1+A2");
        await engine.recalculate();
        return await engine.getCell("B1");
      } finally {
        engine.terminate();
      }
    }, {
      engineEntryUrl
    });

    expect(result?.value).toBe(3);
    expect(cspViolations, `Unexpected CSP violations:\\n${cspViolations.join("\n")}`).toEqual([]);
  });

  test("BrowserExtensionHost can run an extension in a Worker without CSP violations", async ({ page }) => {
    const cspViolations: string[] = [];

    page.on("console", (msg) => {
      if (msg.type() !== "error" && msg.type() !== "warning") {
        return;
      }
      const text = msg.text();
      if (/content security policy/i.test(text)) {
        cspViolations.push(text);
      }
    });

    const response = await page.goto("/");
    const cspHeader = response?.headers()["content-security-policy"];
    expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

    await expect(page.locator("#grid")).toHaveCount(1);

    const manifestUrl = viteFsUrl(path.join(repoRoot, "extensions/sample-hello/package.json"));
    const hostModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-host/src/browser/index.mjs"));

    const result = await page.evaluate(
      async ({ manifestUrl, hostModuleUrl }) => {
        const { BrowserExtensionHost } = await import(hostModuleUrl);

        const writes: Array<{ row: number; col: number; value: unknown }> = [];
        const spreadsheetApi = {
          async getSelection() {
            return {
              startRow: 0,
              startCol: 0,
              endRow: 1,
              endCol: 1,
              values: [
                [1, 2],
                [3, 4]
              ]
            };
          },
          async getCell() {
            return null;
          },
          async setCell(row: number, col: number, value: unknown) {
            writes.push({ row, col, value });
          }
        };

        const host = new BrowserExtensionHost({
          engineVersion: "1.0.0",
          spreadsheetApi,
          permissionPrompt: async () => true,
          // Vite rewrites `@formula/extension-api` into `/@fs/...` URLs, which fail the
          // strict import preflight. The import sandbox is exercised in unit tests; for
          // this e2e suite we disable preflight so we can validate CSP behavior.
          sandbox: { strictImports: false },
        });

        await host.loadExtensionFromUrl(manifestUrl);
        const sum = await host.executeCommand("sampleHello.sumSelection");
        const messages = host.getMessages();
        await host.dispose();

        return { sum, writes, messages };
      },
      { manifestUrl, hostModuleUrl }
    );

    expect(result.sum).toBe(10);
    expect(result.writes).toEqual([{ row: 2, col: 0, value: 10 }]);
    expect(result.messages.some((m: any) => String(m.message).includes("Sum: 10"))).toBe(true);

    expect(cspViolations, `Unexpected CSP violations:\\n${cspViolations.join("\n")}`).toEqual([]);
  });

  test("WebExtensionManager can install and load a v2 marketplace extension (blob/data module URL) without CSP violations", async ({
    page
  }) => {
    const keys = generateEd25519KeyPair();
    const extensionDir = path.join(repoRoot, "extensions", "sample-hello");
    const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });
    const pkgSha256 = crypto.createHash("sha256").update(pkgBytes).digest("hex");
    const pkgSignatureBase64 = readExtensionPackageV2(pkgBytes)?.signature?.signatureBase64 || "";
    expect(pkgSignatureBase64).toMatch(/\S/);

    const keyDer = crypto.createPublicKey(keys.publicKeyPem).export({ type: "spki", format: "der" });
    const publisherKeyId = crypto.createHash("sha256").update(keyDer).digest("hex");

    const cspViolations: string[] = [];

    page.on("console", (msg) => {
      if (msg.type() !== "error" && msg.type() !== "warning") {
        return;
      }
      const text = msg.text();
      if (/content security policy/i.test(text)) {
        cspViolations.push(text);
      }
    });

    await page.route("**/api/extensions/formula.sample-hello", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          id: "formula.sample-hello",
          name: "sample-hello",
          displayName: "Sample Hello",
          publisher: "formula",
          description: "",
          categories: [],
          tags: [],
          screenshots: [],
          verified: true,
          featured: false,
          deprecated: false,
          blocked: false,
          malicious: false,
          downloadCount: 0,
          latestVersion: "1.0.0",
          versions: [{ version: "1.0.0", sha256: pkgSha256, uploadedAt: new Date().toISOString(), yanked: false }],
          readme: "",
          publisherPublicKeyPem: keys.publicKeyPem,
          publisherKeys: [{ id: publisherKeyId, publicKeyPem: keys.publicKeyPem, revoked: false }],
          updatedAt: new Date().toISOString(),
          createdAt: new Date().toISOString()
        })
      });
    });

    await page.route("**/api/extensions/formula.sample-hello/download/1.0.0", async (route) => {
      await route.fulfill({
        status: 200,
        headers: {
          "Content-Type": "application/vnd.formula.extension-package",
          "X-Package-Sha256": pkgSha256,
          "X-Package-Signature": pkgSignatureBase64,
          "X-Package-Format-Version": "2",
          "X-Publisher": "formula",
          "X-Publisher-Key-Id": publisherKeyId
        },
        body: pkgBytes
      });
    });

    const response = await page.goto("/");
    const cspHeader = response?.headers()["content-security-policy"];
    expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

    await expect(page.locator("#grid")).toHaveCount(1);

    const hostModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-host/src/browser/index.mjs"));
    const managerModuleUrl = viteFsUrl(path.join(repoRoot, "apps/web/src/marketplace/WebExtensionManager.ts"));

    const result = await page.evaluate(
      async ({ hostModuleUrl, managerModuleUrl }) => {
        const { BrowserExtensionHost } = await import(hostModuleUrl);
        const { WebExtensionManager } = await import(managerModuleUrl);

        const writes: Array<{ row: number; col: number; value: unknown }> = [];
        const cellMap = new Map<string, unknown>();

        const spreadsheetApi = {
          async getSelection() {
            return {
              startRow: 0,
              startCol: 0,
              endRow: 1,
              endCol: 1,
              values: [
                [1, 2],
                [3, 4]
              ]
            };
          },
          async getCell(row: number, col: number) {
            const key = `${row},${col}`;
            return cellMap.has(key) ? cellMap.get(key) : null;
          },
          async setCell(row: number, col: number, value: unknown) {
            const key = `${row},${col}`;
            cellMap.set(key, value);
            writes.push({ row, col, value });
          }
        };

        const host = new BrowserExtensionHost({
          engineVersion: "1.0.0",
          spreadsheetApi,
          permissionPrompt: async () => true
        });

        const manager = new WebExtensionManager({ host });
        await manager.install("formula.sample-hello");
        const id = await manager.loadInstalled("formula.sample-hello");

        const loadedMainUrl = (manager as any)?._loadedMainUrls?.get(id)?.mainUrl ?? null;

        const sum = await host.executeCommand("sampleHello.sumSelection");
        const outCell = await spreadsheetApi.getCell(2, 0);
        const messages = host.getMessages();
        await manager.dispose();
        await host.dispose();

        return { id, loadedMainUrl, sum, outCell, writes, messages };
      },
      { hostModuleUrl, managerModuleUrl }
    );

    expect(result.id).toBe("formula.sample-hello");
    expect(result.loadedMainUrl).toMatch(/^(blob:|data:)/);
    expect(result.sum).toBe(10);
    expect(result.outCell).toBe(10);
    expect(result.writes).toEqual([{ row: 2, col: 0, value: 10 }]);
    expect(result.messages.some((m: any) => String(m.message).includes("Sum: 10"))).toBe(true);

    expect(cspViolations, `Unexpected CSP violations:\\n${cspViolations.join("\n")}`).toEqual([]);
  });
});
