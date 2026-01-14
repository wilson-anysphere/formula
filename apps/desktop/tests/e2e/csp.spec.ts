import { expect, test } from "@playwright/test";
import crypto from "node:crypto";
import http from "node:http";
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
const { createExtensionPackageV2, readExtensionPackageV2, verifyExtensionPackageV2 } = extensionPackagePkg;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "../../../..");

function viteFsUrl(absPath: string) {
  return `/@fs${absPath}`;
}

function getCspDirectiveSources(cspHeader: string, directive: string): string[] {
  // Very small CSP parser: enough for the e2e parity checks.
  // Example: "default-src 'self'; connect-src 'self' https:".
  const parts = cspHeader
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean);

  const directivePrefix = `${directive} `;
  const directiveValue = parts.find((part) => part === directive || part.startsWith(directivePrefix));
  if (!directiveValue) return [];

  return directiveValue
    .slice(directive.length)
    .trim()
    .split(/\s+/)
    .filter(Boolean);
}

function startHttpServer(
  handler: (req: http.IncomingMessage, res: http.ServerResponse) => void
): Promise<{ origin: string; close: () => Promise<void> }> {
  const server = http.createServer(handler);

  return new Promise((resolve, reject) => {
    server.listen(0, "127.0.0.1", () => {
      const addr = server.address();
      if (!addr || typeof addr === "string") {
        reject(new Error("Failed to bind HTTP server"));
        return;
      }
      resolve({
        origin: `http://127.0.0.1:${addr.port}`,
        close: () =>
          new Promise((resolveClose, rejectClose) => {
            server.close((err) => {
              if (err) rejectClose(err);
              else resolveClose();
            });
          })
      });
    });
    server.on("error", reject);
  });
}

async function gotoDesktopRootWithRetry(
  page: import("@playwright/test").Page,
): Promise<import("@playwright/test").Response | null> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization. If that
  // happens mid-navigation, retry once after the current navigation settles.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      return await page.goto("/", { waitUntil: "domcontentloaded" });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (
        attempt === 0 &&
        (message.includes("Execution context was destroyed") ||
          message.includes("net::ERR_ABORTED") ||
          message.includes("net::ERR_NETWORK_CHANGED") ||
          message.includes("interrupted by another navigation") ||
          message.includes("frame was detached"))
      ) {
        await page.waitForLoadState("domcontentloaded").catch(() => {});
        continue;
      }
      throw err;
    }
  }
  return null;
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

    const response = await gotoDesktopRootWithRetry(page);
    const cspHeader = response?.headers()["content-security-policy"];
    expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

    const connectSrc = getCspDirectiveSources(String(cspHeader), "connect-src");
    // The Tauri CSP is intentionally restrictive (no plain `http:`), but still allows
    // outbound HTTPS + WebSockets for collaboration and extensions, along with the in-memory
    // module URLs used by the extension system (`blob:`/`data:`).
    expect(connectSrc, "CSP `connect-src` should match Tauri config").toContain("'self'");
    expect(connectSrc).toContain("https:");
    expect(connectSrc).toContain("ws:");
    expect(connectSrc).toContain("wss:");
    expect(connectSrc).toContain("blob:");
    expect(connectSrc).toContain("data:");
    expect(connectSrc).not.toContain("http:");

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

  test("marketplace install + extension network.fetch works under production CSP via Tauri proxy", async ({ page }) => {
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

    const keyPair = generateEd25519KeyPair();
    const packageBytes: Buffer = await createExtensionPackageV2(path.join(repoRoot, "extensions/sample-hello"), {
      privateKeyPem: keyPair.privateKeyPem
    });

    const pkgSha256 = crypto.createHash("sha256").update(packageBytes).digest("hex");
    const verified = verifyExtensionPackageV2(packageBytes, keyPair.publicKeyPem);
    const filesSha256 = crypto
      .createHash("sha256")
      .update(JSON.stringify(verified.files || []), "utf8")
      .digest("hex");

    const now = new Date().toISOString();
    const extensionId = "formula.sample-hello";

    const marketplaceServer = await startHttpServer((req, res) => {
      const url = new URL(req.url || "/", "http://localhost");
      const method = req.method || "GET";

      // Minimal Marketplace API used by MarketplaceClient.
      if (method === "GET" && url.pathname === "/api/search") {
        const body = JSON.stringify({
          total: 1,
          results: [
            {
              id: extensionId,
              name: "sample-hello",
              displayName: "Sample Hello",
              publisher: "formula",
              description: "Sample extension",
              latestVersion: "1.0.0",
              verified: true,
              featured: false,
              categories: [],
              tags: [],
              screenshots: [],
              downloadCount: 1,
              updatedAt: now
            }
          ],
          nextCursor: null
        });

        res.statusCode = 200;
        res.setHeader("content-type", "application/json");
        res.end(body);
        return;
      }

      if (method === "GET" && url.pathname === `/api/extensions/${encodeURIComponent(extensionId)}`) {
        const body = JSON.stringify({
          id: extensionId,
          name: "sample-hello",
          displayName: "Sample Hello",
          publisher: "formula",
          description: "Sample extension",
          latestVersion: "1.0.0",
          verified: true,
          featured: false,
          categories: [],
          tags: [],
          screenshots: [],
          downloadCount: 1,
          updatedAt: now,
          versions: [
            {
              version: "1.0.0",
              sha256: pkgSha256,
              uploadedAt: now,
              yanked: false,
              formatVersion: 2
            }
          ],
          readme: "Sample Hello",
          publisherPublicKeyPem: keyPair.publicKeyPem,
          createdAt: now,
          deprecated: false,
          blocked: false,
          malicious: false
        });

        res.statusCode = 200;
        res.setHeader("content-type", "application/json");
        res.end(body);
        return;
      }

      if (
        method === "GET" &&
        url.pathname === `/api/extensions/${encodeURIComponent(extensionId)}/download/1.0.0`
      ) {
        res.statusCode = 200;
        res.setHeader("content-type", "application/octet-stream");
        res.setHeader("x-package-format-version", "2");
        res.setHeader("x-package-sha256", pkgSha256);
        res.setHeader("x-package-signature", verified.signatureBase64);
        res.setHeader("x-publisher", "formula");
        res.setHeader("x-package-scan-status", "passed");
        res.setHeader("x-package-files-sha256", filesSha256);
        res.end(packageBytes);
        return;
      }

      res.statusCode = 404;
      res.setHeader("content-type", "text/plain");
      res.end("not found");
    });

    const networkServer = await startHttpServer((req, res) => {
      const url = new URL(req.url || "/", "http://localhost");
      if (req.method === "GET" && url.pathname === "/hello") {
        res.statusCode = 200;
        res.setHeader("content-type", "text/plain");
        res.end("hello from network");
        return;
      }
      res.statusCode = 404;
      res.setHeader("content-type", "text/plain");
      res.end("not found");
    });

    try {
      await page.exposeFunction("__TAURI_INVOKE__", async (cmd: string, args: any) => {
        const joinMarketplaceUrl = (segments: string[], searchParams?: Record<string, string>) => {
          const baseUrl = String(args?.baseUrl ?? "");
          const u = new URL(baseUrl);
          const basePath = u.pathname.replace(/\/+$/, "");
          u.pathname = `${basePath}/${segments.map((s) => encodeURIComponent(String(s))).join("/")}`;
          u.search = "";
          if (searchParams) {
            for (const [k, v] of Object.entries(searchParams)) {
              if (v === undefined || v === null || String(v).trim().length === 0) continue;
              u.searchParams.set(k, String(v));
            }
          }
          return u.toString();
        };

        switch (cmd) {
          case "network_fetch": {
            const response = await fetch(String(args?.url ?? ""), args?.init ?? undefined);
            const bodyText = await response.text();
            return {
              ok: response.ok,
              status: response.status,
              statusText: response.statusText,
              url: response.url,
              headers: Array.from(response.headers.entries()),
              bodyText
            };
          }

          case "marketplace_search": {
            const url = joinMarketplaceUrl(["search"], {
              q: args?.q,
              category: args?.category,
              tag: args?.tag,
              verified: args?.verified === undefined ? "" : args.verified ? "true" : "false",
              featured: args?.featured === undefined ? "" : args.featured ? "true" : "false",
              sort: args?.sort,
              limit: args?.limit === undefined ? "" : String(args.limit),
              offset: args?.offset === undefined ? "" : String(args.offset),
              cursor: args?.cursor
            });
            const response = await fetch(url);
            if (!response.ok) {
              throw new Error(`marketplace_search failed (${response.status})`);
            }
            return response.json();
          }

          case "marketplace_get_extension": {
            const url = joinMarketplaceUrl(["extensions", String(args?.id ?? "")]);
            const response = await fetch(url);
            if (response.status === 404) return null;
            if (!response.ok) {
              throw new Error(`marketplace_get_extension failed (${response.status})`);
            }
            return response.json();
          }

          case "marketplace_download_package": {
            const url = joinMarketplaceUrl([
              "extensions",
              String(args?.id ?? ""),
              "download",
              String(args?.version ?? "")
            ]);
            const response = await fetch(url);
            if (response.status === 404) return null;
            if (!response.ok) {
              throw new Error(`marketplace_download_package failed (${response.status})`);
            }
            const buf = Buffer.from(await response.arrayBuffer());
            return {
              bytesBase64: buf.toString("base64"),
              signatureBase64: response.headers.get("x-package-signature"),
              sha256: response.headers.get("x-package-sha256"),
              formatVersion: response.headers.get("x-package-format-version")
                ? Number.parseInt(String(response.headers.get("x-package-format-version")), 10)
                : null,
              publisher: response.headers.get("x-publisher"),
              publisherKeyId: response.headers.get("x-publisher-key-id"),
              scanStatus: response.headers.get("x-package-scan-status"),
              filesSha256: response.headers.get("x-package-files-sha256")
            };
          }

          default:
            throw new Error(`Unexpected invoke: ${cmd} ${JSON.stringify(args)}`);
        }
      });

      const response = await gotoDesktopRootWithRetry(page);
      const cspHeader = response?.headers()["content-security-policy"];
      expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

      await expect(page.locator("#grid")).toHaveCount(1);

      const hostModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-host/src/browser/index.mjs"));
      const marketplaceModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-marketplace/src/index.ts"));

      const result = await page.evaluate(
        async ({ hostModuleUrl, marketplaceModuleUrl, marketplaceBaseUrl, networkUrl }) => {
          // Inject a minimal Tauri IPC surface for the marketplace + extension host.
          (window as any).__TAURI__ = {
            core: {
              invoke: (cmd: string, args: any) => (window as any).__TAURI_INVOKE__(cmd, args)
            }
          };

          const { BrowserExtensionHost } = await import(hostModuleUrl);
          const { MarketplaceClient, WebExtensionManager } = await import(marketplaceModuleUrl);

          const spreadsheetApi = {
            async getSelection() {
              return {
                startRow: 0,
                startCol: 0,
                endRow: 0,
                endCol: 0,
                values: [[null]]
              };
            },
            async getCell() {
              return null;
            },
            async setCell() {}
          };

          const host = new BrowserExtensionHost({
            engineVersion: "1.0.0",
            spreadsheetApi,
            permissionPrompt: async () => true,
            // The strict import preflight uses `fetch()` to walk module graphs.
            // Some WebViews can be finicky about `fetch(blob:...)` during module preflight,
            // even when `connect-src` includes `blob:`. The import sandbox is covered in unit
            // tests; disable preflight here so this e2e suite can focus on CSP + Tauri IPC.
            sandbox: { strictImports: false }
          });

          const marketplaceClient = new MarketplaceClient({ baseUrl: marketplaceBaseUrl });
          const manager = new WebExtensionManager({ marketplaceClient, host });

          const id = "formula.sample-hello";
          await manager.install(id);
          await manager.loadInstalled(id);

          const text = await host.executeCommand("sampleHello.fetchText", networkUrl);
          const messages = host.getMessages();

          await manager.dispose();
          await host.dispose();

          return { text, messages };
        },
        {
          hostModuleUrl,
          marketplaceModuleUrl,
          marketplaceBaseUrl: `${marketplaceServer.origin}/api`,
          networkUrl: `${networkServer.origin}/hello`
        }
      );

      expect(result.text).toBe("hello from network");
      expect(result.messages.some((m: any) => String(m.message).includes("Fetched: hello from network"))).toBe(
        true
      );

      expect(cspViolations, `Unexpected CSP violations:\\n${cspViolations.join("\n")}`).toEqual([]);
    } finally {
      await networkServer.close();
      await marketplaceServer.close();
    }
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

    const response = await gotoDesktopRootWithRetry(page);
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

    const response = await gotoDesktopRootWithRetry(page);
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
          sandbox: { strictImports: false }
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

  test("DesktopExtensionHostManager loads the bundled built-in extension (blob module URL) under CSP", async ({ page }) => {
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

    const response = await gotoDesktopRootWithRetry(page);
    const cspHeader = response?.headers()["content-security-policy"];
    expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

    await expect(page.locator("#grid")).toHaveCount(1);

    const managerModuleUrl = viteFsUrl(path.join(repoRoot, "apps/desktop/src/extensions/extensionHostManager.ts"));

    const result = await page.evaluate(
      async ({ managerModuleUrl }) => {
        const { DesktopExtensionHostManager } = await import(managerModuleUrl);

        const originalFetch = typeof fetch === "function" ? fetch.bind(globalThis) : null;
        if (originalFetch) {
          // Built-in extensions should be bundled into the app and loaded via blob/data module URLs.
          // They must not fetch the repo/workspace manifest at runtime (no Vite `/@fs` dependency).
          // Guard against regressions back to `loadExtensionFromUrl(...)`.
          globalThis.fetch = async (input: RequestInfo | URL, init?: RequestInit) => {
            const url =
              typeof input === "string"
                ? input
                : input instanceof URL
                  ? input.toString()
                  : typeof (input as any)?.url === "string"
                    ? String((input as any).url)
                    : String(input);
            if (url.includes("extensions/sample-hello/") && url.includes("package.json")) {
              throw new Error(`Blocked fetch for built-in extension asset: ${url}`);
            }
            return originalFetch(input, init);
          };
        }

        const writes: Array<{ row: number; col: number; value: unknown }> = [];
        const messages: Array<{ message: string; type?: string }> = [];

        const spreadsheetApi = {
          async getSelection() {
            return {
              startRow: 0,
              startCol: 0,
              endRow: 1,
              endCol: 1,
              values: [
                [1, 2],
                [3, 4],
              ],
            };
          },
          async getCell() {
            return null;
          },
          async setCell(row: number, col: number, value: unknown) {
            writes.push({ row, col, value });
          },
        };

        const manager = new DesktopExtensionHostManager({
          engineVersion: "1.0.0",
          spreadsheetApi,
          uiApi: {
            showMessage: async (message: string, type?: string) => {
              messages.push({ message: String(message ?? ""), type: type ? String(type) : undefined });
            },
          },
          permissionPrompt: async () => true,
        });

        await manager.loadBuiltInExtensions();
        const sum = await manager.executeCommand("sampleHello.sumSelection");
        await manager.host.dispose();

        return { sum, writes, messages };
      },
      { managerModuleUrl },
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
          versions: [
            {
              version: "1.0.0",
              sha256: pkgSha256,
              uploadedAt: new Date().toISOString(),
              yanked: false,
              scanStatus: "passed",
            },
          ],
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
          "X-Package-Scan-Status": "passed",
          "X-Publisher": "formula",
          "X-Publisher-Key-Id": publisherKeyId,
        },
        body: pkgBytes
      });
    });

    const networkServer = await startHttpServer((req, res) => {
      const url = new URL(req.url || "/", "http://localhost");
      if (req.method === "GET" && url.pathname === "/hello") {
        res.statusCode = 200;
        res.setHeader("content-type", "text/plain");
        res.end("hello");
        return;
      }
      res.statusCode = 404;
      res.setHeader("content-type", "text/plain");
      res.end("not found");
    });

    try {
      await page.exposeFunction("__TAURI_INVOKE__", async (cmd: string, args: any) => {
        if (cmd !== "network_fetch") {
          throw new Error(`Unexpected invoke: ${cmd} ${JSON.stringify(args)}`);
        }

        const response = await fetch(String(args?.url ?? ""), args?.init ?? undefined);
        const bodyText = await response.text();
        return {
          ok: response.ok,
          status: response.status,
          statusText: response.statusText,
          url: response.url,
          headers: Array.from(response.headers.entries()),
          bodyText
        };
      });

      const response = await gotoDesktopRootWithRetry(page);
      const cspHeader = response?.headers()["content-security-policy"];
      expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

      await expect(page.locator("#grid")).toHaveCount(1);

      const hostModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-host/src/browser/index.mjs"));
      const managerModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-marketplace/src/index.ts"));

      const result = await page.evaluate(
        async ({ hostModuleUrl, managerModuleUrl, networkUrl }) => {
          (window as any).__TAURI__ = {
            core: {
              invoke: (cmd: string, args: any) => (window as any).__TAURI_INVOKE__(cmd, args)
            }
          };

          const { BrowserExtensionHost } = await import(hostModuleUrl);
          const { WebExtensionManager } = await import(managerModuleUrl);

        const canBlobModuleUrls =
          typeof URL !== "undefined" && typeof URL.createObjectURL === "function" && typeof Blob !== "undefined";

        const run = async ({ forceData }: { forceData: boolean }) => {
          const writes: Array<{ row: number; col: number; value: unknown }> = [];
          const cellMap = new Map<string, unknown>();

          const originalProcess = (globalThis as any).process;
          if (forceData) {
            // Force the marketplace loader to fall back to `data:` module URLs so we can validate
            // that the configured CSP permits `data:` entrypoints (some environments do not support
            // `URL.createObjectURL`, and will use `data:` as a fallback).
            try {
              (globalThis as any).process = { versions: { node: "18.0.0" } };
            } catch {
              // ignore
            }
          }

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

          try {
            const manager = new WebExtensionManager({ host, engineVersion: "1.0.0", scanPolicy: "allow" });
            await manager.install("formula.sample-hello");
            const contributedPanelsSeedRaw =
              globalThis.localStorage?.getItem("formula.extensions.contributedPanels.v1") ?? null;
            const contributedPanelsSeed = contributedPanelsSeedRaw ? JSON.parse(String(contributedPanelsSeedRaw)) : null;
            const seededPanel = contributedPanelsSeed?.["sampleHello.panel"] ?? null;
            const id = await manager.loadInstalled("formula.sample-hello");

            const loadedMainUrl = (manager as any)?._loadedMainUrls?.get(id)?.mainUrl ?? null;

            const sum = await host.executeCommand("sampleHello.sumSelection");
            const fetched = await host.executeCommand("sampleHello.fetchText", networkUrl);
            const outCell = await spreadsheetApi.getCell(2, 0);
            const messages = host.getMessages();
            await manager.dispose();
            await host.dispose();

            return { id, loadedMainUrl, seededPanel, sum, fetched, outCell, writes, messages };
          } finally {
            // Ensure we never leak a `process` shim into later runs.
            if (forceData) {
              try {
                if (originalProcess === undefined) {
                  delete (globalThis as any).process;
                } else {
                  (globalThis as any).process = originalProcess;
                }
              } catch {
                // ignore
              }
            }
          }
        };

        const blob = await run({ forceData: false });
        const data = await run({ forceData: true });
          return { canBlobModuleUrls, blob, data };
        },
        { hostModuleUrl, managerModuleUrl, networkUrl: `${networkServer.origin}/hello` }
      );

    expect(result.blob.id).toBe("formula.sample-hello");
    expect(result.blob.loadedMainUrl).toMatch(/^(blob:|data:)/);
    expect(result.blob.seededPanel).toMatchObject({
      extensionId: "formula.sample-hello",
      title: "Sample Hello Panel"
    });
    if (result.canBlobModuleUrls) {
      expect(result.blob.loadedMainUrl).toMatch(/^blob:/);
    }
    expect(result.blob.sum).toBe(10);
    expect(result.blob.fetched).toBe("hello");
    expect(result.blob.outCell).toBe(10);
    expect(result.blob.writes).toEqual([{ row: 2, col: 0, value: 10 }]);
    expect(result.blob.messages.some((m: any) => String(m.message).includes("Sum: 10"))).toBe(true);
    expect(result.blob.messages.some((m: any) => String(m.message).includes("Fetched: hello"))).toBe(true);

    expect(result.data.id).toBe("formula.sample-hello");
    expect(result.data.loadedMainUrl).toMatch(/^(blob:|data:)/);
    expect(result.data.seededPanel).toMatchObject({
      extensionId: "formula.sample-hello",
      title: "Sample Hello Panel"
    });
    expect(result.data.loadedMainUrl).toMatch(/^data:/);
    expect(result.data.sum).toBe(10);
    expect(result.data.fetched).toBe("hello");
    expect(result.data.outCell).toBe(10);
    expect(result.data.writes).toEqual([{ row: 2, col: 0, value: 10 }]);
    expect(result.data.messages.some((m: any) => String(m.message).includes("Sum: 10"))).toBe(true);
      expect(result.data.messages.some((m: any) => String(m.message).includes("Fetched: hello"))).toBe(true);

      expect(cspViolations, `Unexpected CSP violations:\\n${cspViolations.join("\n")}`).toEqual([]);
    } finally {
      await networkServer.close();
    }
  });

  test("extension network.fetch fails when denied under CSP without CSP violations", async ({ page }) => {
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

    let requestCount = 0;
    await page.route("https://example.test/**", async (route) => {
      requestCount += 1;
      await route.fulfill({
        status: 200,
        headers: {
          "Access-Control-Allow-Origin": "*",
          "Content-Type": "text/plain"
        },
        body: "hello"
      });
    });

    const response = await gotoDesktopRootWithRetry(page);
    const cspHeader = response?.headers()["content-security-policy"];
    expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

    await expect(page.locator("#grid")).toHaveCount(1);

    const manifestUrl = viteFsUrl(path.join(repoRoot, "extensions/sample-hello/package.json"));
    const hostModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-host/src/browser/index.mjs"));

    const result = await page.evaluate(
      async ({ manifestUrl, hostModuleUrl }) => {
        const { BrowserExtensionHost } = await import(hostModuleUrl);

        const host = new BrowserExtensionHost({
          engineVersion: "1.0.0",
          spreadsheetApi: {
            async getSelection() {
              return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
            },
            async getCell() {
              return null;
            },
            async setCell() {
              // noop
            }
          },
          permissionPrompt: async ({ permissions }: { permissions: string[] }) => {
            if (permissions.includes("network")) return false;
            return true;
          },
          // Vite rewrites `@formula/extension-api` into `/@fs/...` URLs, which fail the strict import preflight.
          sandbox: { strictImports: false },
          // Avoid leaking persisted grants between tests.
          permissionStorageKey: `formula.extensionHost.permissions.csp.deny.${Date.now()}`
        });

        try {
          await host.loadExtensionFromUrl(manifestUrl);
          await host.executeCommand("sampleHello.fetchText", "https://example.test/hello");
          return { errorMessage: "" };
        } catch (err: any) {
          return { errorMessage: String(err?.message ?? err) };
        } finally {
          await host.dispose();
        }
      },
      { manifestUrl, hostModuleUrl }
    );

    expect(result.errorMessage).toContain("Permission denied");
    expect(requestCount).toBe(0);
    expect(cspViolations, `Unexpected CSP violations:\\n${cspViolations.join("\n")}`).toEqual([]);
  });
  test("extension panels are sandboxed with connect-src 'none' (no network bypass) under CSP", async ({
    page
  }) => {
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

    const response = await gotoDesktopRootWithRetry(page);
    const cspHeader = response?.headers()["content-security-policy"];
    expect(cspHeader, "E2E server should emit Content-Security-Policy header").toBeTruthy();

    await expect(page.locator("#grid")).toHaveCount(1);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHost), undefined, { timeout: 60_000 });
    // Avoid interactive permission dialogs (the prompt waits on user input).
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (window as any).__formulaPermissionPrompt = async () => true;
    });

    await page.evaluate(async () => {
      // Auto-approve extension permission prompts so the test exercises panel CSP
      // behavior rather than hanging on modal UI.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).__formulaPermissionPrompt = async () => true;

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const manager: any = (window as any).__formulaExtensionHostManager;
      if (!manager) throw new Error("Missing window.__formulaExtensionHostManager (desktop e2e harness)");

      // First-run module compilation can exceed the BrowserExtensionHost default activation timeout
      // under CI load. Bump timeouts for this CSP-only test suite so it exercises the iframe
      // sandbox CSP rather than failing due to startup latency.
      manager.host._activationTimeoutMs = 20_000;
      manager.host._commandTimeoutMs = 20_000;

      // Use the built-in (bundled) sample extension so this test doesn't depend on Vite `/@fs`
      // module transforms (which can trip the strict import preflight).
      await manager.loadBuiltInExtensions();
      await manager.executeCommand("sampleHello.openPanel");
    });

    const iframe = page.locator('[data-testid="extension-webview-sampleHello.panel"]');
    await expect(iframe).toHaveCount(1);
    await expect(iframe).toHaveAttribute("sandbox", "allow-scripts");
    await expect(iframe).toHaveAttribute("src", /^(blob:|data:)/);

    const iframeSrc = await iframe.getAttribute("src");
    expect(iframeSrc).toBeTruthy();

    const panelHtml = await page.evaluate(async (src) => {
      const res = await fetch(String(src));
      return await res.text();
    }, iframeSrc);

    expect(panelHtml).toContain("Content-Security-Policy");
    expect(panelHtml).toContain("connect-src 'none'");
    expect(panelHtml).toContain("worker-src 'none'");
    expect(panelHtml).toContain("frame-src 'none'");
    expect(cspViolations, `Unexpected CSP violations:\\n${cspViolations.join("\n")}`).toEqual([]);
  });
}); 
