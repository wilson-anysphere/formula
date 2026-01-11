import { expect, test } from "@playwright/test";
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
const { createExtensionPackageV2 } = extensionPackagePkg;

const repoRoot = fileURLToPath(new URL("../../../../", import.meta.url));

test("install + run marketplace extension in browser (no CSP violations)", async ({ page }) => {
  const keys = generateEd25519KeyPair();
  const extensionDir = path.join(repoRoot, "extensions", "sample-hello");
  const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

  const cspViolations: string[] = [];
  const consoleErrors: string[] = [];
  const pageErrors: string[] = [];
  const requestFailures: string[] = [];
  const notFoundResponses: string[] = [];
  page.on("console", (msg) => {
    const text = msg.text();
    if (msg.type() === "error") {
      consoleErrors.push(text);
      if (text.toLowerCase().includes("content security policy") || text.toLowerCase().includes("csp")) {
        cspViolations.push(text);
      }
    }
  });
  page.on("pageerror", (err) => {
    pageErrors.push(err?.stack ?? err?.message ?? String(err));
  });
  page.on("requestfailed", (req) => {
    const failure = req.failure();
    requestFailures.push(`${req.method()} ${req.url()}${failure?.errorText ? ` (${failure.errorText})` : ""}`);
  });
  page.on("response", (res) => {
    if (res.status() === 404) {
      notFoundResponses.push(res.url());
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
        versions: [{ version: "1.0.0", sha256: "", uploadedAt: new Date().toISOString(), yanked: false }],
        readme: "",
        publisherPublicKeyPem: keys.publicKeyPem,
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
        "X-Package-Format-Version": "2",
        "X-Publisher": "formula"
      },
      body: pkgBytes
    });
  });

  await page.goto("/?extTest=1");
  try {
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionTest), null, { timeout: 30_000 });
  } catch (error) {
    const details = [
      `Timed out waiting for window.__formulaExtensionTest.`,
      `Page URL: ${page.url()}`,
      pageErrors.length ? `Page errors:\n${pageErrors.join("\n\n")}` : "Page errors: <none>",
      consoleErrors.length ? `Console errors:\n${consoleErrors.join("\n")}` : "Console errors: <none>",
      notFoundResponses.length ? `404 responses:\n${notFoundResponses.join("\n")}` : "404 responses: <none>",
      requestFailures.length ? `Request failures:\n${requestFailures.join("\n")}` : "Request failures: <none>"
    ].join("\n\n");
    throw new Error(`${details}\n\nOriginal error:\n${String((error as Error)?.message ?? error)}`);
  }

  const result = await page.evaluate(async () => {
    const api = (window as any).__formulaExtensionTest;
    const id = await api.installSampleHello();

    await api.setCell(0, 0, 1);
    await api.setCell(0, 1, 2);
    await api.setCell(1, 0, 3);
    await api.setCell(1, 1, 4);
    api.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });

    const sum = await api.executeCommand("sampleHello.sumSelection");
    const outCell = api.getCell(2, 0);
    const messages = api.getMessages();
    await api.dispose();
    return { id, sum, outCell, messages };
  });

  expect(result.id).toBe("formula.sample-hello");
  expect(result.sum).toBe(10);
  expect(result.outCell).toBe(10);
  expect(cspViolations).toEqual([]);
});
