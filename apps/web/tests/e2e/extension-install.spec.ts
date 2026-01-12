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

const repoRoot = fileURLToPath(new URL("../../../../", import.meta.url));

type SampleHelloFixture = {
  keys: { publicKeyPem: string; privateKeyPem: string };
  pkgBytes: Uint8Array;
  pkgSha256: string;
  pkgSignatureBase64: string;
  publisherKeyId: string;
};

let sampleHelloFixturePromise: Promise<SampleHelloFixture> | null = null;

async function getSampleHelloFixture(): Promise<SampleHelloFixture> {
  if (sampleHelloFixturePromise) return sampleHelloFixturePromise;
  sampleHelloFixturePromise = (async () => {
    const keys = generateEd25519KeyPair();
    const extensionDir = path.join(repoRoot, "extensions", "sample-hello");
    const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });
    const pkgSha256 = crypto.createHash("sha256").update(pkgBytes).digest("hex");
    const pkgSignatureBase64 = readExtensionPackageV2(pkgBytes)?.signature?.signatureBase64 || "";
    if (!/\S/.test(pkgSignatureBase64)) {
      throw new Error("Failed to read signatureBase64 from built extension package");
    }

    const keyDer = crypto.createPublicKey(keys.publicKeyPem).export({ type: "spki", format: "der" });
    const publisherKeyId = crypto.createHash("sha256").update(keyDer).digest("hex");

    return { keys, pkgBytes, pkgSha256, pkgSignatureBase64, publisherKeyId };
  })();
  return sampleHelloFixturePromise;
}

async function mockSampleHelloMarketplace(
  page: Parameters<typeof test>[0] extends { page: infer P } ? P : any,
  fixture: SampleHelloFixture,
  opts: {
    deprecated?: boolean;
    blocked?: boolean;
    malicious?: boolean;
    publisherRevoked?: boolean;
    scanStatusHeader?: string;
  } = {}
) {
  const scanStatusHeader = typeof opts.scanStatusHeader === "string" ? opts.scanStatusHeader : "passed";

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
        deprecated: Boolean(opts.deprecated),
        blocked: Boolean(opts.blocked),
        malicious: Boolean(opts.malicious),
        publisherRevoked: Boolean(opts.publisherRevoked),
        downloadCount: 0,
        latestVersion: "1.0.0",
        versions: [
          {
            version: "1.0.0",
            sha256: fixture.pkgSha256,
            uploadedAt: new Date().toISOString(),
            yanked: false,
            scanStatus: scanStatusHeader,
            signingKeyId: fixture.publisherKeyId
          }
        ],
        readme: "",
        publisherPublicKeyPem: fixture.keys.publicKeyPem,
        publisherKeys: [{ id: fixture.publisherKeyId, publicKeyPem: fixture.keys.publicKeyPem, revoked: false }],
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
        "X-Package-Sha256": fixture.pkgSha256,
        "X-Package-Signature": fixture.pkgSignatureBase64,
        "X-Package-Scan-Status": scanStatusHeader,
        "X-Package-Format-Version": "2",
        "X-Publisher": "formula",
        "X-Publisher-Key-Id": fixture.publisherKeyId
      },
      body: fixture.pkgBytes
    });
  });
}

test("install + run marketplace extension in browser (no CSP violations)", async ({ page }) => {
  const fixture = await getSampleHelloFixture();

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

  await mockSampleHelloMarketplace(page, fixture, { scanStatusHeader: "passed" });

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

test("blocked extension cannot install", async ({ page }) => {
  const fixture = await getSampleHelloFixture();
  await mockSampleHelloMarketplace(page, fixture, { blocked: true, scanStatusHeader: "passed" });

  await page.goto("/?extTest=1");
  await page.waitForFunction(() => Boolean((window as any).__formulaExtensionTest), null, { timeout: 30_000 });

  const result = await page.evaluate(async () => {
    const api = (window as any).__formulaExtensionTest;
    try {
      await api.installExtension("formula.sample-hello");
      return { ok: true, error: null, installed: await api.listInstalled() };
    } catch (error) {
      return {
        ok: false,
        error: { message: String((error as any)?.message ?? error) },
        installed: await api.listInstalled()
      };
    } finally {
      await api.dispose();
    }
  });

  expect(result.ok).toBe(false);
  expect(result.error?.message).toMatch(/blocked/i);
  expect(result.installed).toEqual([]);
});

test("deprecated extension shows warning and can require confirmation", async ({ page }) => {
  const fixture = await getSampleHelloFixture();
  await mockSampleHelloMarketplace(page, fixture, { deprecated: true, scanStatusHeader: "passed" });

  await page.goto("/?extTest=1");
  await page.waitForFunction(() => Boolean((window as any).__formulaExtensionTest), null, { timeout: 30_000 });

  const result = await page.evaluate(async () => {
    const api = (window as any).__formulaExtensionTest;
    const first = await (async () => {
      try {
        await api.installExtension("formula.sample-hello", null, {
          confirm: () => false,
        });
        return { ok: true };
      } catch (error) {
        return {
          ok: false,
          error: {
            message: String((error as any)?.message ?? error),
          },
          installed: await api.listInstalled(),
        };
      }
    })();

    const second = await api.installExtension("formula.sample-hello", null, {
      confirm: () => true,
    });
    const installed = await api.listInstalled();
    await api.dispose();
    return { first, second, installed };
  });

  expect(result.first.ok).toBe(false);
  expect((result.first as any).error?.message).toMatch(/cancelled/i);
  expect((result.first as any).installed).toEqual([]);
  expect((result.second as any)?.warnings?.some((w: any) => w.kind === "deprecated")).toBe(true);
  expect(result.second).toMatchObject({ id: "formula.sample-hello", version: "1.0.0", scanStatus: "passed" });
  expect((result.second as any)?.signingKeyId).toBe(fixture.publisherKeyId);
  expect(result.installed).toHaveLength(1);
});

test("scanStatus failure blocks install", async ({ page }) => {
  const fixture = await getSampleHelloFixture();
  await mockSampleHelloMarketplace(page, fixture, { scanStatusHeader: "failed" });

  await page.goto("/?extTest=1");
  await page.waitForFunction(() => Boolean((window as any).__formulaExtensionTest), null, { timeout: 30_000 });

  const result = await page.evaluate(async () => {
    const api = (window as any).__formulaExtensionTest;
    try {
      await api.installExtension("formula.sample-hello");
      return { ok: true, error: null, installed: await api.listInstalled() };
    } catch (error) {
      return {
        ok: false,
        error: { message: String((error as any)?.message ?? error) },
        installed: await api.listInstalled()
      };
    } finally {
      await api.dispose();
    }
  });

  expect(result.ok).toBe(false);
  expect(result.error?.message).toMatch(/scan status/i);
  expect(result.error?.message).toMatch(/failed/i);
  expect(result.installed).toEqual([]);
});
