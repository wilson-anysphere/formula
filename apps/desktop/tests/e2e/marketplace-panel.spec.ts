import { expect, test } from "@playwright/test";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

// CJS helpers (shared/* is CommonJS). Playwright's TS loader may not always expose
// named exports for CJS modules, so fall back to `.default`.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const signingImport: any = await import("../../../../shared/crypto/signing.js");
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const extensionPackageImport: any = await import("../../../../shared/extension-package/index.js");

const signingPkg: any = signingImport?.default ?? signingImport;
const extensionPackagePkg: any = extensionPackageImport?.default ?? extensionPackageImport;

const generateEd25519KeyPair: (() => { publicKeyPem: string; privateKeyPem: string }) | undefined =
  signingPkg.generateEd25519KeyPair ?? signingPkg.default?.generateEd25519KeyPair;
const createExtensionPackageV2:
  | ((extensionDir: string, options: { privateKeyPem: string }) => Promise<Uint8Array>)
  | undefined = extensionPackagePkg.createExtensionPackageV2 ?? extensionPackagePkg.default?.createExtensionPackageV2;

if (typeof generateEd25519KeyPair !== "function") {
  throw new Error("Missing generateEd25519KeyPair export from shared/crypto/signing.js");
}
if (typeof createExtensionPackageV2 !== "function") {
  throw new Error("Missing createExtensionPackageV2 export from shared/extension-package/index.js");
}

test.describe("marketplace panel", () => {
  test.beforeEach(async ({ page }) => {
    await page.addInitScript(() => {
      try {
        localStorage.setItem(
          "formula.extensionHost.permissions",
          JSON.stringify({
            "formula.e2e-events": { storage: true },
          }),
        );
      } catch {
        // ignore
      }
    });
  });

  test("opens from the ribbon (View → Panels)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForDesktopReady(page);

    const viewTab = page.getByRole("tab", { name: "View", exact: true });
    await expect(viewTab).toBeVisible();
    await viewTab.click();

    const openMarketplacePanel = page.getByTestId("ribbon-root").getByTestId("open-marketplace-panel");
    await expect(openMarketplacePanel).toBeVisible();
    await openMarketplacePanel.click();

    const panel = page.getByTestId("panel-marketplace");
    await expect(panel).toBeVisible();

    // The command is wired via `toggleDockPanel`, so clicking again should close it.
    await openMarketplacePanel.click();
    await expect(panel).toHaveCount(0);
  });

  test("can render search results (stubbed /api/search)", async ({ page }) => {
    await page.route("**/api/search**", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          total: 1,
          results: [
            {
              id: "acme.hello",
              name: "hello",
              displayName: "Hello Extension",
              publisher: "acme",
              description: "A test extension returned by Playwright stubs.",
              latestVersion: "1.0.0",
              verified: false,
              featured: false,
              categories: [],
              tags: [],
              screenshots: [],
              downloadCount: 0,
              updatedAt: new Date().toISOString(),
            },
          ],
          nextCursor: null,
        }),
      });
    });

    // The panel does a best-effort per-item details fetch (`/api/extensions/:id`). In Vite dev
    // servers this could return HTML; stub it to a clean 404 so the panel consistently falls
    // back to the summary response.
    await page.route("**/api/extensions/**", async (route) => {
      await route.fulfill({ status: 404, body: "" });
    });

    await gotoDesktop(page);
    await waitForDesktopReady(page);

    // The Marketplace panel should be usable without eagerly loading the extension host.
    // (Loading extensions spins up Workers; keep that work lazy until the user opens the
    // Extensions panel or runs a command.)
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager), undefined, {
      timeout: 30_000,
    });
    expect(await page.evaluate(() => (window as any).__formulaExtensionHostManager?.ready)).toBe(false);

    await page.getByRole("tab", { name: "View", exact: true }).click();
    await page.getByTestId("ribbon-root").getByTestId("open-marketplace-panel").click();

    const panel = page.getByTestId("panel-marketplace");
    await expect(panel).toBeVisible();

    await panel.getByPlaceholder("Search extensions…").fill("hello");
    await panel.getByRole("button", { name: "Search", exact: true }).click();

    await expect(panel).toContainText("Hello Extension (acme.hello)");

    expect(await page.evaluate(() => (window as any).__formulaExtensionHostManager?.ready)).toBe(false);
  });

  test("installing via Marketplace does not eagerly start the extension host", async ({ page }) => {
    test.setTimeout(120_000);

    const keys = generateEd25519KeyPair();
    const extensionId = "e2e.marketplace-lazy-install";
    const extensionVersion = "1.0.0";
    const displayName = "Marketplace Lazy Install Test";

    const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "formula-e2e-marketplace-lazy-"));
    const extensionDir = path.join(tmp, "extension");
    await fs.mkdir(path.join(extensionDir, "dist"), { recursive: true });

    const manifest = {
      name: "marketplace-lazy-install",
      displayName,
      version: extensionVersion,
      description: "E2E fixture extension for verifying marketplace lazy-load behavior.",
      publisher: "e2e",
      license: "UNLICENSED",
      main: "./dist/extension.js",
      module: "./dist/extension.mjs",
      browser: "./dist/extension.mjs",
      engines: { formula: "^1.0.0" },
      activationEvents: [],
      contributes: { commands: [] },
      permissions: [],
    };

    await fs.writeFile(path.join(extensionDir, "package.json"), JSON.stringify(manifest, null, 2), "utf8");
    await fs.writeFile(path.join(extensionDir, "dist", "extension.js"), "module.exports = {};\n", "utf8");
    await fs.writeFile(path.join(extensionDir, "dist", "extension.mjs"), "export {};\n", "utf8");

    const pkgBytes: Uint8Array = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });
    const pkgSha256 = crypto.createHash("sha256").update(Buffer.from(pkgBytes)).digest("hex");

    await page.route("**/api/search**", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          total: 1,
          results: [
            {
              id: extensionId,
              name: "marketplace-lazy-install",
              displayName,
              publisher: "e2e",
              description: "E2E fixture extension for verifying marketplace lazy-load behavior.",
              latestVersion: extensionVersion,
              verified: true,
              featured: false,
              categories: [],
              tags: [],
              screenshots: [],
              downloadCount: 0,
              updatedAt: new Date().toISOString(),
            },
          ],
          nextCursor: null,
        }),
      });
    });

    await page.route(`**/api/extensions/${encodeURIComponent(extensionId)}`, async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          id: extensionId,
          name: "marketplace-lazy-install",
          displayName,
          publisher: "e2e",
          description: "E2E fixture extension for verifying marketplace lazy-load behavior.",
          latestVersion: extensionVersion,
          verified: true,
          featured: false,
          categories: [],
          tags: [],
          screenshots: [],
          downloadCount: 0,
          updatedAt: new Date().toISOString(),
          versions: [
            {
              version: extensionVersion,
              sha256: pkgSha256,
              uploadedAt: new Date().toISOString(),
              yanked: false,
              scanStatus: "passed",
              formatVersion: 2,
            },
          ],
          readme: "",
          publisherPublicKeyPem: keys.publicKeyPem,
          createdAt: new Date().toISOString(),
          deprecated: false,
          blocked: false,
          malicious: false,
        }),
      });
    });

    await page.route(
      `**/api/extensions/${encodeURIComponent(extensionId)}/download/${encodeURIComponent(extensionVersion)}`,
      async (route) => {
        await route.fulfill({
          status: 200,
          headers: {
            "content-type": "application/octet-stream",
            "x-package-sha256": pkgSha256,
            "x-package-format-version": "2",
            "x-package-scan-status": "passed",
            "x-publisher": "e2e",
          },
          body: Buffer.from(pkgBytes),
        });
      },
    );

    try {
      await gotoDesktop(page);
      await waitForDesktopReady(page);

      await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager), undefined, {
        timeout: 30_000,
      });

      const initialHostState = await page.evaluate(() => {
        return {
          ready: Boolean((window as any).__formulaExtensionHostManager?.ready),
          loadedCount: Number((window as any).__formulaExtensionHost?.listExtensions?.()?.length ?? 0),
        };
      });
      expect(initialHostState.ready).toBe(false);
      expect(initialHostState.loadedCount).toBe(0);

      await page.getByRole("tab", { name: "View", exact: true }).click();
      await page.getByTestId("ribbon-root").getByTestId("open-marketplace-panel").click();

      const panel = page.getByTestId("panel-marketplace");
      await expect(panel).toBeVisible();

      await panel.getByPlaceholder("Search extensions…").fill("lazy");
      await panel.getByRole("button", { name: "Search", exact: true }).click();

      const resultRow = panel.locator(".marketplace-result").filter({ hasText: extensionId });
      await expect(resultRow).toBeVisible();

      await resultRow.getByRole("button", { name: "Install", exact: true }).click();
      await expect(resultRow).toContainText("Installed", { timeout: 30_000 });

      const afterInstallHostState = await page.evaluate(() => {
        return {
          ready: Boolean((window as any).__formulaExtensionHostManager?.ready),
          loadedCount: Number((window as any).__formulaExtensionHost?.listExtensions?.()?.length ?? 0),
        };
      });
      expect(afterInstallHostState.ready).toBe(false);
      expect(afterInstallHostState.loadedCount).toBe(0);
    } finally {
      await fs.rm(tmp, { recursive: true, force: true });
    }
  });
});
