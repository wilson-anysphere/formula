import { expect, test } from "@playwright/test";
import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { gotoDesktop, openExtensionsPanel } from "./helpers";

// CJS helpers (shared/* is CommonJS).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const signingPkg: any = await import("../../../../shared/crypto/signing.js");
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const extensionPackagePkg: any = await import("../../../../shared/extension-package/index.js");

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

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "../../../..");

test.describe("Desktop Marketplace (browser/IndexedDB install)", () => {
  test("installs sample-hello from the marketplace and can run sampleHello.sumSelection", async ({ page }) => {
    const keys = generateEd25519KeyPair();
    const extensionDir = path.join(repoRoot, "extensions/sample-hello");
    const pkgBytes: Uint8Array = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });

    const sha256 = crypto.createHash("sha256").update(Buffer.from(pkgBytes)).digest("hex");

    // Marketplace routes
    await page.route("**/api/search*", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          total: 1,
          results: [
            {
              id: "formula.sample-hello",
              name: "sample-hello",
              displayName: "Sample Hello",
              publisher: "formula",
              description: "Sample extension demonstrating commands and panels.",
              latestVersion: "1.0.0",
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

    await page.route("**/api/extensions/formula.sample-hello", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          id: "formula.sample-hello",
          name: "sample-hello",
          displayName: "Sample Hello",
          publisher: "formula",
          description: "Sample extension demonstrating commands and panels.",
          latestVersion: "1.0.0",
          verified: true,
          featured: false,
          categories: [],
          tags: [],
          screenshots: [],
          downloadCount: 0,
          updatedAt: new Date().toISOString(),
          versions: [
            {
              version: "1.0.0",
              sha256,
              scanStatus: "passed",
              uploadedAt: new Date().toISOString(),
              yanked: false,
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

    await page.route("**/api/extensions/formula.sample-hello/download/1.0.0", async (route) => {
      await route.fulfill({
        status: 200,
        headers: {
          "content-type": "application/octet-stream",
          "x-package-sha256": sha256,
          "x-package-format-version": "2",
          "x-publisher": "formula",
        },
        body: Buffer.from(pkgBytes),
      });
    });

    await gotoDesktop(page);

    // Pre-grant permissions so the extension can activate + run without an interactive prompt.
    await page.evaluate(() => {
      const extensionId = "formula.sample-hello";
      const key = "formula.extensionHost.permissions";
      const existing = (() => {
        try {
          const raw = localStorage.getItem(key);
          return raw ? JSON.parse(raw) : {};
        } catch {
          return {};
        }
      })();

      existing[extensionId] = {
        ...(existing[extensionId] ?? {}),
        "ui.commands": true,
        "ui.panels": true,
        "cells.read": true,
        "cells.write": true,
        clipboard: true,
      };

      localStorage.setItem(key, JSON.stringify(existing));
    });

    // Seed a selection so sumSelection has something to operate on.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = window.__formulaApp as any;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    // Install via Marketplace UI.
    await page.getByRole("tab", { name: "View", exact: true }).click();
    await page.getByTestId("ribbon-root").getByTestId("open-marketplace-panel").click();
    await expect(page.getByTestId("panel-marketplace")).toBeVisible();

    await page.getByTestId("marketplace-search-input").fill("hello");
    await page.getByTestId("marketplace-search-button").click();

    await expect(page.getByTestId("marketplace-install-formula.sample-hello")).toBeVisible();
    await page.getByTestId("marketplace-install-formula.sample-hello").click();

    // Wait for the UI to reflect the new state.
    await expect(page.getByTestId("marketplace-uninstall-formula.sample-hello")).toBeVisible({ timeout: 60_000 });

    // Verify extension is runnable.
    await page.getByRole("tab", { name: "Home", exact: true }).click();
    await openExtensionsPanel(page);

    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible({ timeout: 60_000 });
    await page.getByTestId("run-command-sampleHello.sumSelection").dispatchEvent("click");

    await expect
      .poll(async () => page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A3")), { timeout: 60_000 })
      .toBe("10");
  });
});
