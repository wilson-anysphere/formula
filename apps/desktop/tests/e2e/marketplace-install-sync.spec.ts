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

const { generateEd25519KeyPair } = signingPkg;
const { createExtensionPackageV2 } = extensionPackagePkg;

test.describe("Marketplace install sync", () => {
  test("installing + uninstalling via Marketplace updates command palette (no reload)", async ({ page }) => {
    test.setTimeout(180_000);

    await page.addInitScript(() => {
      // Avoid permission modal flakiness in this suite; other e2e tests cover explicit
      // permission prompt UI.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).__formulaPermissionPrompt = async () => true;
    });

    const extensionId = "e2e.marketplace-install-sync";
    const extensionVersion = "1.0.0";
    const displayName = "Marketplace Install Sync Test";
    const commandId = "marketplaceInstallSync.hello";
    const commandTitle = `Hello (${extensionId})`;
    const panelId = `${extensionId}.panel`;
    const panelTitle = `Panel (${extensionId})`;

    const keys = generateEd25519KeyPair();

    const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "formula-e2e-marketplace-sync-"));
    const extensionDir = path.join(tmp, "extension");
    await fs.mkdir(path.join(extensionDir, "dist"), { recursive: true });

    const manifest = {
      name: "marketplace-install-sync",
      displayName,
      version: extensionVersion,
      description: "E2E fixture extension for Marketplace → UI resync behavior.",
      publisher: "e2e",
      license: "UNLICENSED",
      main: "./dist/extension.js",
      module: "./dist/extension.mjs",
      browser: "./dist/extension.mjs",
      engines: {
        formula: "^1.0.0",
      },
      activationEvents: ["onStartupFinished", `onView:${panelId}`],
      contributes: {
        commands: [
          {
            command: commandId,
            title: commandTitle,
            category: "Marketplace Test",
          },
        ],
        panels: [
          {
            id: panelId,
            title: panelTitle,
            // Keep this panel out of the left dock so it doesn't collide with the Extensions panel itself.
            defaultDock: "right",
          },
        ],
      },
      permissions: ["ui.commands", "ui.panels"],
    };

    const entrypointSource = `
import { commands, events, ui } from "@formula/extension-api";

export async function activate(context) {
  context.subscriptions.push(
    events.onViewActivated(({ viewId }) => {
      if (viewId !== ${JSON.stringify(panelId)}) return;
      void (async () => {
        const created = await ui.createPanel(${JSON.stringify(panelId)}, { title: ${JSON.stringify(panelTitle)} });
        await created.webview.setHtml(\`<!doctype html><html><body><h1>${panelTitle}</h1></body></html>\`);
        context.subscriptions.push(created);
      })().catch(() => {});
    })
  );

  context.subscriptions.push(
    await commands.registerCommand(${JSON.stringify(commandId)}, async () => {
      await ui.showMessage("Hello from marketplace install sync!");
      return "ok";
    })
  );
}
`.trimStart();

    // `main` is required by the manifest validator but unused in browser builds.
    await fs.writeFile(path.join(extensionDir, "dist", "extension.js"), "module.exports = {};\n", "utf8");
    await fs.writeFile(path.join(extensionDir, "dist", "extension.mjs"), entrypointSource, "utf8");
    await fs.writeFile(path.join(extensionDir, "package.json"), JSON.stringify(manifest, null, 2), "utf8");

    const pkgBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem: keys.privateKeyPem });
    const pkgSha256 = crypto.createHash("sha256").update(pkgBytes).digest("hex");

    // Mock marketplace endpoints used by MarketplaceClient (/api).
    await page.route("**/api/search**", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          total: 1,
          results: [
            {
              id: extensionId,
              name: "marketplace-install-sync",
              displayName,
              publisher: "e2e",
              description: "E2E fixture extension for Marketplace → UI resync behavior.",
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

    // Provide a deterministic 404 fallback for unrelated extension detail requests so we don't
    // accidentally fetch HTML from the dev server.
    //
    // NOTE: Playwright route precedence can be subtle when multiple routes match. Use
    // `route.fallback()` for this test's extension id so the more specific stub handlers
    // (`/api/extensions/:id` + `/download/:version`) always win.
    const extensionApiPrefix = `/api/extensions/${encodeURIComponent(extensionId)}`;
    await page.route("**/api/extensions/**", async (route) => {
      const url = route.request().url();
      if (url.includes(extensionApiPrefix)) {
        await route.fallback();
        return;
      }
      await route.fulfill({ status: 404, body: "" });
    });

    await page.route(`**/api/extensions/${encodeURIComponent(extensionId)}`, async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          id: extensionId,
          name: "marketplace-install-sync",
          displayName,
          publisher: "e2e",
          description: "E2E fixture extension for Marketplace → UI resync behavior.",
          categories: [],
          tags: [],
          screenshots: [],
          verified: true,
          featured: false,
          deprecated: false,
          blocked: false,
          malicious: false,
          downloadCount: 0,
          latestVersion: extensionVersion,
          versions: [],
          readme: "",
          publisherPublicKeyPem: keys.publicKeyPem,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        }),
      });
    });

    await page.route(
      `**/api/extensions/${encodeURIComponent(extensionId)}/download/${encodeURIComponent(extensionVersion)}`,
      async (route) => {
        await route.fulfill({
          status: 200,
          contentType: "application/octet-stream",
          body: Buffer.from(pkgBytes),
          headers: {
            "x-package-sha256": pkgSha256,
            "x-package-format-version": "2",
            "x-package-scan-status": "passed",
            "x-publisher": "e2e",
          },
        });
      },
    );

    try {
      await gotoDesktop(page);
      await waitForDesktopReady(page);

      await page.getByRole("tab", { name: "View", exact: true }).click();
      await page.getByTestId("open-marketplace-panel").click();

      const panel = page.getByTestId("panel-marketplace");
      await expect(panel).toBeVisible();

      await panel.getByPlaceholder("Search extensions…").fill("sync");
      await panel.getByRole("button", { name: "Search", exact: true }).click();

      const resultRow = panel.locator(".marketplace-result").filter({ hasText: extensionId });
      await expect(resultRow).toBeVisible();

      await resultRow.getByRole("button", { name: "Install", exact: true }).click();
      await expect(resultRow).toContainText("Installed");

      const primary = process.platform === "darwin" ? "Meta" : "Control";

      // Verify the install triggers desktop UI re-sync without reloading: the command should show up
      // in the command palette immediately.
      await page.keyboard.press(`${primary}+Shift+P`);
      await expect(page.getByTestId("command-palette-input")).toBeVisible();
      await page.getByTestId("command-palette-input").fill(extensionId);

      const commandItem = page
        .getByTestId("command-palette-list")
        .locator("li.command-palette__item")
        .filter({ hasText: commandTitle });
      await expect(commandItem).toBeVisible({ timeout: 30_000 });
      await commandItem.first().click();

      await expect(page.getByTestId("toast-root")).toContainText("Hello from marketplace install sync!");

      // Verify the install triggers panel contribution sync too: open the Extensions panel and
      // open the contributed view (panel) without reloading.
      await page.getByRole("tab", { name: "Home", exact: true }).click();
      await page.getByTestId("open-extensions-panel").click();
      await expect(page.getByTestId("panel-extensions")).toBeVisible();

      await page.getByTestId("panel-extensions").getByTestId(`open-panel-${panelId}`).click();
      await expect(page.getByTestId(`panel-${panelId}`)).toBeVisible({ timeout: 30_000 });

      const webview = page.getByTestId(`extension-webview-${panelId}`);
      await expect(webview).toBeVisible();
      await expect(page.frameLocator(`[data-testid="extension-webview-${panelId}"]`).getByRole("heading", { name: panelTitle })).toBeVisible({
        timeout: 30_000,
      });

      // Switch back to Marketplace before clicking Uninstall.
      await page.getByTestId("dock-tab-marketplace").click();
      await expect(page.getByTestId("panel-marketplace")).toBeVisible();

      // Re-render the search results so the Marketplace panel can show the uninstall button.
      await panel.getByRole("button", { name: "Search", exact: true }).click();
      const installedRow = panel.locator(".marketplace-result").filter({ hasText: extensionId });
      await expect(installedRow).toBeVisible();

      await installedRow.getByRole("button", { name: "Uninstall", exact: true }).click();
      await expect(installedRow).toContainText("Uninstalled");

      // Uninstall should close any open panels created by the extension.
      await expect(page.getByTestId(`dock-tab-${panelId}`)).toHaveCount(0);

      // Confirm the desktop command registry updates after uninstall (command should disappear).
      await page.keyboard.press(`${primary}+Shift+P`);
      await expect(page.getByTestId("command-palette-input")).toBeVisible();
      await page.getByTestId("command-palette-input").fill(extensionId);
      await expect(page.getByTestId("command-palette-list")).toContainText("No matching commands", { timeout: 30_000 });
    } finally {
      await fs.rm(tmp, { recursive: true, force: true });
    }
  });
});
