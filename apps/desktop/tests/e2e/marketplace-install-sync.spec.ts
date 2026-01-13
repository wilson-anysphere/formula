import { expect, test } from "@playwright/test";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { gotoDesktop, openExtensionsPanel, waitForDesktopReady } from "./helpers";

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
          versions: [
            {
              version: extensionVersion,
              sha256: pkgSha256,
              uploadedAt: new Date().toISOString(),
              yanked: false,
              scanStatus: "passed",
            },
          ],
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

      // Open the Marketplace panel via the CommandRegistry so this test does not depend on the ribbon
      // being mounted/interactive in the web-based desktop harness.
      await page.evaluate(async () => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const registry: any = (window as any).__formulaCommandRegistry;
        if (!registry) throw new Error("Missing window.__formulaCommandRegistry (desktop e2e harness)");
        await registry.executeCommand("view.togglePanel.marketplace");
      });

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
      await openExtensionsPanel(page);

      await page.getByTestId("panel-extensions").getByTestId(`open-panel-${panelId}`).click();
      await expect(page.getByTestId(`panel-${panelId}`)).toBeVisible({ timeout: 30_000 });

      const webview = page.getByTestId(`extension-webview-${panelId}`);
      await expect(webview).toBeVisible();
      await expect(page.frameLocator(`[data-testid="extension-webview-${panelId}"]`).getByRole("heading", { name: panelTitle })).toBeVisible({
        timeout: 30_000,
      });

      // Seed some persisted state owned by the extension so we can assert uninstall cleans it up.
      await page.evaluate(
        ({ extensionId, panelId, panelTitle }) => {
          try {
            localStorage.setItem(
              `formula.extensionHost.storage.${extensionId}`,
              JSON.stringify({ foo: "bar" }),
            );
          } catch {
            // ignore
          }

          try {
            const seedKey = "formula.extensions.contributedPanels.v1";
            let existing: any = {};
            try {
              const raw = localStorage.getItem(seedKey);
              existing = raw ? JSON.parse(raw) : {};
            } catch {
              existing = {};
            }
            if (!existing || typeof existing !== "object" || Array.isArray(existing)) existing = {};
            existing[panelId] = { extensionId, title: panelTitle };
            localStorage.setItem(seedKey, JSON.stringify(existing));
          } catch {
            // ignore
          }
        },
        { extensionId, panelId, panelTitle },
      );

      // Ensure the permission store includes this extension before uninstall so the cleanup assertion
      // is meaningful (the suite also pre-seeds `formula.e2e-events`).
      const permissionsBefore = await page.evaluate(() => localStorage.getItem("formula.extensionHost.permissions"));
      expect(permissionsBefore).not.toBeNull();
      const parsedBefore = permissionsBefore ? JSON.parse(String(permissionsBefore)) : {};
      expect(parsedBefore[extensionId]).toBeTruthy();
      expect(parsedBefore["formula.e2e-events"]).toBeTruthy();

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

      // Verify uninstall cleans up persisted state so a reinstall behaves like a clean slate.
      await page.waitForFunction(
        ({ extensionId, panelId }) => {
          if (localStorage.getItem(`formula.extensionHost.storage.${extensionId}`) !== null) return false;
          const seedsRaw = localStorage.getItem("formula.extensions.contributedPanels.v1");
          if (seedsRaw != null) {
            try {
              const parsed = JSON.parse(seedsRaw);
              if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return false;
              // Ensure this extension's panel seed is removed, but allow other extensions
              // (e.g. built-in desktop e2e fixtures) to keep their seeded panels.
              if (Object.prototype.hasOwnProperty.call(parsed, panelId)) return false;
              for (const seed of Object.values(parsed)) {
                const owner = typeof (seed as any)?.extensionId === "string" ? String((seed as any).extensionId) : "";
                if (owner === extensionId) return false;
              }
            } catch {
              return false;
            }
          }
          try {
            const permissionsRaw = localStorage.getItem("formula.extensionHost.permissions");
            if (!permissionsRaw) return false;
            const parsed = JSON.parse(permissionsRaw);
            if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return false;
            // Ensure only this extension was removed (others should remain).
            if (Object.prototype.hasOwnProperty.call(parsed, extensionId)) return false;
            if (!Object.prototype.hasOwnProperty.call(parsed, "formula.e2e-events")) return false;
          } catch {
            return false;
          }
          return true;
        },
        { extensionId, panelId },
        { timeout: 30_000 },
      );

      const persistedState = await page.evaluate(async ({ extensionId }) => {
        const storageKey = `formula.extensionHost.storage.${extensionId}`;
        const permissionsKey = "formula.extensionHost.permissions";
        const seedKey = "formula.extensions.contributedPanels.v1";

        const readDb = async () => {
          const openReq = indexedDB.open("formula.webExtensions");
          const db = await new Promise<IDBDatabase>((resolve, reject) => {
            openReq.onerror = () => reject(openReq.error ?? new Error("Failed to open IndexedDB"));
            openReq.onsuccess = () => resolve(openReq.result);
          });
          try {
            const tx = db.transaction(["installed", "packages"], "readonly");
            const installedStore = tx.objectStore("installed");
            const packagesStore = tx.objectStore("packages");

            const installedPresent = await new Promise<boolean>((resolve, reject) => {
              const req = installedStore.get(extensionId);
              req.onerror = () => reject(req.error ?? new Error("IndexedDB get failed"));
              req.onsuccess = () => resolve(req.result !== undefined);
            });

            const packagesCount = await new Promise<number>((resolve, reject) => {
              try {
                const index = packagesStore.index("byId");
                const req = index.count(extensionId);
                req.onerror = () => reject(req.error ?? new Error("IndexedDB count failed"));
                req.onsuccess = () => resolve(Number(req.result ?? 0));
              } catch (err) {
                reject(err);
              }
            });

            await new Promise<void>((resolve, reject) => {
              tx.oncomplete = () => resolve();
              tx.onerror = () => reject(tx.error ?? new Error("IndexedDB tx failed"));
              tx.onabort = () => reject(tx.error ?? new Error("IndexedDB tx aborted"));
            });

            return { installedPresent, packagesCount };
          } finally {
            db.close();
          }
        };

        // Poll briefly in case UI updated before the IndexedDB transactions finished.
        const start = Date.now();
        let dbState = { installedPresent: true, packagesCount: -1 };
        for (;;) {
          dbState = await readDb();
          if (!dbState.installedPresent && dbState.packagesCount === 0) break;
          if (Date.now() - start > 3_000) break;
          await new Promise<void>((r) => setTimeout(r, 50));
        }

        let permissions: any = null;
        try {
          const raw = localStorage.getItem(permissionsKey);
          permissions = raw ? JSON.parse(raw) : null;
        } catch {
          permissions = null;
        }

        return {
          storage: localStorage.getItem(storageKey),
          seeds: localStorage.getItem(seedKey),
          permissions,
          dbState,
        };
      }, { extensionId });

      expect(persistedState.storage).toBeNull();
      expect(persistedState.permissions?.[extensionId]).toBeUndefined();
      expect(persistedState.permissions?.["formula.e2e-events"]).toBeTruthy();
      // Contributed panel seeds may exist for built-in extensions (e.g. sample-hello) even when
      // the marketplace-installed extension is removed. Ensure this uninstall cleaned up only
      // the removed extension's seed entries.
      if (persistedState.seeds != null) {
        const seeds = JSON.parse(String(persistedState.seeds));
        expect(seeds?.[panelId]).toBeUndefined();
        for (const seed of Object.values(seeds ?? {})) {
          const owner = typeof (seed as any)?.extensionId === "string" ? String((seed as any).extensionId) : "";
          expect(owner).not.toBe(extensionId);
        }
      }
      expect(persistedState.dbState).toEqual({ installedPresent: false, packagesCount: 0 });
    } finally {
      await fs.rm(tmp, { recursive: true, force: true });
    }
  });
});
