import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { expect, test } from "@playwright/test";

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

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "../../../..");

function viteFsUrl(absPath: string) {
  return `/@fs${absPath}`;
}

test.describe("IndexedDB extension install corruption", () => {
  test("quarantines corrupted installs and repairs via re-download", async ({ page }) => {
    test.setTimeout(180_000);
    await page.addInitScript(() => {
      // Avoid permission modal flakiness in this suite; other e2e tests cover the
      // explicit permission prompt UI.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).__formulaPermissionPrompt = async () => true;
    });
    const extensionId = "test.corrupt-ext";
    const version = "1.0.0";
    const commandId = "corruptExt.hello";

    const keys = generateEd25519KeyPair();
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-corrupt-"));
    const extDir = path.join(tmpRoot, "ext");
    await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

    const manifest = {
      name: "corrupt-ext",
      publisher: "test",
      version,
      main: "./dist/extension.js",
      browser: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      permissions: ["ui.commands"],
      contributes: { commands: [{ command: commandId, title: "Hello (corrupt-ext)" }] },
    };

    await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
    await fs.writeFile(
      path.join(extDir, "dist", "extension.js"),
      `import * as formula from "@formula/extension-api";\nexport async function activate(context) {\n  context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
        commandId,
      )}, async () => {\n    await formula.ui.showMessage("Hello from repaired extension");\n    return "ok";\n  }));\n}\n`,
    );

    const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
    const pkgSha256 = crypto.createHash("sha256").update(pkgBytes).digest("hex");

    // Mock marketplace endpoints used by MarketplaceClient (/api).
    await page.route(`**/api/extensions/${encodeURIComponent(extensionId)}`, async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          id: extensionId,
          name: "corrupt-ext",
          displayName: "corrupt-ext",
          publisher: "test",
          description: "",
          latestVersion: version,
          verified: true,
          featured: false,
          categories: [],
          tags: [],
          screenshots: [],
          downloadCount: 0,
          updatedAt: new Date().toISOString(),
          versions: [],
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
      `**/api/extensions/${encodeURIComponent(extensionId)}/download/${encodeURIComponent(version)}`,
      async (route) => {
        await route.fulfill({
          status: 200,
          contentType: "application/octet-stream",
          body: Buffer.from(pkgBytes),
          headers: {
            "x-package-sha256": pkgSha256,
            "x-package-format-version": "2",
            "x-publisher": "test",
            "x-package-scan-status": "passed",
          },
        });
      },
    );

    await gotoDesktop(page, "/", { waitForIdle: false, appReadyTimeoutMs: 120_000 });

    // Install the extension into IndexedDB.
    const webExtensionManagerUrl = viteFsUrl(
      path.join(repoRoot, "packages/extension-marketplace/src/WebExtensionManager.ts")
    );
    await page.evaluate(
      async ({ webExtensionManagerUrl, extensionId, version: v }) => {
        const { WebExtensionManager } = await import(webExtensionManagerUrl);
        const manager = new WebExtensionManager();
        await manager.install(extensionId, v);
      },
      { webExtensionManagerUrl, extensionId, version },
    );

    // Corrupt the stored package bytes (simulates IndexedDB corruption/partial writes).
    await page.evaluate(
      async ({ extensionId, version: v }) => {
        await new Promise<void>((resolve, reject) => {
          const req = indexedDB.open("formula.webExtensions", 1);
          req.onerror = () => reject(req.error ?? new Error("Failed to open IndexedDB"));
          req.onsuccess = () => {
            const db = req.result;
            const tx = db.transaction(["packages"], "readwrite");
            const store = tx.objectStore("packages");
            const key = `${extensionId}@${v}`;
            const getReq = store.get(key);
            getReq.onsuccess = () => {
              const record: any = getReq.result;
              if (!record) throw new Error(`Missing package record for ${key}`);
              const bytes = new Uint8Array(record.bytes);
              bytes[0] ^= 0xff;
              record.bytes = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
              store.put(record);
            };
            tx.oncomplete = () => {
              db.close();
              resolve();
            };
            tx.onerror = () => {
              const err = tx.error;
              db.close();
              reject(err);
            };
          };
        });
      },
      { extensionId, version },
    );

    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page, { waitForIdle: false, appReadyTimeoutMs: 120_000 });

    // Opening the extensions panel triggers extension host startup + auto-load of installed extensions.
    await openExtensionsPanel(page);
    // Wait for the Extensions panel to finish lazy-loading the extension host manager (built-ins +
    // marketplace-installed extensions). Until `manager.ready` flips, the panel only shows a
    // "Loading extensions…" placeholder and does not render the IndexedDB install status section.
    await expect(page.getByText("Installed (IndexedDB)")).toBeVisible({ timeout: 30_000 });

    // Corrupted extension should be quarantined (not loaded into host).
    await expect(page.locator(`[data-testid="run-command-${commandId}"]`)).toHaveCount(0);

    // UI should reflect corrupted state and expose a repair button.
    await expect(page.getByTestId(`installed-extension-${extensionId}`)).toBeVisible({ timeout: 30_000 });
    await expect(page.getByTestId(`installed-extension-status-${extensionId}`)).toContainText("Corrupted");
    await page.getByTestId(`repair-extension-${extensionId}`).click();

    // Repair should clear corrupted state and load the extension.
    await expect(page.getByTestId(`installed-extension-status-${extensionId}`)).toContainText("OK");
    await expect(page.getByTestId(`run-command-${commandId}`)).toBeVisible();

    // Ensure the repaired extension can execute.
    await page.getByTestId(`run-command-${commandId}`).click();
    await expect(page.getByTestId("toast-root")).toContainText("Hello from repaired extension");

    await fs.rm(tmpRoot, { recursive: true, force: true });
  });
});
