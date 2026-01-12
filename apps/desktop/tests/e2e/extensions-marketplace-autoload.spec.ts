import { expect, test } from "@playwright/test";

import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { gotoDesktop, openExtensionsPanel, waitForDesktopReady } from "./helpers";

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

test.describe("Marketplace-installed extensions auto-load (desktop)", () => {
  // This suite installs a signed extension package, reloads the desktop shell, and waits for the
  // extension host + UI to rehydrate. It can be slow on CI and low-spec runners.
  test.describe.configure({ timeout: 120_000 });

  test("loads IndexedDB-installed extensions when the extension host is initialized after reload", async ({ page }) => {
    // Avoid permission modal flakiness in this suite; other e2e tests cover the explicit
    // permission prompt UI.
    await page.addInitScript(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).__formulaPermissionPrompt = async () => true;
    });

    const extensionId = "e2e.marketplace-autoload";
    const extensionVersion = "1.0.0";
    const commandId = "marketplaceTest.sayHello";
    const displayName = "Marketplace Autoload Test";

    const keys = generateEd25519KeyPair();

    const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "formula-e2e-marketplace-ext-"));
    const extensionDir = path.join(tmp, "extension");
    await fs.mkdir(path.join(extensionDir, "dist"), { recursive: true });

    const manifest = {
      name: "marketplace-autoload",
      displayName,
      version: extensionVersion,
      description: "E2E fixture extension for marketplace autoload.",
      publisher: "e2e",
      license: "UNLICENSED",
      main: "./dist/extension.js",
      module: "./dist/extension.mjs",
      browser: "./dist/extension.mjs",
      engines: {
        formula: "^1.0.0",
      },
      // No onCommand activation event: this extension must be started via onStartupFinished.
      activationEvents: ["onStartupFinished"],
      contributes: {
        commands: [
          {
            command: commandId,
            title: "Say Hello",
            category: "Marketplace Test",
          },
        ],
      },
      permissions: ["ui.commands"],
    };

    const entrypointSource = `
import { commands, ui } from "@formula/extension-api";

export async function activate(context) {
  context.subscriptions.push(
    await commands.registerCommand(${JSON.stringify(commandId)}, async () => {
      await ui.showMessage("Hello from marketplace!");
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
    const pkgSignatureBase64 = readExtensionPackageV2(pkgBytes)?.signature?.signatureBase64 || "";
    expect(pkgSignatureBase64).toMatch(/\S/);

    const keyDer = crypto.createPublicKey(keys.publicKeyPem).export({ type: "spki", format: "der" });
    const publisherKeyId = crypto.createHash("sha256").update(keyDer).digest("hex");

    await page.route(`**/api/extensions/${extensionId}`, async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          id: extensionId,
          name: "marketplace-autoload",
          displayName,
          publisher: "e2e",
          description: "E2E fixture extension for marketplace autoload.",
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
              signingKeyId: publisherKeyId,
              formatVersion: 2,
            },
          ],
          readme: "",
          publisherPublicKeyPem: keys.publicKeyPem,
          publisherKeys: [{ id: publisherKeyId, publicKeyPem: keys.publicKeyPem, revoked: false }],
          updatedAt: new Date().toISOString(),
          createdAt: new Date().toISOString(),
        }),
      });
    });

    await page.route(`**/api/extensions/${extensionId}/download/${extensionVersion}`, async (route) => {
      await route.fulfill({
        status: 200,
        headers: {
          "Content-Type": "application/vnd.formula.extension-package",
          "X-Package-Sha256": pkgSha256,
          "X-Package-Signature": pkgSignatureBase64,
          "X-Package-Format-Version": "2",
          "X-Publisher": "e2e",
          "X-Publisher-Key-Id": publisherKeyId,
          "X-Scan-Status": "passed",
        },
        body: pkgBytes,
      });
    });

    await gotoDesktop(page);

    const managerModuleUrl = viteFsUrl(path.join(repoRoot, "packages/extension-marketplace/src/index.ts"));

    // Install the extension into IndexedDB without opening the Marketplace panel.
    await page.evaluate(
      async ({ managerModuleUrl, extensionId }) => {
        const { WebExtensionManager } = await import(managerModuleUrl);
        const manager = new WebExtensionManager({ engineVersion: "1.0.0" });
        await manager.install(extensionId);

        // Pre-grant permissions so activation doesn't block on the UI prompt.
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
        };
        localStorage.setItem(key, JSON.stringify(existing));
      },
      { managerModuleUrl, extensionId },
    );

    // Reload to simulate desktop restart. The extension should be auto-loaded from IndexedDB
    // when the extension host boots (lazy, when opening the Extensions panel).
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    await openExtensionsPanel(page);
    await expect(page.getByTestId(`extension-card-${extensionId}`)).toBeVisible({ timeout: 30_000 });

    const runBtn = page.getByTestId(`run-command-${commandId}`);
    await expect(runBtn).toBeVisible({ timeout: 30_000 });
    await runBtn.click();

    await expect(page.getByTestId("toast-root")).toContainText("Hello from marketplace!");
  });
});
