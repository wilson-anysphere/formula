import { expect, test } from "@playwright/test";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";

const requireFromHere = createRequire(import.meta.url);

// shared/* is CommonJS (shared/package.json sets type=commonjs).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const signingPkg: any = requireFromHere("../../../../shared/crypto/signing.js");
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const extensionPackagePkg: any = requireFromHere("../../../../shared/extension-package/index.js");

const { generateEd25519KeyPair } = signingPkg;
const { createExtensionPackageV2, readExtensionPackageV2 } = extensionPackagePkg;

test("uninstall clears permission grants + extension storage so reinstall behaves like a clean install", async ({ page }) => {
  const keys = generateEd25519KeyPair();

  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-web-ext-uninstall-"));
  try {
    const extDir = path.join(tmpRoot, "ext");
    await fs.mkdir(path.join(extDir, "dist"), { recursive: true });

    const extensionId = "test.state-test";
    const manifest = {
      name: "state-test",
      displayName: "State Test",
      version: "1.0.0",
      description: "",
      publisher: "test",
      main: "./dist/extension.js",
      browser: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCommand:stateTest.set", "onCommand:stateTest.get"],
      contributes: {
        commands: [
          { command: "stateTest.set", title: "Set Storage (Test)" },
          { command: "stateTest.get", title: "Get Storage (Test)" },
        ],
      },
      permissions: ["ui.commands", "storage"],
    };

    await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2));
    await fs.writeFile(
      path.join(extDir, "dist", "extension.js"),
      `import * as formula from "@formula/extension-api";
 export async function activate(context) {
   context.subscriptions.push(await formula.commands.registerCommand("stateTest.set", async (key, value) => {
      await formula.storage.set(String(key), value);
     return true;
   }));

   context.subscriptions.push(await formula.commands.registerCommand("stateTest.get", async (key) => {
     const value = await formula.storage.get(String(key));
     return value === undefined ? null : value;
   }));
 }
 `,
      "utf8",
    );

    const pkgBytes = await createExtensionPackageV2(extDir, { privateKeyPem: keys.privateKeyPem });
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
          name: "state-test",
          displayName: "State Test",
          publisher: "test",
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
          createdAt: new Date().toISOString(),
        }),
      });
    });

    await page.route(`**/api/extensions/${extensionId}/download/1.0.0`, async (route) => {
      await route.fulfill({
        status: 200,
        headers: {
          "Content-Type": "application/vnd.formula.extension-package",
          "X-Package-Sha256": pkgSha256,
          "X-Package-Signature": pkgSignatureBase64,
          "X-Package-Format-Version": "2",
          "X-Package-Scan-Status": "passed",
          "X-Publisher": "test",
          "X-Publisher-Key-Id": publisherKeyId,
        },
        body: pkgBytes,
      });
    });

    await page.goto("/?extTest=1");
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionTest), null, { timeout: 30_000 });

    const result = await page.evaluate(async (id) => {
      const api = (window as any).__formulaExtensionTest;
      api.clearPermissionPrompts();

      await api.installExtension(id);
      await api.executeCommand("stateTest.set", "k", "v1");
      const before = await api.executeCommand("stateTest.get", "k");
      const promptsBefore = api.getPermissionPrompts();

      api.clearPermissionPrompts();
      await api.uninstallExtension(id);

      await api.installExtension(id);
      const after = await api.executeCommand("stateTest.get", "k");
      const promptsAfter = api.getPermissionPrompts();

      await api.dispose();

      return { before, after, promptsBefore, promptsAfter };
    }, extensionId);

    expect(result.before).toBe("v1");
    expect(result.after).toBeNull();

    const promptPerms = (items: any[]) => items.map((p) => (Array.isArray(p?.permissions) ? p.permissions : []));

    // First install should prompt at least once for storage and/or ui.commands.
    expect(promptPerms(result.promptsBefore).flat()).toContain("storage");

    // After uninstall + reinstall, permission prompts should be required again.
    expect(promptPerms(result.promptsAfter).flat()).toContain("storage");
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});
