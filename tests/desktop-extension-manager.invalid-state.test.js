import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";

import { ExtensionManager } from "../apps/desktop/src/marketplace/extensionManager.js";
import extensionPackagePkg from "../shared/extension-package/index.js";

const { createExtensionPackageV2 } = extensionPackagePkg;

function sha256Hex(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

test("ExtensionManager.install tolerates invalid JSON state file", async (t) => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-ext-manager-invalid-state-"));
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const extensionDir = path.join(tmpRoot, "ext");
  await fs.mkdir(path.join(extensionDir, "dist"), { recursive: true });
  await fs.writeFile(
    path.join(extensionDir, "package.json"),
    JSON.stringify(
      {
        name: "hello",
        publisher: "formula",
        version: "1.0.0",
        main: "./dist/extension.js",
        engines: { formula: "^1.0.0" },
        activationEvents: ["onCommand:hello.world"],
        contributes: {
          commands: [{ command: "hello.world", title: "Hello World" }],
        },
        permissions: ["ui.commands"],
      },
      null,
      2,
    ),
  );
  await fs.writeFile(
    path.join(extensionDir, "dist", "extension.js"),
    [
      'const formula = require("@formula/extension-api");',
      "async function activate(context) {",
      '  context.subscriptions.push(await formula.commands.registerCommand("hello.world", async () => "hi"));',
      "}",
      "module.exports = { activate };",
      "",
    ].join("\n"),
  );

  // Corrupt the state file before installing.
  await fs.mkdir(path.dirname(statePath), { recursive: true });
  await fs.writeFile(statePath, "{", "utf8");

  const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
  const publicKeyPem = publicKey.export({ type: "spki", format: "pem" });
  const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" });

  const packageBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem });
  const packageSha256 = sha256Hex(packageBytes);

  const extensionId = "formula.hello";
  /** @type {any} */
  const marketplaceClient = {
    async getExtension(id) {
      if (id !== extensionId) return null;
      return { id, latestVersion: "1.0.0", publisherPublicKeyPem: publicKeyPem };
    },
    async downloadPackage(id, version) {
      if (id !== extensionId || version !== "1.0.0") return null;
      return {
        bytes: Buffer.from(packageBytes),
        sha256: packageSha256,
        formatVersion: 2,
        signatureBase64: null,
        publisher: "formula",
      };
    },
  };

  const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  const record = await manager.install(extensionId);
  assert.equal(record.id, extensionId);
  assert.equal(record.version, "1.0.0");
  assert.ok(Array.isArray(record.files));

  const installed = await manager.getInstalled(extensionId);
  assert.ok(installed);
  assert.equal(installed.version, "1.0.0");

  // State should have been rewritten as valid JSON.
  const state = JSON.parse(await fs.readFile(statePath, "utf8"));
  assert.ok(state.installed?.[extensionId]);
});

