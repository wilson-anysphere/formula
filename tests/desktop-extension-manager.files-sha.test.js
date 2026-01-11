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

async function pathExists(filePath) {
  try {
    await fs.stat(filePath);
    return true;
  } catch (error) {
    if (error && (error.code === "ENOENT" || error.code === "ENOTDIR")) return false;
    throw error;
  }
}

test("ExtensionManager.install rejects mismatched filesSha256 and cleans up extracted install dir", async (t) => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-ext-files-sha-"));
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
      },
      null,
      2
    )
  );
  await fs.writeFile(path.join(extensionDir, "dist", "extension.js"), 'module.exports = { activate() {} };\n', "utf8");

  const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
  const publicKeyPem = publicKey.export({ type: "spki", format: "pem" });
  const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" });

  const packageBytes = await createExtensionPackageV2(extensionDir, { privateKeyPem });
  const packageSha256 = sha256Hex(packageBytes);

  const extensionId = "formula.hello";
  const installDir = path.join(extensionsDir, extensionId);

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
        scanStatus: "passed",
        filesSha256: "0".repeat(64),
      };
    },
  };

  const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  await assert.rejects(() => manager.install(extensionId), /files sha256 mismatch/i);

  assert.equal(await pathExists(installDir), false, "installDir should be removed after failed install");
  assert.equal(await pathExists(statePath), false, "state file should not be created for failed install");
});

