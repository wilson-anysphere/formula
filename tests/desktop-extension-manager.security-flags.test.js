import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { ExtensionManager } from "../apps/desktop/tools/marketplace/extensionManager.js";
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

async function createFixture() {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-ext-sec-"));
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

  return {
    tmpRoot,
    extensionsDir,
    statePath,
    extensionId,
    installDir,
    publicKeyPem,
    packageBytes,
    packageSha256,
  };
}

test("ExtensionManager.install refuses blocked extensions", async (t) => {
  const fixture = await createFixture();
  const {
    tmpRoot,
    extensionsDir,
    statePath,
    extensionId,
    installDir,
    publicKeyPem,
    packageBytes,
    packageSha256,
  } = fixture;

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  /** @type {any} */
  const marketplaceClient = {
    async getExtension(id) {
      if (id !== extensionId) return null;
      return { id, latestVersion: "1.0.0", publisherPublicKeyPem: publicKeyPem, blocked: true };
    },
    async downloadPackage() {
      // Should not be reached, but provide it anyway.
      return {
        bytes: Buffer.from(packageBytes),
        sha256: packageSha256,
        formatVersion: 2,
        signatureBase64: null,
        publisher: "formula",
        scanStatus: "passed",
        filesSha256: null,
      };
    },
  };

  const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
  await assert.rejects(() => manager.install(extensionId), /blocked/i);

  assert.equal(await pathExists(installDir), false);
  assert.equal(await pathExists(statePath), false);
});

test("ExtensionManager.install refuses malicious extensions", async (t) => {
  const fixture = await createFixture();
  const {
    tmpRoot,
    extensionsDir,
    statePath,
    extensionId,
    installDir,
    publicKeyPem,
    packageBytes,
    packageSha256,
  } = fixture;

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  /** @type {any} */
  const marketplaceClient = {
    async getExtension(id) {
      if (id !== extensionId) return null;
      return { id, latestVersion: "1.0.0", publisherPublicKeyPem: publicKeyPem, malicious: true };
    },
    async downloadPackage() {
      return {
        bytes: Buffer.from(packageBytes),
        sha256: packageSha256,
        formatVersion: 2,
        signatureBase64: null,
        publisher: "formula",
        scanStatus: "passed",
        filesSha256: null,
      };
    },
  };

  const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
  await assert.rejects(() => manager.install(extensionId), /malicious/i);

  assert.equal(await pathExists(installDir), false);
  assert.equal(await pathExists(statePath), false);
});

test("ExtensionManager.install refuses extensions whose publisher is revoked", async (t) => {
  const fixture = await createFixture();
  const {
    tmpRoot,
    extensionsDir,
    statePath,
    extensionId,
    installDir,
    publicKeyPem,
    packageBytes,
    packageSha256,
  } = fixture;

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  /** @type {any} */
  const marketplaceClient = {
    async getExtension(id) {
      if (id !== extensionId) return null;
      return {
        id,
        latestVersion: "1.0.0",
        publisherPublicKeyPem: publicKeyPem,
        publisherRevoked: true,
      };
    },
    async downloadPackage() {
      return {
        bytes: Buffer.from(packageBytes),
        sha256: packageSha256,
        formatVersion: 2,
        signatureBase64: null,
        publisher: "formula",
        scanStatus: "passed",
        filesSha256: null,
      };
    },
  };

  const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
  await assert.rejects(() => manager.install(extensionId), /revoked/i);

  assert.equal(await pathExists(installDir), false);
  assert.equal(await pathExists(statePath), false);
});

test("ExtensionManager.install refuses extensions when all publisher signing keys are revoked", async (t) => {
  const fixture = await createFixture();
  const {
    tmpRoot,
    extensionsDir,
    statePath,
    extensionId,
    installDir,
    publicKeyPem,
    packageBytes,
    packageSha256,
  } = fixture;

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  /** @type {any} */
  const marketplaceClient = {
    async getExtension(id) {
      if (id !== extensionId) return null;
      return {
        id,
        latestVersion: "1.0.0",
        publisherPublicKeyPem: publicKeyPem,
        publisherKeys: [{ id: "key-1", publicKeyPem, revoked: true }],
      };
    },
    async downloadPackage() {
      return {
        bytes: Buffer.from(packageBytes),
        sha256: packageSha256,
        formatVersion: 2,
        signatureBase64: null,
        publisher: "formula",
        scanStatus: "passed",
        filesSha256: null,
      };
    },
  };

  const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
  await assert.rejects(() => manager.install(extensionId), /revoked/i);

  assert.equal(await pathExists(installDir), false);
  assert.equal(await pathExists(statePath), false);
});

test("ExtensionManager.install enforces scanStatus when configured", async (t) => {
  const fixture = await createFixture();
  const {
    tmpRoot,
    extensionsDir,
    statePath,
    extensionId,
    installDir,
    publicKeyPem,
    packageBytes,
    packageSha256,
  } = fixture;

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

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
        scanStatus: "pending",
        filesSha256: null,
      };
    },
  };

  const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
  await assert.rejects(() => manager.install(extensionId, null, { scanPolicy: "enforce" }), /scan status/i);

  assert.equal(await pathExists(installDir), false);
  assert.equal(await pathExists(statePath), false);
});

test("ExtensionManager.install can allow non-passed scanStatus when configured", async (t) => {
  const fixture = await createFixture();
  const { tmpRoot, extensionsDir, statePath, extensionId, installDir, publicKeyPem, packageBytes, packageSha256 } =
    fixture;

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

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
        scanStatus: "pending",
        filesSha256: null,
      };
    },
  };

  const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
  const record = await manager.install(extensionId, null, { scanPolicy: "allow" });
  assert.equal(record.id, extensionId);
  assert.equal(record.version, "1.0.0");
  assert.equal(record.scanStatus, "pending");
  assert.ok(Array.isArray(record.warnings));
  assert.ok(record.warnings.some((w) => w && w.kind === "scanStatus" && String(w.scanStatus) === "pending"));
  assert.equal(await pathExists(installDir), true);
  assert.equal(await pathExists(statePath), true);
});

test("ExtensionManager.install can be cancelled via confirm() when scanStatus is non-passed and policy=allow", async (t) => {
  const fixture = await createFixture();
  const { tmpRoot, extensionsDir, statePath, extensionId, installDir, publicKeyPem, packageBytes, packageSha256 } =
    fixture;

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

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
        scanStatus: "pending",
        filesSha256: null,
      };
    },
  };

  const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
  let confirmCalled = false;
  await assert.rejects(
    () =>
      manager.install(extensionId, null, {
        scanPolicy: "allow",
        confirm: async (warning) => {
          confirmCalled = true;
          assert.equal(warning.kind, "scanStatus");
          assert.equal(String(warning.scanStatus), "pending");
          return false;
        },
      }),
    /cancel/i
  );
  assert.equal(confirmCalled, true);
  assert.equal(await pathExists(installDir), false);
  assert.equal(await pathExists(statePath), false);
});

test("ExtensionManager.install warns when installing a deprecated extension", async (t) => {
  const fixture = await createFixture();
  const { tmpRoot, extensionsDir, statePath, extensionId, installDir, publicKeyPem, packageBytes, packageSha256 } =
    fixture;

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  /** @type {any} */
  const marketplaceClient = {
    async getExtension(id) {
      if (id !== extensionId) return null;
      return { id, latestVersion: "1.0.0", publisherPublicKeyPem: publicKeyPem, deprecated: true };
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
        filesSha256: null,
      };
    },
  };

  const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
  const record = await manager.install(extensionId);
  assert.equal(record.id, extensionId);
  assert.equal(record.version, "1.0.0");
  assert.equal(record.deprecated, true);
  assert.ok(Array.isArray(record.warnings));
  assert.ok(record.warnings.some((w) => w && w.kind === "deprecated"));
  assert.equal(await pathExists(installDir), true);
  assert.equal(await pathExists(statePath), true);
});

test("ExtensionManager.install can be cancelled via confirm() when installing a deprecated extension", async (t) => {
  const fixture = await createFixture();
  const { tmpRoot, extensionsDir, statePath, extensionId, installDir, publicKeyPem, packageBytes, packageSha256 } =
    fixture;

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  /** @type {any} */
  const marketplaceClient = {
    async getExtension(id) {
      if (id !== extensionId) return null;
      return { id, latestVersion: "1.0.0", publisherPublicKeyPem: publicKeyPem, deprecated: true };
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
        filesSha256: null,
      };
    },
  };

  const manager = new ExtensionManager({ marketplaceClient, extensionsDir, statePath });
  let confirmCalled = false;
  await assert.rejects(
    () =>
      manager.install(extensionId, null, {
        confirm: async (warning) => {
          confirmCalled = true;
          assert.equal(warning.kind, "deprecated");
          return false;
        },
      }),
    /cancel/i
  );
  assert.equal(confirmCalled, true);
  assert.equal(await pathExists(installDir), false);
  assert.equal(await pathExists(statePath), false);
});
