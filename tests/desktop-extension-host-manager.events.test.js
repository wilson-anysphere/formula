import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";

import { ExtensionHostManager } from "../apps/desktop/src/extensions/ExtensionHostManager.js";

const EXTENSION_TIMEOUT_MS = 20_000;

function sha256Hex(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function listFilesWithIntegrity(rootDir) {
  /** @type {{ path: string, sha256: string, size: number }[]} */
  const out = [];

  async function visit(dir) {
    const entries = await fs.readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
      const absPath = path.join(dir, entry.name);
      const relPath = path.relative(rootDir, absPath).replace(/\\/g, "/");
      if (relPath === "" || relPath.startsWith("..")) continue;

      if (entry.isDirectory()) {
        await visit(absPath);
        continue;
      }
      if (!entry.isFile()) continue;

      // eslint-disable-next-line no-await-in-loop
      const bytes = await fs.readFile(absPath);
      out.push({ path: relPath, sha256: sha256Hex(bytes), size: bytes.length });
    }
  }

  await visit(rootDir);
  out.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));
  return out;
}

class FakeExtensionManager {
  constructor() {
    /** @type {Set<(event: any) => void>} */
    this._listeners = new Set();
  }

  onDidChange(listener) {
    this._listeners.add(listener);
    return { dispose: () => this._listeners.delete(listener) };
  }

  emit(event) {
    for (const listener of [...this._listeners]) {
      listener(event);
    }
  }
}

test("ExtensionHostManager can auto-sync from ExtensionManager change events", async (t) => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-ext-events-"));
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const extensionId = "formula.good-ext";
  const extensionDir = path.join(extensionsDir, extensionId);
  await fs.mkdir(path.join(extensionDir, "dist"), { recursive: true });

  await fs.writeFile(
    path.join(extensionDir, "package.json"),
    JSON.stringify(
      {
        name: "good-ext",
        publisher: "formula",
        version: "1.0.0",
        main: "./dist/extension.js",
        engines: { formula: "^1.0.0" },
        activationEvents: ["onCommand:good.hello"],
        contributes: {
          commands: [{ command: "good.hello", title: "Good Hello" }],
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
      '  context.subscriptions.push(await formula.commands.registerCommand("good.hello", async () => "hi"));',
      "}",
      "module.exports = { activate };",
      "",
    ].join("\n"),
  );

  await fs.mkdir(path.dirname(statePath), { recursive: true });
  await fs.writeFile(statePath, JSON.stringify({ installed: {} }, null, 2));

  const fakeManager = new FakeExtensionManager();
  const runtime = new ExtensionHostManager({
    extensionsDir,
    statePath,
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(tmpRoot, "permissions.json"),
    extensionStoragePath: path.join(tmpRoot, "storage.json"),
    permissionPrompt: async () => true,
    activationTimeoutMs: EXTENSION_TIMEOUT_MS,
    commandTimeoutMs: EXTENSION_TIMEOUT_MS,
    extensionManager: fakeManager,
  });

  t.after(async () => {
    await runtime.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  let syncCalls = 0;
  let lastSync = null;
  const originalSync = runtime.syncInstalledExtensions.bind(runtime);
  runtime.syncInstalledExtensions = (...args) => {
    syncCalls += 1;
    lastSync = originalSync(...args);
    return lastSync;
  };

  await runtime.startup();

  const files = await listFilesWithIntegrity(extensionDir);
  await fs.writeFile(
    statePath,
    JSON.stringify(
      {
        installed: {
          [extensionId]: { id: extensionId, version: "1.0.0", installedAt: new Date().toISOString(), files },
        },
      },
      null,
      2,
    ),
  );

  fakeManager.emit({ action: "install", id: extensionId });
  assert.ok(lastSync);
  await lastSync;
  assert.equal(syncCalls, 1);

  assert.equal(await runtime.executeCommand("good.hello"), "hi");

  await fs.writeFile(statePath, JSON.stringify({ installed: {} }, null, 2));
  fakeManager.emit({ action: "uninstall", id: extensionId });
  await lastSync;
  assert.equal(syncCalls, 2);

  await assert.rejects(() => runtime.executeCommand("good.hello"), /Unknown command/i);
});
