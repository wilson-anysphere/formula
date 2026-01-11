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

async function writeJson(filePath, data) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(data, null, 2));
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

test("ExtensionHostManager serializes command execution vs sync/unload operations", async (t) => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-ext-concurrency-"));
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const extensionId = "formula.slow-ext";
  const extensionDir = path.join(extensionsDir, extensionId);
  await fs.mkdir(path.join(extensionDir, "dist"), { recursive: true });

  await fs.writeFile(
    path.join(extensionDir, "package.json"),
    JSON.stringify(
      {
        name: "slow-ext",
        publisher: "formula",
        version: "1.0.0",
        main: "./dist/extension.js",
        engines: { formula: "^1.0.0" },
        activationEvents: ["onCommand:slow.wait"],
        contributes: {
          commands: [{ command: "slow.wait", title: "Wait (Slow)" }],
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
      '  context.subscriptions.push(await formula.commands.registerCommand("slow.wait", async (ms = 50) => {',
      "    await new Promise((resolve) => setTimeout(resolve, Number(ms) || 0));",
      '    return "done";',
      "  }));",
      "}",
      "module.exports = { activate };",
      "",
    ].join("\n"),
  );

  const files = await listFilesWithIntegrity(extensionDir);
  await writeJson(statePath, {
    installed: {
      [extensionId]: { id: extensionId, version: "1.0.0", installedAt: new Date().toISOString(), files },
    },
  });

  const runtime = new ExtensionHostManager({
    extensionsDir,
    statePath,
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(tmpRoot, "permissions.json"),
    extensionStoragePath: path.join(tmpRoot, "storage.json"),
    permissionPrompt: async () => true,
    activationTimeoutMs: EXTENSION_TIMEOUT_MS,
    commandTimeoutMs: EXTENSION_TIMEOUT_MS,
  });

  t.after(async () => {
    await runtime.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  await runtime.startup();

  // Kick off a slow command then immediately sync away the extension. With the host operation
  // queue, the sync should wait until the command completes (rather than racing an unload).
  const commandPromise = runtime.executeCommand("slow.wait", 75);

  await writeJson(statePath, { installed: {} });
  const syncPromise = runtime.syncInstalledExtensions();

  assert.equal(await commandPromise, "done");
  await syncPromise;

  await assert.rejects(() => runtime.executeCommand("slow.wait", 1), /Unknown command/i);
});

