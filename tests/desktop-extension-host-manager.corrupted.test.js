import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";

import { ExtensionHostManager } from "../apps/desktop/src/extensions/ExtensionHostManager.js";

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

test("ExtensionHostManager.startup skips corrupted installs instead of throwing", async (t) => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-ext-corrupted-"));
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const goodId = "formula.good-ext";
  const goodDir = path.join(extensionsDir, goodId);
  await fs.mkdir(path.join(goodDir, "dist"), { recursive: true });
  await fs.writeFile(
    path.join(goodDir, "package.json"),
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
      2
    )
  );
  await fs.writeFile(
    path.join(goodDir, "dist", "extension.js"),
    [
      'const formula = require("@formula/extension-api");',
      "async function activate(context) {",
      '  context.subscriptions.push(await formula.commands.registerCommand("good.hello", async () => "hi"));',
      "}",
      "module.exports = { activate };",
      "",
    ].join("\n")
  );

  const goodFiles = await listFilesWithIntegrity(goodDir);

  // Corrupted install: missing integrity metadata (no `files` list).
  const badId = "formula.broken-ext";

  await fs.writeFile(
    statePath,
    JSON.stringify(
      {
        installed: {
          [goodId]: { id: goodId, version: "1.0.0", installedAt: new Date().toISOString(), files: goodFiles },
          [badId]: { id: badId, version: "1.0.0", installedAt: new Date().toISOString() },
        },
      },
      null,
      2
    )
  );

  const runtime = new ExtensionHostManager({
    extensionsDir,
    statePath,
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(tmpRoot, "permissions.json"),
    extensionStoragePath: path.join(tmpRoot, "storage.json"),
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await runtime.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  // Should not throw even though one installed record is corrupt.
  await runtime.startup();

  assert.equal(await runtime.executeCommand("good.hello"), "hi");

  // Sync should also not throw and should keep the corrupted install quarantined.
  await assert.doesNotReject(() => runtime.syncInstalledExtensions());

  const updated = JSON.parse(await fs.readFile(statePath, "utf8"));
  assert.equal(updated.installed[badId].corrupted, true);
  assert.match(String(updated.installed[badId].corruptedReason), /Missing integrity metadata/i);
});
