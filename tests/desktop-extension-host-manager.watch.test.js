import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";

import { ExtensionHostManager } from "../apps/desktop/src/extensions/ExtensionHostManager.js";

const EXTENSION_TIMEOUT_MS = 20_000;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");

async function copyDir(srcDir, destDir) {
  await fs.mkdir(destDir, { recursive: true });
  const entries = await fs.readdir(srcDir, { withFileTypes: true });
  for (const entry of entries) {
    const src = path.join(srcDir, entry.name);
    const dest = path.join(destDir, entry.name);
    if (entry.isDirectory()) {
      await copyDir(src, dest);
      continue;
    }
    if (entry.isFile()) {
      await fs.copyFile(src, dest);
    }
  }
}

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

async function waitForValue(getValue, predicate, { timeoutMs = 5_000, intervalMs = 25 } = {}) {
  const deadline = Date.now() + timeoutMs;
  // eslint-disable-next-line no-constant-condition
  while (true) {
    let value = null;
    try {
      value = getValue();
    } catch {
      value = null;
    }
    if (predicate(value)) return value;
    if (Date.now() >= deadline) {
      throw new Error("Timed out waiting for condition");
    }
    // eslint-disable-next-line no-await-in-loop
    await new Promise((r) => setTimeout(r, intervalMs));
  }
}

test("ExtensionHostManager can watch statePath and auto-sync on changes", async (t) => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-ext-watch-"));
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
  const manifest = JSON.parse(await fs.readFile(path.join(sampleExtensionSrc, "package.json"), "utf8"));
  const extensionId = `${manifest.publisher}.${manifest.name}`;

  const installedPath = path.join(extensionsDir, extensionId);
  await copyDir(sampleExtensionSrc, installedPath);
  const files = await listFilesWithIntegrity(installedPath);

  await fs.mkdir(path.dirname(statePath), { recursive: true });
  await fs.writeFile(statePath, JSON.stringify({ installed: {} }, null, 2));

  const runtime = new ExtensionHostManager({
    extensionsDir,
    statePath,
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(tmpRoot, "permissions.json"),
    extensionStoragePath: path.join(tmpRoot, "storage.json"),
    permissionPrompt: async () => true,
    activationTimeoutMs: EXTENSION_TIMEOUT_MS,
    commandTimeoutMs: EXTENSION_TIMEOUT_MS,
    customFunctionTimeoutMs: EXTENSION_TIMEOUT_MS,
    dataConnectorTimeoutMs: EXTENSION_TIMEOUT_MS,
  });

  t.after(async () => {
    await runtime.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  let lastSync = null;
  const originalSync = runtime.syncInstalledExtensions.bind(runtime);
  runtime.syncInstalledExtensions = (...args) => {
    lastSync = originalSync(...args);
    return lastSync;
  };

  await runtime.startup();
  await runtime.watchStateFile({ debounceMs: 10 });
  // Give the underlying fs watcher a moment to attach before mutating the state file.
  await new Promise((r) => setTimeout(r, 25));

  // Install by mutating state file.
  lastSync = null;
  await fs.writeFile(
    statePath,
    JSON.stringify(
      {
        installed: {
          [extensionId]: { id: extensionId, version: manifest.version, installedAt: new Date().toISOString(), files },
        },
      },
      null,
      2,
    ),
  );

  const firstSync = await waitForValue(
    () => lastSync,
    (value) => Boolean(value && typeof value.then === "function"),
  );
  await firstSync;

  runtime.spreadsheet.setCell(0, 0, 1);
  runtime.spreadsheet.setCell(0, 1, 2);
  runtime.spreadsheet.setCell(1, 0, 3);
  runtime.spreadsheet.setCell(1, 1, 4);
  runtime.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
  assert.equal(await runtime.executeCommand("sampleHello.sumSelection"), 10);

  // Uninstall by mutating state file.
  lastSync = null;
  await fs.writeFile(statePath, JSON.stringify({ installed: {} }, null, 2));
  const secondSync = await waitForValue(
    () => lastSync,
    (value) => Boolean(value && typeof value.then === "function"),
  );
  await secondSync;

  await assert.rejects(() => runtime.executeCommand("sampleHello.sumSelection"), /Unknown command/i);
});
