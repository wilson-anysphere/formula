import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";

import { fileURLToPath } from "node:url";

import { ExtensionHostManager } from "../apps/desktop/src/extensions/ExtensionHostManager.js";

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

async function writeJson(filePath, data) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(data, null, 2));
}

test("ExtensionHostManager.syncInstalledExtensions loads/reloads/unloads based on statePath", async (t) => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-ext-sync-"));
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
  const manifest = JSON.parse(await fs.readFile(path.join(sampleExtensionSrc, "package.json"), "utf8"));
  const extensionId = `${manifest.publisher}.${manifest.name}`;

  const installedPath = path.join(extensionsDir, extensionId);
  await copyDir(sampleExtensionSrc, installedPath);

  await writeJson(statePath, {
    installed: {
      [extensionId]: { id: extensionId, version: "1.0.0", installedAt: new Date().toISOString() },
    },
  });

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

  await runtime.startup();

  runtime.spreadsheet.setCell(0, 0, 1);
  runtime.spreadsheet.setCell(0, 1, 2);
  runtime.spreadsheet.setCell(1, 0, 3);
  runtime.spreadsheet.setCell(1, 1, 4);
  runtime.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
  assert.equal(await runtime.executeCommand("sampleHello.sumSelection"), 10);

  // Simulate an update written to disk + state file: version changes and command behavior changes.
  const manifestPath = path.join(installedPath, "package.json");
  const installedManifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
  installedManifest.version = "1.1.0";
  await fs.writeFile(manifestPath, JSON.stringify(installedManifest, null, 2));

  const distPath = path.join(installedPath, "dist", "extension.js");
  const distText = await fs.readFile(distPath, "utf8");
  const sumValuesIdx = distText.indexOf("function sumValues");
  assert.ok(sumValuesIdx >= 0);
  const sumValuesEnd = distText.indexOf("\n\nasync function getSelectionSum", sumValuesIdx);
  assert.ok(sumValuesEnd >= 0);
  const sumFn = distText.slice(sumValuesIdx, sumValuesEnd);
  const needle = "return sum;";
  const lastReturnIdx = sumFn.lastIndexOf(needle);
  assert.ok(lastReturnIdx >= 0);
  const sumFnPatched = sumFn.slice(0, lastReturnIdx) + "return sum + 1;" + sumFn.slice(lastReturnIdx + needle.length);
  assert.notEqual(sumFnPatched, sumFn);
  await fs.writeFile(distPath, distText.slice(0, sumValuesIdx) + sumFnPatched + distText.slice(sumValuesEnd));

  await writeJson(statePath, {
    installed: {
      [extensionId]: { id: extensionId, version: "1.1.0", installedAt: new Date().toISOString() },
    },
  });

  await runtime.syncInstalledExtensions();

  runtime.spreadsheet.setCell(0, 0, 1);
  runtime.spreadsheet.setCell(0, 1, 2);
  runtime.spreadsheet.setCell(1, 0, 3);
  runtime.spreadsheet.setCell(1, 1, 4);
  runtime.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
  assert.equal(await runtime.executeCommand("sampleHello.sumSelection"), 11);

  // Simulate uninstall by removing the entry from state.
  await writeJson(statePath, { installed: {} });
  await runtime.syncInstalledExtensions();

  await assert.rejects(() => runtime.executeCommand("sampleHello.sumSelection"), /Unknown command/);
});
