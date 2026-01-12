import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";

import extensionHostPkg from "../packages/extension-host/src/index.js";

const { ExtensionHost } = extensionHostPkg;

const EXTENSION_TIMEOUT_MS = 20_000;

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

test("ExtensionHost.unloadExtension removes contributed commands/custom functions/menus/data connectors", async (t) => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-extension-host-unload-"));
  const permissionsStoragePath = path.join(tmpRoot, "permissions.json");
  const extensionStoragePath = path.join(tmpRoot, "storage.json");

  const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
  const extensionDir = path.join(tmpRoot, "sample-hello");
  await copyDir(sampleExtensionSrc, extensionDir);

  const manifestPath = path.join(extensionDir, "package.json");
  const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
  const extensionId = `${manifest.publisher}.${manifest.name}`;
  manifest.contributes = manifest.contributes ?? {};
  manifest.contributes.menus = {
    ...(manifest.contributes.menus ?? {}),
    "cell/context": [{ command: "sampleHello.sumSelection", when: "cellHasValue" }],
  };
  await fs.writeFile(manifestPath, JSON.stringify(manifest, null, 2));

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath,
    extensionStoragePath,
    permissionPrompt: async () => true,
    activationTimeoutMs: EXTENSION_TIMEOUT_MS,
    commandTimeoutMs: EXTENSION_TIMEOUT_MS,
    customFunctionTimeoutMs: EXTENSION_TIMEOUT_MS,
    dataConnectorTimeoutMs: EXTENSION_TIMEOUT_MS,
  });

  t.after(async () => {
    await host.dispose();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  await host.loadExtension(extensionDir);

  host.spreadsheet.setCell(0, 0, 1);
  host.spreadsheet.setCell(0, 1, 2);
  host.spreadsheet.setCell(1, 0, 3);
  host.spreadsheet.setCell(1, 1, 4);
  host.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });

  assert.equal(await host.executeCommand("sampleHello.sumSelection"), 10);

  const connectorResult = await host.invokeDataConnector("sampleHello.connector", "testConnection", {});
  assert.deepEqual(connectorResult, { success: true });

  const menus = host.getContributedMenus();
  assert.ok(menus["cell/context"]);
  assert.equal(menus["cell/context"].length, 1);

  assert.ok(host.getContributedDataConnectors().some((c) => c.id === "sampleHello.connector"));

  await host.unloadExtension(extensionId);

  assert.deepEqual(host.listExtensions(), []);
  assert.deepEqual(host.getContributedCommands(), []);
  assert.deepEqual(host.getContributedMenus(), {});
  assert.deepEqual(host.getContributedDataConnectors(), []);

  await assert.rejects(() => host.executeCommand("sampleHello.sumSelection"), /Unknown command/);
  await assert.rejects(() => host.invokeCustomFunction("SAMPLEHELLO_DOUBLE", 2), /Unknown custom function/);
  await assert.rejects(() => host.invokeDataConnector("sampleHello.connector", "testConnection", {}), /Unknown data connector/);
});
