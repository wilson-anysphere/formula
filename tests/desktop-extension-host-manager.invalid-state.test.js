import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";

import { ExtensionHostManager } from "../apps/desktop/src/extensions/ExtensionHostManager.js";

const EXTENSION_TIMEOUT_MS = 20_000;

test("ExtensionHostManager.startup tolerates invalid JSON state file", async (t) => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-ext-invalid-state-"));
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  await fs.mkdir(extensionsDir, { recursive: true });
  await fs.writeFile(statePath, "{", "utf8");

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
  await assert.doesNotReject(() => runtime.syncInstalledExtensions());

  const contributions = runtime.listContributions();
  assert.deepEqual(contributions.commands, []);
  assert.deepEqual(contributions.customFunctions, []);
  assert.deepEqual(contributions.dataConnectors, []);
});

