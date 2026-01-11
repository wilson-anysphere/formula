import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";

import { ExtensionHostManager } from "../apps/desktop/src/extensions/ExtensionHostManager.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");

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

async function writeJson(filePath, data) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(data, null, 2));
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

test("ExtensionHostManager routes commands/panels/custom functions/data connectors", async (t) => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-desktop-ext-runtime-"));
  const extensionsDir = path.join(tmpRoot, "installed-extensions");
  const statePath = path.join(tmpRoot, "extensions-state.json");

  const sampleExtensionSrc = path.join(repoRoot, "extensions", "sample-hello");
  const manifest = JSON.parse(await fs.readFile(path.join(sampleExtensionSrc, "package.json"), "utf8"));
  const extensionId = `${manifest.publisher}.${manifest.name}`;

  const installedPath = path.join(extensionsDir, extensionId);
  await copyDir(sampleExtensionSrc, installedPath);
  const files = await listFilesWithIntegrity(installedPath);

  await writeJson(statePath, {
    installed: {
      [extensionId]: { id: extensionId, version: manifest.version, installedAt: new Date().toISOString(), files },
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
    customFunctionTimeoutMs: EXTENSION_TIMEOUT_MS,
    dataConnectorTimeoutMs: EXTENSION_TIMEOUT_MS,
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
  assert.equal(await runtime.invokeCustomFunction("SAMPLEHELLO_DOUBLE", 5), 10);

  const connectorResult = await runtime.invokeDataConnector("sampleHello.connector", "testConnection", {});
  assert.deepEqual(connectorResult, { success: true });

  await runtime.executeCommand("sampleHello.openPanel");
  const panel = runtime.getPanel("sampleHello.panel");
  assert.ok(panel);
  assert.ok(panel.html.includes("Sample Hello Panel"));

  runtime.dispatchPanelMessage("sampleHello.panel", { type: "ping" });

  const deadline = Date.now() + 5_000;
  while (Date.now() < deadline) {
    const outgoing = runtime.getPanelOutgoingMessages("sampleHello.panel");
    if (outgoing.some((m) => m && m.type === "pong")) return;
    await new Promise((r) => setTimeout(r, 10));
  }

  assert.fail("Timed out waiting for pong message from extension panel");
});
