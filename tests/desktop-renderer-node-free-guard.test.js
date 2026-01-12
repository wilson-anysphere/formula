import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const desktopSrcDir = path.join(repoRoot, "apps", "desktop", "src");
const nodeFreeGuardScript = path.join(repoRoot, "tools", "ci", "check-desktop-renderer-node-free.mjs");
const checkNoNodeScript = path.join(repoRoot, "apps", "desktop", "scripts", "check-no-node.mjs");

function runNode(scriptPath) {
  return spawnSync(process.execPath, [scriptPath], { cwd: repoRoot, encoding: "utf8" });
}

test("desktop renderer Node-free guards fail on node:* imports in apps/desktop/src", async () => {
  const filename = `__node_guard_test__.${process.pid}.${Date.now()}.ts`;
  const filePath = path.join(desktopSrcDir, filename);
  const relToRepo = `apps/desktop/src/${filename}`;
  const relToDesktop = `src/${filename}`;

  await fs.writeFile(
    filePath,
    [
      "// intentionally invalid for the desktop renderer bundle",
      'import fs from "node:fs";',
      "export const x = fs;",
      "",
    ].join("\n"),
  );

  try {
    const guard = runNode(nodeFreeGuardScript);
    assert.notEqual(guard.status, 0, "expected tools/ci/check-desktop-renderer-node-free.mjs to fail");
    assert.match(guard.stderr ?? "", new RegExp(relToRepo.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));

    const checkNoNode = runNode(checkNoNodeScript);
    assert.notEqual(checkNoNode.status, 0, "expected apps/desktop/scripts/check-no-node.mjs to fail");
    assert.match(checkNoNode.stderr ?? "", new RegExp(relToDesktop.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  } finally {
    await fs.rm(filePath, { force: true }).catch(() => {});
  }
});

test("desktop renderer Node-free guards fail on imports from apps/desktop/tools", async () => {
  const filename = `__node_tool_import_test__.${process.pid}.${Date.now()}.ts`;
  const filePath = path.join(desktopSrcDir, filename);
  const relToRepo = `apps/desktop/src/${filename}`;

  await fs.writeFile(
    filePath,
    [
      "// intentionally invalid for the desktop renderer bundle",
      'import { MarketplaceClient } from "../tools/marketplace/client.js";',
      "export const x = MarketplaceClient;",
      "",
    ].join("\n"),
  );

  try {
    const guard = runNode(nodeFreeGuardScript);
    assert.notEqual(guard.status, 0, "expected tools/ci/check-desktop-renderer-node-free.mjs to fail");
    assert.match(guard.stderr ?? "", new RegExp(relToRepo.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
    assert.match(guard.stderr ?? "", /renderer-imports-node-only-tooling/);

    const checkNoNode = runNode(checkNoNodeScript);
    assert.notEqual(checkNoNode.status, 0, "expected apps/desktop/scripts/check-no-node.mjs to fail");
    assert.match(checkNoNode.stderr ?? "", /imports Node-only module "tools\/marketplace\/client\.js"/);
  } finally {
    await fs.rm(filePath, { force: true }).catch(() => {});
  }
});

