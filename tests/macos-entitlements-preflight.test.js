import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const SCRIPT = fileURLToPath(new URL("../scripts/check-macos-entitlements.mjs", import.meta.url));

/**
 * @param {string} root
 * @param {string} relativePath
 * @param {string} contents
 */
async function writeFixtureFile(root, relativePath, contents) {
  const fullPath = path.join(root, relativePath);
  await fs.mkdir(path.dirname(fullPath), { recursive: true });
  await fs.writeFile(fullPath, contents, "utf8");
}

/**
 * @param {string} rootDir
 */
function runCli(rootDir) {
  return spawnSync(process.execPath, [SCRIPT, "--root", rootDir], {
    encoding: "utf8",
  });
}

test("macOS entitlements preflight passes when required JIT keys are present", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "macos-entitlements-pass-"));
  try {
    await writeFixtureFile(
      tmpRoot,
      "apps/desktop/src-tauri/tauri.conf.json",
      JSON.stringify({ bundle: { macOS: { entitlements: "entitlements.plist" } } }, null, 2) + "\n",
    );
    await writeFixtureFile(
      tmpRoot,
      "apps/desktop/src-tauri/entitlements.plist",
      [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">',
        '<plist version="1.0">',
        "  <dict>",
        "    <key>com.apple.security.network.client</key>",
        "    <true/>",
        "    <key>com.apple.security.cs.allow-jit</key>",
        "    <true/>",
        "    <key>com.apple.security.cs.allow-unsigned-executable-memory</key>",
        "    <true/>",
        "  </dict>",
        "</plist>",
        "",
      ].join("\n"),
    );

    const proc = runCli(tmpRoot);
    assert.equal(proc.status, 0, proc.stderr || proc.stdout);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("macOS entitlements preflight fails when allow-jit entitlement is missing", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "macos-entitlements-fail-"));
  try {
    await writeFixtureFile(
      tmpRoot,
      "apps/desktop/src-tauri/tauri.conf.json",
      JSON.stringify({ bundle: { macOS: { entitlements: "entitlements.plist" } } }, null, 2) + "\n",
    );
    await writeFixtureFile(
      tmpRoot,
      "apps/desktop/src-tauri/entitlements.plist",
      [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">',
        '<plist version="1.0">',
        "  <dict>",
        "    <key>com.apple.security.network.client</key>",
        "    <true/>",
        "    <key>com.apple.security.cs.allow-unsigned-executable-memory</key>",
        "    <true/>",
        "  </dict>",
        "</plist>",
        "",
      ].join("\n"),
    );

    const proc = runCli(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /allow-jit/);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("macOS entitlements preflight fails when network.client entitlement is missing", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "macos-entitlements-network-fail-"));
  try {
    await writeFixtureFile(
      tmpRoot,
      "apps/desktop/src-tauri/tauri.conf.json",
      JSON.stringify({ bundle: { macOS: { entitlements: "entitlements.plist" } } }, null, 2) + "\n",
    );
    await writeFixtureFile(
      tmpRoot,
      "apps/desktop/src-tauri/entitlements.plist",
      [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">',
        '<plist version="1.0">',
        "  <dict>",
        "    <key>com.apple.security.cs.allow-jit</key>",
        "    <true/>",
        "    <key>com.apple.security.cs.allow-unsigned-executable-memory</key>",
        "    <true/>",
        "  </dict>",
        "</plist>",
        "",
      ].join("\n"),
    );

    const proc = runCli(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /network\.client/);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("macOS entitlements preflight fails when forbidden entitlements are enabled", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "macos-entitlements-forbidden-fail-"));
  try {
    await writeFixtureFile(
      tmpRoot,
      "apps/desktop/src-tauri/tauri.conf.json",
      JSON.stringify({ bundle: { macOS: { entitlements: "entitlements.plist" } } }, null, 2) + "\n",
    );
    await writeFixtureFile(
      tmpRoot,
      "apps/desktop/src-tauri/entitlements.plist",
      [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">',
        '<plist version="1.0">',
        "  <dict>",
        "    <key>com.apple.security.network.client</key>",
        "    <true/>",
        "    <key>com.apple.security.cs.allow-jit</key>",
        "    <true/>",
        "    <key>com.apple.security.cs.allow-unsigned-executable-memory</key>",
        "    <true/>",
        "    <key>com.apple.security.cs.disable-library-validation</key>",
        "    <true/>",
        "  </dict>",
        "</plist>",
        "",
      ].join("\n"),
    );

    const proc = runCli(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /disable-library-validation/);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});
