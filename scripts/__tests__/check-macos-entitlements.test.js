import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "check-macos-entitlements.mjs");

/**
 * @param {Array<[string, string]>} entries
 */
function buildEntitlementsPlist(entries) {
  const body = entries
    .map(([key, value]) => `    <key>${key}</key>\n    ${value}`)
    .join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n  <dict>\n${body}\n  </dict>\n</plist>\n`;
}

/**
 * @param {import("node:test").TestContext} t
 * @param {string} plistXml
 */
function runCheck(t, plistXml) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "formula-entitlements-test-"));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const entitlementsPath = path.join(dir, "entitlements.plist");
  fs.writeFileSync(entitlementsPath, plistXml, "utf8");

  return spawnSync(process.execPath, [scriptPath, "--path", entitlementsPath], {
    cwd: repoRoot,
    encoding: "utf8",
  });
}

test("check-macos-entitlements: accepts minimal allowlisted entitlements", (t) => {
  const child = runCheck(
    t,
    buildEntitlementsPlist([
      ["com.apple.security.network.client", "<true/>"],
      ["com.apple.security.cs.allow-jit", "<true/>"],
      ["com.apple.security.cs.allow-unsigned-executable-memory", "<true/>"],
    ]),
  );

  assert.equal(child.status, 0, `expected exit 0, got ${child.status}\n${child.stderr}`);
  assert.match(child.stdout, /preflight passed/i);
});

test("check-macos-entitlements: rejects missing required WKWebView JIT entitlements", (t) => {
  const child = runCheck(
    t,
    buildEntitlementsPlist([["com.apple.security.network.client", "<true/>"]]),
  );

  assert.notEqual(child.status, 0);
  assert.match(child.stderr, /Missing required entitlement/i);
  assert.match(child.stderr, /com\.apple\.security\.cs\.allow-jit/);
  assert.match(child.stderr, /com\.apple\.security\.cs\.allow-unsigned-executable-memory/);
});

test("check-macos-entitlements: rejects forbidden entitlements", (t) => {
  const child = runCheck(
    t,
    buildEntitlementsPlist([
      ["com.apple.security.network.client", "<true/>"],
      ["com.apple.security.cs.allow-jit", "<true/>"],
      ["com.apple.security.cs.allow-unsigned-executable-memory", "<true/>"],
      ["com.apple.security.cs.disable-library-validation", "<true/>"],
    ]),
  );

  assert.notEqual(child.status, 0);
  assert.match(child.stderr, /Disallowed macOS entitlements enabled/i);
  assert.match(child.stderr, /com\.apple\.security\.cs\.disable-library-validation/);
});

test("check-macos-entitlements: rejects unexpected extra entitlements", (t) => {
  const child = runCheck(
    t,
    buildEntitlementsPlist([
      ["com.apple.security.network.client", "<true/>"],
      ["com.apple.security.cs.allow-jit", "<true/>"],
      ["com.apple.security.cs.allow-unsigned-executable-memory", "<true/>"],
      ["com.apple.security.device.camera", "<true/>"],
    ]),
  );

  assert.notEqual(child.status, 0);
  assert.match(child.stderr, /Unexpected macOS entitlements enabled/i);
  assert.match(child.stderr, /com\.apple\.security\.device\.camera/);
});

test("check-macos-entitlements: rejects duplicate keys", (t) => {
  const child = runCheck(
    t,
    buildEntitlementsPlist([
      ["com.apple.security.network.client", "<true/>"],
      ["com.apple.security.network.client", "<true/>"],
      ["com.apple.security.cs.allow-jit", "<true/>"],
      ["com.apple.security.cs.allow-unsigned-executable-memory", "<true/>"],
    ]),
  );

  assert.notEqual(child.status, 0);
  assert.match(child.stderr, /Duplicate keys detected/i);
  assert.match(child.stderr, /com\.apple\.security\.network\.client/);
});

test("check-macos-entitlements: rejects non-true values (false)", (t) => {
  const child = runCheck(
    t,
    buildEntitlementsPlist([
      ["com.apple.security.network.client", "<true/>"],
      ["com.apple.security.cs.allow-jit", "<false/>"],
      ["com.apple.security.cs.allow-unsigned-executable-memory", "<true/>"],
    ]),
  );

  assert.notEqual(child.status, 0);
  assert.match(child.stderr, /must be set to boolean <true\/>/i);
  assert.match(child.stderr, /com\.apple\.security\.cs\.allow-jit/);
});

test("check-macos-entitlements: requires network.server when App Sandbox is enabled", (t) => {
  const child = runCheck(
    t,
    buildEntitlementsPlist([
      ["com.apple.security.network.client", "<true/>"],
      ["com.apple.security.cs.allow-jit", "<true/>"],
      ["com.apple.security.cs.allow-unsigned-executable-memory", "<true/>"],
      ["com.apple.security.app-sandbox", "<true/>"],
    ]),
  );

  assert.notEqual(child.status, 0);
  assert.match(child.stderr, /com\.apple\.security\.network\.server/);
});

test("check-macos-entitlements: accepts App Sandbox + network.server (optional future mode)", (t) => {
  const child = runCheck(
    t,
    buildEntitlementsPlist([
      ["com.apple.security.network.client", "<true/>"],
      ["com.apple.security.cs.allow-jit", "<true/>"],
      ["com.apple.security.cs.allow-unsigned-executable-memory", "<true/>"],
      ["com.apple.security.app-sandbox", "<true/>"],
      ["com.apple.security.network.server", "<true/>"],
    ]),
  );

  assert.equal(child.status, 0, `expected exit 0, got ${child.status}\n${child.stderr}`);
  assert.match(child.stdout, /preflight passed/i);
});

