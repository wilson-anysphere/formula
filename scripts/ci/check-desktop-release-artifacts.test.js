import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-desktop-release-artifacts.mjs");

/**
 * @param {string} p
 */
function touch(p) {
  fs.mkdirSync(path.dirname(p), { recursive: true });
  fs.writeFileSync(p, "x");
}

/**
 * @param {Record<string, string | undefined>} env
 * @param {string[]} args
 */
function run(env, args) {
  const proc = spawnSync(process.execPath, [scriptPath, ...args], {
    encoding: "utf8",
    cwd: repoRoot,
    env: {
      ...process.env,
      ...env,
    },
  });
  if (proc.error) throw proc.error;
  return proc;
}

test("fork mode: macOS unsigned build passes without updater signatures", () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "formula-release-artifacts-"));
  const bundleDir = path.join(tmp, "target", "release", "bundle");

  touch(path.join(bundleDir, "dmg", "Formula.dmg"));

  const proc = run(
    {
      RUNNER_OS: "macOS",
      FORMULA_REQUIRE_TAURI_UPDATER_SIGNATURES: "false",
      FORMULA_HAS_TAURI_UPDATER_KEY: "false",
    },
    ["--bundle-dir", bundleDir],
  );
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /passed/i);
});

test("check-desktop-release-artifacts bounds fallback repo walk (perf guardrail)", () => {
  const raw = fs.readFileSync(scriptPath, "utf8");
  const idx = raw.indexOf("Fallback: search for src-tauri directories");
  assert.ok(idx >= 0, "Expected check-desktop-release-artifacts.mjs to have a fallback target dir discovery.");
  const snippet = raw.slice(idx, idx + 1200);
  assert.match(
    snippet,
    /maxDepth\s*=\s*\d+/,
    `Expected fallback repo walk to define a maxDepth bound.\nSaw snippet:\n${snippet}`,
  );
  assert.match(
    snippet,
    /depth:\s*0/,
    `Expected fallback repo walk to track traversal depth (avoid unbounded scans).\nSaw snippet:\n${snippet}`,
  );
});

test("signed mode: macOS requires latest.json + installer/archive signatures", () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "formula-release-artifacts-"));
  const bundleDir = path.join(tmp, "target", "release", "bundle");

  touch(path.join(bundleDir, "latest.json"));
  touch(path.join(bundleDir, "latest.json.sig"));
  touch(path.join(bundleDir, "dmg", "Formula.dmg"));
  touch(path.join(bundleDir, "dmg", "Formula.dmg.sig"));
  touch(path.join(bundleDir, "macos", "Formula_universal.tgz"));
  touch(path.join(bundleDir, "macos", "Formula_universal.tgz.sig"));

  const proc = run(
    {
      RUNNER_OS: "macOS",
      FORMULA_REQUIRE_TAURI_UPDATER_SIGNATURES: "true",
      FORMULA_HAS_TAURI_UPDATER_KEY: "true",
    },
    ["--bundle-dir", bundleDir],
  );
  assert.equal(proc.status, 0, proc.stderr);
});

test("fork mode: Windows requires installer artifacts but not updater signatures", () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "formula-release-artifacts-"));
  const bundleDir = path.join(tmp, "target", "release", "bundle");

  // Must match the script's stricter path filters:
  touch(path.join(bundleDir, "msi", "Formula_x64.msi"));
  touch(path.join(bundleDir, "nsis", "Formula_x64.exe"));
  // A WebView2 bootstrapper should not count as the shipped installer.
  touch(path.join(bundleDir, "nsis", "MicrosoftEdgeWebView2Setup.exe"));

  const proc = run(
    {
      RUNNER_OS: "Windows",
      FORMULA_REQUIRE_TAURI_UPDATER_SIGNATURES: "false",
      FORMULA_HAS_TAURI_UPDATER_KEY: "false",
    },
    ["--bundle-dir", bundleDir],
  );
  assert.equal(proc.status, 0, proc.stderr);
});

test("fork mode: Linux requires AppImage + deb + rpm but not updater signatures", () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "formula-release-artifacts-"));
  const bundleDir = path.join(tmp, "target", "release", "bundle");

  touch(path.join(bundleDir, "appimage", "Formula.AppImage"));
  touch(path.join(bundleDir, "deb", "Formula.deb"));
  touch(path.join(bundleDir, "rpm", "Formula.rpm"));

  const proc = run(
    {
      RUNNER_OS: "Linux",
      FORMULA_REQUIRE_TAURI_UPDATER_SIGNATURES: "false",
      FORMULA_HAS_TAURI_UPDATER_KEY: "false",
    },
    ["--bundle-dir", bundleDir],
  );
  assert.equal(proc.status, 0, proc.stderr);
});

test("signed mode: Linux requires latest.json + signatures for AppImage/deb/rpm", () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "formula-release-artifacts-"));
  const bundleDir = path.join(tmp, "target", "release", "bundle");

  touch(path.join(bundleDir, "latest.json"));
  touch(path.join(bundleDir, "latest.json.sig"));

  touch(path.join(bundleDir, "appimage", "Formula.AppImage"));
  touch(path.join(bundleDir, "appimage", "Formula.AppImage.sig"));
  touch(path.join(bundleDir, "deb", "Formula.deb"));
  touch(path.join(bundleDir, "deb", "Formula.deb.sig"));
  touch(path.join(bundleDir, "rpm", "Formula.rpm"));
  touch(path.join(bundleDir, "rpm", "Formula.rpm.sig"));

  const proc = run(
    {
      RUNNER_OS: "Linux",
      FORMULA_REQUIRE_TAURI_UPDATER_SIGNATURES: "true",
      FORMULA_HAS_TAURI_UPDATER_KEY: "true",
    },
    ["--bundle-dir", bundleDir],
  );
  assert.equal(proc.status, 0, proc.stderr);
});
