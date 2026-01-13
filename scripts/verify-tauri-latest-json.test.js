import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "verify-tauri-latest-json.mjs");

/**
 * @param {any} manifest
 * @param {{ includeSig?: boolean; sigContent?: string }} [opts]
 */
function runLocal(manifest, opts = {}) {
  const includeSig = opts.includeSig ?? true;
  const sigContent = opts.sigContent ?? "dummy-signature";

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "verify-tauri-latest-json-"));
  const manifestPath = path.join(tmpDir, "latest.json");
  const sigPath = path.join(tmpDir, includeSig ? "latest.json.sig" : "missing.sig");

  fs.writeFileSync(manifestPath, JSON.stringify(manifest), "utf8");
  if (includeSig) {
    fs.writeFileSync(sigPath, sigContent, "utf8");
  }

  const proc = spawnSync(
    process.execPath,
    [scriptPath, "--manifest", manifestPath, "--sig", sigPath],
    {
      encoding: "utf8",
      cwd: repoRoot,
    },
  );

  fs.rmSync(tmpDir, { recursive: true, force: true });

  if (proc.error) throw proc.error;
  return proc;
}

test("passes for macOS universal (per-arch keys) + windows x64/arm64 + linux x64/arm64", () => {
  const proc = runLocal({
    version: "0.0.0",
    platforms: {
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "linux-x86_64": { url: "https://example.com/Formula_x86_64.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
    },
  });

  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /verification passed/i);
});

test("fails for alias platform keys (Rust target triples) instead of Tauri updater keys", () => {
  const proc = runLocal({
    version: "0.0.0",
    platforms: {
      // These keys are intentionally *not* accepted by the validator; we pin the expected
      // `latest.json.platforms` keys to the values documented in docs/desktop-updater-target-mapping.md.
      "x86_64-apple-darwin": { url: "https://example.com/app-x64.tar.gz", signature: "sig" },
      "aarch64-apple-darwin": { url: "https://example.com/app-arm.tar.gz", signature: "sig" },
      "x86_64-pc-windows-msvc": { url: "https://example.com/app.msi", signature: "sig" },
      "aarch64-pc-windows-msvc": { url: "https://example.com/app-arm.msi", signature: "sig" },
      "x86_64-unknown-linux-gnu": { url: "https://example.com/app_x86_64.AppImage", signature: "sig" },
      "aarch64-unknown-linux-gnu": { url: "https://example.com/app_arm64.AppImage", signature: "sig" },
    },
  });

  assert.notEqual(proc.status, 0);
  assert.ok(
    proc.stderr.includes("Missing required latest.json.platforms keys"),
    `stderr did not include expected message; got:\n${proc.stderr}`,
  );
});

test("finds nested platforms objects", () => {
  const proc = runLocal({
    meta: { something: true },
    data: {
      platforms: {
        "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
        "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
        "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
        "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
        "linux-x86_64": { url: "https://example.com/Formula_x86_64.AppImage", signature: "sig" },
        "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      },
    },
  });

  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when windows-arm64 is missing", () => {
  const proc = runLocal({
    version: "0.0.0",
    platforms: {
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "linux-x86_64": { url: "https://example.com/Formula_x86_64.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
    },
  });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /windows-aarch64/);
});

test("fails when updater asset types do not match platform expectations (Linux .deb)", () => {
  const proc = runLocal({
    version: "0.0.0",
    platforms: {
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      // Linux updater should point to an AppImage, not a deb/rpm package.
      "linux-x86_64": { url: "https://example.com/app.deb", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
    },
  });

  assert.notEqual(proc.status, 0);
  assert.ok(
    proc.stderr.includes("Updater asset type mismatch in latest.json.platforms"),
    `stderr did not include expected asset type mismatch; got:\n${proc.stderr}`,
  );
});

test("fails when multiple targets share the same updater URL (collision)", () => {
  const url = "https://example.com/app.msi";
  const proc = runLocal({
    version: "0.0.0",
    platforms: {
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "windows-x86_64": { url, signature: "sig" },
      "windows-aarch64": { url, signature: "sig" },
      "linux-x86_64": { url: "https://example.com/Formula_x86_64.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
    },
  });

  assert.notEqual(proc.status, 0);
  assert.ok(
    proc.stderr.includes("Duplicate updater URLs across required targets"),
    `stderr did not include expected collision error; got:\n${proc.stderr}`,
  );
});

test("fails when latest.json.sig is missing", () => {
  const proc = runLocal(
    {
      version: "0.0.0",
      platforms: {
        "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
        "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
        "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
        "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
        "linux-x86_64": { url: "https://example.com/Formula_x86_64.AppImage", signature: "sig" },
        "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      },
    },
    { includeSig: false },
  );

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /missing signature file/i);
});
