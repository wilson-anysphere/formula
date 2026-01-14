import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "export-updater-manifest.mjs");

/**
 * @param {string} cwd
 * @param {Record<string, string | undefined>} env
 * @param {string[]} args
 */
function run(cwd, env, args) {
  const proc = spawnSync(process.execPath, [scriptPath, ...args], {
    encoding: "utf8",
    cwd,
    env: {
      ...process.env,
      ...env,
    },
  });
  if (proc.error) throw proc.error;
  return proc;
}

test("export-updater-manifest copies latest.json from repo root when present", () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "formula-export-updater-manifest-"));
  const rootManifest = path.join(tmp, "latest.json");
  fs.writeFileSync(rootManifest, '{"version":"0.0.0"}\n', "utf8");

  const outPath = path.join(tmp, "out", "manifest.json");
  const proc = run(tmp, {}, [outPath]);
  assert.equal(proc.status, 0, proc.stderr);
  assert.ok(fs.existsSync(outPath), "Expected output file to be created");
  assert.equal(fs.readFileSync(outPath, "utf8"), fs.readFileSync(rootManifest, "utf8"));
});

test("export-updater-manifest falls back to bundle dir discovery (release/bundle/latest.json)", () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "formula-export-updater-manifest-"));
  const targetDir = path.join(tmp, "cargo-target");
  const bundleDir = path.join(targetDir, "release", "bundle");
  fs.mkdirSync(bundleDir, { recursive: true });

  const manifest = path.join(bundleDir, "latest.json");
  fs.writeFileSync(manifest, '{"version":"0.0.1"}\n', "utf8");

  const outPath = path.join(tmp, "out", "manifest.json");
  const proc = run(tmp, { CARGO_TARGET_DIR: path.relative(tmp, targetDir) }, [outPath]);
  assert.equal(proc.status, 0, proc.stderr);
  assert.ok(fs.existsSync(outPath), "Expected output file to be created");
  assert.equal(fs.readFileSync(outPath, "utf8"), fs.readFileSync(manifest, "utf8"));
});

test("export-updater-manifest can find latest.json one level under the bundle dir", () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "formula-export-updater-manifest-"));
  const targetDir = path.join(tmp, "cargo-target");
  const bundleDir = path.join(targetDir, "release", "bundle");
  const macosDir = path.join(bundleDir, "macos");
  fs.mkdirSync(macosDir, { recursive: true });

  const manifest = path.join(macosDir, "latest.json");
  fs.writeFileSync(manifest, '{"version":"0.0.2"}\n', "utf8");

  const outPath = path.join(tmp, "out", "manifest.json");
  const proc = run(tmp, { CARGO_TARGET_DIR: targetDir }, [outPath]);
  assert.equal(proc.status, 0, proc.stderr);
  assert.ok(fs.existsSync(outPath), "Expected output file to be created");
  assert.equal(fs.readFileSync(outPath, "utf8"), fs.readFileSync(manifest, "utf8"));
});

