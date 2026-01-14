import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const smokeTestPath = path.join(repoRoot, "scripts", "release-smoke-test.mjs");

function currentDesktopTag() {
  const tauriConfPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
  const config = JSON.parse(fs.readFileSync(tauriConfPath, "utf8"));
  const version = typeof config?.version === "string" ? config.version.trim() : "";
  assert.ok(version, "Expected tauri.conf.json to contain a non-empty version");
  return version.startsWith("v") ? version : `v${version}`;
}

test("release-smoke-test: --help prints usage and exits 0", () => {
  const child = spawnSync(process.execPath, [smokeTestPath, "--help"], {
    cwd: repoRoot,
    encoding: "utf8",
  });
  assert.equal(child.status, 0, `expected exit 0, got ${child.status}\n${child.stderr}`);
  assert.match(child.stdout, /Release smoke test/i);
});

test("release-smoke-test: runs required steps and can forward --help to verifier", () => {
  const tag = currentDesktopTag();
  const child = spawnSync(process.execPath, [smokeTestPath, "--tag", tag, "--repo", "owner/repo", "--", "--help"], {
    cwd: repoRoot,
    encoding: "utf8",
  });

  assert.equal(
    child.status,
    0,
    `expected exit 0, got ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
  );
  assert.match(child.stdout, /Release smoke test PASSED/i);
});

test("release-smoke-test: defaults --tag from GITHUB_REF_NAME", () => {
  const tag = currentDesktopTag();
  const child = spawnSync(process.execPath, [smokeTestPath, "--repo", "owner/repo", "--", "--help"], {
    cwd: repoRoot,
    env: { ...process.env, GITHUB_REF_NAME: tag },
    encoding: "utf8",
  });

  assert.equal(
    child.status,
    0,
    `expected exit 0, got ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
  );
  assert.match(child.stdout, /Release smoke test PASSED/i);
});

