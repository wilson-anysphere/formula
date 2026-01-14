import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-toml-validity.py");

const gitProbe = spawnSync("git", ["--version"], { encoding: "utf8" });
const hasGit = !gitProbe.error && gitProbe.status === 0;

const pythonProbe = spawnSync("python3", ["--version"], { encoding: "utf8" });
const hasPython3 = !pythonProbe.error && pythonProbe.status === 0;

const canRun = hasGit && hasPython3;

/**
 * @param {Record<string, string>} files
 * @returns {ReturnType<typeof spawnSync>}
 */
function run(files) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-toml-validity-"));
  try {
    let proc = spawnSync("git", ["init"], { cwd: tmpdir, encoding: "utf8" });
    assert.equal(proc.status, 0, proc.stderr);

    for (const [name, content] of Object.entries(files)) {
      const filePath = path.join(tmpdir, name);
      writeFileSync(filePath, `${content}\n`, "utf8");
      proc = spawnSync("git", ["add", name], { cwd: tmpdir, encoding: "utf8" });
      assert.equal(proc.status, 0, proc.stderr);
    }

    proc = spawnSync("python3", [scriptPath], { cwd: tmpdir, encoding: "utf8" });
    return proc;
  } finally {
    rmSync(tmpdir, { recursive: true, force: true });
  }
}

test("skips when no tracked TOML files exist", { skip: !canRun }, () => {
  const proc = run({});
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stderr, /No tracked \.toml files found/i);
});

test("passes when tracked TOML files are valid", { skip: !canRun }, () => {
  const proc = run({
    "ok.toml": `
[package]
name = "ok"
version = "1.0.0"
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /TOML parse: OK/i);
});

test("fails when tracked TOML contains duplicate keys", { skip: !canRun }, () => {
  const proc = run({
    "bad.toml": `
[package]
name = "a"
name = "b"
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /failed to parse TOML/i);
  assert.match(proc.stderr, /bad\.toml/);
});

