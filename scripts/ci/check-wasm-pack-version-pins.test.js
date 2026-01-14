import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const sourceScriptPath = path.join(repoRoot, "scripts", "ci", "check-wasm-pack-version-pins.sh");
const scriptContents = readFileSync(sourceScriptPath, "utf8");

const bashProbe = spawnSync("bash", ["--version"], { encoding: "utf8" });
const hasBash = !bashProbe.error && bashProbe.status === 0;

const gitProbe = spawnSync("git", ["--version"], { encoding: "utf8" });
const hasGit = !gitProbe.error && gitProbe.status === 0;

const canRun = hasBash && hasGit;

/**
 * Runs the wasm-pack pin guard in a temporary git repo.
 * @param {Record<string, string>} files
 */
function run(files) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-wasm-pack-pins-"));
  try {
    let proc = spawnSync("git", ["init"], { cwd: tmpdir, encoding: "utf8" });
    assert.equal(proc.status, 0, proc.stderr);

    mkdirSync(path.join(tmpdir, ".github", "workflows"), { recursive: true });
    mkdirSync(path.join(tmpdir, "scripts", "ci"), { recursive: true });

    const tmpScriptPath = path.join(tmpdir, "scripts", "ci", "check-wasm-pack-version-pins.sh");
    writeFileSync(tmpScriptPath, scriptContents, "utf8");

    for (const [name, content] of Object.entries(files)) {
      const filePath = path.join(tmpdir, name);
      mkdirSync(path.dirname(filePath), { recursive: true });
      writeFileSync(filePath, `${content}\n`, "utf8");
      proc = spawnSync("git", ["add", name], { cwd: tmpdir, encoding: "utf8" });
      assert.equal(proc.status, 0, proc.stderr);
    }

    return spawnSync("bash", [tmpScriptPath], { cwd: tmpdir, encoding: "utf8" });
  } finally {
    rmSync(tmpdir, { recursive: true, force: true });
  }
}

test("passes when all workflows pin the same wasm-pack version", { skip: !canRun }, () => {
  const proc = run({
    ".github/workflows/ci.yml": `
name: CI
env:
  WASM_PACK_VERSION: 0.13.1
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`,
    ".github/workflows/other.yml": `
name: Other
env:
  WASM_PACK_VERSION: 0.13.1
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /wasm-pack version pins match/i);
});

test("passes when workflows use a v/V prefix for the wasm-pack version", { skip: !canRun }, () => {
  const proc = run({
    ".github/workflows/ci.yml": `
name: CI
env:
  WASM_PACK_VERSION: V0.13.1
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`,
    ".github/workflows/other.yml": `
name: Other
env:
  WASM_PACK_VERSION: v0.13.1
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /wasm-pack version pins match/i);
});

test("ignores WASM_PACK_VERSION strings inside YAML block scalars", { skip: !canRun }, () => {
  const proc = run({
    ".github/workflows/ci.yml": `
name: CI
env:
  WASM_PACK_VERSION: 0.13.1
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`,
    ".github/workflows/other.yml": `
name: Other
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: Script mentions WASM_PACK_VERSION yaml-ish text
        run: |
          # Script content; should not count as workflow YAML.
          WASM_PACK_VERSION: 0.13.2
          echo ok
env:
  WASM_PACK_VERSION: 0.13.1
`,
  });

  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /wasm-pack version pins match/i);
});

test("fails when workflows pin different wasm-pack versions", { skip: !canRun }, () => {
  const proc = run({
    ".github/workflows/ci.yml": `
name: CI
env:
  WASM_PACK_VERSION: 0.13.1
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`,
    ".github/workflows/other.yml": `
name: Other
env:
  WASM_PACK_VERSION: 0.13.2
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /mismatch/i);
});
