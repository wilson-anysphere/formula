import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const sourceScriptPath = path.join(repoRoot, "scripts", "ci", "check-node-version-pins.sh");
const scriptContents = readFileSync(sourceScriptPath, "utf8");

const bashProbe = spawnSync("bash", ["--version"], { encoding: "utf8" });
const hasBash = !bashProbe.error && bashProbe.status === 0;

const gitProbe = spawnSync("git", ["--version"], { encoding: "utf8" });
const hasGit = !gitProbe.error && gitProbe.status === 0;

const canRun = hasBash && hasGit;

/**
 * Runs the Node pin guard in a temporary git repo.
 * @param {Record<string, string>} files
 */
function run(files) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-node-pins-"));
  try {
    let proc = spawnSync("git", ["init"], { cwd: tmpdir, encoding: "utf8" });
    assert.equal(proc.status, 0, proc.stderr);

    mkdirSync(path.join(tmpdir, ".github", "workflows"), { recursive: true });
    mkdirSync(path.join(tmpdir, "scripts", "ci"), { recursive: true });

    const tmpScriptPath = path.join(tmpdir, "scripts", "ci", "check-node-version-pins.sh");
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

test("passes when node major pins are consistent and setup-node uses env.NODE_VERSION", { skip: !canRun }, () => {
  const proc = run({
    ".nvmrc": "22",
    ".github/workflows/ci.yml": `
name: CI
env:
  NODE_VERSION: 22
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/setup-node@v4
        with:
          node-version: \${{ env.NODE_VERSION }}
      - run: node -v
`,
    ".github/workflows/release.yml": `
name: Release
env:
  NODE_VERSION: 22
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/setup-node@v4
        with:
          node-version: \${{ env.NODE_VERSION }}
      - run: node -v
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Node version pins match/i);
});

test("ignores node-version strings inside YAML block scalars", { skip: !canRun }, () => {
  const proc = run({
    ".nvmrc": "22",
    ".github/workflows/ci.yml": `
name: CI
env:
  NODE_VERSION: 22
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/setup-node@v4
        with:
          node-version: \${{ env.NODE_VERSION }}
      - name: Script mentions node-version in a block scalar
        run: |
          # This is script content and should not count as workflow configuration.
          node-version: 999
          echo ok
`,
    ".github/workflows/release.yml": `
name: Release
env:
  NODE_VERSION: 22
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/setup-node@v4
        with:
          node-version: \${{ env.NODE_VERSION }}
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Node version pins match/i);
});

test("ignores NODE_VERSION strings inside YAML block scalars", { skip: !canRun }, () => {
  const proc = run({
    ".nvmrc": "22",
    ".github/workflows/ci.yml": `
name: CI
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/setup-node@v4
        with:
          node-version: \${{ env.NODE_VERSION }}
      - name: Script mentions NODE_VERSION yaml
        run: |
          # Script content; should not count as workflow YAML.
          NODE_VERSION: 999
          echo ok
env:
  NODE_VERSION: 22
`,
    ".github/workflows/release.yml": `
name: Release
env:
  NODE_VERSION: 22
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/setup-node@v4
        with:
          node-version: \${{ env.NODE_VERSION }}
`,
  });

  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Node version pins match/i);
});

test("fails when NODE_VERSION differs between CI and release workflows", { skip: !canRun }, () => {
  const proc = run({
    ".nvmrc": "22",
    ".github/workflows/ci.yml": `
name: CI
env:
  NODE_VERSION: 22
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/setup-node@v4
        with:
          node-version: \${{ env.NODE_VERSION }}
`,
    ".github/workflows/release.yml": `
name: Release
env:
  NODE_VERSION: 23
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/setup-node@v4
        with:
          node-version: \${{ env.NODE_VERSION }}
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Node major pin mismatch/i);
});

test("fails when setup-node uses a non-env node-version pin", { skip: !canRun }, () => {
  const proc = run({
    ".nvmrc": "22",
    ".github/workflows/ci.yml": `
name: CI
env:
  NODE_VERSION: 22
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/setup-node@v4
        with:
          node-version: 22
`,
    ".github/workflows/release.yml": `
name: Release
env:
  NODE_VERSION: 22
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/setup-node@v4
        with:
          node-version: \${{ env.NODE_VERSION }}
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /node-version/i);
});
