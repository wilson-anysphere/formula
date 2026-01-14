import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const sourceScriptPath = path.join(repoRoot, "scripts", "ci", "check-pnpm-version-pins.sh");
const scriptContents = readFileSync(sourceScriptPath, "utf8");

const bashProbe = spawnSync("bash", ["--version"], { encoding: "utf8" });
const hasBash = !bashProbe.error && bashProbe.status === 0;

const gitProbe = spawnSync("git", ["--version"], { encoding: "utf8" });
const hasGit = !gitProbe.error && gitProbe.status === 0;

const canRun = hasBash && hasGit;

/**
 * Runs the pnpm pin guard in a temporary git repo.
 * @param {Record<string, string>} files
 */
function run(files) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-pnpm-pins-"));
  try {
    let proc = spawnSync("git", ["init"], { cwd: tmpdir, encoding: "utf8" });
    assert.equal(proc.status, 0, proc.stderr);

    mkdirSync(path.join(tmpdir, ".github", "workflows"), { recursive: true });
    mkdirSync(path.join(tmpdir, "scripts", "ci"), { recursive: true });

    // Install the script under test in the temp repo so its repo_root resolution works.
    const tmpScriptPath = path.join(tmpdir, "scripts", "ci", "check-pnpm-version-pins.sh");
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

test("passes when pnpm pins match package.json (action + corepack)", { skip: !canRun }, () => {
  const proc = run({
    "package.json": `{
  "name": "formula",
  "private": true,
  "packageManager": "pnpm@9.0.0"
 }`,
    ".github/workflows/ci.yml": `
name: CI
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: pnpm/action-setup@v4
        with:
          version: 9.0.0
`,
    ".github/workflows/security.yml": `
name: Security
jobs:
  scan:
    runs-on: ubuntu-24.04
    steps:
      - run: corepack prepare pnpm@9.0.0 --activate
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /pnpm version pins match package\.json/i);
});

test("ignores commented-out pnpm/action-setup steps", { skip: !canRun }, () => {
  const proc = run({
    "package.json": `{
  "name": "formula",
  "private": true,
  "packageManager": "pnpm@9.0.0"
 }`,
    ".github/workflows/ci.yml": `
name: CI
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: pnpm/action-setup@v4
        with:
          version: 9.0.0
      # - uses: pnpm/action-setup@v4
      #   with:
      #     version: 8.0.0
`,
    ".github/workflows/security.yml": `
name: Security
jobs:
  scan:
    runs-on: ubuntu-24.04
    steps:
      - run: corepack prepare pnpm@9.0.0 --activate
`,
  });

  assert.equal(proc.status, 0, proc.stderr);
});

test("ignores pnpm/action-setup strings inside YAML block scalars", { skip: !canRun }, () => {
  const proc = run({
    "package.json": `{
  "name": "formula",
  "private": true,
  "packageManager": "pnpm@9.0.0"
 }`,
    ".github/workflows/ci.yml": `
name: CI
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: pnpm/action-setup@v4
        with:
          version: 9.0.0
      - name: Run script mentions action-setup YAML
        run: |
          # Script content; should not count as workflow YAML.
          - uses: pnpm/action-setup@v4
          echo ok
`,
    ".github/workflows/security.yml": `
name: Security
jobs:
  scan:
    runs-on: ubuntu-24.04
    steps:
      - run: corepack prepare pnpm@9.0.0 --activate
`,
  });

  assert.equal(proc.status, 0, proc.stderr);
});

test("ignores pnpm/action-setup strings inside inline run steps", { skip: !canRun }, () => {
  const proc = run({
    "package.json": `{
  "name": "formula",
  "private": true,
  "packageManager": "pnpm@9.0.0"
 }`,
    ".github/workflows/ci.yml": `
name: CI
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: pnpm/action-setup@v4
        with:
          version: 9.0.0
      - name: Inline run echoes a YAML-ish string
        run: echo "uses: pnpm/action-setup@v4"
`,
  });

  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when pnpm/action-setup pin mismatches package.json", { skip: !canRun }, () => {
  const proc = run({
    "package.json": `{
  "name": "formula",
  "private": true,
  "packageManager": "pnpm@9.0.0"
}`,
    ".github/workflows/ci.yml": `
name: CI
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: pnpm/action-setup@v4
        with:
          version: 9.0.1
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /pnpm version pin mismatch/i);
});

test("fails when corepack prepare pnpm@... mismatches package.json", { skip: !canRun }, () => {
  const proc = run({
    "package.json": `{
  "name": "formula",
  "private": true,
  "packageManager": "pnpm@9.0.0"
}`,
    ".github/workflows/ci.yml": `
name: CI
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: pnpm/action-setup@v4
        with:
          version: 9.0.0
`,
    ".github/workflows/security.yml": `
name: Security
jobs:
  scan:
    runs-on: ubuntu-24.04
    steps:
      - run: corepack prepare pnpm@8.9.0 --activate
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /corepack prepare/i);
});

test("fails when mise.toml pnpm pin mismatches package.json", { skip: !canRun }, () => {
  const proc = run({
    "package.json": `{
  "name": "formula",
  "private": true,
  "packageManager": "pnpm@9.0.0"
}`,
    "mise.toml": `
[tools]
pnpm = "8.0.0"
`,
    ".github/workflows/ci.yml": `
name: CI
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: pnpm/action-setup@v4
        with:
          version: 9.0.0
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /mise\.toml/i);
});
