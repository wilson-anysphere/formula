import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-gha-runner-pins.sh");

/**
 * @param {string} yaml
 */
function run(yaml) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-runner-pins-"));
  const workflowPath = path.join(tmpdir, "workflow.yml");
  writeFileSync(workflowPath, `${yaml}\n`, "utf8");

  const proc = spawnSync("bash", [scriptPath, workflowPath], {
    cwd: repoRoot,
    encoding: "utf8",
  });

  rmSync(tmpdir, { recursive: true, force: true });
  if (proc.error) throw proc.error;
  return proc;
}

test("passes when runner images are pinned", () => {
  const proc = run(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when runs-on uses ubuntu-latest", () => {
  const proc = run(`
jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - run: echo ok
`);
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /forbidden/i);
  assert.match(proc.stderr, /ubuntu-latest/);
});

test("fails when a matrix includes windows-latest", () => {
  const proc = run(`
jobs:
  build:
    strategy:
      matrix:
        include:
          - platform: windows-latest
    runs-on: \${{ matrix.platform }}
    steps:
      - run: echo ok
`);
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /windows-latest/);
});

test("ignores comment-only occurrences of *-latest", () => {
  const proc = run(`
# ubuntu-latest
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("ignores *-latest occurrences inside YAML block scalars", () => {
  const proc = run(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: |
          echo ubuntu-latest
          echo windows-latest
`);
  assert.equal(proc.status, 0, proc.stderr);
});

