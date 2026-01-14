import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripHashComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-gha-runner-pins.sh");

const bashProbe = spawnSync("bash", ["--version"], { encoding: "utf8" });
const hasBash = !bashProbe.error && bashProbe.status === 0;

test("check-gha-runner-pins bounds directory scans (perf guardrail)", () => {
  const script = stripHashComments(readFileSync(scriptPath, "utf8"));
  const idx = script.indexOf('find "$path"');
  assert.ok(idx >= 0, "Expected script to enumerate workflow directories via find \"$path\".");
  const snippet = script.slice(idx, idx + 120);
  assert.ok(
    snippet.includes("-maxdepth"),
    `Expected workflow file discovery to be bounded with -maxdepth.\nSaw snippet:\n${snippet}`,
  );
});

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

test("passes when runner images are pinned", { skip: !hasBash }, () => {
  const proc = run(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when runs-on uses ubuntu-latest", { skip: !hasBash }, () => {
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

test("fails when a matrix includes windows-latest", { skip: !hasBash }, () => {
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

test("ignores comment-only occurrences of *-latest", { skip: !hasBash }, () => {
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

test("ignores *-latest occurrences inside YAML block scalars", { skip: !hasBash }, () => {
  const proc = run(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: |
          echo ubuntu-latest
          echo windows-latest
          
          echo macos-latest
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("ignores *-latest occurrences inside single-line run commands", { skip: !hasBash }, () => {
  const proc = run(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ubuntu-latest
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("ignores *-latest occurrences inside block scalars with indentation/chomp indicators", { skip: !hasBash }, () => {
  const proc = run(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: |2-
          echo ubuntu-latest
          echo windows-latest
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("supports scanning directories of workflow files", { skip: !hasBash }, () => {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-runner-pins-dir-"));
  writeFileSync(
    path.join(tmpdir, "pinned.yml"),
    "jobs:\n  build:\n    runs-on: ubuntu-24.04\n    steps:\n      - run: echo ok\n",
    "utf8",
  );
  writeFileSync(
    path.join(tmpdir, "bad.yml"),
    "jobs:\n  build:\n    runs-on: ubuntu-latest\n    steps:\n      - run: echo ok\n",
    "utf8",
  );

  const proc = spawnSync("bash", [scriptPath, tmpdir], { cwd: repoRoot, encoding: "utf8" });
  rmSync(tmpdir, { recursive: true, force: true });
  if (proc.error) throw proc.error;

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /ubuntu-latest/);
  assert.match(proc.stderr, /bad\.yml/);
});
