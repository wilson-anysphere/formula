import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const sourceScriptPath = path.join(repoRoot, "scripts", "ci", "check-workflow-step-name-colons.sh");
const scriptContents = readFileSync(sourceScriptPath, "utf8");

const bashProbe = spawnSync("bash", ["--version"], { encoding: "utf8" });
const hasBash = !bashProbe.error && bashProbe.status === 0;

const gitProbe = spawnSync("git", ["--version"], { encoding: "utf8" });
const hasGit = !gitProbe.error && gitProbe.status === 0;

const canRun = hasBash && hasGit;

/**
 * Runs the colon guard script against a temporary repository with a single workflow file.
 * @param {string} workflowYaml
 */
function run(workflowYaml) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-step-name-colons-"));
  try {
    let proc = spawnSync("git", ["init"], { cwd: tmpdir, encoding: "utf8" });
    assert.equal(proc.status, 0, proc.stderr);

    mkdirSync(path.join(tmpdir, ".github", "workflows"), { recursive: true });
    mkdirSync(path.join(tmpdir, "scripts", "ci"), { recursive: true });

    const workflowPath = path.join(tmpdir, ".github", "workflows", "workflow.yml");
    writeFileSync(workflowPath, `${workflowYaml}\n`, "utf8");

    const tmpScriptPath = path.join(tmpdir, "scripts", "ci", "check-workflow-step-name-colons.sh");
    writeFileSync(tmpScriptPath, scriptContents, "utf8");

    proc = spawnSync("git", ["add", ".github/workflows/workflow.yml"], { cwd: tmpdir, encoding: "utf8" });
    assert.equal(proc.status, 0, proc.stderr);

    return spawnSync("bash", [tmpScriptPath], { cwd: tmpdir, encoding: "utf8" });
  } finally {
    rmSync(tmpdir, { recursive: true, force: true });
  }
}

test("passes when colon-containing name fields are quoted", { skip: !canRun }, () => {
  const proc = run(`
name: "Guard: OK"
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: "Guard: Step"
        run: echo ok
`);
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /OK/i);
});

test("fails when workflow name contains an unquoted colon", { skip: !canRun }, () => {
  const proc = run(`
name: Guard: Bad
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`);
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /unquoted/i);
  assert.match(proc.stderr, /workflow\.yml/);
});

test("passes when colon is part of a token (e.g. node:test)", { skip: !canRun }, () => {
  const proc = run(`
name: node:test
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: node:test
        run: echo ok
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("ignores name: strings inside YAML block scalars", { skip: !canRun }, () => {
  const proc = run(`
name: "OK"
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - name: "Step"
        run: |
          # These lines are inside a YAML block scalar and should NOT be interpreted as workflow keys.
          name: Guard: Bad
          echo ok
`);
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /OK/i);
});
