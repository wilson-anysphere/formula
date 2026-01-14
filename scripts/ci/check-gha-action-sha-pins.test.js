import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-gha-action-sha-pins.sh");

const bashProbe = spawnSync("bash", ["--version"], { encoding: "utf8" });
const hasBash = !bashProbe.error && bashProbe.status === 0;

/**
 * @param {string} yaml
 * @returns {ReturnType<typeof spawnSync>}
 */
function runYaml(yaml) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-action-sha-pins-"));
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

/**
 * @param {Record<string, string>} files
 * @returns {ReturnType<typeof spawnSync>}
 */
function runDir(files) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-action-sha-pins-dir-"));
  for (const [name, content] of Object.entries(files)) {
    writeFileSync(path.join(tmpdir, name), `${content}\n`, "utf8");
  }
  const proc = spawnSync("bash", [scriptPath, tmpdir], { cwd: repoRoot, encoding: "utf8" });
  rmSync(tmpdir, { recursive: true, force: true });
  if (proc.error) throw proc.error;
  return proc;
}

test("passes when actions are pinned to a SHA and have a version comment", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/checkout@34e114876b0b11c390a56381ad16ebd13914f8d5 # v4
      - run: echo ok
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when actions use a floating tag (e.g. @v4)", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/checkout@v4
`);
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /must pin/i);
  assert.match(proc.stderr, /actions\/checkout@v4/);
});

test("fails when SHA pin is missing a trailing version comment", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/checkout@34e114876b0b11c390a56381ad16ebd13914f8d5
`);
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /missing.*comment/i);
});

test("fails when version comment does not start with a semver-ish tag", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/checkout@34e114876b0b11c390a56381ad16ebd13914f8d5 # latest
`);
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /comment should start with a version tag/i);
});

test("ignores uses: strings inside YAML block scalars", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: |
          echo "uses: actions/checkout@v4"
          echo "uses: actions/setup-node@v4"
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("ignores uses: strings inside block scalars with indentation/chomp indicators", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: |2-
          echo "uses: actions/checkout@v4"
          echo "uses: actions/setup-node@v4"
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("accepts directories and fails if any contained workflow is not SHA pinned", { skip: !hasBash }, () => {
  const proc = runDir({
    "ok.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/checkout@34e114876b0b11c390a56381ad16ebd13914f8d5 # v4
`,
    "bad.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/checkout@v4
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /actions\/checkout@v4/);
});
