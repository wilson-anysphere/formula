import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-gha-action-sha-pins.sh");

const bashProbe = spawnSync("bash", ["--version"], { encoding: "utf8" });
const hasBash = !bashProbe.error && bashProbe.status === 0;

test("check-gha-action-sha-pins bounds directory scans (perf guardrail)", () => {
  const script = readFileSync(scriptPath, "utf8");
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
    const filePath = path.join(tmpdir, name);
    mkdirSync(path.dirname(filePath), { recursive: true });
    writeFileSync(filePath, `${content}\n`, "utf8");
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

test("allows branch-style comment tags for actions that don't publish semver tags (dtolnay/rust-toolchain)", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: dtolnay/rust-toolchain@f7ccc83f9ed1e5b9c81d8a67d7ad1a747e22a561 # master
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

test("fails when a reusable workflow uses a floating ref", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  reuse:
    uses: some-org/some-repo/.github/workflows/reusable.yml@v1
`);
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /must pin/i);
  assert.match(proc.stderr, /some-org\/some-repo\/\.github\/workflows\/reusable\.yml@v1/);
});

test("allows local reusable workflow references by path", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  reuse:
    uses: ./.github/workflows/reusable.yml
`);
  assert.equal(proc.status, 0, proc.stderr);
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

test("allows non-semver ref comments (e.g. # master)", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/checkout@34e114876b0b11c390a56381ad16ebd13914f8d5 # master
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when version comment is not a tag/branch token (e.g. # latest)", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/checkout@34e114876b0b11c390a56381ad16ebd13914f8d5 # latest
`);
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /comment should start/i);
});

test("ignores uses: strings inside YAML block scalars", { skip: !hasBash }, () => {
  const proc = runYaml(`
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: |
          # These lines are inside a YAML block scalar and should NOT be interpreted as workflow steps.
          - uses: actions/checkout@v4
          - uses: actions/setup-node@v4
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
          - uses: actions/checkout@v4
          - uses: actions/setup-node@v4
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

test("directory scan is recursive", { skip: !hasBash }, () => {
  const proc = runDir({
    "nested/ok.yml": `
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/checkout@34e114876b0b11c390a56381ad16ebd13914f8d5 # v4
`,
    "nested/deeper/bad.yml": `
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
