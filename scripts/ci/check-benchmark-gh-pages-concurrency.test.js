import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripHashComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-benchmark-gh-pages-concurrency.sh");

const bashProbe = spawnSync("bash", ["--version"], { encoding: "utf8" });
const hasBash = !bashProbe.error && bashProbe.status === 0;

const canRun = hasBash;

test("check-benchmark-gh-pages-concurrency bounds directory scans (perf guardrail)", () => {
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
 * @returns {ReturnType<typeof spawnSync>}
 */
function runYaml(yaml) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-benchmark-gh-pages-"));
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
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-benchmark-gh-pages-dir-"));
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

test("passes when benchmark-action does not auto-push", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  pr:
    runs-on: ubuntu-24.04
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: false
  other:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("passes when auto-push is explicitly disabled via expression", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  pr:
    runs-on: ubuntu-24.04
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: \${{ false }}
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when benchmark-action auto-pushes without shared concurrency", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  publish:
    runs-on: ubuntu-24.04
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: true
`);
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /serialize pushes/i);
});

test("passes when benchmark-action auto-pushes with shared concurrency", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  publish:
    runs-on: ubuntu-24.04
    concurrency:
      group: benchmark-gh-pages-publish
      cancel-in-progress: false
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: true
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when concurrency group exists but is attached to a different job", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  other:
    runs-on: ubuntu-24.04
    concurrency:
      group: benchmark-gh-pages-publish
      cancel-in-progress: false
    steps:
      - run: echo ok
  publish:
    runs-on: ubuntu-24.04
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: true
`);
  assert.notEqual(proc.status, 0);
});

test("passes when concurrency group is quoted", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  publish:
    runs-on: ubuntu-24.04
    concurrency:
      group: "benchmark-gh-pages-publish"
      cancel-in-progress: false
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: true
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("passes when workflow-level concurrency serializes publishes", { skip: !canRun }, () => {
  const proc = runYaml(`
concurrency:
  group: benchmark-gh-pages-publish
  cancel-in-progress: false
jobs:
  publish:
    runs-on: ubuntu-24.04
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: true
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("passes when workflow-level concurrency uses inline mapping form", { skip: !canRun }, () => {
  const proc = runYaml(`
concurrency: { group: benchmark-gh-pages-publish, cancel-in-progress: false }
jobs:
  publish:
    runs-on: ubuntu-24.04
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: true
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("passes when job-level concurrency uses scalar form", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  publish:
    runs-on: ubuntu-24.04
    concurrency: benchmark-gh-pages-publish
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: true
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("passes when job-level concurrency uses inline mapping form", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  publish:
    runs-on: ubuntu-24.04
    concurrency: { group: benchmark-gh-pages-publish, cancel-in-progress: false }
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: true
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("requires concurrency when auto-push is a truthy expression", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  publish:
    runs-on: ubuntu-24.04
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: \${{ github.event_name == 'push' }}
`);
  assert.notEqual(proc.status, 0);
});

test("detects auto-push inside with inline mapping", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  publish:
    runs-on: ubuntu-24.04
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with: { auto-push: true }
`);
  assert.notEqual(proc.status, 0);
});

test("passes when with inline mapping auto-push is disabled", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  publish:
    runs-on: ubuntu-24.04
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with: { auto-push: false }
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("ignores auto-push occurrences inside YAML block scalars", { skip: !canRun }, () => {
  const proc = runYaml(`
jobs:
  publish:
    runs-on: ubuntu-24.04
    steps:
      - run: |
          # This is inside a YAML block scalar and should be ignored.
          auto-push: true
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: false
`);
  assert.equal(proc.status, 0, proc.stderr);
});

test("directory scan fails if any workflow auto-pushes without concurrency", { skip: !canRun }, () => {
  const proc = runDir({
    "ok.yml": `
jobs:
  publish:
    runs-on: ubuntu-24.04
    concurrency:
      group: benchmark-gh-pages-publish
      cancel-in-progress: false
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: true
`,
    "bad.yml": `
jobs:
  publish:
    runs-on: ubuntu-24.04
    steps:
      - uses: benchmark-action/github-action-benchmark@v1
        with:
          auto-push: true
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /bad\.yml/);
});
