import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripHashComments, stripYamlBlockScalarBodies } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function readWorkflow(repoRoot, name) {
  const p = path.join(repoRoot, ".github", "workflows", name);
  return fs.readFileSync(p, "utf8");
}

/**
 * Extract `run:` steps (inline + block scalars) from a workflow, while ignoring YAML block scalar
 * bodies that are *not* run scripts (e.g. env vars, github-script inputs).
 *
 * This prevents YAML-ish strings like `run: pnpm ...` embedded inside other multiline scalars from
 * satisfying or failing assertions in this test.
 *
 * @param {string} workflowText
 * @returns {Array<{ line: number, script: string }>}
 */
function extractWorkflowRunSteps(workflowText) {
  const stripped = stripHashComments(workflowText);
  const lines = stripped.split(/\r?\n/);
  /** @type {Array<{ line: number, script: string }>} */
  const runSteps = [];

  let inBlock = false;
  let blockIndent = 0;
  let blockIsRun = false;
  let runBlockStartLine = 0;
  /** @type {string[]} */
  let runBlockBody = [];

  const blockRe = /:[\t ]*[>|][0-9+-]*[\t ]*$/;

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i] ?? "";
    const trimmed = line.trim();
    const indent = line.match(/^\s*/)?.[0]?.length ?? 0;

    if (inBlock) {
      if (trimmed === "") {
        // Blank lines can appear in block scalars with any indentation.
        if (blockIsRun) runBlockBody.push("");
        continue;
      }
      if (indent > blockIndent) {
        if (blockIsRun) runBlockBody.push(line);
        continue;
      }

      // Block scalar ended; flush any collected run script.
      if (blockIsRun) {
        runSteps.push({ line: runBlockStartLine, script: runBlockBody.join("\n") });
        runBlockBody = [];
      }
      inBlock = false;
      blockIsRun = false;
    }

    const isBlockScalarHeader = blockRe.test(line.trimEnd());
    if (isBlockScalarHeader) {
      // Track block scalars so we can skip non-run scalar bodies.
      inBlock = true;
      blockIndent = indent;
      blockIsRun = /^\s*-?\s*run:\s*[>|]/.test(line);
      runBlockStartLine = i + 1;
      runBlockBody = [];
      continue;
    }

    const m = line.match(/^\s*-?\s*run:\s*(.+)$/);
    if (!m) continue;
    const rest = (m[1] ?? "").trimEnd();

    // Ignore `run:` keys with no command (rare/invalid in workflows).
    if (rest === "") continue;

    runSteps.push({ line: i + 1, script: rest });
  }

  if (inBlock && blockIsRun) {
    runSteps.push({ line: runBlockStartLine, script: runBlockBody.join("\n") });
  }

  return runSteps;
}

function assertWorkflowUsesNoBuildForCoi(workflowName, text) {
  const runSteps = extractWorkflowRunSteps(text);

  const needle = "pnpm -C apps/desktop check:coi";
  const matches = runSteps.filter((step) => step.script.includes(needle));

  assert.ok(matches.length > 0, `expected ${workflowName} to invoke pnpm -C apps/desktop check:coi in a run step`);

  for (const { script, line } of matches) {
    assert.ok(
      script.includes("--no-build") || script.includes("FORMULA_COI_NO_BUILD"),
      `expected ${workflowName}:${line} COI invocation to use --no-build (or FORMULA_COI_NO_BUILD).\nFound script:\n${script}`,
    );
  }

  // Heuristic: ensure a Tauri build step exists somewhere before the first COI invocation so the workflow
  // can reuse already-built artifacts.
  const searchLines = stripYamlBlockScalarBodies(stripHashComments(text)).split(/\r?\n/);
  const firstCoi = (matches[0]?.line ?? 1) - 1;
  const tauriActionRe = /^\s*uses:\s*tauri-apps\/tauri-action\b/;
  const firstTauriAction = searchLines.findIndex((l) => tauriActionRe.test(l ?? ""));
  assert.ok(
    firstTauriAction !== -1 && firstTauriAction < firstCoi,
    `expected ${workflowName} to include a tauri-apps/tauri-action step before the COI check`,
  );
}

test("release + dry-run workflows run COI smoke checks against prebuilt artifacts", () => {
  const repoRoot = path.join(__dirname, "..", "..", "..");

  const release = readWorkflow(repoRoot, "release.yml");
  assertWorkflowUsesNoBuildForCoi("release.yml", release);

  const dryRun = readWorkflow(repoRoot, "desktop-bundle-dry-run.yml");
  assertWorkflowUsesNoBuildForCoi("desktop-bundle-dry-run.yml", dryRun);
});

test("extractWorkflowRunSteps ignores run-like strings inside non-run YAML block scalars", () => {
  const workflow = `
name: Example
jobs:
  build:
    runs-on: ubuntu-24.04
    env:
      NOTES: |
        run: pnpm -C apps/desktop check:coi
    steps:
      - run: echo ok
      - run: |
          pnpm -C apps/desktop check:coi --no-build
`;

  const steps = extractWorkflowRunSteps(workflow);
  assert.ok(steps.some((s) => s.script.includes("echo ok")));

  const coiSteps = steps.filter((s) => s.script.includes("pnpm -C apps/desktop check:coi"));
  assert.equal(coiSteps.length, 1, `expected exactly one COI run step; got:\n${coiSteps.map((s) => s.script).join("\n\n")}`);
});
