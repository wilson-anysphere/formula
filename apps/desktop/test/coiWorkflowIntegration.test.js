import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { extractYamlRunSteps, stripHashComments, stripYamlBlockScalarBodies } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function readWorkflow(repoRoot, name) {
  const p = path.join(repoRoot, ".github", "workflows", name);
  return fs.readFileSync(p, "utf8");
}

function assertWorkflowUsesNoBuildForCoi(workflowName, text) {
  const runSteps = extractYamlRunSteps(text);

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
