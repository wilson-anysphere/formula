import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function readWorkflow(repoRoot, name) {
  const p = path.join(repoRoot, ".github", "workflows", name);
  return fs.readFileSync(p, "utf8");
}

function assertWorkflowUsesNoBuildForCoi(workflowName, text) {
  const matches = [...text.matchAll(/pnpm\s+-C\s+apps\/desktop\s+check:coi[^\n]*/g)].map((m) => m[0]);
  assert.ok(matches.length > 0, `expected ${workflowName} to invoke pnpm -C apps/desktop check:coi`);

  for (const m of matches) {
    assert.ok(
      m.includes("--no-build") || m.includes("FORMULA_COI_NO_BUILD"),
      `expected ${workflowName} COI invocation to use --no-build (or FORMULA_COI_NO_BUILD). Found: ${m}`,
    );
  }

  // Heuristic: ensure a Tauri build step exists somewhere before the COI invocation so the workflow
  // can reuse already-built artifacts.
  const firstCoi = text.indexOf(matches[0] ?? "");
  const tauriActionRe = /^\s*uses:\s*tauri-apps\/tauri-action\b/m;
  const firstTauriAction = text.search(tauriActionRe);
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
