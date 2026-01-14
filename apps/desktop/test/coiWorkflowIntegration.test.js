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
  /** @type {Array<{ index: number, snippet: string }>} */
  const matches = [];
  const re = /pnpm\s+-C\s+apps\/desktop\s+check:coi/g;
  for (const m of text.matchAll(re)) {
    const idx = m.index ?? -1;
    if (idx < 0) continue;
    const lineEnd = text.indexOf("\n", idx);
    const snippet = (lineEnd === -1 ? text.slice(idx) : text.slice(idx, lineEnd)).trim();
    matches.push({ index: idx, snippet });
  }

  assert.ok(matches.length > 0, `expected ${workflowName} to invoke pnpm -C apps/desktop check:coi`);

  for (const { index, snippet } of matches) {
    // Look ahead a short window so wrapped/multiline YAML `run: |` commands still match.
    const window = text.slice(index, Math.min(text.length, index + 200));
    assert.ok(
      window.includes("--no-build") || window.includes("FORMULA_COI_NO_BUILD"),
      `expected ${workflowName} COI invocation to use --no-build (or FORMULA_COI_NO_BUILD). Found: ${snippet}`,
    );
  }

  // Heuristic: ensure a Tauri build step exists somewhere before the first COI invocation so the workflow
  // can reuse already-built artifacts.
  const firstCoi = matches[0]?.index ?? -1;
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
