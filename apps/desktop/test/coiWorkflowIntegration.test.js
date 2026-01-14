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
  const lines = text.split(/\r?\n/);

  /** @type {Array<{ line: number, snippet: string, window: string }>} */
  const matches = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] ?? "";
    const m = line.match(/^(\s*)run:\s*(.*)$/);
    if (!m) continue;

    const indent = m[1]?.length ?? 0;
    const rest = (m[2] ?? "").trimEnd();

    const isBlock = rest === "|" || rest === "|-" || rest === ">" || rest === ">-";
    let blockText = rest;
    if (isBlock) {
      const body = [];
      for (let j = i + 1; j < lines.length; j++) {
        const bodyLine = lines[j] ?? "";
        const bodyIndent = bodyLine.match(/^\s*/)?.[0]?.length ?? 0;
        if (bodyLine.trim() !== "" && bodyIndent <= indent) break;
        body.push(bodyLine);
      }
      blockText = body.join("\n");
    }

    if (!blockText.includes("pnpm -C apps/desktop check:coi")) continue;

    // Capture a small window (line + a few following lines) so we can enforce --no-build even
    // when a command is wrapped across multiple lines.
    const window = [line, ...(lines.slice(i + 1, i + 8) ?? [])].join("\n");
    matches.push({
      line: i + 1,
      snippet: line.trim(),
      window,
    });
  }

  assert.ok(matches.length > 0, `expected ${workflowName} to invoke pnpm -C apps/desktop check:coi in a run step`);

  for (const { window, snippet, line } of matches) {
    assert.ok(
      window.includes("--no-build") || window.includes("FORMULA_COI_NO_BUILD"),
      `expected ${workflowName}:${line} COI invocation to use --no-build (or FORMULA_COI_NO_BUILD). Found: ${snippet}`,
    );
  }

  // Heuristic: ensure a Tauri build step exists somewhere before the first COI invocation so the workflow
  // can reuse already-built artifacts.
  const firstCoi = (matches[0]?.line ?? 1) - 1;
  const tauriActionRe = /^\s*uses:\s*tauri-apps\/tauri-action\b/m;
  const firstTauriAction = lines.findIndex((l) => tauriActionRe.test(l ?? ""));
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
