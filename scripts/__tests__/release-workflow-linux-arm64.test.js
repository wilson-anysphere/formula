import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { stripHashComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const releaseWorkflowPath = path.join(repoRoot, ".github", "workflows", "release.yml");

async function readReleaseWorkflow() {
  return stripHashComments(await readFile(releaseWorkflowPath, "utf8"));
}

/**
 * Extracts a YAML list item's block by scanning forward until either:
 * - the next list item at the same indentation, or
 * - an outdent (indentation decreases), which indicates the list ended.
 *
 * This keeps tests resilient to harmless workflow formatting churn (e.g. inserting
 * extra keys like `env:` before an `if:` guard) without needing a YAML parser.
 *
 * @param {string[]} lines
 * @param {number} startIdx
 */
function yamlListItemBlock(lines, startIdx) {
  const startLine = lines[startIdx] ?? "";
  const indent = startLine.match(/^\s*/)?.[0]?.length ?? 0;
  const nextItemRe = new RegExp(`^\\s{${indent}}-\\s+`);
 
  let endIdx = startIdx + 1;
  for (; endIdx < lines.length; endIdx += 1) {
    const line = lines[endIdx] ?? "";
    if (line.trim() === "") continue;
    const lineIndent = line.match(/^\s*/)?.[0]?.length ?? 0;
    if (lineIndent < indent) break;
    if (nextItemRe.test(line)) break;
  }
  return lines.slice(startIdx, endIdx).join("\n");
}

test("release workflow includes a Linux ARM64 (aarch64) build runner in the matrix", async () => {
  const text = await readReleaseWorkflow();

  const armRunnerRe = /^\s*-\s*platform:\s*ubuntu-24\.04-arm(?:64)?\b/m;
  assert.match(
    text,
    armRunnerRe,
    `Expected ${path.relative(repoRoot, releaseWorkflowPath)} to include an Ubuntu ARM64 runner (ubuntu-24.04-arm64).`,
  );

  const lines = text.split(/\r?\n/);
  const idx = lines.findIndex((line) => /^\s*-\s*platform:\s*ubuntu-24\.04-arm(?:64)?\b/.test(line));
  assert.ok(idx >= 0);
  const snippet = yamlListItemBlock(lines, idx);
  assert.match(snippet, /cache_target:\s*aarch64-unknown-linux-gnu\b/);
});

test("release workflow validates .deb packages on all Linux runners (x86_64 + arm64)", async () => {
  const text = await readReleaseWorkflow();
  const lines = text.split(/\r?\n/);

  const stepNeedle = "Verify Linux .deb package (deps + ldd + desktop integration)";
  const idx = lines.findIndex((line) => line.includes(stepNeedle));
  assert.ok(
    idx >= 0,
    `Expected ${path.relative(repoRoot, releaseWorkflowPath)} to contain a step named: ${stepNeedle}`,
  );

  const snippet = yamlListItemBlock(lines, idx);
  const runnerOsGuard = /if:\s*runner\.os\s*==\s*['"]Linux['"]/;
  const ubuntuPrefixGuard = /if:\s*startsWith\(matrix\.platform,\s*['"]ubuntu-24\.04['"]\)/;
  const x86OnlyGuard = /if:\s*matrix\.platform\s*==\s*['"]ubuntu-24\.04['"]/;

  // This step must run for both Ubuntu x86_64 and Ubuntu ARM64 runs. Historically it regressed to
  // `matrix.platform == 'ubuntu-24.04'`, which skipped the ARM64 build. Accept either a broad Linux
  // guard (`runner.os == 'Linux'`) or an Ubuntu prefix guard (`startsWith(matrix.platform, 'ubuntu-24.04')`).
  assert.doesNotMatch(
    snippet,
    x86OnlyGuard,
    `Expected the Linux .deb validation step NOT to be gated to only ubuntu-24.04 (x86_64).`,
  );
  assert.ok(
    runnerOsGuard.test(snippet) || ubuntuPrefixGuard.test(snippet),
    `Expected the Linux .deb validation step to run for both ubuntu-24.04 and ubuntu-24.04-arm64.\nSaw snippet:\n${snippet}`,
  );
});

/**
 * @param {string[]} lines
 * @param {RegExp} needle
 * @param {number} windowSize
 */
function snippetAfter(lines, needle, windowSize = 40) {
  const idx = lines.findIndex((line) => needle.test(line));
  assert.ok(idx >= 0, `Expected to find line matching ${needle}`);
  return yamlListItemBlock(lines, idx);
}

/**
 * @param {string} snippet
 */
function parseTauriArgsFromSnippet(snippet) {
  const m = snippet.match(/tauri_args:\s*["']([^"']+)["']/);
  assert.ok(m, `Expected to find a tauri_args: \"...\" line.\nSaw snippet:\n${snippet}`);
  return m[1];
}

test("release workflow builds Linux bundles (.AppImage + .deb + .rpm) on x86_64 and ARM64", async () => {
  const text = await readReleaseWorkflow();
  const lines = text.split(/\r?\n/);

  const x64Snippet = snippetAfter(lines, /^\s*-\s*platform:\s*ubuntu-24\.04\b/);
  const arm64Snippet = snippetAfter(lines, /^\s*-\s*platform:\s*ubuntu-24\.04-arm(?:64)?\b/);

  for (const [label, snippet] of [
    ["ubuntu-24.04", x64Snippet],
    ["ubuntu-24.04-arm64", arm64Snippet],
  ]) {
    const tauriArgs = parseTauriArgsFromSnippet(snippet);

    assert.match(
      tauriArgs,
      /--bundles\b/i,
      `Expected ${label} matrix entry to pass --bundles to Tauri (explicit Linux bundle formats).`,
    );
    assert.match(
      tauriArgs,
      /\bappimage\b/i,
      `Expected ${label} matrix entry to build an AppImage bundle.`,
    );
    assert.match(tauriArgs, /\bdeb\b/i, `Expected ${label} matrix entry to build a .deb package.`);
    assert.match(tauriArgs, /\brpm\b/i, `Expected ${label} matrix entry to build a .rpm package.`);
  }
});

test("release workflow verifies Linux artifacts include .AppImage + .deb + .rpm (not x86-only)", async () => {
  const text = await readReleaseWorkflow();
  const lines = text.split(/\r?\n/);

  const stepNeedle = "Verify Linux release artifacts (AppImage + deb + rpm + signatures)";
  const idx = lines.findIndex((line) => line.includes(stepNeedle));
  assert.ok(
    idx >= 0,
    `Expected ${path.relative(repoRoot, releaseWorkflowPath)} to contain a step named: ${stepNeedle}`,
  );

  const snippet = yamlListItemBlock(lines, idx);
  const runnerOsGuard = /if:\s*runner\.os\s*==\s*['"]Linux['"]/;
  const ubuntuPrefixGuard = /if:\s*startsWith\(matrix\.platform,\s*['"]ubuntu-24\.04['"]\)/;
  const x86OnlyGuard = /if:\s*matrix\.platform\s*==\s*['"]ubuntu-24\.04['"]/;

  assert.doesNotMatch(
    snippet,
    x86OnlyGuard,
    `Expected the Linux artifact verification step NOT to be gated to only ubuntu-24.04 (x86_64).`,
  );
  assert.ok(
    runnerOsGuard.test(snippet) || ubuntuPrefixGuard.test(snippet),
    `Expected the Linux artifact verification step to run for both ubuntu-24.04 and ubuntu-24.04-arm64.\nSaw snippet:\n${snippet}`,
  );
});

test("release workflow installs required Linux build deps for both x86_64 and ARM64 runners", async () => {
  const text = await readReleaseWorkflow();
  const lines = text.split(/\r?\n/);

  const stepNeedle = "Install Linux dependencies";
  const idx = lines.findIndex((line) => line.includes(stepNeedle));
  assert.ok(
    idx >= 0,
    `Expected ${path.relative(repoRoot, releaseWorkflowPath)} to contain a step named: ${stepNeedle}`,
  );

  const snippet = yamlListItemBlock(lines, idx);

  // This must run for both ubuntu-24.04 (x86_64) and ubuntu-24.04-arm64.
  const runnerOsGuard = /if:\s*runner\.os\s*==\s*['"]Linux['"]/;
  const ubuntuPrefixGuard = /if:\s*startsWith\(matrix\.platform,\s*['"]ubuntu-24\.04['"]\)/;
  const x86OnlyGuard = /if:\s*matrix\.platform\s*==\s*['"]ubuntu-24\.04['"]/;

  assert.doesNotMatch(
    snippet,
    x86OnlyGuard,
    `Expected the Linux dependency install step NOT to be gated to only ubuntu-24.04 (x86_64).`,
  );
  assert.ok(
    runnerOsGuard.test(snippet) || ubuntuPrefixGuard.test(snippet),
    `Expected the Linux dependency install step to run for both ubuntu-24.04 and ubuntu-24.04-arm64.\nSaw snippet:\n${snippet}`,
  );

  // Keep this lightweight: assert the key GUI/webview deps remain present.
  assert.match(snippet, /\blibgtk-3-dev\b/);
  assert.match(snippet, /\blibwebkit2gtk-4\.1-dev\b/);
});
