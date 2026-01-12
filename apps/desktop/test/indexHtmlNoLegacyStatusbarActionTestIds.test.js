import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function extractTestIdsFromIndexHtml(html) {
  const dataTestIdRegex = /\bdata-testid\s*=\s*(["'])(.*?)\1/g;
  /** @type {Set<string>} */
  const ids = new Set();
  for (const match of html.matchAll(dataTestIdRegex)) {
    ids.add(match[2]);
  }
  return ids;
}

function findDuplicateTestIdsInIndexHtml(html) {
  const dataTestIdRegex = /\bdata-testid\s*=\s*(["'])(.*?)\1/g;
  /** @type {Map<string, number>} */
  const counts = new Map();
  for (const match of html.matchAll(dataTestIdRegex)) {
    const testId = match[2];
    counts.set(testId, (counts.get(testId) ?? 0) + 1);
  }

  return [...counts.entries()]
    .filter(([, count]) => count > 1)
    .map(([testId, count]) => `${testId} (${count})`)
    .sort((a, b) => a.localeCompare(b));
}

function extractRibbonTestIdsFromSource(source) {
  const testIdRegex = /\btestId\s*:\s*(["'])(.*?)\1/g;
  // Heuristic: only treat `data-testid="..."` occurrences preceded by whitespace as "definitions".
  // This avoids false positives from selector strings like `[data-testid="foo"]`.
  const dataTestIdRegex = /(?<=\s)data-testid\s*=\s*(["'])([^"']+)\1(?=\s|>|\/)/g;
  /** @type {string[]} */
  const ids = [];
  for (const match of source.matchAll(testIdRegex)) {
    ids.push(match[2]);
  }
  for (const match of source.matchAll(dataTestIdRegex)) {
    ids.push(match[2]);
  }
  return ids;
}

function collectRibbonTestIds() {
  const ribbonDir = path.join(__dirname, "..", "src", "ribbon");
  /** @type {string[]} */
  const ids = [];

  const walk = (dir) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      // Skip tests; they often embed `data-testid="..."` strings for snapshots.
      if (entry.isDirectory() && entry.name === "__tests__") continue;

      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(fullPath);
        continue;
      }
      if (!entry.isFile()) continue;
      if (!/\.(ts|tsx)$/.test(entry.name)) continue;

      const source = fs.readFileSync(fullPath, "utf8");
      ids.push(...extractRibbonTestIdsFromSource(source));
    }
  };

  walk(ribbonDir);
  return ids;
}

test("desktop index.html does not hardcode ribbon action testids (avoid Playwright strict locator collisions)", () => {
  const htmlPath = path.join(__dirname, "..", "index.html");
  const html = fs.readFileSync(htmlPath, "utf8");

  const indexDuplicates = findDuplicateTestIdsInIndexHtml(html);
  assert.deepEqual(
    indexDuplicates,
    [],
    `apps/desktop/index.html contains duplicate data-testid values (Playwright strict locators would fail):\\n${indexDuplicates
      .map((id) => `- ${id}`)
      .join("\\n")}`,
  );

  const indexTestIds = extractTestIdsFromIndexHtml(html);
  const ribbonTestIds = collectRibbonTestIds();

  // Ensure Ribbon itself does not ship duplicate test IDs (Playwright strict mode would
  // fail even without any static HTML collisions).
  const ribbonTestIdCounts = new Map();
  for (const testId of ribbonTestIds) {
    ribbonTestIdCounts.set(testId, (ribbonTestIdCounts.get(testId) ?? 0) + 1);
  }
  const ribbonTestIdDuplicates = [...ribbonTestIdCounts.entries()]
    .filter(([, count]) => count > 1)
    .map(([testId, count]) => `${testId} (${count})`);

  assert.deepEqual(
    ribbonTestIdDuplicates,
    [],
    `Ribbon contains duplicate test id values (should be unique):\\n${ribbonTestIdDuplicates
      .map((id) => `- ${id}`)
      .join("\\n")}`,
  );

  // Ensure the static shell does not hardcode any of the Ribbon action hooks. Those
  // should render exactly once (in the ribbon), otherwise Playwright strict locators
  // like page.getByTestId(...).click() throw.
  const overlap = [...new Set(ribbonTestIds)].filter((testId) => indexTestIds.has(testId));

  assert.deepEqual(
    overlap,
    [],
    `apps/desktop/index.html includes ribbon action testids that must be owned by the ribbon (avoid collisions):\\n${overlap
      .map((id) => `- data-testid="${id}"`)
      .join("\\n")}`,
  );
});
