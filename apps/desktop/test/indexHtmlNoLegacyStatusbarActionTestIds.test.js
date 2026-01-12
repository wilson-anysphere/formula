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

function extractTestIdsFromRibbonSchema(schema) {
  const testIdRegex = /\btestId\s*:\s*(["'])(.*?)\1/g;
  /** @type {string[]} */
  const ids = [];
  for (const match of schema.matchAll(testIdRegex)) {
    ids.push(match[2]);
  }
  return ids;
}

test("desktop index.html does not hardcode ribbon action testids (avoid Playwright strict locator collisions)", () => {
  const htmlPath = path.join(__dirname, "..", "index.html");
  const html = fs.readFileSync(htmlPath, "utf8");

  const schemaPath = path.join(__dirname, "..", "src", "ribbon", "ribbonSchema.ts");
  const schema = fs.readFileSync(schemaPath, "utf8");

  const indexTestIds = extractTestIdsFromIndexHtml(html);
  const ribbonTestIds = extractTestIdsFromRibbonSchema(schema);

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
    `Ribbon schema contains duplicate testId values (should be unique):\\n${ribbonTestIdDuplicates
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
