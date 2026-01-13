import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function collectStringPropertyValues(source, propertyName) {
  const values = [];
  // Match `propertyName: "value"` with permissive whitespace/newlines.
  const re = new RegExp(`${propertyName}\\\\s*:\\\\s*\"([^\"]+)\"`, "g");
  let match;
  while ((match = re.exec(source))) {
    values.push(match[1]);
  }
  return values;
}

function findDuplicates(values) {
  const counts = new Map();
  for (const value of values) counts.set(value, (counts.get(value) ?? 0) + 1);
  return [...counts.entries()]
    .filter(([, count]) => count > 1)
    .map(([value, count]) => ({ value, count }))
    .sort((a, b) => b.count - a.count || a.value.localeCompare(b.value));
}

test("ribbon schema and File backstage expose stable, unique test ids", () => {
  const ribbonSchemaDir = path.join(__dirname, "..", "src", "ribbon", "schema");
  const ribbonSchemaFiles = fs
    .readdirSync(ribbonSchemaDir)
    .filter((entry) => entry.endsWith(".ts"))
    .sort((a, b) => a.localeCompare(b));
  const ribbonSchema = ribbonSchemaFiles
    .map((file) => fs.readFileSync(path.join(ribbonSchemaDir, file), "utf8"))
    .join("\n");

  const ribbonTestIds = collectStringPropertyValues(ribbonSchema, "testId");
  const ribbonDuplicates = findDuplicates(ribbonTestIds);
  assert.deepEqual(
    ribbonDuplicates,
    [],
    `apps/desktop/src/ribbon/schema/*.ts contains duplicate testId values (breaks Playwright strict-mode):\n${ribbonDuplicates
      .map(({ value, count }) => `- ${value} (${count}x)`)
      .join("\n")}`,
  );

  const requiredRibbonTestIds = [
    // Home tab: core e2e hooks.
    "open-panel-ai-chat",
    "open-panel-ai-audit",
    "open-ai-audit-panel",
    "open-inline-ai-edit",
    "open-extensions-panel",
    "open-comments-panel",
    "open-macros-panel",
    "open-python-panel",
    "open-script-editor-panel",
    "audit-precedents",
    "audit-dependents",
    "audit-transitive",
    "split-vertical",
    "split-horizontal",
    "split-none",
    "freeze-panes",
    "freeze-top-row",
    "freeze-first-column",
    "unfreeze-panes",
    "ribbon-find",
    "ribbon-replace",
    "ribbon-goto",
    // View tab: e2e hooks.
    "open-marketplace-panel",
    "open-version-history-panel",
    "open-branch-manager-panel",
    "theme-selector",
    "ribbon-show-formulas",
    "ribbon-perf-stats",
  ];

  const missingRibbonTestIds = requiredRibbonTestIds.filter((id) => !ribbonSchema.includes(`testId: "${id}"`));
  assert.deepEqual(
    missingRibbonTestIds,
    [],
    `apps/desktop/src/ribbon/schema/*.ts is missing required test ids:\n${missingRibbonTestIds.map((id) => `- ${id}`).join("\n")}`,
  );

  const fileBackstagePath = path.join(__dirname, "..", "src", "ribbon", "FileBackstage.tsx");
  const fileBackstage = fs.readFileSync(fileBackstagePath, "utf8");
  const backstageTestIds = collectStringPropertyValues(fileBackstage, "testId");
  const backstageDuplicates = findDuplicates(backstageTestIds);
  assert.deepEqual(
    backstageDuplicates,
    [],
    `apps/desktop/src/ribbon/FileBackstage.tsx contains duplicate testId values (breaks Playwright strict-mode):\n${backstageDuplicates
      .map(({ value, count }) => `- ${value} (${count}x)`)
      .join("\n")}`,
  );

  const requiredBackstageTestIds = ["file-new", "file-open", "file-quit"];
  const missingBackstageTestIds = requiredBackstageTestIds.filter((id) => !fileBackstage.includes(`testId: "${id}"`));
  assert.deepEqual(
    missingBackstageTestIds,
    [],
    `apps/desktop/src/ribbon/FileBackstage.tsx is missing required test ids:\n${missingBackstageTestIds
      .map((id) => `- ${id}`)
      .join("\n")}`,
  );
});
