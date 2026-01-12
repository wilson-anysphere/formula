import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("desktop index.html does not hardcode legacy statusbar debug action testids (ribbon owns them)", () => {
  const htmlPath = path.join(__dirname, "..", "index.html");
  const html = fs.readFileSync(htmlPath, "utf8");

  const forbiddenActionTestIds = [
    // Auditing
    "audit-precedents",
    "audit-dependents",
    "audit-transitive",

    // Split view
    "split-vertical",
    "split-horizontal",
    "split-none",

    // Freeze panes
    "freeze-panes",
    "freeze-top-row",
    "freeze-first-column",
    "unfreeze-panes",

    // Panels
    "open-panel-ai-chat",
    "open-panel-ai-audit",
    "open-data-queries-panel",
    "open-macros-panel",
    "open-script-editor-panel",
    "open-python-panel",
    "open-extensions-panel",
    "open-vba-migrate-panel",
    "open-comments-panel",
  ];

  const dataTestIdRegex = /\bdata-testid\s*=\s*(["'])(.*?)\1/g;
  /** @type {Set<string>} */
  const presentTestIds = new Set();
  for (const match of html.matchAll(dataTestIdRegex)) {
    presentTestIds.add(match[2]);
  }

  const found = forbiddenActionTestIds.filter((testId) => presentTestIds.has(testId));

  assert.deepEqual(
    found,
    [],
    `apps/desktop/index.html includes legacy action hooks that collide with ribbon data-testids:\\n${found
      .map((id) => `- data-testid="${id}"`)
      .join("\\n")}`,
  );
});
