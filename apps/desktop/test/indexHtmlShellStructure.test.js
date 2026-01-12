import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("desktop index.html exposes required shell containers and testids", () => {
  const htmlPath = path.join(__dirname, "..", "index.html");
  const html = fs.readFileSync(htmlPath, "utf8");

  const requiredIds = [
    // Shell roots
    "app",
    "titlebar-root",
    "ribbon",
    "formula-bar",
    "workspace",
    "grid-split",
    "grid",
    "grid-secondary",
    "dock-left",
    "dock-right",
    "dock-bottom",
    "floating-root",
    "sheet-tabs",
    "toast-root",
  ];

  const requiredTestIds = [
    "titlebar",
    "ribbon",
    "sheet-tabs",
    "toast-root",

    // Status bar (e2e relies on these)
    "status-mode",
    "active-cell",
    "selection-range",
    "active-value",
    "collab-status",
    "selection-sum",
    "selection-avg",
    "selection-count",
    "sheet-switcher",
    "zoom-control",
    "status-zoom",
    "sheet-position",

    // Debug/test controls (kept in static HTML for Playwright)
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

  const missingIds = requiredIds
    .filter((id) => !new RegExp(`\\bid=["']${escapeRegExp(id)}["']`).test(html))
    .map((id) => `id="${id}"`);
  const missingTestIds = requiredTestIds
    .filter((testId) => !html.includes(`data-testid="${testId}"`))
    .map((testId) => `data-testid="${testId}"`);
  const missing = [...missingIds, ...missingTestIds];

  assert.deepEqual(
    missing,
    [],
    `apps/desktop/index.html is missing required shell markup:\\n${missing.map((m) => `- ${m}`).join("\\n")}`,
  );

  // Ensure the base app shell styling hooks stay intact.
  assert.match(
    html,
    /\bid="app"[^>]*\bclass="[^"]*\bformula-app\b[^"]*"/,
    'Expected #app to include class="formula-app" so base UI styles apply',
  );

  // Sheet tabs are styled via ui.css under `#sheet-tabs.sheet-bar`.
  assert.match(
    html,
    /\bid="sheet-tabs"[^>]*\bclass="[^"]*\bsheet-bar\b[^"]*"/,
    'Expected #sheet-tabs to include class="sheet-bar" for sheet bar styling',
  );

  // The grid must remain focusable for keyboard navigation.
  assert.match(
    html,
    /\bid="grid"[^>]*\btabindex="0"/,
    'Expected #grid to have tabindex="0" so it can receive focus',
  );
  assert.match(
    html,
    /\bid="grid-secondary"[^>]*\btabindex="0"/,
    'Expected #grid-secondary to have tabindex="0" so it can receive focus',
  );
});
