import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("desktop index.html exposes required shell containers and testids", () => {
  const htmlPath = path.join(__dirname, "..", "index.html");
  const html = fs.readFileSync(htmlPath, "utf8");

  const requiredSnippets = [
    // Shell roots
    'id="app"',
    'id="titlebar-root"',
    'id="ribbon"',
    'id="formula-bar"',
    'id="workspace"',
    'id="grid-split"',
    'id="grid"',
    'id="grid-secondary"',
    'id="dock-left"',
    'id="dock-right"',
    'id="dock-bottom"',
    'id="floating-root"',
    'id="sheet-tabs"',
    'data-testid="sheet-tabs"',
    'id="toast-root"',
    'data-testid="toast-root"',

    // Status bar (e2e relies on these)
    'data-testid="active-cell"',
    'data-testid="selection-range"',
    'data-testid="active-value"',
    'data-testid="sheet-switcher"',
    'data-testid="zoom-control"',
    'data-testid="sheet-position"',
  ];

  const missing = requiredSnippets.filter((snippet) => !html.includes(snippet));
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
    "Expected #grid to have tabindex=\"0\" so it can receive focus",
  );
  assert.match(
    html,
    /\bid="grid-secondary"[^>]*\btabindex="0"/,
    "Expected #grid-secondary to have tabindex=\"0\" so it can receive focus",
  );
});
