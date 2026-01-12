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

    // Debug/utility buttons (kept for now; used by some e2e flows)
    'data-testid="audit-precedents"',
    'data-testid="audit-dependents"',
    'data-testid="audit-transitive"',
    'data-testid="split-vertical"',
    'data-testid="split-horizontal"',
    'data-testid="split-none"',
    'data-testid="freeze-panes"',
    'data-testid="freeze-top-row"',
    'data-testid="freeze-first-column"',
    'data-testid="unfreeze-panes"',
  ];

  const missing = requiredSnippets.filter((snippet) => !html.includes(snippet));
  assert.deepEqual(
    missing,
    [],
    `apps/desktop/index.html is missing required shell markup:\\n${missing.map((m) => `- ${m}`).join("\\n")}`,
  );
});

