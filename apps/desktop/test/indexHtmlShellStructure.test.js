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
    "dock-left",
    "dock-right",
    "dock-bottom",
    "floating-root",
    "grid-split",
    "grid-secondary",

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

  // Playwright's `getByTestId` is strict by default and will fail if the same test id is present
  // multiple times in the DOM. Catch accidental duplication early (in the static HTML layer).
  const duplicateTestIds = requiredTestIds
    .map((testId) => {
      const matches = html.match(new RegExp(`data-testid=["']${escapeRegExp(testId)}["']`, "g")) ?? [];
      return { testId, count: matches.length };
    })
    .filter(({ count }) => count > 1)
    .map(({ testId, count }) => `${testId} (${count}x)`);

  assert.deepEqual(
    duplicateTestIds,
    [],
    `apps/desktop/index.html contains duplicate data-testid values (breaks Playwright strict-mode):\\n${duplicateTestIds
      .map((m) => `- ${m}`)
      .join("\\n")}`,
  );

  // Also enforce uniqueness across *all* static test ids, not just the required shell ones.
  // This keeps `page.getByTestId(...)` deterministic even as new shell markup gets added.
  const htmlWithoutScripts = html.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, "");
  const allTestIdMatches = [...htmlWithoutScripts.matchAll(/data-testid=["']([^"']+)["']/g)].map((m) => m[1]);
  const allTestIdCounts = new Map();
  for (const testId of allTestIdMatches) {
    allTestIdCounts.set(testId, (allTestIdCounts.get(testId) ?? 0) + 1);
  }
  const allDuplicateTestIds = [...allTestIdCounts.entries()]
    .filter(([, count]) => count > 1)
    .map(([testId, count]) => `${testId} (${count}x)`)
    .sort();
  assert.deepEqual(
    allDuplicateTestIds,
    [],
    `apps/desktop/index.html contains duplicate data-testid values:\\n${allDuplicateTestIds.map((m) => `- ${m}`).join("\\n")}`,
  );

  // Similar guard for HTML ids: the app boot code relies on `getElementById` for key shell
  // mounts (grid, ribbon, docks). Duplicate ids can cause subtle startup bugs.
  const allIdMatches = [...htmlWithoutScripts.matchAll(/\bid=["']([^"']+)["']/g)].map((m) => m[1]);
  const idCounts = new Map();
  for (const id of allIdMatches) idCounts.set(id, (idCounts.get(id) ?? 0) + 1);
  const duplicateIds = [...idCounts.entries()]
    .filter(([, count]) => count > 1)
    .map(([id, count]) => `${id} (${count}x)`)
    .sort();
  assert.deepEqual(
    duplicateIds,
    [],
    `apps/desktop/index.html contains duplicate id attributes:\\n${duplicateIds.map((m) => `- ${m}`).join("\\n")}`,
  );

  // The collaboration indicator is part of the visible status bar (it should not be
  // hidden inside `.statusbar__debug`, which is display:none in production styles).
  const collabStatusIndex = html.indexOf('data-testid="collab-status"');
  const debugIndex = html.indexOf('class="statusbar__debug"');
  assert.ok(collabStatusIndex >= 0, "Expected data-testid=\"collab-status\" to exist in index.html");
  assert.ok(debugIndex >= 0, "Expected .statusbar__debug section to exist in index.html");
  assert.ok(
    collabStatusIndex < debugIndex,
    "Expected collab-status indicator to appear in the visible statusbar section (before .statusbar__debug)",
  );

  // A11y: collab status updates should be announced politely by screen readers.
  assert.match(
    html,
    /data-testid="collab-status"[^>]*\brole="status"/,
    'Expected collab-status element to include role="status" for accessibility',
  );
  assert.match(
    html,
    /data-testid="collab-status"[^>]*\baria-label="Collaboration status"/,
    'Expected collab-status element to include aria-label="Collaboration status" for accessibility',
  );
  assert.match(
    html,
    /data-testid="collab-status"[^>]*\baria-live="polite"/,
    'Expected collab-status element to include aria-live="polite" for accessibility',
  );
  assert.match(
    html,
    /data-testid="collab-status"[^>]*\baria-atomic="true"/,
    'Expected collab-status element to include aria-atomic="true" for accessibility',
  );

  // A11y: sheet position updates should be announced politely by screen readers.
  assert.match(
    html,
    /data-testid="sheet-position"[^>]*\brole="status"/,
    'Expected sheet-position element to include role="status" for accessibility',
  );
  assert.match(
    html,
    /data-testid="sheet-position"[^>]*\baria-label="Sheet position"/,
    'Expected sheet-position element to include aria-label="Sheet position" for accessibility',
  );
  assert.match(
    html,
    /data-testid="sheet-position"[^>]*\baria-live="polite"/,
    'Expected sheet-position element to include aria-live="polite" for accessibility',
  );
  assert.match(
    html,
    /data-testid="sheet-position"[^>]*\baria-atomic="true"/,
    'Expected sheet-position element to include aria-atomic="true" for accessibility',
  );

  // Sheet position: expose machine-readable numbers for tests/tooling so assertions don't
  // need to depend on the translated text content.
  assert.match(
    html,
    /data-testid="sheet-position"[^>]*\bdata-sheet-position=["'][^"']+["']/,
    'Expected sheet-position element to include data-sheet-position',
  );
  assert.match(
    html,
    /data-testid="sheet-position"[^>]*\bdata-sheet-total=["'][^"']+["']/,
    'Expected sheet-position element to include data-sheet-total',
  );

  // Debug controls should live in the ribbon (React) rather than being duplicated in the
  // static `index.html` status bar. Duplicating them here causes Playwright strict-mode
  // failures because `getByTestId(...)` matches multiple elements with the same test id.
  const forbiddenSnippets = [
    'data-testid="open-ai-panel"',
    'data-testid="open-ai-audit-panel"',
    'data-testid="open-panel-ai-chat"',
    'data-testid="open-panel-ai-audit"',
    'data-testid="open-data-queries-panel"',
    'data-testid="open-macros-panel"',
    'data-testid="open-script-editor-panel"',
    'data-testid="open-python-panel"',
    'data-testid="open-extensions-panel"',
    'data-testid="open-vba-migrate-panel"',
    'data-testid="open-comments-panel"',
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

    // Other ribbon-mounted controls referenced directly by Playwright tests.
    // Keep them out of `index.html` so `page.getByTestId(...)` remains unambiguous.
    'data-testid="open-inline-ai-edit"',
    'data-testid="open-marketplace-panel"',
    'data-testid="theme-selector"',

    // Ribbon submenu items / backstage actions are rendered by React. If these appear in the
    // static HTML they can also cause strict-mode collisions.
    'data-testid="ribbon-find"',
    'data-testid="ribbon-replace"',
    'data-testid="ribbon-goto"',
    'data-testid="ribbon-show-formulas"',
    'data-testid="ribbon-perf-stats"',
    'data-testid="file-new"',
    'data-testid="file-open"',
    'data-testid="file-quit"',
  ];

  const forbiddenPresent = forbiddenSnippets.filter((snippet) => html.includes(snippet));
  assert.deepEqual(
    forbiddenPresent,
    [],
    `apps/desktop/index.html should not include legacy debug buttons (they belong in the ribbon):\\n${forbiddenPresent
      .map((m) => `- ${m}`)
      .join("\\n")}`,
  );

  // Collaboration panels should be discoverable via always-visible status bar buttons.
  const versionHistoryIndex = html.indexOf('data-testid="open-version-history-panel"');
  const branchManagerIndex = html.indexOf('data-testid="open-branch-manager-panel"');
  assert.ok(
    versionHistoryIndex >= 0,
    "Expected data-testid=\"open-version-history-panel\" to exist in index.html",
  );
  assert.ok(
    branchManagerIndex >= 0,
    "Expected data-testid=\"open-branch-manager-panel\" to exist in index.html",
  );
  assert.ok(
    versionHistoryIndex < debugIndex,
    "Expected open-version-history-panel button to appear in the visible statusbar section (before .statusbar__debug)",
  );
  assert.ok(
    branchManagerIndex < debugIndex,
    "Expected open-branch-manager-panel button to appear in the visible statusbar section (before .statusbar__debug)",
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
