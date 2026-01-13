import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("ContextMenu is styled via CSS classes (no inline style.* except positioning)", () => {
  const filePath = path.join(__dirname, "..", "src", "menus", "contextMenu.ts");
  const content = fs.readFileSync(filePath, "utf8");

  const matches = [...content.matchAll(/\.style\.([a-zA-Z]+)/g)];
  const disallowed = matches.filter((m) => !["left", "top"].includes(m[1]));

  assert.deepEqual(
    disallowed.map((m) => m[0]),
    [],
    "ContextMenu should not set inline styles (move static styling into src/styles/context-menu.css)",
  );

  // Sanity-check that the overlay/menu use CSS classes.
  for (const cls of ["context-menu-overlay", "context-menu__item", "context-menu__separator"]) {
    assert.ok(content.includes(cls), `Expected ContextMenu implementation to reference .${cls} for styling`);
  }

  // Guardrails for context-menu.css: use radius tokens (not hardcoded large radii)
  // and include high-contrast/forced-colors affordances so the menu remains usable
  // when shadows/subtle hover colors are neutralized to system colors.
  const cssPath = path.join(__dirname, "..", "src", "styles", "context-menu.css");
  assert.equal(fs.existsSync(cssPath), true, "Expected apps/desktop/src/styles/context-menu.css to exist");
  const css = fs.readFileSync(cssPath, "utf8");

  assert.match(css, /border-radius:\s*var\(--radius\)/, "Context menu container should use --radius token");
  assert.match(css, /border-radius:\s*var\(--radius-sm\)/, "Context menu items should use --radius-sm token");
  assert.doesNotMatch(css, /border-radius:\s*10px\b/, "Avoid hardcoded 10px radii in context-menu.css");
  assert.doesNotMatch(css, /border-radius:\s*8px\b/, "Avoid hardcoded 8px radii in context-menu.css");

  assert.ok(
    css.includes("@media (forced-colors: active)") || css.includes("data-theme=\"high-contrast\""),
    "context-menu.css should include forced-colors/high-contrast tweaks",
  );
});
