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
});

