import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { readRibbonSchemaSource } from "./ribbonSchemaSource.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function countMatches(source, pattern) {
  const re = pattern instanceof RegExp ? pattern : new RegExp(String(pattern), "g");
  const matches = source.match(re);
  return matches ? matches.length : 0;
}

test("Ribbon schema aligns Home → Editing AutoSum/Fill ids with CommandRegistry ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  // Canonical command ids.
  const requiredIds = ["edit.autoSum", "edit.fillDown", "edit.fillRight"];
  for (const id of requiredIds) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }

  // AutoSum should be used for both the dropdown button id and the default "Sum" menu item.
  assert.ok(
    countMatches(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp("edit.autoSum")}["']`, "g")) >= 2,
    "Expected edit.autoSum to appear at least twice (button + menu item)",
  );

  // Legacy ids should not be present.
  const legacyIds = ["home.editing.autoSum", "home.editing.autoSum.sum", "home.editing.fill.down", "home.editing.fill.right"];
  for (const id of legacyIds) {
    assert.doesNotMatch(
      schema,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`),
      `Expected homeTab.ts to not include legacy id ${id}`,
    );
  }
});

test("Desktop main.ts routes canonical Editing ribbon commands through the CommandRegistry (no legacy mapping)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = fs.readFileSync(builtinsPath, "utf8");

  // Canonical editing ids should be registered as builtin commands so ribbon, command palette,
  // and keybindings share the same execution path (via createRibbonActionsFromCommands).
  const expects = ["edit.autoSum", "edit.fillDown", "edit.fillRight"];
  for (const id of expects) {
    assert.match(
      builtins,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerBuiltinCommands.ts to register ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
  }

  // Ensure the old ribbon-only ids are no longer mapped in main.ts.
  const legacyCases = ["home.editing.autoSum", "home.editing.autoSum.sum", "home.editing.fill.down", "home.editing.fill.right"];
  for (const id of legacyCases) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts not to contain legacy case ${id}`,
    );
  }

  // Sanity check: the ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActionsFromCommands\(/);
});
