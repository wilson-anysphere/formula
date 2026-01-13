import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

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
  const schemaPath = path.join(__dirname, "..", "src", "ribbon", "schema", "homeTab.ts");
  const schema = fs.readFileSync(schemaPath, "utf8");

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
      `Expected ribbonSchema.ts to not include legacy id ${id}`,
    );
  }
});

test("Desktop main.ts handles canonical Editing ribbon commands directly (no legacy mapping)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const expects = ["edit.autoSum", "edit.fillDown", "edit.fillRight"];
  for (const id of expects) {
    assert.match(main, new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`), `Expected main.ts to handle ${id}`);
    assert.match(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:[\\s\\S]*?\\bexecuteBuiltinCommand\\(commandId\\);`),
      `Expected main.ts to execute builtin command for ${id}`,
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
});
