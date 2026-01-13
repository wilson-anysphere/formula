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

test("Ribbon schema includes View → Zoom command ids", () => {
  const schema = readRibbonSchemaSource("viewTab.ts");

  const ids = [
    // View → Zoom controls.
    "view.zoom.zoom",
    "view.zoom.zoom100",
    "view.zoom.zoomToSelection",

    // Zoom dropdown menu items.
    "view.zoom.zoom400",
    "view.zoom.zoom200",
    "view.zoom.zoom150",
    "view.zoom.zoom100",
    "view.zoom.zoom75",
    "view.zoom.zoom50",
    "view.zoom.zoom25",
    "view.zoom.openPicker",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected ribbon schema to include ${id}`);
  }

  // Ensure the primary zoom control is a dropdown.
  assert.match(schema, /\bid:\s*["']view\.zoom\.zoom["'][\s\S]*?\bkind:\s*["']dropdown["']/);
});

test("Desktop main.ts delegates View → Zoom ribbon commands to CommandRegistry", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // Zoom commands should not be hardcoded through the ribbon's `onCommand` switch.
  // They are either:
  // - registered as builtin commands (executed through CommandRegistry), or
  // - routed via a prefix-based handler for dynamic zoom menu items.
  assert.doesNotMatch(main, /\bcase\s+["']view\.zoom\.zoom100["']:/);
  assert.doesNotMatch(main, /\bcase\s+["']view\.zoom\.zoomToSelection["']:/);

  // Ensure ribbon onCommand delegates registered ids through CommandRegistry.
  assert.match(main, /\bcommandRegistry\.getCommand\(commandId\)/);
  assert.match(main, /\bexecuteBuiltinCommand\(\s*commandId/);
});
