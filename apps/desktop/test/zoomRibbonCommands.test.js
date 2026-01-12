import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema includes View → Zoom command ids", () => {
  const schemaPath = path.join(__dirname, "..", "src", "ribbon", "ribbonSchema.ts");
  const schema = fs.readFileSync(schemaPath, "utf8");

  const ids = [
    // View → Zoom controls.
    "view.zoom.zoom",
    "view.zoom.zoom100",
    "view.zoom.zoomToSelection",

    // Zoom dropdown menu items.
    "view.zoom.zoom.400",
    "view.zoom.zoom.200",
    "view.zoom.zoom.150",
    "view.zoom.zoom.100",
    "view.zoom.zoom.75",
    "view.zoom.zoom.50",
    "view.zoom.zoom.25",
    "view.zoom.zoom.custom",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected ribbon schema to include ${id}`);
  }

  // Ensure the primary zoom control is a dropdown.
  assert.match(schema, /\bid:\s*["']view\.zoom\.zoom["'][\s\S]*?\bkind:\s*["']dropdown["']/);
});

test("Desktop main.ts wires View → Zoom ribbon commands to SpreadsheetApp zoom", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // Fixed action buttons should be explicit `case` handlers.
  const buttonCases = [
    { id: "view.zoom.zoom100", pattern: /\bapp\.setZoom\(\s*1\s*\)/ },
    { id: "view.zoom.zoomToSelection", pattern: /\bapp\.zoomToSelection\(\)/ },
    { id: "view.zoom.zoom", pattern: /\bopenCustomZoomQuickPick\b/ },
  ];

  for (const { id, pattern } of buttonCases) {
    assert.match(main, new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`), `Expected main.ts to handle ${id}`);
    assert.match(main, pattern, `Expected main.ts to invoke ${String(pattern)} for ${id}`);
  }

  // The dropdown menu items (e.g. view.zoom.zoom.200) are handled via prefix parsing.
  assert.match(main, /\bconst\s+zoomMenuItemPrefix\s*=\s*["']view\.zoom\.zoom\.\s*["'];/);
  assert.match(main, /\bcommandId\.startsWith\(zoomMenuItemPrefix\)/);
  assert.match(main, /\bsuffix\s*===\s*["']custom["']/);
  assert.match(main, /\bconst\s+percent\s*=\s*Number\(suffix\)/);
  assert.match(main, /\bapp\.setZoom\(\s*percent\s*\/\s*100\s*\)/);
  assert.match(main, /\bsyncZoomControl\(\)/);
  assert.match(main, /\bapp\.focus\(\)/);
});

