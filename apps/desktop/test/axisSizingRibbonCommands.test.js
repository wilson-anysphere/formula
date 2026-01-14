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

test("Ribbon schema includes Home → Cells → Format sizing command ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");
  const ids = ["home.cells.format.rowHeight", "home.cells.format.columnWidth"];
  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }
});

test("Axis sizing ribbon commands are registered in CommandRegistry and not handled via main.ts switch cases", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const axisCommandsPath = path.join(__dirname, "..", "src", "commands", "registerAxisSizingCommands.ts");
  const axisCommands = fs.readFileSync(axisCommandsPath, "utf8");

  const ids = ["home.cells.format.rowHeight", "home.cells.format.columnWidth"];
  for (const id of ids) {
    assert.match(
      axisCommands,
      new RegExp(`\\b${escapeRegExp(id)}\\b`),
      `Expected registerAxisSizingCommands.ts to reference command id ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+[\"']${escapeRegExp(id)}[\"']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
  }

  // Guardrail: ensure these ids are actually registered (not just defined).
  assert.match(axisCommands, /\bregisterBuiltinCommand\(\s*AXIS_SIZING_COMMAND_IDS\.rowHeight/);
  assert.match(axisCommands, /\bregisterBuiltinCommand\(\s*AXIS_SIZING_COMMAND_IDS\.columnWidth/);

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActionsFromCommands\(/);
});

