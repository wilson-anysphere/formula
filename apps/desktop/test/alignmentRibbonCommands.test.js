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

test("Ribbon schema uses canonical Home → Alignment command ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  const requiredIds = [
    // Horizontal alignment
    "format.alignLeft",
    "format.alignCenter",
    "format.alignRight",

    // Vertical alignment
    "format.alignTop",
    "format.alignMiddle",
    "format.alignBottom",

    // Indent
    "format.increaseIndent",
    "format.decreaseIndent",

    // Orientation dropdown menu items
    "format.textRotation.angleCounterclockwise",
    "format.textRotation.angleClockwise",
    "format.textRotation.verticalText",
    "format.textRotation.rotateUp",
    "format.textRotation.rotateDown",
    "format.openAlignmentDialog",
  ];

  for (const id of requiredIds) {
    assert.match(
      schema,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`),
      `Expected homeTab.ts to include ${id}`,
    );
  }

  // Orientation dropdown trigger remains a ribbon-specific id (it's a menu container).
  assert.match(schema, /\bid:\s*["']home\.alignment\.orientation["']/);
  assert.match(schema, /\bkind:\s*["']dropdown["']/);

  // Legacy alignment ids (previously wired directly in main.ts) should not exist in the schema.
  const legacyIds = [
    "home.alignment.alignLeft",
    "home.alignment.center",
    "home.alignment.alignRight",
    "home.alignment.topAlign",
    "home.alignment.middleAlign",
    "home.alignment.bottomAlign",
    "home.alignment.increaseIndent",
    "home.alignment.decreaseIndent",
    "home.alignment.orientation.angleCounterclockwise",
    "home.alignment.orientation.angleClockwise",
    "home.alignment.orientation.verticalText",
    "home.alignment.orientation.rotateUp",
    "home.alignment.orientation.rotateDown",
    "home.alignment.orientation.formatCellAlignment",
  ];
  for (const id of legacyIds) {
    assert.doesNotMatch(
      schema,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`),
      `Expected homeTab.ts to not include legacy id ${id}`,
    );
  }
});

test("Alignment ribbon commands are registered in CommandRegistry and not handled via main.ts switch cases", () => {
  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerFormatAlignmentCommands.ts");
  const commands = fs.readFileSync(commandsPath, "utf8");

  const commandIds = [
    "format.alignLeft",
    "format.alignCenter",
    "format.alignRight",
    "format.alignTop",
    "format.alignMiddle",
    "format.alignBottom",
    "format.increaseIndent",
    "format.decreaseIndent",
    "format.textRotation.angleCounterclockwise",
    "format.textRotation.angleClockwise",
    "format.textRotation.verticalText",
    "format.textRotation.rotateUp",
    "format.textRotation.rotateDown",
    "format.openAlignmentDialog",
  ];

  for (const id of commandIds) {
    assert.match(
      commands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerFormatAlignmentCommands.ts to register ${id}`,
    );
  }

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // These ids should be dispatched by `createRibbonActionsFromCommands` via CommandRegistry;
  // main.ts may reference them for pressed/disabled state, but should not handle them in its
  // ribbon fallback switch.
  for (const id of commandIds) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be executed via CommandRegistry)`,
    );
  }
});

