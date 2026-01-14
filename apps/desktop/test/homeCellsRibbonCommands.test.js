import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { readRibbonSchemaSource } from "./ribbonSchemaSource.js";
import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema includes Home → Cells command ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  const ids = [
    "home.cells.format.organizeSheets",
    "home.cells.insert.insertCells",
    "home.cells.insert.insertSheetRows",
    "home.cells.insert.insertSheetColumns",
    "home.cells.insert.insertSheet",
    "home.cells.delete.deleteCells",
    "home.cells.delete.deleteSheetRows",
    "home.cells.delete.deleteSheetColumns",
    "home.cells.delete.deleteSheet",
  ];
  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }
});

test("Home → Cells ribbon commands are registered in CommandRegistry and not handled via main.ts switch cases", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));
  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const commands = stripComments(fs.readFileSync(commandsPath, "utf8"));

  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = stripComments(fs.readFileSync(disablingPath, "utf8"));

  const sheetStructureIds = [
    "home.cells.format.organizeSheets",
    "home.cells.insert.insertSheet",
    "home.cells.delete.deleteSheet",
  ];
  for (const id of sheetStructureIds) {
    assert.match(
      commands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );

    // Sheet structure mutations should be guarded in read-only sessions (viewers/commenters) so
    // CommandRegistry surfaces (e.g. command palette) can't bypass ribbon disabling.
    const idx = commands.indexOf(`\"${id}\"`);
    assert.notEqual(idx, -1, `Expected registerDesktopCommands.ts to include ${id} literal`);
    assert.match(
      commands.slice(idx, idx + 600),
      /\bisReadOnly\(\)/,
      `Expected ${id} command handler to guard read-only mode`,
    );
  }

  const insertDeleteCellsIds = ["home.cells.insert.insertCells", "home.cells.delete.deleteCells"];
  for (const id of insertDeleteCellsIds) {
    assert.match(
      commands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );

    // Insert/Delete Cells are structural edits and should respect `registerDesktopCommands`'s
    // `isEditing` override (split-view secondary editor state).
    const idx = commands.indexOf(`\"${id}\"`);
    assert.notEqual(idx, -1, `Expected registerDesktopCommands.ts to include ${id} literal`);
    assert.match(
      commands.slice(idx, idx + 600),
      /\bisEditingFn\(\)/,
      `Expected ${id} command handler to guard edit mode via isEditingFn()`,
    );
  }

  const sheetRowColumnIds = [
    "home.cells.insert.insertSheetRows",
    "home.cells.insert.insertSheetColumns",
    "home.cells.delete.deleteSheetRows",
    "home.cells.delete.deleteSheetColumns",
  ];
  for (const id of sheetRowColumnIds) {
    assert.match(
      commands,
      new RegExp(`\\bregisterCellsStructuralCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id} via registerCellsStructuralCommand`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
  }

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActions\(/);
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
  assert.match(
    main,
    /\bsheetStructureHandlers\s*:/,
    "Expected main.ts to pass sheetStructureHandlers to registerDesktopCommands so sheet-structure ribbon ids are registered",
  );
});
