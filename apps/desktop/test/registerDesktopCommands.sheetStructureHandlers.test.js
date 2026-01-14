import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("main.ts wires sheetStructureHandlers + autoFilterHandlers into registerDesktopCommands", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));

  const start = main.indexOf("registerDesktopCommands({");
  assert.notEqual(start, -1, "Expected main.ts to call registerDesktopCommands({ ... })");

  const end = main.indexOf("workbenchFileHandlers", start);
  assert.notEqual(end, -1, "Expected main.ts to pass workbenchFileHandlers into registerDesktopCommands");

  const segment = main.slice(start, end);

  assert.match(
    segment,
    /\bsheetStructureHandlers\s*:\s*{/,
    "Expected main.ts to pass sheetStructureHandlers into registerDesktopCommands",
  );
  assert.match(
    segment,
    /\binsertSheet\s*:\s*handleAddSheet\b/,
    "Expected sheetStructureHandlers.insertSheet to be wired to handleAddSheet",
  );
  assert.match(
    segment,
    /\bdeleteActiveSheet\s*:\s*handleDeleteActiveSheet\b/,
    "Expected sheetStructureHandlers.deleteActiveSheet to be wired to handleDeleteActiveSheet",
  );
  assert.match(
    segment,
    /\bopenOrganizeSheets\b/,
    "Expected sheetStructureHandlers.openOrganizeSheets to be provided by main.ts",
  );

  assert.match(
    segment,
    /\bautoFilterHandlers\s*:\s*ribbonAutoFilterHandlers\b/,
    "Expected main.ts to pass autoFilterHandlers into registerDesktopCommands",
  );

  assert.match(main, /\bconst\s+ribbonAutoFilterHandlers\s*=\s*{/, "Expected main.ts to define ribbonAutoFilterHandlers");
  assert.match(
    main,
    /\bconst\s+ribbonAutoFilterHandlers\s*=\s*{[\s\S]*?\btoggle\s*:\s*(?:async\s*)?(?:\([^)]*\)|[a-zA-Z_$][\w$]*)\s*=>\s*\{/m,
    "Expected ribbonAutoFilterHandlers.toggle to be wired in main.ts",
  );
  assert.match(
    main,
    /\bconst\s+ribbonAutoFilterHandlers\s*=\s*{[\s\S]*?\btoggle\s*:\s*(?:async\s*)?(?:\([^)]*\)|[a-zA-Z_$][\w$]*)\s*=>\s*\{[\s\S]*?\bclearRibbonAutoFiltersForActiveSheet\(\)/m,
    "Expected ribbonAutoFilterHandlers.toggle to clear ribbon AutoFilters when disabling",
  );
  assert.match(
    main,
    /\bribbonAutoFilterStore\.hasAny\b/,
    "Expected ribbonAutoFilterHandlers.toggle to consult ribbonAutoFilterStore.hasAny",
  );
  assert.match(
    main,
    /\bapplyRibbonAutoFilterFromSelection\b/,
    "Expected ribbonAutoFilterHandlers.toggle to apply ribbon AutoFilter from selection",
  );
  assert.match(
    main,
    /\bclear\s*:\s*\(\)\s*=>\s*clearRibbonAutoFilterCriteriaForActiveSheet\b/,
    "Expected ribbonAutoFilterHandlers.clear to be wired to clearRibbonAutoFilterCriteriaForActiveSheet",
  );
  assert.match(
    main,
    /\breapply\s*:\s*\(\)\s*=>\s*reapplyRibbonAutoFiltersForActiveSheet\b/,
    "Expected ribbonAutoFilterHandlers.reapply to be wired to reapplyRibbonAutoFiltersForActiveSheet",
  );

  // These ribbon command ids are now registered in CommandRegistry; keep main.ts from
  // special-casing them in the `onUnknownCommand` switch.
  const ids = ["home.cells.insert.insertSheet", "home.cells.delete.deleteSheet", "home.cells.format.organizeSheets"];
  for (const id of ids) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be routed via the ribbon command router)`,
    );
  }
});
