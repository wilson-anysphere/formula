import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema includes Page Layout command ids", () => {
  let schema = "";
  try {
    // The ribbon schema is modularized; Page Layout lives in its own schema module.
    const schemaPath = path.join(__dirname, "..", "src", "ribbon", "schema", "pageLayoutTab.ts");
    schema = stripComments(fs.readFileSync(schemaPath, "utf8"));
  } catch {
    // Back-compat: older versions kept all tab definitions in ribbonSchema.ts.
    const schemaPath = path.join(__dirname, "..", "src", "ribbon", "ribbonSchema.ts");
    schema = stripComments(fs.readFileSync(schemaPath, "utf8"));
  }

  const ids = [
    "pageLayout.pageSetup.pageSetupDialog",
    "pageLayout.pageSetup.margins.normal",
    "pageLayout.pageSetup.margins.wide",
    "pageLayout.pageSetup.margins.narrow",
    "pageLayout.pageSetup.margins.custom",
    "pageLayout.pageSetup.orientation.portrait",
    "pageLayout.pageSetup.orientation.landscape",
    "pageLayout.pageSetup.size.letter",
    "pageLayout.pageSetup.size.a4",
    "pageLayout.pageSetup.size.more",
    "pageLayout.printArea.setPrintArea",
    "pageLayout.printArea.clearPrintArea",
    "pageLayout.pageSetup.printArea.set",
    "pageLayout.pageSetup.printArea.clear",
    "pageLayout.pageSetup.printArea.addTo",
    "pageLayout.export.exportPdf",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected ribbon schema to include ${id}`);
  }
});

test("CommandRegistry registers Page Layout ribbon ids as builtin commands", () => {
  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerPageLayoutCommands.ts");
  const source = stripComments(fs.readFileSync(commandsPath, "utf8"));

  const ids = [
    "pageLayout.pageSetup.pageSetupDialog",
    "pageLayout.pageSetup.margins.normal",
    "pageLayout.pageSetup.margins.wide",
    "pageLayout.pageSetup.margins.narrow",
    "pageLayout.pageSetup.margins.custom",
    "pageLayout.pageSetup.orientation.portrait",
    "pageLayout.pageSetup.orientation.landscape",
    "pageLayout.pageSetup.size.letter",
    "pageLayout.pageSetup.size.a4",
    "pageLayout.pageSetup.size.more",
    "pageLayout.printArea.setPrintArea",
    "pageLayout.printArea.clearPrintArea",
    "pageLayout.pageSetup.printArea.set",
    "pageLayout.pageSetup.printArea.clear",
    "pageLayout.pageSetup.printArea.addTo",
    "pageLayout.export.exportPdf",
  ];

  for (const id of ids) {
    assert.match(source, new RegExp(`["']${escapeRegExp(id)}["']`), `Expected Page Layout command registration to include ${id}`);
  }

  // The Page Setup dropdown uses schema-scoped Print Area ids that are aliases of the primary
  // Print Area group commands. Keep them registered for ribbon coverage, but hide them from
  // the command palette to avoid duplicate entries.
  assert.match(source, /\bPAGE_LAYOUT_COMMANDS\.printArea\.set\b[\s\S]*?\bwhen:\s*["']false["']/);
  assert.match(source, /\bPAGE_LAYOUT_COMMANDS\.printArea\.clear\b[\s\S]*?\bwhen:\s*["']false["']/);
});

test("Desktop main.ts does not special-case Page Layout ribbon actions in the ribbon switch", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));

  // Page Layout actions should be routed through CommandRegistry (not handled as ad-hoc ribbon switch cases).
  assert.ok(!/\bcase\s+["']pageLayout\./.test(main), "Expected main.ts to avoid pageLayout.* case handlers in ribbon switch");

  // Ensure the desktop shell wires Page Layout ids to the existing helper functions via registerDesktopCommands.
  assert.match(main, /\bpageLayoutHandlers\s*:\s*{/, "Expected main.ts to pass pageLayoutHandlers into registerDesktopCommands");

  const mappings = [
    { key: "openPageSetupDialog", fn: "handleRibbonPageSetup" },
    { key: "updatePageSetup", fn: "handleRibbonUpdatePageSetup" },
    { key: "setPrintArea", fn: "handleRibbonSetPrintArea" },
    { key: "clearPrintArea", fn: "handleRibbonClearPrintArea" },
    { key: "addToPrintArea", fn: "handleRibbonAddToPrintArea" },
    { key: "exportPdf", fn: "handleRibbonExportPdf" },
  ];

  for (const { key, fn } of mappings) {
    assert.match(
      main,
      new RegExp(`\\b${escapeRegExp(key)}\\b[\\s\\S]{0,200}\\b${escapeRegExp(fn)}\\b`),
      `Expected main.ts to route pageLayoutHandlers.${key} through ${fn}`,
    );
  }

  // Page layout commands are disabled in edit-mode + read-only sessions in the ribbon. Ensure
  // the underlying handlers also guard those states so CommandRegistry surfaces (command palette,
  // keybindings) can't bypass the ribbon UI disabling.
  const guardedFns = [
    "handleRibbonPageSetup",
    "handleRibbonUpdatePageSetup",
    "handleRibbonSetPrintArea",
    "handleRibbonClearPrintArea",
    "handleRibbonAddToPrintArea",
  ];
  for (const fn of guardedFns) {
    const start = main.indexOf(`function ${fn}`);
    assert.notEqual(start, -1, `Expected main.ts to define ${fn}`);
    const slice = main.slice(start, start + 500);
    assert.match(slice, /\bisSpreadsheetEditing\b/, `Expected ${fn} to guard edit mode via isSpreadsheetEditing()`);
    assert.match(slice, /\bapp\.isReadOnly\b/, `Expected ${fn} to guard read-only mode via app.isReadOnly()`);
  }
});
