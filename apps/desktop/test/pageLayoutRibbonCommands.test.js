import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema includes Page Layout command ids", () => {
  const schemaPath = path.join(__dirname, "..", "src", "ribbon", "ribbonSchema.ts");
  const schema = fs.readFileSync(schemaPath, "utf8");

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
  const source = fs.readFileSync(commandsPath, "utf8");

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
});

test("Desktop main.ts does not special-case Page Layout ribbon actions in the ribbon switch", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

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
});
