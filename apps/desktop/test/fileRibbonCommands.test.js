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

test("Ribbon schema includes File tab command ids", () => {
  const schema = readRibbonSchemaSource("fileTab.ts");

  const ids = [
    "file.new.new",
    "file.new.blankWorkbook",
    "file.open.open",
    "file.save.save",
    "file.save.saveAs",
    "file.save.saveAs.copy",
    "file.save.saveAs.download",
    "file.save.autoSave",
    "file.info.manageWorkbook.versions",
    "file.info.manageWorkbook.branches",
    "file.print.print",
    "file.print.printPreview",
    "file.print.pageSetup",
    "file.print.pageSetup.printTitles",
    "file.print.pageSetup.margins",
    "file.export.createPdf",
    "file.export.export.pdf",
    "file.export.export.csv",
    "file.export.export.xlsx",
    "file.export.changeFileType.pdf",
    "file.export.changeFileType.csv",
    "file.export.changeFileType.tsv",
    "file.export.changeFileType.xlsx",
    "file.options.close",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected fileTab.ts to include ${id}`);
  }
});

test("File tab ribbon ids are registered in CommandRegistry (no exemptions needed)", () => {
  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const commands = stripComments(fs.readFileSync(commandsPath, "utf8"));

  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = stripComments(fs.readFileSync(disablingPath, "utf8"));

  const ids = [
    "file.new.new",
    "file.new.blankWorkbook",
    "file.open.open",
    "file.save.save",
    "file.save.saveAs",
    "file.save.saveAs.copy",
    "file.save.saveAs.download",
    "file.save.autoSave",
    "file.print.print",
    "file.print.printPreview",
    "file.options.close",
    "file.export.export.xlsx",
    "file.export.export.csv",
    "file.export.changeFileType.tsv",
  ];

  for (const id of ids) {
    assert.match(
      commands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id}`,
    );
    assert.match(
      commands,
      new RegExp(
        `\\bregisterBuiltinCommand\\([\\s\\S]*?["']${escapeRegExp(id)}["'][\\s\\S]*?\\bwhen:\\s*["']false["']`,
        "m",
      ),
      `Expected ${id} to be hidden from the command palette via when: "false" (ribbon-only alias)`,
    );
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
  }

  // These file ids depend on Page Layout handlers being provided (mirrors `main.ts`).
  const pageLayoutDependent = [
    "file.print.pageSetup",
    "file.print.pageSetup.printTitles",
    "file.print.pageSetup.margins",
    "file.export.createPdf",
    "file.export.export.pdf",
    "file.export.changeFileType.pdf",
    // These are always registered (they export without the desktop print backend), but keep
    // them in this list so the test covers both export paths.
    "file.export.changeFileType.csv",
    "file.export.changeFileType.xlsx",
  ];
  for (const id of pageLayoutDependent) {
    assert.match(
      commands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id} (when pageLayoutHandlers are present)`,
    );
    assert.match(
      commands,
      new RegExp(
        `\\bregisterBuiltinCommand\\([\\s\\S]*?["']${escapeRegExp(id)}["'][\\s\\S]*?\\bwhen:\\s*["']false["']`,
        "m",
      ),
      `Expected ${id} to be hidden from the command palette via when: "false" (ribbon-only alias)`,
    );
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
  }

  // These file ids are registered as hidden aliases that route to canonical view/panel commands.
  const panelAliases = ["file.info.manageWorkbook.versions", "file.info.manageWorkbook.branches"];
  for (const id of panelAliases) {
    assert.match(
      commands,
      new RegExp(`\\bregisterPanelAlias\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id} via registerPanelAlias`,
    );
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
  }

  // Ensure the panel-alias helper keeps File-tab ids hidden from the command palette.
  // These ids are ribbon-only aliases of canonical `view.togglePanel.*` commands.
  assert.match(commands, /\bregisterPanelAlias\b[\s\S]*?\bwhen:\s*["']false["']/m);
});
