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

test("Ribbon schema includes Home → Number command ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  const requiredIds = [
    // Number format presets
    "format.numberFormat.general",
    "format.numberFormat.number",
    "format.numberFormat.currency",
    "format.numberFormat.accounting",
    "format.numberFormat.shortDate",
    "format.numberFormat.longDate",
    "format.numberFormat.time",
    "format.numberFormat.percent",
    "format.numberFormat.fraction",
    "format.numberFormat.scientific",
    "format.numberFormat.text",

    // Quick actions
    "format.numberFormat.commaStyle",
    "format.numberFormat.increaseDecimal",
    "format.numberFormat.decreaseDecimal",

    // Accounting symbol dropdown menu items
    "format.numberFormat.accounting.usd",
    "format.numberFormat.accounting.eur",
    "format.numberFormat.accounting.gbp",
    "format.numberFormat.accounting.jpy",

    // Format Cells / custom formats
    "format.openFormatCells",
    "home.number.moreFormats.custom",
  ];

  for (const id of requiredIds) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }

  // Dropdown triggers remain ribbon-specific ids (menu containers).
  assert.match(schema, /\bid:\s*["']home\.number\.numberFormat["']/);
  assert.match(schema, /\bid:\s*["']home\.number\.moreFormats["']/);

  // Legacy Home → Number ids (previously routed via ribbon handlers) should not exist in the schema.
  const legacyIds = [
    "home.number.percent",
    "home.number.accounting",
    "home.number.date",
    "home.number.comma",
    "home.number.increaseDecimal",
    "home.number.decreaseDecimal",
    "home.number.formatCells",
    "home.number.moreFormats.formatCells",
    "home.cells.format.formatCells",
  ];
  for (const id of legacyIds) {
    assert.doesNotMatch(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to not include legacy id ${id}`);
  }
});

test("Home → Number ribbon commands are registered in CommandRegistry and not handled via main.ts switch cases", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const desktopCommandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const desktopCommands = fs.readFileSync(desktopCommandsPath, "utf8");

  const numberFormatPath = path.join(__dirname, "..", "src", "commands", "registerNumberFormatCommands.ts");
  const numberFormatCommands = fs.readFileSync(numberFormatPath, "utf8");

  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = fs.readFileSync(disablingPath, "utf8");

  // Ensure number formats are wired through the desktop command catalog so ribbon enable/disable
  // can rely on CommandRegistry registration.
  assert.match(desktopCommands, /\bregisterNumberFormatCommands\(/, "Expected registerDesktopCommands.ts to invoke registerNumberFormatCommands");
  assert.match(desktopCommands, /\bregisterBuiltinCommand\(\s*["']home\.number\.moreFormats\.custom["']/, "Expected registerDesktopCommands.ts to register home.number.moreFormats.custom");
  assert.match(desktopCommands, /\bregisterBuiltinCommand\(\s*["']format\.openFormatCells["']/, "Expected registerDesktopCommands.ts to register format.openFormatCells");

  const numberFormatIds = [
    "format.numberFormat.general",
    "format.numberFormat.number",
    "format.numberFormat.currency",
    "format.numberFormat.accounting",
    "format.numberFormat.shortDate",
    "format.numberFormat.longDate",
    "format.numberFormat.time",
    "format.numberFormat.percent",
    "format.numberFormat.fraction",
    "format.numberFormat.scientific",
    "format.numberFormat.text",
    "format.numberFormat.commaStyle",
    "format.numberFormat.increaseDecimal",
    "format.numberFormat.decreaseDecimal",
  ];
  for (const id of numberFormatIds) {
    assert.match(
      numberFormatCommands,
      new RegExp(`\\bregister\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerNumberFormatCommands.ts to register ${id}`,
    );
  }

  // Accounting currency symbol menu items are registered via an array loop.
  assert.match(numberFormatCommands, /\bformat\.numberFormat\.accounting\./, "Expected registerNumberFormatCommands.ts to register accounting symbol commands");
  for (const currency of ["usd", "eur", "gbp", "jpy"]) {
    assert.match(numberFormatCommands, new RegExp(`\\bid:\\s*["']${escapeRegExp(currency)}["']`), `Expected registerNumberFormatCommands.ts to include accounting symbol ${currency}`);
  }

  const implementedIds = [
    ...numberFormatIds,
    "format.numberFormat.accounting.usd",
    "format.numberFormat.accounting.eur",
    "format.numberFormat.accounting.gbp",
    "format.numberFormat.accounting.jpy",
    "format.openFormatCells",
    "home.number.moreFormats.custom",
  ];
  for (const id of implementedIds) {
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
  }

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActions\(/);
});
