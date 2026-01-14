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

test("Ribbon schema includes Home → Styles command ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  const ids = [
    "home.styles.cellStyles.goodBadNeutral",
    "home.styles.formatAsTable.light",
    "home.styles.formatAsTable.medium",
    "home.styles.formatAsTable.dark",
  ];
  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }

  // Sanity: ensure the dropdown triggers remain present in the schema (even if they are not registered commands).
  assert.match(schema, /\bid:\s*["']home\.styles\.cellStyles["']/);
  assert.match(schema, /\bid:\s*["']home\.styles\.formatAsTable["']/);
});

test("Home → Styles ribbon commands are registered in CommandRegistry and not handled via main.ts switch cases", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));

  const desktopCommandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const desktopCommands = fs.readFileSync(desktopCommandsPath, "utf8");

  const homeStylesPath = path.join(__dirname, "..", "src", "commands", "registerHomeStylesCommands.ts");
  const homeStylesCommands = fs.readFileSync(homeStylesPath, "utf8");

  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = fs.readFileSync(disablingPath, "utf8");
  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  // Ensure the desktop command catalog wires in the Home Styles registrations.
  assert.match(desktopCommands, /\bregisterHomeStylesCommands\(/, "Expected registerDesktopCommands.ts to invoke registerHomeStylesCommands");

  const ids = [
    "home.styles.cellStyles.goodBadNeutral",
    "home.styles.formatAsTable.light",
    "home.styles.formatAsTable.medium",
    "home.styles.formatAsTable.dark",
  ];
  for (const id of ids) {
    assert.match(
      homeStylesCommands,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Expected registerHomeStylesCommands.ts to reference command id ${id}`,
    );
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

  // Sanity check: main.ts should mount the ribbon through the ribbon command router, which in turn
  // delegates registered ribbon ids to the CommandRegistry bridge (`createRibbonActionsFromCommands`).
  assert.match(main, /\bcreateRibbonActions\(/);
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
});
