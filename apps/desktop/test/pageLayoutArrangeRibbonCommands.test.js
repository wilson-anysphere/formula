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

test("Ribbon schema includes Page Layout → Arrange command ids", () => {
  const schema = readRibbonSchemaSource("pageLayoutTab.ts");

  const ids = [
    "pageLayout.arrange.bringForward",
    "pageLayout.arrange.sendBackward",
    "pageLayout.arrange.selectionPane",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected pageLayoutTab.ts to include ${id}`);
  }
});

test("Page Layout → Arrange ribbon commands are registered in CommandRegistry (no exemptions / no main.ts switch cases)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));

  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  const desktopCommandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const desktopCommands = stripComments(fs.readFileSync(desktopCommandsPath, "utf8"));

  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = stripComments(fs.readFileSync(builtinsPath, "utf8"));

  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = stripComments(fs.readFileSync(disablingPath, "utf8"));

  // Drawing z-order commands are registered in the desktop command catalog.
  const desktopIds = ["pageLayout.arrange.bringForward", "pageLayout.arrange.sendBackward"];
  for (const id of desktopIds) {
    assert.match(
      desktopCommands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id}`,
    );
  }

  // Selection Pane is registered as a builtin panel command.
  assert.match(
    builtins,
    /\bregisterBuiltinCommand\(\s*["']pageLayout\.arrange\.selectionPane["']/,
    "Expected registerBuiltinCommands.ts to register pageLayout.arrange.selectionPane",
  );

  const implementedIds = [...desktopIds, "pageLayout.arrange.selectionPane"];
  for (const id of implementedIds) {
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActions)`,
    );
  }

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActions\(/);
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
});

