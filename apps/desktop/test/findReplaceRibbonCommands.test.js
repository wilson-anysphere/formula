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

test("Ribbon schema includes Home → Find & Select command ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  // The Home → Find dropdown should dispatch the canonical CommandRegistry ids so ribbon, command palette,
  // and keybindings share the same execution path.
  const requiredIds = ["edit.find", "edit.replace", "navigation.goTo"];
  for (const id of requiredIds) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }
});

test("Desktop main.ts routes Find/Replace ribbon commands through the CommandRegistry", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));
  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  const desktopCommandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const desktopCommands = stripComments(fs.readFileSync(desktopCommandsPath, "utf8"));

  // Find/Replace/Go To are registered as commands in registerDesktopCommands (overriding the builtin no-op
  // registrations) and should not be handled by `handleRibbonCommand` switch cases.
  const ids = ["edit.find", "edit.replace", "navigation.goTo"];
  for (const id of ids) {
    assert.match(
      desktopCommands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be routed via the ribbon command router)`,
    );
    assert.doesNotMatch(
      router,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected ribbonCommandRouter.ts to not handle ${id} via switch case (should dispatch via CommandRegistry)`,
    );
    assert.doesNotMatch(
      router,
      new RegExp(`\\bcommandId\\s*===\\s*["']${escapeRegExp(id)}["']`),
      `Expected ribbonCommandRouter.ts to not special-case ${id} via commandId === checks (should dispatch via CommandRegistry)`,
    );
    assert.doesNotMatch(
      router,
      new RegExp(`\\bcommandOverrides:\\s*\\{[\\s\\S]*?["']${escapeRegExp(id)}["']\\s*:`),
      `Expected ribbonCommandRouter.ts to not special-case ${id} via commandOverrides (should dispatch via CommandRegistry)`,
    );
  }

  // Sanity check: main.ts should mount the ribbon through the shared router.
  assert.match(main, /\bcreateRibbonActions\(/);
  // And the router should delegate registered commands to the CommandRegistry bridge.
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
});
