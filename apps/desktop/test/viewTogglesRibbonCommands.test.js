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

test("Ribbon schema includes View → toggle command ids (Show/Window)", () => {
  const schema = readRibbonSchemaSource("viewTab.ts");

  const ids = [
    "view.toggleShowFormulas",
    "view.togglePerformanceStats",
    "view.toggleSplitView",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected viewTab.ts to include ${id}`);
  }
});

test("View toggle ribbon commands are registered in CommandRegistry (no exemptions / no main.ts switch cases)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));

  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = stripComments(fs.readFileSync(builtinsPath, "utf8"));

  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = stripComments(fs.readFileSync(disablingPath, "utf8"));

  const ids = ["view.toggleShowFormulas", "view.togglePerformanceStats", "view.toggleSplitView"];
  for (const id of ids) {
    assert.match(
      builtins,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerBuiltinCommands.ts to register ${id}`,
    );
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

    // Guardrail: avoid reintroducing bespoke routing paths in the ribbon command router.
    // These ids are registered commands and should dispatch via CommandRegistry.
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
      new RegExp(`\\btoggleOverrides:\\s*\\{[\\s\\S]*?["']${escapeRegExp(id)}["']\\s*:`),
      `Expected ribbonCommandRouter.ts to not special-case ${id} via toggleOverrides (should dispatch via CommandRegistry)`,
    );
    assert.doesNotMatch(
      router,
      new RegExp(`\\bcommandOverrides:\\s*\\{[\\s\\S]*?["']${escapeRegExp(id)}["']\\s*:`),
      `Expected ribbonCommandRouter.ts to not special-case ${id} via commandOverrides (should dispatch via CommandRegistry)`,
    );
  }

  // Pressed state should be computed by main.ts (not in the router). Verify the ribbon
  // toggle ids are present in the pressed-state mapping so the UI stays in sync.
  const pressedByIdStart = main.indexOf("const pressedById");
  assert.ok(pressedByIdStart !== -1, "Expected main.ts to define a pressedById mapping for ribbon toggles");
  const pressedByIdEnd = main.indexOf("const numberFormatLabel", pressedByIdStart);
  assert.ok(pressedByIdEnd !== -1, "Expected to find end of pressedById mapping in main.ts");
  const pressedByIdBlock = main.slice(pressedByIdStart, pressedByIdEnd);
  for (const id of ids) {
    assert.match(
      pressedByIdBlock,
      new RegExp(`["']${escapeRegExp(id)}["']\\s*:`),
      `Expected main.ts pressedById mapping to include ${id}`,
    );
  }

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActions\(/);
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
});
