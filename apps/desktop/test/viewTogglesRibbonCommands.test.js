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
  }

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActions\(/);
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
});

