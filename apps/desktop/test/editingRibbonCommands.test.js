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

function countMatches(source, pattern) {
  const re = pattern instanceof RegExp ? pattern : new RegExp(String(pattern), "g");
  const matches = source.match(re);
  return matches ? matches.length : 0;
}

test("Ribbon schema aligns Home → Editing AutoSum/Fill ids with CommandRegistry ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  // Canonical command ids.
  const requiredIds = ["edit.autoSum", "edit.fillDown", "edit.fillRight", "edit.fillUp", "edit.fillLeft"];
  for (const id of requiredIds) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }

  // AutoSum should be used for both the dropdown button id and the default "Sum" menu item.
  assert.ok(
    countMatches(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp("edit.autoSum")}["']`, "g")) >= 2,
    "Expected edit.autoSum to appear at least twice (button + menu item)",
  );

  // Legacy ids should not be present.
  const legacyIds = [
    "home.editing.autoSum",
    "home.editing.autoSum.sum",
    "home.editing.fill.down",
    "home.editing.fill.right",
    "home.editing.fill.up",
    "home.editing.fill.left",
  ];
  for (const id of legacyIds) {
    assert.doesNotMatch(
      schema,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`),
      `Expected homeTab.ts to not include legacy id ${id}`,
    );
  }
});

test("Desktop main.ts routes canonical Editing ribbon commands through the CommandRegistry (no legacy mapping)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));
  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = stripComments(fs.readFileSync(builtinsPath, "utf8"));
  const desktopCommandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const desktopCommands = stripComments(fs.readFileSync(desktopCommandsPath, "utf8"));

  // Canonical editing ids should be registered as builtin commands so ribbon, command palette,
  // and keybindings share the same execution path (via the ribbon command router).
  const expects = [
    "edit.autoSum",
    "edit.fillDown",
    "edit.fillRight",
    "edit.fillUp",
    "edit.fillLeft",
    // Ribbon-specific AutoSum dropdown variants should be registered and dispatched via CommandRegistry.
    "home.editing.autoSum.average",
    "home.editing.autoSum.countNumbers",
    "home.editing.autoSum.max",
    "home.editing.autoSum.min",
  ];
  for (const id of expects) {
    assert.match(
      builtins,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerBuiltinCommands.ts to register ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be routed via the ribbon command router)`,
    );
  }

  // The ribbon schema now uses the canonical `edit.fillUp`/`edit.fillLeft` ids, so the legacy
  // `home.editing.fill.*` aliases should not be present in the builtin command catalog.
  const disallowedBuiltinIds = ["home.editing.fill.up", "home.editing.fill.left"];
  for (const id of disallowedBuiltinIds) {
    assert.doesNotMatch(
      builtins,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerBuiltinCommands.ts to not register legacy alias ${id}`,
    );
  }

  // Ribbon-only "Fill → Series…" is registered in registerDesktopCommands so ribbon execution goes through CommandRegistry.
  assert.match(
    desktopCommands,
    /\bregisterBuiltinCommand\(\s*["']home\.editing\.fill\.series["']/,
    "Expected registerDesktopCommands.ts to register home.editing.fill.series",
  );
  assert.doesNotMatch(
    main,
    /\bcase\s+["']home\.editing\.fill\.series["']:/,
    "Expected main.ts to not handle home.editing.fill.series via switch case (should be routed via the ribbon command router)",
  );

  // Ensure the old ribbon-only ids are no longer mapped in main.ts.
  const legacyCases = [
    "home.editing.autoSum",
    "home.editing.autoSum.sum",
    "home.editing.fill.down",
    "home.editing.fill.right",
    "home.editing.fill.up",
    "home.editing.fill.left",
  ];
  for (const id of legacyCases) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts not to contain legacy case ${id}`,
    );
  }

  // Sanity check: main.ts should mount the ribbon through the shared router.
  assert.match(main, /\bcreateRibbonActions\(/);
  // And the router should delegate registered commands to the CommandRegistry bridge.
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
});
