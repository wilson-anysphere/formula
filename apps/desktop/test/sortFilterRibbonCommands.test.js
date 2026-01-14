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

test("Ribbon schema includes Sort & Filter command ids (Home/Data tabs)", () => {
  const schema = readRibbonSchemaSource(["homeTab.ts", "dataTab.ts"]);

  const ids = [
    "data.sortFilter.sortAtoZ",
    "data.sortFilter.sortZtoA",
    "home.editing.sortFilter.customSort",
    "data.sortFilter.sort.customSort",
    "data.sortFilter.filter",
    "data.sortFilter.clear",
    "data.sortFilter.reapply",
    "data.sortFilter.advanced.clearFilter",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected ribbon schema to include ${id}`);
  }
});

test("Sort & Filter ribbon commands are registered in CommandRegistry (no exemptions / no main.ts switch cases)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));
  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  const desktopCommandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const desktopCommands = stripComments(fs.readFileSync(desktopCommandsPath, "utf8"));

  const autoFilterCommandsPath = path.join(__dirname, "..", "src", "commands", "registerRibbonAutoFilterCommands.ts");
  const autoFilterCommands = fs.readFileSync(autoFilterCommandsPath, "utf8");

  const sortFilterCommandsPath = path.join(__dirname, "..", "src", "commands", "registerSortFilterCommands.ts");
  const sortFilterCommands = stripComments(fs.readFileSync(sortFilterCommandsPath, "utf8"));

  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = stripComments(fs.readFileSync(disablingPath, "utf8"));

  // Guardrail: AutoFilter is a registered CommandRegistry toggle command, so it should not be
  // special-cased as a ribbon `toggleOverrides` handler (it should dispatch through CommandRegistry).
  const autoFilterOverrideIds = ["data.sortFilter.filter", "data.sortFilter.clear", "data.sortFilter.reapply", "data.sortFilter.advanced.clearFilter"];
  for (const id of autoFilterOverrideIds) {
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
  }
  assert.doesNotMatch(
    router,
    /\bcommandId\.startsWith\(\s*["']data\.sortFilter\./,
    "Did not expect ribbonCommandRouter.ts to add bespoke data.sortFilter.* prefix routing (dispatch should go through CommandRegistry)",
  );

  // MVP AutoFilter commands are registered via the shared helper (invoked by registerDesktopCommands,
  // with host implementations injected from main.ts).
  const autoFilterIds = [
    "data.sortFilter.filter",
    "data.sortFilter.clear",
    "data.sortFilter.reapply",
    "data.sortFilter.advanced.clearFilter",
  ];

  assert.match(
    desktopCommands,
    /\bregisterRibbonAutoFilterCommands\(/,
    "Expected registerDesktopCommands.ts to invoke registerRibbonAutoFilterCommands",
  );

  for (const id of autoFilterIds) {
    assert.match(
      autoFilterCommands,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Expected registerRibbonAutoFilterCommands.ts to reference ${id}`,
    );
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (ribbon commands should be routed via the ribbon command router)`,
    );
  }

  // Sort + Custom Sort commands are registered via registerSortFilterCommands (invoked by registerDesktopCommands).
  assert.match(
    desktopCommands,
    /\bregisterSortFilterCommands\(/,
    "Expected registerDesktopCommands.ts to invoke registerSortFilterCommands",
  );

  const sortFilterIds = [
    ["sortAtoZ", "data.sortFilter.sortAtoZ"],
    ["sortZtoA", "data.sortFilter.sortZtoA"],
    ["homeCustomSort", "home.editing.sortFilter.customSort"],
    ["dataCustomSort", "data.sortFilter.sort.customSort"],
  ];
  for (const [key, id] of sortFilterIds) {
    assert.match(
      sortFilterCommands,
      new RegExp(`\\b${escapeRegExp(key)}\\s*:\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerSortFilterCommands.ts to define ${key} as ${id}`,
    );
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (ribbon commands should be routed via the ribbon command router)`,
    );
  }

  // Sanity check: main.ts should mount the ribbon through the shared router.
  assert.match(main, /\bcreateRibbonActions\(/);
  // And the router should delegate registered commands to the CommandRegistry bridge.
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
});
