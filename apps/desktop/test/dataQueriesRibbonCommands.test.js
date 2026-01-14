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

test("Ribbon schema includes Data → Queries & Connections controls", () => {
  const schema = readRibbonSchemaSource("dataTab.ts");

  // Toggle button.
  assert.match(schema, /\bid:\s*["']data\.queriesConnections\.queriesConnections["']/);
  assert.match(schema, /\bkind:\s*["']toggle["']/);

  // Refresh All dropdown + key menu items.
  assert.match(schema, /\bid:\s*["']data\.queriesConnections\.refreshAll["']/);
  assert.match(schema, /\bkind:\s*["']dropdown["']/);

  const refreshMenuIds = [
    "data.queriesConnections.refreshAll",
    "data.queriesConnections.refreshAll.refresh",
    "data.queriesConnections.refreshAll.refreshAllConnections",
    "data.queriesConnections.refreshAll.refreshAllQueries",
  ];
  for (const id of refreshMenuIds) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`));
  }
});

test("Data → Queries & Connections ribbon commands are registered in CommandRegistry (not wired only in main.ts)", () => {
  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerDataQueriesCommands.ts");
  const commands = fs.readFileSync(commandsPath, "utf8");

  const commandIds = [
    "data.queriesConnections.queriesConnections",
    "data.queriesConnections.refreshAll",
    "data.queriesConnections.refreshAll.refresh",
    "data.queriesConnections.refreshAll.refreshAllConnections",
    "data.queriesConnections.refreshAll.refreshAllQueries",
  ];

  // Ensure all ribbon ids have corresponding built-in command registrations.
  for (const id of commandIds) {
    assert.match(
      commands,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Expected registerDataQueriesCommands.ts to reference command id ${id}`,
    );
  }
  assert.match(commands, /\bregisterBuiltinCommand\(/);

  // Toggle command should open/close the DATA_QUERIES panel based on layout placement.
  assert.match(commands, /\bgetPanelPlacement\(/);
  assert.match(commands, /\bopenPanel\(PanelIds\.DATA_QUERIES\)/);
  assert.match(commands, /\bclosePanel\(PanelIds\.DATA_QUERIES\)/);

  // Refresh commands should ultimately call PowerQueryService.refreshAll().
  assert.match(commands, /\bservice\.refreshAll\(\)/);

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // main.ts should register the commands and avoid ribbon-only wiring. Registration is
  // centralized in `registerDesktopCommands` so the desktop shell shares a single command catalog.
  assert.match(main, /\bregisterDesktopCommands\(/);
  assert.match(main, /\bdataQueriesHandlers\s*:/);
  assert.doesNotMatch(main, /\bregisterDataQueriesCommands\(/);
  assert.doesNotMatch(main, /\btoggleOverrides:\s*\{[\s\S]*?["']data\.queriesConnections\.queriesConnections["']\s*:/m);
  for (const id of commandIds.slice(1)) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcommandId\\s*===\\s*["']${escapeRegExp(id)}["']`),
      `Did not expect main.ts to special-case refresh command id ${id}`,
    );
  }

  // Pressed state sync should reflect whether the Data Queries panel is open.
  assert.match(main, /"data\.queriesConnections\.queriesConnections":\s*isPanelOpen\(\s*PanelIds\.DATA_QUERIES\s*\)/);

  // Since these ids are now real commands, they should not be kept in the ribbon
  // CommandRegistry exemption list (that list is for ribbon-only wiring).
  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = fs.readFileSync(disablingPath, "utf8");
  for (const id of commandIds) {
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
  }
});
