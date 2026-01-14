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

test("Ribbon schema includes Home → Alignment → Merge & Center command ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  const ids = [
    "home.alignment.mergeCenter.mergeCenter",
    "home.alignment.mergeCenter.mergeAcross",
    "home.alignment.mergeCenter.mergeCells",
    "home.alignment.mergeCenter.unmergeCells",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }

  // Ensure the dropdown trigger is still present.
  assert.match(schema, /\bid:\s*["']home\.alignment\.mergeCenter["']/);
  assert.match(schema, /\bkind:\s*["']dropdown["']/);
});

test("Merge & Center ribbon commands are registered in CommandRegistry (not exempted as ribbon-only)", () => {
  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const commands = fs.readFileSync(commandsPath, "utf8");

  const commandIds = [
    "home.alignment.mergeCenter.mergeCenter",
    "home.alignment.mergeCenter.mergeAcross",
    "home.alignment.mergeCenter.mergeCells",
    "home.alignment.mergeCenter.unmergeCells",
  ];

  for (const id of commandIds) {
    assert.match(
      commands,
      new RegExp(`\\bregisterMergeCommand\\(\\{[\\s\\S]*?\\bid:\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register merge command id ${id}`,
    );
  }

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
