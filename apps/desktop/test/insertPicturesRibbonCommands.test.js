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

test("Ribbon schema includes Insert → Pictures command ids", () => {
  const schema = readRibbonSchemaSource("insertTab.ts");

  const ids = [
    "insert.illustrations.pictures",
    "insert.illustrations.pictures.thisDevice",
    "insert.illustrations.pictures.stockImages",
    "insert.illustrations.pictures.onlinePictures",
    "insert.illustrations.onlinePictures",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected insertTab.ts to include ${id}`);
  }
});

test("Insert → Pictures ribbon commands are registered in CommandRegistry and not handled via main.ts switch cases", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const commands = fs.readFileSync(commandsPath, "utf8");

  const ids = [
    "insert.illustrations.pictures",
    "insert.illustrations.pictures.thisDevice",
    "insert.illustrations.pictures.stockImages",
    "insert.illustrations.pictures.onlinePictures",
    "insert.illustrations.onlinePictures",
  ];

  for (const id of ids) {
    assert.match(
      commands,
      new RegExp(`\\bregisterInsertPicturesCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id} via registerInsertPicturesCommand`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
  }

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActionsFromCommands\(/);
});

