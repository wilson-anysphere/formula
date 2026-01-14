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
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));
  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const commands = stripComments(fs.readFileSync(commandsPath, "utf8"));

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

  // `insert.illustrations.pictures` is the canonical command (currently maps to "This Device").
  // Keep the more specific ribbon menu ids registered for schema coverage, but hide them from the
  // command palette to avoid duplicate/overlapping entries.
  const hiddenIds = [
    "insert.illustrations.pictures.thisDevice",
    "insert.illustrations.pictures.stockImages",
    "insert.illustrations.pictures.onlinePictures",
    "insert.illustrations.onlinePictures",
  ];
  for (const id of hiddenIds) {
    const idx = commands.indexOf(`registerInsertPicturesCommand(\"${id}\"`);
    assert.ok(idx >= 0, `Expected to find registerInsertPicturesCommand(...) call for ${id}`);
    const snippet = commands.slice(idx, idx + 900);
    assert.match(snippet, /\bwhen:\s*["']false["']/, `Expected ${id} to be hidden via when: "false"`);
  }

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActions\(/);
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
});
