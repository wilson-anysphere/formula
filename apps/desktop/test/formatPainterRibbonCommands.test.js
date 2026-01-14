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

test("Ribbon schema includes Format Painter command id (Home → Clipboard)", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  assert.match(schema, /\bid:\s*["']format\.toggleFormatPainter["']/);
});

test("Format Painter ribbon command is registered in CommandRegistry (no exemptions / no main.ts switch cases)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));

  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  const desktopCommandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const desktopCommands = stripComments(fs.readFileSync(desktopCommandsPath, "utf8"));

  const formatPainterPath = path.join(__dirname, "..", "src", "commands", "formatPainterCommand.ts");
  const formatPainter = stripComments(fs.readFileSync(formatPainterPath, "utf8"));

  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = stripComments(fs.readFileSync(disablingPath, "utf8"));

  // Ensure the command id stays canonical and is registered via the format painter command helper.
  assert.match(formatPainter, /\bFORMAT_PAINTER_COMMAND_ID\s*=\s*["']format\.toggleFormatPainter["']/);
  assert.match(desktopCommands, /\bregisterFormatPainterCommand\(/, "Expected registerDesktopCommands.ts to invoke registerFormatPainterCommand");

  const id = "format.toggleFormatPainter";
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

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActions\(/);
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
});

