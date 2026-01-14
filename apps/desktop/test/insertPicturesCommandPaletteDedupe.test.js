import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Insert → Pictures ribbon aliases are hidden from the command palette", () => {
  const sourcePath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const source = stripComments(fs.readFileSync(sourcePath, "utf8"));

  // Guardrail: ensure the Insert → Pictures helper passes through `when` so hiding
  // menu-item aliases from the command palette is effective.
  assert.match(
    source,
    /\bregisterInsertPicturesCommand\b[\s\S]*?\bwhen:\s*options\.when\s*\?\?\s*null/,
    "Expected registerInsertPicturesCommand to forward options.when into registerBuiltinCommand(...)",
  );

  const extractInvocation = (commandId) => {
    const re = new RegExp(
      String.raw`\bregisterInsertPicturesCommand\(\s*["']${escapeRegExp(commandId)}["'][\s\S]*?\n?\s*\);\n?`,
    );
    return source.match(re)?.[0] ?? null;
  };

  const hiddenIds = [
    "insert.illustrations.pictures.thisDevice",
    "insert.illustrations.pictures.stockImages",
    "insert.illustrations.pictures.onlinePictures",
    "insert.illustrations.onlinePictures",
  ];

  for (const id of hiddenIds) {
    const invocation = extractInvocation(id);
    assert.ok(invocation, `Expected registerDesktopCommands.ts to register ${id}`);
    assert.match(invocation, /\bwhen:\s*["']false["']/, `Expected ${id} to be hidden via when: "false"`);
  }

  const canonical = extractInvocation("insert.illustrations.pictures");
  assert.ok(canonical, "Expected registerDesktopCommands.ts to register insert.illustrations.pictures");
  assert.doesNotMatch(canonical, /\bwhen:\s*["']false["']/, "Expected insert.illustrations.pictures to remain visible");

  // Ensure the "This Device…" menu item is routed through the canonical Pictures command so
  // command-palette recents tracking sees the canonical id.
  const thisDevice = extractInvocation("insert.illustrations.pictures.thisDevice");
  assert.ok(thisDevice, "Expected registerDesktopCommands.ts to register insert.illustrations.pictures.thisDevice");
  assert.match(
    thisDevice,
    /\bexecuteCommand\(\s*["']insert\.illustrations\.pictures["']\s*\)/,
    "Expected insert.illustrations.pictures.thisDevice to delegate to insert.illustrations.pictures",
  );
});
