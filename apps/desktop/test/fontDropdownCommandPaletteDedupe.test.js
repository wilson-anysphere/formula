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

test("Home → Font ribbon alias ids are hidden from the command palette", () => {
  const sourcePath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const source = stripComments(fs.readFileSync(sourcePath, "utf8"));

  const aliases = [
    { id: "home.font.borders", canonical: "format.borders.all" },
    { id: "home.font.fillColor", canonical: "format.fillColor" },
    { id: "home.font.fontColor", canonical: "format.fontColor" },
    { id: "home.font.fontSize", canonical: "format.fontSize.set" },
  ];

  for (const { id, canonical } of aliases) {
    const idx = source.indexOf(`registerBuiltinCommand(\"${id}\"`);
    assert.ok(idx >= 0, `Expected registerDesktopCommands.ts to register ${id}`);
    const snippet = source.slice(idx, idx + 250);
    assert.match(
      snippet,
      /\bwhen:\s*["']false["']/,
      `Expected ${id} to be hidden via when: "false" (avoid duplicate command palette entries)`,
    );
    assert.match(
      snippet,
      new RegExp(`\\bexecuteCommand\\(["']${escapeRegExp(canonical)}["']\\)`),
      `Expected ${id} to delegate to canonical command ${canonical}`,
    );
  }
});
