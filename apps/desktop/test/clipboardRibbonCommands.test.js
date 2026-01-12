import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema includes Home → Clipboard command ids", () => {
  const schemaPath = path.join(__dirname, "..", "src", "ribbon", "ribbonSchema.ts");
  const schema = fs.readFileSync(schemaPath, "utf8");

  const ids = [
    // Clipboard group core actions.
    "home.clipboard.cut",
    "home.clipboard.copy",
    "home.clipboard.paste",
    "home.clipboard.pasteSpecial",

    // Paste dropdown menu items.
    "home.clipboard.paste.default",
    "home.clipboard.paste.values",
    "home.clipboard.paste.formulas",
    "home.clipboard.paste.formats",
    "home.clipboard.paste.transpose",

    // Paste Special dropdown menu items.
    "home.clipboard.pasteSpecial.dialog",
    "home.clipboard.pasteSpecial.values",
    "home.clipboard.pasteSpecial.formulas",
    "home.clipboard.pasteSpecial.formats",
    "home.clipboard.pasteSpecial.transpose",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected ribbonSchema.ts to include ${id}`);
  }

  // Ensure the two primary controls are dropdowns.
  assert.match(schema, /\bid:\s*["']home\.clipboard\.paste["'][\s\S]*?\bkind:\s*["']dropdown["']/);
  assert.match(schema, /\bid:\s*["']home\.clipboard\.pasteSpecial["'][\s\S]*?\bkind:\s*["']dropdown["']/);
});

test("Desktop main.ts wires Home → Clipboard ribbon commands to SpreadsheetApp clipboard ops", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const expects = [
    { id: "home.clipboard.cut", pattern: /\bclipboardCut\(\)/ },
    { id: "home.clipboard.copy", pattern: /\bclipboardCopy\(\)/ },
    { id: "home.clipboard.paste", pattern: /\bclipboardPaste\(\)/ },
    { id: "home.clipboard.paste.default", pattern: /\bclipboardPaste\(\)/ },
    { id: "home.clipboard.paste.values", pattern: /\bclipboardPasteSpecial\(\s*["']values["']\s*\)/ },
    { id: "home.clipboard.paste.formulas", pattern: /\bclipboardPasteSpecial\(\s*["']formulas["']\s*\)/ },
    { id: "home.clipboard.paste.formats", pattern: /\bclipboardPasteSpecial\(\s*["']formats["']\s*\)/ },
    { id: "home.clipboard.pasteSpecial.values", pattern: /\bclipboardPasteSpecial\(\s*["']values["']\s*\)/ },
    { id: "home.clipboard.pasteSpecial.formulas", pattern: /\bclipboardPasteSpecial\(\s*["']formulas["']\s*\)/ },
    { id: "home.clipboard.pasteSpecial.formats", pattern: /\bclipboardPasteSpecial\(\s*["']formats["']\s*\)/ },
  ];

  for (const { id, pattern } of expects) {
    assert.match(main, new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`), `Expected main.ts to handle ${id}`);
    assert.match(main, pattern, `Expected main.ts to invoke ${String(pattern)} for ${id}`);
  }

  // Paste special should surface the existing Paste Special menu items through a quick pick.
  assert.match(main, /\bcase\s+["']home\.clipboard\.pasteSpecial\.dialog["']:/);
  assert.match(main, /\bgetPasteSpecialMenuItems\(\)/);
  assert.match(main, /\bshowQuickPick\(/);

  // Transpose should invoke the spreadsheet paste pipeline (Excel-style).
  assert.match(main, /\bcase\s+["']home\.clipboard\.paste\.transpose["']:/);
  assert.match(main, /\bcase\s+["']home\.clipboard\.pasteSpecial\.transpose["']:/);
  assert.match(main, /\bclipboardPasteSpecial\(\s*["']all["']\s*,\s*\{\s*transpose:\s*true\s*\}\s*\)/);
  assert.doesNotMatch(main, /Paste Transpose not implemented/);
});
