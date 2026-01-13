import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Ribbon AI buttons map to canonical command ids for keybinding tooltip display", () => {
  const shortcutsPath = path.join(__dirname, "..", "src", "ribbon", "ribbonShortcuts.ts");
  const source = fs.readFileSync(shortcutsPath, "utf8");

  // The ribbon buttons now emit canonical command ids directly (matching keybindings/command palette),
  // so the ribbon shortcut lookup table should use those ids as keys.
  assert.match(source, /["']view\.togglePanel\.aiChat["']\s*:\s*["']view\.togglePanel\.aiChat["']/);
  assert.match(source, /["']ai\.inlineEdit["']\s*:\s*["']ai\.inlineEdit["']/);

  // Legacy ribbon-only ids should no longer appear as shortcut keys (the corresponding buttons use
  // stable Playwright testIds instead).
  assert.doesNotMatch(source, /["']open-panel-ai-chat["']\s*:/);
  assert.doesNotMatch(source, /["']open-inline-ai-edit["']\s*:/);
});
