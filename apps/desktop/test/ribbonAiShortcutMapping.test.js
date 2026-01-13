import test from "node:test";
import assert from "node:assert/strict";

import { deriveRibbonShortcutById } from "../src/ribbon/ribbonShortcuts.ts";

test("Ribbon AI buttons map to canonical command ids for keybinding tooltip display", () => {
  const index = new Map([
    ["view.togglePanel.aiChat", ["Ctrl+Shift+A"]],
    ["ai.inlineEdit", ["Ctrl+K"]],
  ]);

  const shortcuts = deriveRibbonShortcutById(index);
  assert.equal(shortcuts["view.togglePanel.aiChat"], "Ctrl+Shift+A");
  assert.equal(shortcuts["ai.inlineEdit"], "Ctrl+K");
});
