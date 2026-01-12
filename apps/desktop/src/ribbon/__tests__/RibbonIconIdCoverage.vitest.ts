import { describe, it, expect } from "vitest";

import type { RibbonButtonDefinition } from "../ribbonSchema";
import { defaultRibbonSchema } from "../ribbonSchema";

function collectButtonsBySize(size: "icon" | "large"): RibbonButtonDefinition[] {
  const buttons: RibbonButtonDefinition[] = [];
  for (const tab of defaultRibbonSchema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        if (button.size === size) {
          buttons.push(button);
        }
      }
    }
  }
  return buttons;
}

describe("Ribbon schema iconId coverage", () => {
  it('assigns an iconId for every schema button with size: "icon"', () => {
    const iconButtons = collectButtonsBySize("icon");

    // Guard against a broken traversal so the test can't pass vacuously.
    expect(iconButtons.map((button) => button.id)).toContain("home.font.bold");

    const missing = iconButtons.filter((button) => !button.iconId).map((button) => button.id);
    expect(missing).toEqual([]);
  });

  it('assigns an iconId for every schema button with size: "large"', () => {
    const largeButtons = collectButtonsBySize("large");

    // Guard against a broken traversal so the test can't pass vacuously.
    expect(largeButtons.map((button) => button.id)).toContain("file.save.save");

    const missing = largeButtons.filter((button) => !button.iconId).map((button) => button.id);
    expect(missing).toEqual([]);
  });
});
