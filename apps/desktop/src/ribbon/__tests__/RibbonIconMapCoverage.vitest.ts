import { describe, it, expect } from "vitest";

import { defaultRibbonSchema } from "../ribbonSchema";
import { ribbonIconMap } from "../../ui/icons/ribbonIconMap";

function collectButtonIdsBySize(size: "icon" | "large"): string[] {
  const ids: string[] = [];
  for (const tab of defaultRibbonSchema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        if (button.size === size) {
          ids.push(button.id);
        }
      }
    }
  }
  return ids;
}

describe("ribbonIconMap coverage", () => {
  it('includes mappings for every schema button with size: "icon"', () => {
    const iconButtonIds = collectButtonIdsBySize("icon");

    // Guard against a broken traversal so the test can't pass vacuously.
    expect(iconButtonIds).toContain("home.font.bold");

    const missing = iconButtonIds.filter((id) => !Object.prototype.hasOwnProperty.call(ribbonIconMap, id));
    expect(missing).toEqual([]);
  });

  it('includes mappings for every schema button with size: "large"', () => {
    const largeButtonIds = collectButtonIdsBySize("large");

    // Guard against a broken traversal so the test can't pass vacuously.
    expect(largeButtonIds).toContain("file.save.save");

    const missing = largeButtonIds.filter((id) => !Object.prototype.hasOwnProperty.call(ribbonIconMap, id));
    expect(missing).toEqual([]);
  });
});

