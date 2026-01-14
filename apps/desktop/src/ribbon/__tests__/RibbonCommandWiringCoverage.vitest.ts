import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import type { RibbonSchema } from "../ribbonSchema";
import { defaultRibbonSchema } from "../ribbonSchema";

function collectRibbonCommandIds(schema: RibbonSchema): string[] {
  const ids = new Set<string>();
  for (const tab of schema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        ids.add(button.id);
        for (const item of button.menuItems ?? []) {
          ids.add(item.id);
        }
      }
    }
  }
  return [...ids].sort();
}

describe("Ribbon command wiring coverage (Home → Font dropdowns)", () => {
  it("uses canonical `format.*` ids for Font dropdown menu items", () => {
    const ids = collectRibbonCommandIds(defaultRibbonSchema);

    // Guard against a broken traversal so the test can't pass vacuously.
    expect(ids).toContain("home.font.fillColor");
    expect(ids).toContain("home.font.fontColor");
    expect(ids).toContain("home.font.borders");
    expect(ids).toContain("home.font.clearFormatting");

    // Font dropdown menu items were historically wired via `home.font.*` prefixes in `main.ts`.
    // These actions are now canonical `format.*` commands so ribbon/command-palette/keybindings
    // share a single command surface.
    const legacyMenuItemPrefixes = [
      "home.font.fillColor.",
      "home.font.fontColor.",
      "home.font.borders.",
      "home.font.clearFormatting.",
    ] as const;

    const legacyMenuItemIds = ids.filter((id) => legacyMenuItemPrefixes.some((prefix) => id.startsWith(prefix)));
    expect(
      legacyMenuItemIds,
      `Legacy Home→Font menu item ids should not exist in the ribbon schema:\n${legacyMenuItemIds.map((id) => `- ${id}`).join("\n")}`,
    ).toEqual([]);

    // Representative new ids (the complete set is covered by CommandRegistry + ribbon schema tests).
    expect(ids).toContain("format.fillColor.none");
    expect(ids).toContain("format.fontColor.black");
    expect(ids).toContain("format.borders.top");
    expect(ids).toContain("format.clearFormats");
  });

  it("does not use `home.font.*` prefix parsing for font dropdown menu items in main.ts", () => {
    const mainTsPath = fileURLToPath(new URL("../../main.ts", import.meta.url));
    const source = readFileSync(mainTsPath, "utf8");

    // Ensure the old prefix-parsing blocks were removed. (The dropdown trigger ids
    // like `home.font.fillColor` may still exist as fallbacks; only the menu item
    // prefix parsing is disallowed.)
    expect(source).not.toContain("home.font.fillColor.");
    expect(source).not.toContain("home.font.fontColor.");
    expect(source).not.toContain("home.font.borders.");
    expect(source).not.toContain("home.font.clearFormatting.");
  });
});
