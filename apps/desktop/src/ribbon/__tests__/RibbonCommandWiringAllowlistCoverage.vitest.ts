import { describe, expect, it } from "vitest";

import type { RibbonSchema } from "../ribbonSchema";
import { defaultRibbonSchema } from "../ribbonSchema";
import { handledRibbonCommandIds, unimplementedRibbonCommandIds } from "../ribbonCommandRouter";

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
  return Array.from(ids).sort();
}

describe("Ribbon command wiring allowlists", () => {
  it("classifies every command id in defaultRibbonSchema as handled vs intentionally unimplemented", () => {
    const schemaCommandIds = collectRibbonCommandIds(defaultRibbonSchema);
    const schemaIdSet = new Set(schemaCommandIds);

    // Guard against a broken traversal so the test can't pass vacuously.
    expect(schemaCommandIds).toContain("clipboard.copy");
    expect(schemaCommandIds).toContain("format.toggleBold");
    expect(schemaCommandIds).toContain("file.open.open");
    expect(schemaCommandIds).toContain("view.zoom.zoom100");

    const overlap = Array.from(handledRibbonCommandIds).filter((id) => unimplementedRibbonCommandIds.has(id));
    expect(overlap, `Command ids cannot be both handled and unimplemented: ${overlap.join(", ")}`).toEqual([]);

    const unknown = schemaCommandIds.filter((id) => !handledRibbonCommandIds.has(id) && !unimplementedRibbonCommandIds.has(id));
    expect(
      unknown,
      `Found schema command ids with no wiring classification (add to handledRibbonCommandIds or unimplementedRibbonCommandIds):\n${unknown
        .map((id) => `- ${id}`)
        .join("\n")}`,
    ).toEqual([]);

    const extraHandled = Array.from(handledRibbonCommandIds)
      .filter((id) => !schemaIdSet.has(id))
      .sort();
    expect(
      extraHandled,
      `handledRibbonCommandIds contains ids that are not in defaultRibbonSchema:\n${extraHandled.map((id) => `- ${id}`).join("\n")}`,
    ).toEqual([]);

    const extraUnimplemented = Array.from(unimplementedRibbonCommandIds)
      .filter((id) => !schemaIdSet.has(id))
      .sort();
    expect(
      extraUnimplemented,
      `unimplementedRibbonCommandIds contains ids that are not in defaultRibbonSchema:\n${extraUnimplemented
        .map((id) => `- ${id}`)
        .join("\n")}`,
    ).toEqual([]);

    const unionSize = new Set([...handledRibbonCommandIds, ...unimplementedRibbonCommandIds]).size;
    expect(unionSize, "Expected handled+unimplemented allowlists to cover the entire ribbon schema").toBe(schemaIdSet.size);
  });
});
