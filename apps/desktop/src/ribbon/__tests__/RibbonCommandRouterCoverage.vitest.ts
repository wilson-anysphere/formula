import { describe, expect, it } from "vitest";

import { defaultRibbonSchema } from "../ribbonSchema";
import { handledRibbonCommandIds, unimplementedRibbonCommandIds } from "../ribbonCommandRouter";

function collectRibbonCommandIdsFromSchema(): Set<string> {
  const ids = new Set<string>();
  for (const tab of defaultRibbonSchema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        ids.add(button.id);
        for (const item of button.menuItems ?? []) {
          ids.add(item.id);
        }
      }
    }
  }
  return ids;
}

describe("ribbon command router coverage", () => {
  it("ensures every schema command id has an explicit handler or reviewed fallback", () => {
    const schemaIds = collectRibbonCommandIdsFromSchema();
    const missing: string[] = [];

    for (const id of schemaIds) {
      if (handledRibbonCommandIds.has(id)) continue;
      if (unimplementedRibbonCommandIds.has(id)) continue;
      missing.push(id);
    }

    missing.sort();
    expect(missing, `Found ribbon schema command ids without handlers: ${missing.join(", ")}`).toEqual([]);
  });

  it("keeps handled and unimplemented allowlists disjoint", () => {
    const overlap: string[] = [];
    for (const id of handledRibbonCommandIds) {
      if (unimplementedRibbonCommandIds.has(id)) overlap.push(id);
    }
    overlap.sort();
    expect(overlap, `Command ids should not be both handled and unimplemented: ${overlap.join(", ")}`).toEqual([]);
  });
});

