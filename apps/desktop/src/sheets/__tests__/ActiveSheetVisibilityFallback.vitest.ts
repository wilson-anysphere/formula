import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { stripComments } from "../../__tests__/sourceTextUtils";

describe("active sheet visibility fallbacks", () => {
  it("prefers adjacent visible sheets when the active sheet is missing from the visible list", () => {
    const mainTsPath = fileURLToPath(new URL("../../main.ts", import.meta.url));
    const main = stripComments(readFileSync(mainTsPath, "utf8"));

    // syncSheetUi path
    const syncMarker = "renderSheetPosition(sheets, nextActiveId);";
    const syncIdx = main.indexOf(syncMarker);
    expect(syncIdx).toBeGreaterThanOrEqual(0);
    expect(main.slice(Math.max(0, syncIdx - 1600), syncIdx)).toContain("pickAdjacentVisibleSheetId(");

    // sheet-store subscription path (keeps grid + sheet switcher consistent)
    const subscriptionMarker = "renderSheetSwitcher(sheets, activeId);";
    const subIdx = main.indexOf(subscriptionMarker);
    expect(subIdx).toBeGreaterThanOrEqual(0);
    expect(main.slice(Math.max(0, subIdx - 1600), subIdx)).toContain("pickAdjacentVisibleSheetId(");
  });
});

