import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { stripComments } from "../../__tests__/sourceTextUtils";

describe("extension spreadsheetApi deleteSheet activation", () => {
  it("uses Excel-like adjacent sheet activation when deleting the active sheet", () => {
    const mainTsPath = fileURLToPath(new URL("../../main.ts", import.meta.url));
    const main = stripComments(readFileSync(mainTsPath, "utf8"));

    const match = main.match(/\basync\s+deleteSheet\s*\(/);
    expect(match).not.toBeNull();
    const idx = match?.index ?? -1;
    expect(idx).toBeGreaterThanOrEqual(0);

    // Keep the scan window bounded so the assertion doesn't pass due to another callsite elsewhere
    // in main.ts (there should only be one, but guard against future refactors).
    const block = main.slice(idx, idx + 1400);
    expect(block).toMatch(/\bpickAdjacentVisibleSheetId\s*\(/);
  });
});

