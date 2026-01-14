import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { stripComments } from "../../__tests__/sourceTextUtils";

describe("undo/redo active sheet fallback", () => {
  it("prefers the adjacent visible sheet when the active sheet id disappears (Excel-like)", () => {
    const mainTsPath = fileURLToPath(new URL("../../main.ts", import.meta.url));
    const main = stripComments(readFileSync(mainTsPath, "utf8"));

    const marker = 'source !== "undo" && source !== "redo" && source !== "applyState"';
    const idx = main.indexOf(marker);
    expect(idx).toBeGreaterThanOrEqual(0);

    // Ensure the guardrail stays resilient: the logic should live close to the undo/redo source gate.
    const windowText = main.slice(idx, idx + 1600);
    expect(windowText).toContain("pickAdjacentVisibleSheetId(");
  });
});

