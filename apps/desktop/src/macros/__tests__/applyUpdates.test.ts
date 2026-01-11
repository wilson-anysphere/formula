import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { applyMacroCellUpdates } from "../applyUpdates";
import type { MacroCellUpdate } from "../types";

describe("applyMacroCellUpdates", () => {
  it("writes formulas (canonicalized) and preserves literal values", () => {
    const doc = new DocumentController();

    const updates: MacroCellUpdate[] = [
      // Formula should win over value, and be stored with leading "=".
      { sheetId: "Sheet1", row: 0, col: 0, value: 999, formula: "SUM(1,2)", displayValue: "3" },
      // Value should be treated as a literal even if it starts with "=".
      { sheetId: "Sheet1", row: 1, col: 0, value: "=not a formula", formula: null, displayValue: "=not a formula" },
    ];

    applyMacroCellUpdates(doc, updates);

    const a1 = doc.getCell("Sheet1", { row: 0, col: 0 }) as { value: unknown; formula: string | null };
    expect(a1.formula).toBe("=SUM(1,2)");
    expect(a1.value).toBeNull();

    const a2 = doc.getCell("Sheet1", { row: 1, col: 0 }) as { value: unknown; formula: string | null };
    expect(a2.formula).toBeNull();
    expect(a2.value).toBe("=not a formula");
  });
});

