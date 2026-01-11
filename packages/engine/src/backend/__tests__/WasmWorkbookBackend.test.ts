import { describe, expect, it, vi } from "vitest";

import type { EngineClient } from "../../client";
import { colToName, toA1, toA1Range } from "../a1";
import { normalizeFormulaText } from "../formula";
import { WasmWorkbookBackend } from "../WasmWorkbookBackend";

describe("A1 helpers", () => {
  it("converts 0-based columns to Excel column names", () => {
    expect(colToName(0)).toBe("A");
    expect(colToName(25)).toBe("Z");
    expect(colToName(26)).toBe("AA");
    expect(colToName(27)).toBe("AB");
    expect(colToName(51)).toBe("AZ");
    expect(colToName(52)).toBe("BA");
    expect(colToName(701)).toBe("ZZ");
    expect(colToName(702)).toBe("AAA");
  });

  it("converts 0-based row/col coords to A1 addresses", () => {
    expect(toA1(0, 0)).toBe("A1");
    expect(toA1(1, 0)).toBe("A2");
    expect(toA1(0, 1)).toBe("B1");
    expect(toA1(9, 25)).toBe("Z10");
  });

  it("formats an A1 range (collapsing to a single cell when needed)", () => {
    expect(toA1Range(0, 0, 0, 0)).toBe("A1");
    expect(toA1Range(0, 0, 1, 1)).toBe("A1:B2");
  });
});

describe("formula normalization", () => {
  it("ensures formulas start with '=' and strips leading whitespace", () => {
    expect(normalizeFormulaText("A1*2")).toBe("=A1*2");
    expect(normalizeFormulaText(" =A1*2")).toBe("=A1*2");
    expect(normalizeFormulaText("=A1*2")).toBe("=A1*2");
  });
});

describe("WasmWorkbookBackend", () => {
  it("translates setRange row/col rectangles into engine A1 range calls (with formula normalization)", async () => {
    const engine: EngineClient = {
      init: vi.fn(async () => {}),
      newWorkbook: vi.fn(async () => {}),
      loadWorkbookFromJson: vi.fn(async () => {}),
      toJson: vi.fn(async () => "{}"),
      getCell: vi.fn(async () => ({ sheet: "Sheet1", address: "A1", input: null, value: null })),
      getRange: vi.fn(async () => []),
      setCell: vi.fn(async () => {}),
      setCells: vi.fn(async () => {}),
      setRange: vi.fn(async () => {}),
      recalculate: vi.fn(async () => []),
      terminate: vi.fn(),
    };

    const backend = new WasmWorkbookBackend(engine);

    await backend.setRange({
      sheetId: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
      values: [
        [
          { value: 1, formula: null },
          { value: 123, formula: " A1*2" },
        ],
        [
          { value: true, formula: null },
          { value: { text: "Hello", runs: [] }, formula: null },
        ],
      ],
    });

    expect(engine.setRange).toHaveBeenCalledTimes(1);
    expect(engine.setRange).toHaveBeenCalledWith(
      "A1:B2",
      [
        [1, "=A1*2"],
        [true, "Hello"],
      ],
      "Sheet1",
    );

    expect(engine.recalculate).toHaveBeenCalledTimes(1);
    expect(engine.recalculate).toHaveBeenCalledWith("Sheet1");
  });
});

