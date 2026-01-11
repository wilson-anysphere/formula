import { describe, expect, it, vi } from "vitest";

import { evaluateFormula, type AiFunctionEvaluator } from "./evaluateFormula.js";

describe("evaluateFormula AI provenance", () => {
  it("captures referenced cells and ranges per AI argument", () => {
    const ai: AiFunctionEvaluator = {
      evaluateAiFunction: vi.fn(() => "ok"),
    };

    const getCellValue = (addr: string) => {
      if (addr === "A1") return "hello";
      if (addr === "B2") return "b2";
      if (addr === "C2") return "c2";
      if (addr === "B3") return "b3";
      if (addr === "C3") return "c3";
      return null;
    };

    const value = evaluateFormula('=AI("summarize", A1, B2:C3)', getCellValue, { ai, cellAddress: "Sheet1!D1" });
    expect(value).toBe("ok");
    expect(ai.evaluateAiFunction).toHaveBeenCalledTimes(1);

    const call = (ai.evaluateAiFunction as any).mock.calls[0]?.[0];
    expect(call?.name).toBe("AI");
    expect(call?.cellAddress).toBe("Sheet1!D1");
    expect(call?.argProvenance).toHaveLength(3);

    const provenance = call.argProvenance;
    expect(provenance[0]).toEqual({ cells: [], ranges: [] }); // "summarize" literal
    expect(provenance[1]?.cells).toEqual(["Sheet1!A1"]);
    expect(provenance[1]?.ranges).toEqual([]);
    expect(provenance[2]?.cells).toEqual([]);
    expect(provenance[2]?.ranges).toEqual(["Sheet1!B2:C3"]);
  });
});
