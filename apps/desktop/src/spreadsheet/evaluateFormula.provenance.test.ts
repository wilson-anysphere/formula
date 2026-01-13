import { describe, expect, it } from "vitest";

import { evaluateFormula } from "./evaluateFormula.js";

describe("evaluateFormula (AI provenance)", () => {
  it("wraps direct cell/range references passed to AI functions with __cellRef metadata", () => {
    const calls: any[] = [];
    const ai = {
      evaluateAiFunction: (params: any) => {
        calls.push(params);
        return "ok";
      },
    };

    const getCellValue = (addr: string) => {
      if (addr === "A1") return "hello";
      if (addr === "A2") return "world";
      return null;
    };

    const result = evaluateFormula('=AI("summarize", A1:A2)', getCellValue, { ai, cellAddress: "Sheet1!B1" });
    expect(result).toBe("ok");
    expect(calls).toHaveLength(1);

    const args = calls[0]?.args;
    expect(args[0]).toBe("summarize");
    expect(args[1]).toHaveLength(2);
    expect(args[1][0]).toEqual({ __cellRef: "Sheet1!A1", value: "hello" });
    expect(args[1][1]).toEqual({ __cellRef: "Sheet1!A2", value: "world" });
    expect((args[1] as any).__rangeRef).toBe("Sheet1!A1:A2");
    expect((args[1] as any).__totalCells).toBe(2);
  });

  it("preserves provenance on derived values inside AI arguments (e.g. SUM)", () => {
    const calls: any[] = [];
    const ai = {
      evaluateAiFunction: (params: any) => {
        calls.push(params);
        return "ok";
      },
    };

    const getCellValue = (addr: string) => (addr === "A1" ? 1 : addr === "A2" ? 2 : null);

    // Nested SUM should evaluate normally inside AI arguments, while still preserving provenance for DLP enforcement.
    const result = evaluateFormula('=AI("sum", SUM(A1:A2))', getCellValue, { ai, cellAddress: "Sheet1!B1" });
    expect(result).toBe("ok");
    expect(calls).toHaveLength(1);
    expect(calls[0]?.args?.[1]).toEqual({ __cellRef: "Sheet1!A1:A2", value: 3 });
  });

  it("samples large range references passed directly to AI functions", () => {
    const calls: any[] = [];
    const ai = {
      evaluateAiFunction: (params: any) => {
        calls.push(params);
        return "ok";
      },
    };

    let readCount = 0;
    const readAddrs: string[] = [];
    const getCellValue = (addr: string) => {
      readCount += 1;
      readAddrs.push(addr);
      return 1;
    };

    const result = evaluateFormula('=AI("summarize", A1:A500)', getCellValue, { ai, cellAddress: "Sheet1!B1" });
    expect(result).toBe("ok");
    expect(readCount).toBe(200);
    // Ensure sampling isn't just the first N cells: we expect to read from deeper into the range.
    const maxRow = Math.max(
      ...readAddrs.map((addr) => {
        const match = /(\d+)$/.exec(addr);
        return match ? Number(match[1]) : 0;
      }),
    );
    expect(maxRow).toBeGreaterThan(200);

    const rangeArg = calls[0]?.args?.[1];
    expect(Array.isArray(rangeArg)).toBe(true);
    expect(rangeArg).toHaveLength(200);
    expect((rangeArg as any).__rangeRef).toBe("Sheet1!A1:A500");
    expect((rangeArg as any).__totalCells).toBe(500);
  });

  it("samples direct AI range arguments deterministically (same inputs -> same sampled addresses)", () => {
    const ai = { evaluateAiFunction: () => "ok" };

    const evalOnce = () => {
      const readAddrs: string[] = [];
      const getCellValue = (addr: string) => {
        readAddrs.push(addr);
        return 1;
      };
      const result = evaluateFormula('=AI("summarize", A1:A500)', getCellValue, { ai, cellAddress: "Sheet1!B1" });
      expect(result).toBe("ok");
      expect(readAddrs).toHaveLength(200);
      return readAddrs;
    };

    const first = evalOnce();
    const second = evalOnce();
    expect(second).toEqual(first);
  });

  it("uses the AI evaluator's rangeSampleLimit when provided", () => {
    const calls: any[] = [];
    const ai = {
      rangeSampleLimit: 10,
      evaluateAiFunction: (params: any) => {
        calls.push(params);
        return "ok";
      },
    };

    let readCount = 0;
    const getCellValue = (_addr: string) => {
      readCount += 1;
      return 1;
    };

    const result = evaluateFormula('=AI("summarize", A1:A50)', getCellValue, { ai, cellAddress: "Sheet1!B1" });
    expect(result).toBe("ok");
    expect(readCount).toBe(10);
    expect(calls[0]?.args?.[1]).toHaveLength(10);
    expect((calls[0]?.args?.[1] as any).__totalCells).toBe(50);
  });

  it("does not sample ranges inside nested non-AI functions (e.g. SUM) within AI arguments", () => {
    const calls: any[] = [];
    const ai = {
      evaluateAiFunction: (params: any) => {
        calls.push(params);
        return "ok";
      },
    };

    let readCount = 0;
    const getCellValue = (_addr: string) => {
      readCount += 1;
      return 1;
    };

    const result = evaluateFormula('=AI("sum", SUM(A1:A500))', getCellValue, { ai, cellAddress: "Sheet1!B1" });
    expect(result).toBe("ok");
    expect(readCount).toBe(500);
    expect(calls[0]?.args?.[1]).toEqual({ __cellRef: "Sheet1!A1:A500", value: 500 });
  });

  it("preserves provenance through nested expressions inside AI arguments (e.g. IF)", () => {
    const calls: any[] = [];
    const ai = {
      evaluateAiFunction: (params: any) => {
        calls.push(params);
        return "ok";
      },
    };

    const getCellValue = (addr: string) => (addr === "A1" ? "secret" : null);

    const result = evaluateFormula('=AI("summarize", IF(TRUE, A1, "x"))', getCellValue, { ai, cellAddress: "Sheet1!B1" });
    expect(result).toBe("ok");
    expect(calls).toHaveLength(1);
    expect(calls[0]?.args?.[1]).toEqual({ __cellRef: "Sheet1!A1", value: "secret" });
  });

  it("taints conditional results derived from referenced cells (e.g. IF(A1,\"Y\",\"N\"))", () => {
    const calls: any[] = [];
    const ai = {
      evaluateAiFunction: (params: any) => {
        calls.push(params);
        return "ok";
      },
    };

    const getCellValue = (addr: string) => (addr === "A1" ? 1 : null);

    const result = evaluateFormula('=AI("summarize", IF(A1, "Y", "N"))', getCellValue, { ai, cellAddress: "Sheet1!B1" });
    expect(result).toBe("ok");
    expect(calls).toHaveLength(1);
    expect(calls[0]?.args?.[1]).toEqual({ __cellRef: "Sheet1!A1", value: "Y" });
  });
});
