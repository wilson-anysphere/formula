import { describe, expect, it } from "vitest";

import { ensureFormulaWasmNodeBuild, formulaWasmNodeEntryUrl } from "../../../../scripts/build-formula-wasm-node.mjs";

const skipWasmBuild = process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true";
const describeWasm = skipWasmBuild ? describe.skip : describe;

async function loadFormulaWasm() {
  // Ensure the nodejs wasm-pack build exists when running this test in isolation.
  ensureFormulaWasmNodeBuild();

  const entry = formulaWasmNodeEntryUrl();
  // wasm-pack `--target nodejs` outputs CommonJS. Under ESM dynamic import, the exports are
  // exposed on `default`.
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore - `@vite-ignore` is required for runtime-defined file URLs.
  const mod = await import(/* @vite-ignore */ entry);
  return (mod as any).default ?? mod;
}

describeWasm("formula-wasm INFO() metadata (wasm)", () => {
  it("surfaces host-provided EngineInfo + origin overrides via INFO()", async () => {
    const wasm = await loadFormulaWasm();
    const wb = new wasm.WasmWorkbook();

    wb.setCell("A1", '=INFO("system")', "Sheet1");
    wb.setCell("A2", '=INFO("osversion")', "Sheet1");
    wb.setCell("A3", '=INFO("release")', "Sheet1");
    wb.setCell("A4", '=INFO("version")', "Sheet1");
    wb.setCell("A5", '=INFO("memavail")', "Sheet1");
    wb.setCell("A6", '=INFO("totmem")', "Sheet1");
    wb.setCell("A7", '=INFO("directory")', "Sheet1");
    wb.setCell("A8", '=INFO("origin")', "Sheet1");
    wb.setCell("A1", '=INFO("origin")', "Sheet2");

    wb.setEngineInfo({
      system: "mac",
      osversion: "14.0",
      release: "sonoma",
      version: "1.2.3",
      memavail: 123.5,
      totmem: 456.25,
      directory: "/tmp",
    });
    // `INFO("origin")` returns the top-left visible cell for the sheet. The engine is deterministic,
    // so hosts provide this view state explicitly.
    wb.setSheetOrigin("Sheet1", "B2");

    wb.recalculate();

    expect(wb.getCell("A1", "Sheet1").value).toBe("mac");
    expect(wb.getCell("A2", "Sheet1").value).toBe("14.0");
    expect(wb.getCell("A3", "Sheet1").value).toBe("sonoma");
    expect(wb.getCell("A4", "Sheet1").value).toBe("1.2.3");
    expect(wb.getCell("A5", "Sheet1").value).toBe(123.5);
    expect(wb.getCell("A6", "Sheet1").value).toBe(456.25);
    // Excel-compatible directory results include a trailing path separator.
    expect(wb.getCell("A7", "Sheet1").value).toBe("/tmp/");
    expect(wb.getCell("A8", "Sheet1").value).toBe("$B$2");

    // Sheet2 uses the default origin (`$A$1`) when unset.
    expect(wb.getCell("A1", "Sheet2").value).toBe("$A$1");
  });

  it("treats empty strings as unset and rejects non-finite mem numbers", async () => {
    const wasm = await loadFormulaWasm();
    const wb = new wasm.WasmWorkbook();

    wb.setCell("A1", '=INFO("origin")', "Sheet1");
    wb.setCell("A2", '=INFO("memavail")', "Sheet1");
    wb.setCell("A3", '=INFO("system")', "Sheet1");

    wb.setSheetOrigin("Sheet1", "B2");
    wb.setEngineInfo({ memavail: 10, system: "mac" });
    wb.recalculate();
    expect(wb.getCell("A1", "Sheet1").value).toBe("$B$2");
    expect(wb.getCell("A2", "Sheet1").value).toBe(10);
    expect(wb.getCell("A3", "Sheet1").value).toBe("mac");

    // Empty strings clear, falling back to `$A$1`.
    wb.setSheetOrigin("Sheet1", "");
    wb.recalculate();
    expect(wb.getCell("A1", "Sheet1").value).toBe("$A$1");

    wb.setEngineInfo({ system: "" });
    wb.recalculate();
    expect(wb.getCell("A3", "Sheet1").value).toBe("pcdos");

    expect(() => wb.setEngineInfo({ memavail: Number.POSITIVE_INFINITY })).toThrow(/memavail/i);
  });
});
