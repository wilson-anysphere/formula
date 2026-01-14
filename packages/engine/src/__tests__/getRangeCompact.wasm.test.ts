import { describe, expect, it } from "vitest";

import { formulaWasmNodeEntryUrl } from "../../../../scripts/build-formula-wasm-node.mjs";

const skipWasmBuild = process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true";
const describeWasm = skipWasmBuild ? describe.skip : describe;

async function loadFormulaWasm() {
  const entry = formulaWasmNodeEntryUrl();
  // wasm-pack `--target nodejs` outputs CommonJS. Under ESM dynamic import, the exports
  // are exposed on `default`.
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore - `@vite-ignore` is required for runtime-defined file URLs.
  try {
    const mod = await import(/* @vite-ignore */ entry);
    return (mod as any).default ?? mod;
  } catch (err) {
    throw new Error(
      `Failed to import formula-wasm Node build (${entry}). ` +
        `Run \`node scripts/build-formula-wasm-node.mjs\` (or rerun vitest without FORMULA_SKIP_WASM_BUILD).\n\n` +
        `Original error: ${err instanceof Error ? err.message : String(err)}`,
    );
  }
}

describeWasm("WasmWorkbook.getRangeCompact", () => {
  it("matches getRange input/value scalars for mixed formulas + literals", async () => {
    const wasm = await loadFormulaWasm();
    const wb = new (wasm as any).WasmWorkbook();

    wb.setCell("A1", 1);
    wb.setCell("B1", "  =A1*2  ");
    wb.setCell("A2", null);
    // Quote prefix forces literal text even if it looks like an error code.
    wb.setCell("B2", "'#FIELD!");
    wb.recalculate();

    const legacy = wb.getRange("A1:B2");
    const compact = wb.getRangeCompact("A1:B2");

    expect(compact).toEqual(
      legacy.map((row: any[]) => row.map((cell: any) => [cell.input ?? null, cell.value ?? null])),
    );

    // Sanity-check: legacy includes redundant sheet/address per cell.
    expect(legacy[0][0]).toMatchObject({ sheet: "Sheet1", address: "A1", input: 1, value: 1 });
    // And compact payload is just the tuple.
    expect(compact[0][0]).toEqual([1, 1]);
  });
});
