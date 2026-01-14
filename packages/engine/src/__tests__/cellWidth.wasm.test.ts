import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { formulaWasmNodeEntryUrl } from "../../../../scripts/build-formula-wasm-node.mjs";

import { DocumentController } from "../../../../apps/desktop/src/document/documentController.js";
import { engineHydrateFromDocument } from "../documentControllerSync.ts";

async function loadFormulaWasm() {
  const entry = formulaWasmNodeEntryUrl();
  // wasm-pack `--target nodejs` outputs CommonJS. Under ESM dynamic import, the exports
  // are exposed on `default`.
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore - `@vite-ignore` is required for runtime-defined file URLs.
  const mod = await import(/* @vite-ignore */ entry);
  return (mod as any).default ?? mod;
}

const skipWasmBuild = process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true";
const describeWasm = skipWasmBuild ? describe.skip : describe;

describeWasm("CELL(\"width\") + column width/hidden metadata (wasm)", () => {
  it("applies DocumentController colWidths + hidden cols during engineHydrateFromDocument()", async () => {
    const doc = new DocumentController();

    // Column B: 25 Excel character units => 180px.
    doc.setColWidth("Sheet1", 1, 180);
    // Hidden column C.
    (doc as any).__sheetHiddenCols = { Sheet1: [2] };

    doc.setCellFormula("Sheet1", "A1", '=CELL("width",B1)');
    doc.setCellFormula("Sheet1", "A2", '=CELL("width",C1)');

    expect(doc.getSheetView("Sheet1").colWidths).toEqual({ "1": 180 });

    const wasm = await loadFormulaWasm();
    let wb: any | null = null;

    // Minimal EngineSyncTarget wrapper around WasmWorkbook so we can use the same hydration
    // code path as the desktop app.
    const engine = {
      loadWorkbookFromJson: (json: string) => {
        wb = wasm.WasmWorkbook.fromJson(json);
      },
      setColWidth: (sheet: string, col: number, widthChars: number | null) => {
        wb!.setColWidth(sheet, col, widthChars);
      },
      setColHidden: (sheet: string, col: number, hidden: boolean) => {
        wb!.setColHidden(sheet, col, hidden);
      },
      setCell: (address: string, value: any, sheet?: string) => {
        wb!.setCell(address, value, sheet);
      },
      recalculate: (_sheet?: string) => {
        wb!.recalculate();
        return [];
      },
    };

    await engineHydrateFromDocument(engine as any, doc);

    // Excel-compatible semantics: integer width + 0.1 indicates a custom width.
    expect(wb!.getCell("A1").value).toBeCloseTo(25.1, 6);
    expect(wb!.getCell("A2").value).toBe(0);
  });

  it("imports column widths + hidden metadata from XLSX so CELL(\"width\") matches Excel immediately", async () => {
    const wasm = await loadFormulaWasm();

    const fixturePath = fileURLToPath(new URL("../../../../fixtures/xlsx/basic/row-col-attrs.xlsx", import.meta.url));
    const bytes = new Uint8Array(readFileSync(fixturePath));

    const wb = wasm.WasmWorkbook.fromXlsxBytes(bytes);

    // The fixture has column B width=25 and column C hidden.
    wb.setCell("A1", '=CELL("width",B1)');
    wb.setCell("A2", '=CELL("width",C1)');
    wb.recalculate();

    expect(wb.getCell("A1").value).toBeCloseTo(25.1, 6);
    expect(wb.getCell("A2").value).toBe(0);
  });
});
