import { rm, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import * as ts from "typescript";
import { expect, test } from "vitest";

test("index.js is fully typed for TS consumers", async () => {
  // Vitest/Vite does not typecheck TypeScript by default (it uses esbuild). Run a
  // targeted TypeScript compilation of a small snippet that imports from the
  // public entrypoint and asserts the expected types.
  const entryFile = fileURLToPath(new URL("./.ai-context-index-typecheck.ts", import.meta.url));

  const source = `\
 import {
   EXCEL_MAX_COLS,
   EXCEL_MAX_ROWS,
   ContextManager,
   classifyText,
   extractWorkbookSchema,
   summarizeWorkbookSchema,
   parseA1Range,
   RagIndex,
   isLikelyHeaderRow,
   headSampleRows,
   tailSampleRows,
   systematicSampleRows,
   randomSampleRows,
   stratifiedSampleRows,
   scoreRegionForQuery,
   pickBestRegionForQuery,
   type RegionType,
   type RegionRef,
   type SystematicSamplingOptions,
 } from "../src/index.js";
 import type { SheetSchema } from "../src/schema.js";
 
type IsAny<T> = 0 extends (1 & T) ? true : false;
type Assert<T extends true> = T;
 
// --- Constants ---
type _ExcelMaxRows_NotAny = Assert<IsAny<typeof EXCEL_MAX_ROWS> extends false ? true : false>;
type _ExcelMaxCols_NotAny = Assert<IsAny<typeof EXCEL_MAX_COLS> extends false ? true : false>;
type _ExcelMaxRows_IsNumber = Assert<typeof EXCEL_MAX_ROWS extends number ? true : false>;
type _ExcelMaxCols_IsNumber = Assert<typeof EXCEL_MAX_COLS extends number ? true : false>;
  
// --- A1 parsing ---
type ParsedRange = ReturnType<typeof parseA1Range>;
type _ParseA1Range_NotAny = Assert<IsAny<ParsedRange> extends false ? true : false>;
type _ParseA1Range_Shape = Assert<
  ParsedRange extends { sheetName?: string; startRow: number; startCol: number; endRow: number; endCol: number } ? true : false
>;

// --- DLP ---
type DlpResult = ReturnType<typeof classifyText>;
type _Dlp_NotAny = Assert<IsAny<DlpResult> extends false ? true : false>;
type _Dlp_Shape = Assert<DlpResult extends { level: "public" | "sensitive"; findings: Array<string> } ? true : false>;
type DlpFinding = DlpResult["findings"][number];
// Ensure we preserve the discriminated literal union rather than widening to string.
type _DlpFinding_NotString = Assert<string extends DlpFinding ? false : true>;
type _DlpFinding_HasPhone = Assert<"phone_number" extends DlpFinding ? true : false>;
type _DlpFinding_HasApiKey = Assert<"api_key" extends DlpFinding ? true : false>;
type _DlpFinding_HasIban = Assert<"iban" extends DlpFinding ? true : false>;
type _DlpFinding_HasPrivateKey = Assert<"private_key" extends DlpFinding ? true : false>;

// --- Header heuristics ---
type _HeaderRow_NotAny = Assert<IsAny<ReturnType<typeof isLikelyHeaderRow>> extends false ? true : false>;
type _HeaderRow_Return = Assert<ReturnType<typeof isLikelyHeaderRow> extends boolean ? true : false>;
// --- RAG indexing ---
type IndexSheetResult = Awaited<ReturnType<RagIndex["indexSheet"]>>;
type _IndexSheet_NotAny = Assert<IsAny<IndexSheetResult> extends false ? true : false>;
type _IndexSheet_Shape = Assert<IndexSheetResult extends { schema: SheetSchema; chunkCount: number } ? true : false>;
type _IndexSheet_Schema_NotAny = Assert<IsAny<IndexSheetResult["schema"]> extends false ? true : false>;

// --- Query-aware scoring ---
type _RegionType_NotAny = Assert<IsAny<RegionType> extends false ? true : false>;
type _RegionRef_NotAny = Assert<IsAny<RegionRef> extends false ? true : false>;
type _Score_ReturnType = Assert<ReturnType<typeof scoreRegionForQuery> extends number ? true : false>;
type PickedRegion = ReturnType<typeof pickBestRegionForQuery>;
type _PickBest_NotAny = Assert<IsAny<PickedRegion> extends false ? true : false>;
type _PickBest_Shape = Assert<PickedRegion extends { type: RegionType; index: number; range: string } | null ? true : false>;

// --- Sampling helpers ---
type HeadSampled = ReturnType<typeof headSampleRows<number>>;
type _HeadSample_NotAny = Assert<IsAny<HeadSampled> extends false ? true : false>;
type _HeadSample_Shape = Assert<HeadSampled extends number[] ? true : false>;
type _TailSample_Shape = Assert<ReturnType<typeof tailSampleRows<number>> extends number[] ? true : false>;
type _SystematicOpts_NotAny = Assert<IsAny<SystematicSamplingOptions> extends false ? true : false>;
type _SystematicSample_Shape = Assert<ReturnType<typeof systematicSampleRows<number>> extends number[] ? true : false>;
type _RandomSample_Shape = Assert<ReturnType<typeof randomSampleRows<number>> extends number[] ? true : false>;
type _StratifiedSample_Shape = Assert<
  ReturnType<typeof stratifiedSampleRows<{ k: string }>> extends Array<{ k: string }> ? true : false
>;

// --- Workbook schema extraction ---
type WorkbookSchema = ReturnType<typeof extractWorkbookSchema>;
type _WorkbookSchema_NotAny = Assert<IsAny<WorkbookSchema> extends false ? true : false>;
 type _WorkbookSchema_Shape = Assert<
   WorkbookSchema extends { id: string; sheets: Array<{ name: string }>; tables: unknown[]; namedRanges: unknown[] } ? true : false
 >;
 type _SummarizeWorkbook_Return = Assert<ReturnType<typeof summarizeWorkbookSchema> extends string ? true : false>;

// Basic runtime sanity checks (also ensures the compiler doesn't tree-shake the imports).
void EXCEL_MAX_ROWS;
void EXCEL_MAX_COLS;
const parsed = parseA1Range("$A$1:B2");
 const dlp = classifyText("test@example.com");
 void dlp;
 const index = new RagIndex();
 void index;
 void headSampleRows([1, 2, 3], 2);
 void tailSampleRows([1, 2, 3], 2);
 void systematicSampleRows([1, 2, 3, 4], 2, { seed: 1 });
 void randomSampleRows([1, 2, 3, 4], 2, { seed: 1 });
 void stratifiedSampleRows([{ k: "a" }, { k: "b" }], 1, { getStratum: (r) => r.k, seed: 1 });
 const cm = new ContextManager();
 void cm.buildContext({ sheet: { name: "Sheet1", values: [[1]] }, query: "hi", samplingStrategy: "systematic" });
 // @ts-expect-error - samplingStrategy must be a supported strategy string.
 void cm.buildContext({ sheet: { name: "Sheet1", values: [[1]] }, query: "hi", samplingStrategy: "bogus" });
 const wbSchema = extractWorkbookSchema({
   id: "wb1",
   sheets: [{ name: "Sheet1", cells: [[{ v: "Header" }], [{ v: 1 }]] }],
   tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 0 } }],
 });
 wbSchema.tables[0]?.rangeA1;
 void summarizeWorkbookSchema(wbSchema);
 const schema: SheetSchema = { name: "Sheet1", tables: [], namedRanges: [], dataRegions: [] };
 const ref: RegionRef = { type: "table", index: 0 };
 scoreRegionForQuery(ref, schema, "revenue");
 pickBestRegionForQuery(schema, "revenue");
 void parsed;
`;

  await writeFile(entryFile, source, "utf8");

  try {
    // `ts.createProgram()` expects resolved lib file names (e.g. `lib.es2022.d.ts`),
    // while tsconfig-style options use human-friendly lib names (e.g. `"ES2022"`).
    // Use TypeScript's JSON conversion helper so this test stays resilient across
    // TypeScript versions.
    const converted = ts.convertCompilerOptionsFromJson(
      {
        // Keep the compilation as close to the repo defaults as possible.
        target: "ES2022",
        module: "ESNext",
        moduleResolution: "Bundler",
        lib: ["ES2022", "DOM", "DOM.Iterable"],
        types: ["node"],
        strict: true,
        allowJs: true,
        checkJs: false,
        skipLibCheck: true,
        noEmit: true,
      },
      process.cwd(),
    );

    if (converted.errors.length > 0) {
      const host: ts.FormatDiagnosticsHost = {
        getCanonicalFileName: (fileName) => fileName,
        getCurrentDirectory: () => process.cwd(),
        getNewLine: () => "\n",
      };
      const formatted = ts.formatDiagnosticsWithColorAndContext(converted.errors, host);
      throw new Error(`TypeScript compiler option parsing failed:\n${formatted}`);
    }

    const program = ts.createProgram([entryFile], converted.options);
    const diagnostics = ts.getPreEmitDiagnostics(program);

    if (diagnostics.length > 0) {
      const host: ts.FormatDiagnosticsHost = {
        getCanonicalFileName: (fileName) => fileName,
        getCurrentDirectory: () => process.cwd(),
        getNewLine: () => "\n",
      };
      const formatted = ts.formatDiagnosticsWithColorAndContext(diagnostics, host);
      throw new Error(`TypeScript typecheck failed:\n${formatted}`);
    }

    // If compilation succeeds, the type surface is usable for TS consumers.
    expect(diagnostics).toHaveLength(0);
  } finally {
    await rm(entryFile, { force: true });
  }
});
