import { writeFileSync, unlinkSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import ts from "typescript";

describe("ai-context TS entrypoint", () => {
  it("is importable from TypeScript with full types", () => {
    const testDir = dirname(fileURLToPath(import.meta.url));
    const tmpFile = join(testDir, `.__index-types.${process.pid}.${Date.now()}.ts`);

    // Keep the file next to this test so `../src/index.js` matches real-world usage.
      writeFileSync(
      tmpFile,
      [
        'import { ContextManager, EXCEL_MAX_COLS, EXCEL_MAX_ROWS, RagIndex, classifyText, chunkSheetByRegions, chunkSheetByRegionsWithSchema, extractSheetSchema, extractWorkbookSchema, isLikelyHeaderRow, parseA1Range, summarizeRegion, summarizeSheetSchema, summarizeWorkbookSchema, trimMessagesToBudget, valuesRangeToTsv } from "../src/index.js";',
        'import { headSampleRows, randomSampleRows, stratifiedSampleRows, systematicSampleRows, tailSampleRows } from "../src/index.js";',
        "",
        "EXCEL_MAX_ROWS satisfies number;",
        "EXCEL_MAX_COLS satisfies number;",
        'valuesRangeToTsv([["a"]], { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }, { maxRows: 1 }) satisfies string;',
        "",
        'isLikelyHeaderRow(["Header"], ["Value"]) satisfies boolean;',
        "",
        'const range = parseA1Range("Sheet1!A1:B2");',
        "range.startRow satisfies number;",
        "",
        "const schema = extractSheetSchema({",
        '  name: "Sheet1",',
        "  origin: { row: 10, col: 3 },",
        '  values: [["Header", "Value"], ["A", 1]],',
        "});",
        "schema.dataRegions[0]?.range satisfies string;",
        "",
        "const summary = summarizeSheetSchema(schema);",
        "summary satisfies string;",
        "summarizeSheetSchema(schema, { maxNamedRanges: 1, includeNamedRanges: false });",
        "const regionSummary = summarizeRegion(schema.tables[0]!);",
        "regionSummary satisfies string;",
        "",
        'const chunks = chunkSheetByRegions({ name: "Sheet1", values: [[1]] });',
        'chunks[0]?.metadata.type satisfies "region";',
        "chunks[0]?.metadata.sheetName satisfies string;",
        "chunks[0]?.metadata.regionRange satisfies string;",
        "",
        'const withSchema = chunkSheetByRegionsWithSchema({ name: "Sheet1", origin: { row: 0, col: 0 }, values: [[1]] });',
        "withSchema.schema.dataRegions[0]?.range satisfies string;",
        'withSchema.chunks[0]?.metadata.type satisfies "region";',
        "withSchema.chunks[0]?.metadata.regionRange satisfies string;",
        "",
        'void trimMessagesToBudget({ messages: [{ role: "user", content: "hi" }], maxTokens: 128 });',
        'void trimMessagesToBudget({ messages: [{ role: "user", content: "hi" }], maxTokens: 128, preserveToolCallPairs: true, dropToolMessagesFirst: true });',
        'void trimMessagesToBudget({ messages: [{ role: "user", content: "hi" }], maxTokens: 128, preserveToolCallPairs: false });',
        "",
        "const cm = new ContextManager();",
        'void cm.buildContext({ sheet: { name: "Sheet1", values: [[1]] }, query: "hi" });',
        'void cm.buildContext({ sheet: { name: "Sheet1", values: [[1]] }, query: "hi", samplingStrategy: "systematic" });',
        'void cm.buildContext({ sheet: { name: "Sheet1", values: [[1]] }, query: "hi", samplingStrategy: "tail" });',
        'void cm.buildContext({ sheet: { name: "Sheet1", values: [[1]], tables: [{ name: "T", range: "Sheet1!A1:A1" }] }, query: "hi" });',
        "",
        "// --- Sampling helpers (should be exported + typed) ---",
        "headSampleRows([1, 2, 3], 2)[0] satisfies number;",
        "tailSampleRows([1, 2, 3], 2)[0] satisfies number;",
        "systematicSampleRows([1, 2, 3, 4], 2, { seed: 1 })[0] satisfies number;",
        "randomSampleRows([1, 2, 3, 4], 2, { seed: 1 })[0] satisfies number;",
        'stratifiedSampleRows([{ k: "a" }, { k: "b" }], 1, { getStratum: (r) => r.k, seed: 1 })[0]?.k satisfies string;',
        "",
        "const cmConfigured = new ContextManager({",
        "  maxContextRows: 10,",
        "  maxContextCells: 100,",
        "  maxChunkRows: 5,",
        "  splitRegions: true,",
        "  chunkRowOverlap: 3,",
        "  maxChunksPerRegion: 50,",
        "  sheetRagTopK: 2,",
        "});",
        "void cmConfigured.buildContext({",
        '  sheet: { name: "Sheet1", values: [[1]] },',
        '  query: "hi",',
        "  limits: { maxContextRows: 1, maxContextCells: 1, maxChunkRows: 1, splitRegions: true, chunkRowOverlap: 3, maxChunksPerRegion: 50 },",
        "});",
        "",
        "const workbookSchema = extractWorkbookSchema({",
        '  id: "wb1",',
        '  sheets: [{ name: "Sheet1", cells: [["H1"], [1]] }],',
        '  tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 0 } }],',
        "});",
        "workbookSchema.tables[0]!.rangeA1 satisfies string;",
        "const workbookSummary = summarizeWorkbookSchema(workbookSchema);",
        "workbookSummary satisfies string;",
        "",
        'const dlp = classifyText("test@example.com");',
        'dlp.level satisfies "public" | "sensitive";',
        '"phone_number" satisfies typeof dlp.findings[number];',
        '"api_key" satisfies typeof dlp.findings[number];',
        '"iban" satisfies typeof dlp.findings[number];',
        '"private_key" satisfies typeof dlp.findings[number];',
        "",
        "const ragIndex = new RagIndex();",
        "const { schema: indexedSchema, chunkCount } = await ragIndex.indexSheet(",
        '  { name: "Sheet1", origin: { row: 0, col: 0 }, values: [[1]] },',
        "  { maxChunkRows: 10, splitRegions: true, chunkRowOverlap: 0, maxChunksPerRegion: 10 },",
        ");",
        "indexedSchema.dataRegions[0]?.range satisfies string;",
        "chunkCount satisfies number;",
        "",
        "// @ts-expect-error - valuesRangeToTsv requires maxRows in options.",
        'valuesRangeToTsv([["a"]], { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }, {});',
        "",
        "// @ts-expect-error - parseA1Range expects a string.",
        "parseA1Range(123);",
        "",
        "// @ts-expect-error - extractWorkbookSchema requires an id.",
        "extractWorkbookSchema({ sheets: [] });",
        "",
        "// @ts-expect-error - chunkSheetByRegions requires a 2D matrix.",
        'chunkSheetByRegions({ name: "Sheet1", values: [1] });',
        "",
        "// @ts-expect-error - maxTokens must be a number.",
        'trimMessagesToBudget({ messages: [], maxTokens: "128" });',
        "",
        "// @ts-expect-error - values must be a 2D array.",
        'extractSheetSchema({ name: "Sheet1", values: [1] });',
        "",
        "// @ts-expect-error - classifyText expects a string.",
        "classifyText(1);",
        "",
        "// @ts-expect-error - summarizeSheetSchema expects a SheetSchema.",
        "summarizeSheetSchema({});",
        "",
        "// @ts-expect-error - summarizeRegion expects a TableSchema or DataRegionSchema.",
        "summarizeRegion({});",
        "",
        "// @ts-expect-error - buildContext requires a query string.",
        'cm.buildContext({ sheet: { name: "Sheet1", values: [[1]] } });',
        "",
        "// @ts-expect-error - samplingStrategy must be a supported strategy string.",
        'cm.buildContext({ sheet: { name: "Sheet1", values: [[1]] }, query: "hi", samplingStrategy: "bogus" });',
        "",
      ].join("\n"),
      "utf8",
    );

    try {
      const converted = ts.convertCompilerOptionsFromJson(
        {
          target: "ES2022",
          module: "ESNext",
          moduleResolution: "Bundler",
          lib: ["ES2022", "DOM", "DOM.Iterable"],
          types: ["node"],
          strict: true,
          skipLibCheck: true,
          noEmit: true,
        },
        testDir,
      );

      // If the options are invalid, surface that as a test failure with details.
      if (converted.errors.length > 0) {
        const host: ts.FormatDiagnosticsHost = {
          getCurrentDirectory: ts.sys.getCurrentDirectory,
          getCanonicalFileName: (f) => f,
          getNewLine: () => "\n",
        };
        const message = ts.formatDiagnosticsWithColorAndContext(converted.errors, host);
        throw new Error(message);
      }

      const program = ts.createProgram([tmpFile], converted.options);
      const diagnostics = ts.getPreEmitDiagnostics(program);

      const host: ts.FormatDiagnosticsHost = {
        getCurrentDirectory: ts.sys.getCurrentDirectory,
        getCanonicalFileName: (f) => f,
        getNewLine: () => "\n",
      };

      const formatted = diagnostics.length ? ts.formatDiagnosticsWithColorAndContext(diagnostics, host) : "";
      expect(diagnostics, formatted).toHaveLength(0);
    } finally {
      unlinkSync(tmpFile);
    }
  });
});
