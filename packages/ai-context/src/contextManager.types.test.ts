import { rm, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import * as ts from "typescript";

import { ContextManager, type SheetSchema } from "./index.js";

describe("ContextManager types", () => {
  it("typechecks the public ContextManager declaration surface (no `any` leakage)", async () => {
    // Vitest/Vite does not typecheck TS by default (it uses esbuild). Run a targeted
    // TypeScript compilation of a small snippet that asserts the expected types.
    const entryFile = fileURLToPath(new URL(`./.__contextManager-typecheck.${process.pid}.${Date.now()}.ts`, import.meta.url));

    const source = `\
import {
  ContextManager,
  RagIndex,
  type BuildContextResult,
  type BuildWorkbookContextResult,
  type ContextSheet,
  type DlpClassificationRecord,
  type DlpOptions,
  type DlpOptionsInput,
  type RetrievedSheetChunk,
  type RetrievedWorkbookChunk,
  type WorkbookChunkMetadata,
  type WorkbookIndexStats,
  type WorkbookRagTable,
  type WorkbookRagNamedRange,
  type SheetNameResolverLike,
  type SpreadsheetApiLike,
  type SpreadsheetApiWithNonEmptyCells,
  type WorkbookRagVectorStore,
  type WorkbookRagWorkbook,
} from "./index.js";
import type { SheetSchema } from "./schema.js";

type IsAny<T> = 0 extends (1 & T) ? true : false;
type Assert<T extends true> = T;

type _SchemaIsSheetSchema = Assert<BuildContextResult["schema"] extends SheetSchema ? true : false>;
type _SchemaNotAny = Assert<IsAny<BuildContextResult["schema"]> extends false ? true : false>;
type _RetrievedNotAny = Assert<IsAny<BuildContextResult["retrieved"][number]> extends false ? true : false>;
type _RetrievedShape = Assert<BuildContextResult["retrieved"][number] extends RetrievedSheetChunk ? true : false>;

type _RagIndexPropNotAny = Assert<IsAny<ContextManager["ragIndex"]> extends false ? true : false>;
type _RagIndexPropShape = Assert<ContextManager["ragIndex"] extends RagIndex ? true : false>;

type _WorkbookRetrievedNotAny = Assert<IsAny<BuildWorkbookContextResult["retrieved"][number]> extends false ? true : false>;
type _WorkbookRetrievedShape = Assert<
  BuildWorkbookContextResult["retrieved"][number] extends RetrievedWorkbookChunk ? true : false
>;
type _WorkbookChunkMetadataNotAny = Assert<IsAny<WorkbookChunkMetadata> extends false ? true : false>;
type _WorkbookChunkMetadataTextIsUndefined = Assert<WorkbookChunkMetadata["text"] extends undefined ? true : false>;
type _WorkbookIndexStatsNotAny = Assert<IsAny<BuildWorkbookContextResult["indexStats"]> extends false ? true : false>;
type _WorkbookIndexStatsShape = Assert<
  BuildWorkbookContextResult["indexStats"] extends WorkbookIndexStats | null ? true : false
>;

type _ContextSheetNotAny = Assert<IsAny<ContextSheet> extends false ? true : false>;
type _DlpClassificationRecordNotAny = Assert<IsAny<DlpClassificationRecord> extends false ? true : false>;
type _DlpOptionsNotAny = Assert<IsAny<DlpOptions> extends false ? true : false>;
type _DlpOptionsInputNotAny = Assert<IsAny<DlpOptionsInput> extends false ? true : false>;
type _SheetNameResolverNotAny = Assert<IsAny<SheetNameResolverLike> extends false ? true : false>;
type _SpreadsheetNotAny = Assert<IsAny<SpreadsheetApiLike> extends false ? true : false>;
type _SpreadsheetWithCellsNotAny = Assert<IsAny<SpreadsheetApiWithNonEmptyCells> extends false ? true : false>;
type _WorkbookRagTableNotAny = Assert<IsAny<WorkbookRagTable> extends false ? true : false>;
type _WorkbookRagNamedRangeNotAny = Assert<IsAny<WorkbookRagNamedRange> extends false ? true : false>;
type _VectorStoreNotAny = Assert<IsAny<WorkbookRagVectorStore> extends false ? true : false>;
type _WorkbookNotAny = Assert<IsAny<WorkbookRagWorkbook> extends false ? true : false>;
type _VectorStoreListContentHashesShape = Assert<
  WorkbookRagVectorStore["listContentHashes"] extends
    | ((opts?: { workbookId?: string; signal?: AbortSignal }) => Promise<
        Array<{ id: string; contentHash: string | null; metadataHash: string | null }>
      >)
    | undefined
    ? true
    : false
>;
type _ClearCacheOptionsNotAny = Assert<
  IsAny<Parameters<ContextManager["clearSheetIndexCache"]>[0]> extends false ? true : false
>;

 const cm = new ContextManager({
   // Ensure the single-sheet cache knobs are part of the public surface.
   cacheSheetIndex: true,
   sheetIndexCacheLimit: 32,
   // Ensure the wide-sheet safety cap is part of the public surface.
   maxContextCols: 500,
   workbookRag: {
     vectorStore: { query: async () => [] },
     embedder: { embedTexts: async () => [new Float32Array(1)] },
   },
 });

const result = await cm.buildContext({
  sheet: {
    name: "Sheet1",
    values: [[1]],
    tables: [{ name: "T", range: "A1:B2", id: "tbl-1" }],
    namedRanges: [{ name: "NR", range: "A1:A1", id: "nr-1" }],
  },
  query: "hi",
   limits: { maxContextCols: 10 },
});

// If \`schema\` were \`any\`, this assignment would still typecheck. The \`IsAny\`
// assertions above ensure it is not \`any\`.
const schema: SheetSchema = result.schema;
void schema;

const dlp: DlpOptions = { documentId: "doc-1", policy: {} };
void dlp;

// DlpOptionsInput requires a document id (camelCase or snake_case).
// @ts-expect-error - documentId/document_id is required when using DlpOptionsInput.
const _dlpMissingId: DlpOptionsInput = { policy: {} };
const _dlpCamel: DlpOptionsInput = { documentId: "doc-1", policy: {} };
const _dlpSnake: DlpOptionsInput = { document_id: "doc-1", policy: {} };
void _dlpCamel;
void _dlpSnake;

// auditLogger can be sync or async.
const _dlpAsyncLogger: DlpOptionsInput = {
  documentId: "doc-1",
  policy: {},
  auditLogger: { log: async () => {} },
};
void _dlpAsyncLogger;
const _dlpNullLogger: DlpOptionsInput = { documentId: "doc-1", policy: {}, auditLogger: null };
void _dlpNullLogger;

// Workbook RAG types are also checked (retrieval-only path).
await cm.buildWorkbookContext({
  workbook: { id: "wb-1", sheets: [] },
  query: "hi",
  skipIndexing: true,
});

// SpreadsheetApi wrapper types are checked too (cheap path).
await cm.buildWorkbookContextFromSpreadsheetApi({
  spreadsheet: { listSheets: () => [] },
  workbookId: "wb-1",
  query: "hi",
  skipIndexing: true,
});

// DLP + skipIndexing without skipIndexingWithDlp still requires listNonEmptyCells (indexing must run).
// @ts-expect-error - spreadsheet.listNonEmptyCells is required when DLP indexing cannot be skipped.
await cm.buildWorkbookContextFromSpreadsheetApi({
  spreadsheet: { listSheets: () => [] },
  workbookId: "wb-1",
  query: "hi",
  skipIndexing: true,
  dlp: { documentId: "doc-1", policy: {} },
});

await cm.buildWorkbookContextFromSpreadsheetApi({
  spreadsheet: { listSheets: () => [], listNonEmptyCells: () => [] },
  workbookId: "wb-1",
  query: "hi",
  skipIndexing: true,
  dlp: { documentId: "doc-1", policy: {} },
});

// DLP-safe cheap path: caller asserts the workbook is already indexed with DLP applied.
await cm.buildWorkbookContextFromSpreadsheetApi({
  spreadsheet: { listSheets: () => [] },
  workbookId: "wb-1",
  query: "hi",
  skipIndexing: true,
  skipIndexingWithDlp: true,
  dlp: { documentId: "doc-1", policy: {} },
});

// Default path requires listNonEmptyCells.
// @ts-expect-error - spreadsheet.listNonEmptyCells is required when skipIndexing is not true.
await cm.buildWorkbookContextFromSpreadsheetApi({
  spreadsheet: { listSheets: () => [] },
  workbookId: "wb-1",
  query: "hi",
});

// Explicit cache clearing API should be available and typed.
await cm.clearSheetIndexCache({ clearStore: true });
`;

    await writeFile(entryFile, source, "utf8");

    try {
      const converted = ts.convertCompilerOptionsFromJson(
        {
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

      expect(diagnostics).toHaveLength(0);
    } finally {
      await rm(entryFile, { force: true });
    }
  });

  it("buildContext().schema is a SheetSchema at runtime (and not `any` at compile time)", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 10_000,
      // Avoid redaction so prompt strings are stable for snapshots/debugging.
      redactor: (text: string) => text,
    });

    const result = await cm.buildContext({
      sheet: { name: "Sheet1", values: [["A"], ["B"]] },
      query: "A",
    });

    // Runtime sanity check.
    expect(result.schema.name).toBe("Sheet1");

    // Compile-time check (reinforces the intent of this test file).
    const _schema: SheetSchema = result.schema;
    expect(_schema.tables).toBeDefined();
  });
});
