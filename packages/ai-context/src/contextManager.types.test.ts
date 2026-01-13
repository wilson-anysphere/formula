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
  type BuildContextResult,
  type BuildWorkbookContextResult,
  type ContextSheet,
  type DlpOptions,
  type RetrievedSheetChunk,
  type RetrievedWorkbookChunk,
  type WorkbookIndexStats,
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

type _WorkbookRetrievedNotAny = Assert<IsAny<BuildWorkbookContextResult["retrieved"][number]> extends false ? true : false>;
type _WorkbookRetrievedShape = Assert<
  BuildWorkbookContextResult["retrieved"][number] extends RetrievedWorkbookChunk ? true : false
>;
type _WorkbookIndexStatsNotAny = Assert<IsAny<BuildWorkbookContextResult["indexStats"]> extends false ? true : false>;
type _WorkbookIndexStatsShape = Assert<
  BuildWorkbookContextResult["indexStats"] extends WorkbookIndexStats | null ? true : false
>;

type _ContextSheetNotAny = Assert<IsAny<ContextSheet> extends false ? true : false>;
type _DlpOptionsNotAny = Assert<IsAny<DlpOptions> extends false ? true : false>;
type _SheetNameResolverNotAny = Assert<IsAny<SheetNameResolverLike> extends false ? true : false>;
type _SpreadsheetNotAny = Assert<IsAny<SpreadsheetApiLike> extends false ? true : false>;
type _SpreadsheetWithCellsNotAny = Assert<IsAny<SpreadsheetApiWithNonEmptyCells> extends false ? true : false>;
type _VectorStoreNotAny = Assert<IsAny<WorkbookRagVectorStore> extends false ? true : false>;
type _WorkbookNotAny = Assert<IsAny<WorkbookRagWorkbook> extends false ? true : false>;

const cm = new ContextManager({
  workbookRag: {
    vectorStore: { query: async () => [] },
    embedder: { embedTexts: async () => [new Float32Array(1)] },
  },
});

const result = await cm.buildContext({
  sheet: { name: "Sheet1", values: [[1]] },
  query: "hi",
});

// If \`schema\` were \`any\`, this assignment would still typecheck. The \`IsAny\`
// assertions above ensure it is not \`any\`.
const schema: SheetSchema = result.schema;
void schema;

const dlp: DlpOptions = { documentId: "doc-1", policy: {} };
void dlp;

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
