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
  parseA1Range,
  RagIndex,
  scoreRegionForQuery,
  pickBestRegionForQuery,
  type RegionType,
  type RegionRef,
} from "../src/index.js";
import type { SheetSchema } from "../src/schema.js";

type IsAny<T> = 0 extends (1 & T) ? true : false;
type Assert<T extends true> = T;

// --- A1 parsing ---
type ParsedRange = ReturnType<typeof parseA1Range>;
type _ParseA1Range_NotAny = Assert<IsAny<ParsedRange> extends false ? true : false>;
type _ParseA1Range_Shape = Assert<
  ParsedRange extends { sheetName?: string; startRow: number; startCol: number; endRow: number; endCol: number } ? true : false
>;

// --- RAG indexing ---
type IndexSheetResult = Awaited<ReturnType<RagIndex["indexSheet"]>>;
type _IndexSheet_NotAny = Assert<IsAny<IndexSheetResult> extends false ? true : false>;
type _IndexSheet_Shape = Assert<IndexSheetResult extends { schema: SheetSchema; chunkCount: number } ? true : false>;

// --- Query-aware scoring ---
type _RegionType_NotAny = Assert<IsAny<RegionType> extends false ? true : false>;
type _RegionRef_NotAny = Assert<IsAny<RegionRef> extends false ? true : false>;
type _Score_ReturnType = Assert<ReturnType<typeof scoreRegionForQuery> extends number ? true : false>;
type PickedRegion = ReturnType<typeof pickBestRegionForQuery>;
type _PickBest_NotAny = Assert<IsAny<PickedRegion> extends false ? true : false>;
type _PickBest_Shape = Assert<PickedRegion extends { type: RegionType; index: number; range: string } | null ? true : false>;

// Basic runtime sanity checks (also ensures the compiler doesn't tree-shake the imports).
const parsed = parseA1Range("$A$1:B2");
const index = new RagIndex();
void index;
const schema: SheetSchema = { name: "Sheet1", tables: [], namedRanges: [], dataRegions: [] };
const ref: RegionRef = { type: "table", index: 0 };
scoreRegionForQuery(ref, schema, "revenue");
pickBestRegionForQuery(schema, "revenue");
void parsed;
`;

  await writeFile(entryFile, source, "utf8");

  try {
    const options: ts.CompilerOptions = {
      // Keep the compilation as close to the repo defaults as possible.
      target: ts.ScriptTarget.ES2022,
      module: ts.ModuleKind.ESNext,
      moduleResolution: ts.ModuleResolutionKind.Bundler,
      lib: ["ES2022", "DOM", "DOM.Iterable"],
      types: ["node"],
      strict: true,
      allowJs: true,
      checkJs: false,
      skipLibCheck: true,
      noEmit: true,
    };

    const program = ts.createProgram([entryFile], options);
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
