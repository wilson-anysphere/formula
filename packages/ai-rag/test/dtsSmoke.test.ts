import { test } from "vitest";
import * as ts from "typescript";
import { readFile } from "node:fs/promises";
import { dirname } from "node:path";
import { fileURLToPath } from "node:url";

function formatDiagnostics(diagnostics: readonly ts.Diagnostic[]) {
  return ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: (fileName) => fileName,
    getCurrentDirectory: ts.sys.getCurrentDirectory,
    getNewLine: () => ts.sys.newLine,
  });
}

function collectNamedExports(sourceFile: ts.SourceFile) {
  const exports: string[] = [];

  function hasExportModifier(node: ts.Node & { modifiers?: ts.NodeArray<ts.ModifierLike> }) {
    return Boolean(node.modifiers?.some((m) => m.kind === ts.SyntaxKind.ExportKeyword));
  }

  function isConstEnum(node: ts.EnumDeclaration) {
    return Boolean(node.modifiers?.some((m) => m.kind === ts.SyntaxKind.ConstKeyword));
  }

  for (const node of sourceFile.statements) {
    // `export { ... } from "..."`
    if (ts.isExportDeclaration(node)) {
      if (node.isTypeOnly) continue;
      if (!node.exportClause) continue;
      if (ts.isNamedExports(node.exportClause)) {
        for (const specifier of node.exportClause.elements) {
          if (specifier.isTypeOnly) continue;
          exports.push(specifier.name.text);
        }
      } else if (ts.isNamespaceExport(node.exportClause)) {
        exports.push(node.exportClause.name.text);
      }
      continue;
    }

    // `export class Foo {}`, `export function foo() {}`, `export const x = ...`
    if (!hasExportModifier(node as any)) continue;

    if (ts.isTypeAliasDeclaration(node) || ts.isInterfaceDeclaration(node)) {
      // Type-only exports; ignore for runtime alignment.
      continue;
    }

    if (ts.isFunctionDeclaration(node) || ts.isClassDeclaration(node)) {
      if (node.name) exports.push(node.name.text);
      continue;
    }

    if (ts.isEnumDeclaration(node)) {
      // const enums are erased at runtime, so treat them as type-only for alignment.
      if (isConstEnum(node)) continue;
      exports.push(node.name.text);
      continue;
    }

    if (ts.isVariableStatement(node)) {
      for (const decl of node.declarationList.declarations) {
        if (ts.isIdentifier(decl.name)) exports.push(decl.name.text);
      }
      continue;
    }
  }

  return exports;
}

function collectNamedImports(sourceFile: ts.SourceFile, moduleSpecifier: string) {
  const imports: string[] = [];
  for (const node of sourceFile.statements) {
    if (!ts.isImportDeclaration(node)) continue;
    if (!ts.isStringLiteral(node.moduleSpecifier)) continue;
    if (node.moduleSpecifier.text !== moduleSpecifier) continue;
    const clause = node.importClause;
    if (clause?.isTypeOnly) continue;
    if (!clause?.namedBindings) continue;
    if (!ts.isNamedImports(clause.namedBindings)) continue;
    for (const specifier of clause.namedBindings.elements) {
      if (specifier.isTypeOnly) continue;
      imports.push(specifier.name.text);
    }
  }
  return imports;
}

function stableSorted(arr: Iterable<string>) {
  return Array.from(new Set(arr)).sort((a, b) => a.localeCompare(b));
}

function assertSameSet(label: string, a: Iterable<string>, b: Iterable<string>) {
  const left = stableSorted(a);
  const right = stableSorted(b);
  const leftOnly = left.filter((x) => !right.includes(x));
  const rightOnly = right.filter((x) => !left.includes(x));
  if (leftOnly.length === 0 && rightOnly.length === 0) return;
  throw new Error(
    [
      `${label} mismatch:`,
      leftOnly.length ? `  Only in left: ${leftOnly.join(", ")}` : null,
      rightOnly.length ? `  Only in right: ${rightOnly.join(", ")}` : null,
    ]
      .filter(Boolean)
      .join("\n"),
  );
}

test("ai-rag index exports match index.d.ts (API surface consistency)", async () => {
  const runtimeModule = await import("../src/index.js");
  const runtimeExports = Object.keys(runtimeModule).filter((key) => key !== "default");

  const indexDtsPath = fileURLToPath(new URL("../src/index.d.ts", import.meta.url));
  const indexDtsText = await readFile(indexDtsPath, "utf8");
  const indexDts = ts.createSourceFile(indexDtsPath, indexDtsText, ts.ScriptTarget.Latest, true);
  const declaredExports = collectNamedExports(indexDts);

  assertSameSet("index exports", runtimeExports, declaredExports);

  const smokeProgramPath = fileURLToPath(new URL("./dtsSmokeProgram.ts", import.meta.url));
  const smokeProgramText = await readFile(smokeProgramPath, "utf8");
  const smokeProgram = ts.createSourceFile(smokeProgramPath, smokeProgramText, ts.ScriptTarget.Latest, true);
  const smokeImports = collectNamedImports(smokeProgram, "../src/index.js");

  assertSameSet("dtsSmokeProgram imports", runtimeExports, smokeImports);
});

test("ai-rag module exports match adjacent .d.ts files", async () => {
  const modules: Array<{ js: string; dts: string }> = [
    { js: "../src/pipeline/indexWorkbook.js", dts: "../src/pipeline/indexWorkbook.d.ts" },
    { js: "../src/workbook/chunkWorkbook.js", dts: "../src/workbook/chunkWorkbook.d.ts" },
    { js: "../src/workbook/chunkToText.js", dts: "../src/workbook/chunkToText.d.ts" },
    { js: "../src/store/binaryStorage.js", dts: "../src/store/binaryStorage.d.ts" },
    { js: "../src/store/inMemoryVectorStore.js", dts: "../src/store/inMemoryVectorStore.d.ts" },
    { js: "../src/store/jsonVectorStore.js", dts: "../src/store/jsonVectorStore.d.ts" },
    { js: "../src/store/sqliteVectorStore.js", dts: "../src/store/sqliteVectorStore.d.ts" },
  ];

  for (const { js, dts } of modules) {
    const runtimeModule = await import(js);
    const runtimeExports = Object.keys(runtimeModule).filter((key) => key !== "default");

    const dtsPath = fileURLToPath(new URL(dts, import.meta.url));
    const dtsText = await readFile(dtsPath, "utf8");
    const dtsSource = ts.createSourceFile(dtsPath, dtsText, ts.ScriptTarget.Latest, true);
    const declaredExports = collectNamedExports(dtsSource);

    assertSameSet(`${js} exports`, runtimeExports, declaredExports);
  }
});

test("ai-rag d.ts smoke test (public API stays in sync)", () => {
  const configPath = fileURLToPath(new URL("./tsconfig.dts-smoke.json", import.meta.url));
  const configDir = dirname(configPath);

  const configFile = ts.readConfigFile(configPath, ts.sys.readFile);
  if (configFile.error) {
    throw new Error(formatDiagnostics([configFile.error]));
  }

  const parsed = ts.parseJsonConfigFileContent(configFile.config, ts.sys, configDir);
  const program = ts.createProgram({ rootNames: parsed.fileNames, options: parsed.options });

  const diagnostics = ts.getPreEmitDiagnostics(program);
  if (diagnostics.length === 0) return;

  throw new Error(formatDiagnostics(diagnostics));
});
