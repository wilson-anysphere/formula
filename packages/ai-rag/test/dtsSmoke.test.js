import assert from "node:assert/strict";
import test from "node:test";
import { createRequire } from "node:module";
import { readdir, readFile, stat } from "node:fs/promises";
import { dirname, join, relative } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const require = createRequire(import.meta.url);

function hasTypeScriptDependency() {
  try {
    require.resolve("typescript");
    return true;
  } catch {
    return false;
  }
}

const hasTypeScript = hasTypeScriptDependency();
/** @type {import("typescript") | null} */
const ts = hasTypeScript ? require("typescript") : null;

function formatDiagnostics(diagnostics) {
  assert.ok(ts);
  return ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: (fileName) => fileName,
    getCurrentDirectory: ts.sys.getCurrentDirectory,
    getNewLine: () => ts.sys.newLine,
  });
}

function collectNamedExports(sourceFile) {
  assert.ok(ts);
  const exports = [];

  function hasExportModifier(node) {
    return Boolean(node.modifiers?.some((m) => m.kind === ts.SyntaxKind.ExportKeyword));
  }

  function isConstEnum(node) {
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
    if (!hasExportModifier(node)) continue;

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

function collectNamedImports(sourceFile, moduleSpecifier) {
  assert.ok(ts);
  const imports = [];
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

function stableSorted(arr) {
  return Array.from(new Set(arr)).sort((a, b) => a.localeCompare(b));
}

function assertSameSet(label, a, b) {
  const left = stableSorted(a);
  const right = stableSorted(b);
  const leftOnly = left.filter((x) => !right.includes(x));
  const rightOnly = right.filter((x) => !left.includes(x));
  if (leftOnly.length === 0 && rightOnly.length === 0) return;
  assert.fail(
    [
      `${label} mismatch:`,
      leftOnly.length ? `  Only in left: ${leftOnly.join(", ")}` : null,
      rightOnly.length ? `  Only in right: ${rightOnly.join(", ")}` : null,
    ]
      .filter(Boolean)
      .join("\n"),
  );
}

test("ai-rag index exports match index.d.ts (API surface consistency)", { skip: !hasTypeScript }, async () => {
  assert.ok(ts);
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

test("ai-rag module exports match adjacent .d.ts files", { skip: !hasTypeScript }, async () => {
  assert.ok(ts);
  const srcDir = fileURLToPath(new URL("../src/", import.meta.url));

  async function collectDtsFiles(dir) {
    const out = [];
    const entries = await readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = join(dir, entry.name);
      if (entry.isDirectory()) {
        out.push(...(await collectDtsFiles(fullPath)));
        continue;
      }
      if (!entry.isFile()) continue;
      if (!entry.name.endsWith(".d.ts")) continue;
      out.push(fullPath);
    }
    return out;
  }

  const dtsFiles = await collectDtsFiles(srcDir);
  for (const dtsPath of dtsFiles) {
    const jsPath = dtsPath.slice(0, -".d.ts".length) + ".js";
    try {
      const jsStats = await stat(jsPath);
      if (!jsStats.isFile()) continue;
    } catch {
      // Ignore declaration files that don't have an adjacent runtime module.
      continue;
    }

    const runtimeModule = await import(pathToFileURL(jsPath).href);
    const runtimeExports = Object.keys(runtimeModule).filter((key) => key !== "default");

    const dtsText = await readFile(dtsPath, "utf8");
    const dtsSource = ts.createSourceFile(dtsPath, dtsText, ts.ScriptTarget.Latest, true);
    const declaredExports = collectNamedExports(dtsSource);

    const label = relative(srcDir, jsPath).split("\\").join("/");
    assertSameSet(`${label} exports`, runtimeExports, declaredExports);
  }
});

test("ai-rag d.ts smoke test (public API stays in sync)", { skip: !hasTypeScript }, () => {
  assert.ok(ts);
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
