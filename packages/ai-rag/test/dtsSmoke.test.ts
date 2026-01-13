import { test } from "vitest";
import * as ts from "typescript";
import { dirname } from "node:path";
import { fileURLToPath } from "node:url";

test("ai-rag d.ts smoke test (public API stays in sync)", () => {
  const configPath = fileURLToPath(new URL("./tsconfig.dts-smoke.json", import.meta.url));
  const configDir = dirname(configPath);

  const configFile = ts.readConfigFile(configPath, ts.sys.readFile);
  if (configFile.error) {
    const message = ts.formatDiagnosticsWithColorAndContext([configFile.error], {
      getCanonicalFileName: (fileName) => fileName,
      getCurrentDirectory: ts.sys.getCurrentDirectory,
      getNewLine: () => ts.sys.newLine,
    });
    throw new Error(message);
  }

  const parsed = ts.parseJsonConfigFileContent(configFile.config, ts.sys, configDir);
  const program = ts.createProgram({ rootNames: parsed.fileNames, options: parsed.options });

  const diagnostics = ts.getPreEmitDiagnostics(program);
  if (diagnostics.length === 0) return;

  const message = ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: (fileName) => fileName,
    getCurrentDirectory: ts.sys.getCurrentDirectory,
    getNewLine: () => ts.sys.newLine,
  });
  throw new Error(message);
});

