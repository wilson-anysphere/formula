import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import ts from "typescript";

const desktopRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const mainFile = path.join(desktopRoot, "src", "main.ts");

const tsconfigPath = ts.findConfigFile(desktopRoot, ts.sys.fileExists, "tsconfig.json");
if (!tsconfigPath) {
  console.error("Failed to find apps/desktop/tsconfig.json");
  process.exit(1);
}

const configFile = ts.readConfigFile(tsconfigPath, ts.sys.readFile);
if (configFile.error) {
  console.error(formatDiagnostics([configFile.error]));
  process.exit(1);
}

const parsedConfig = ts.parseJsonConfigFileContent(configFile.config, ts.sys, path.dirname(tsconfigPath));

// Note: `apps/desktop` intentionally relies on Vite/esbuild for transpilation, and
// a full `tsc -p tsconfig.json` run can produce unrelated type errors (the repo
// does not currently gate desktop builds on full TypeScript typechecking).
//
// This script is a lightweight guard to ensure the desktop entrypoint remains
// syntactically valid for TypeScript, including rejecting duplicate object-literal
// keys (TS1117), which can otherwise slip through esbuild-based builds.
const options = {
  ...parsedConfig.options,
  noEmit: true,
  // Keep the check fast: we only care about diagnostics for the entrypoint itself.
  // Resolving and typechecking the full import graph is intentionally out-of-scope here.
  noResolve: true,
};

const program = ts.createProgram([mainFile], options);
const sourceFile = program.getSourceFile(mainFile);
if (!sourceFile) {
  console.error(`Failed to load source file: ${mainFile}`);
  process.exit(1);
}

const diagnostics = [
  ...program.getOptionsDiagnostics(),
  ...program.getSyntacticDiagnostics(sourceFile),
  // TS1117 ("duplicate object-literal key") is reported as a semantic diagnostic.
  ...program.getSemanticDiagnostics(sourceFile).filter((d) => d.code === 1117),
];

if (diagnostics.length > 0) {
  console.error(formatDiagnostics(diagnostics));
  process.exit(1);
}

process.stdout.write("Desktop TypeScript entrypoint check passed.\n");

/**
 * @param {readonly ts.Diagnostic[]} diags
 */
function formatDiagnostics(diags) {
  const formatHost = {
    getCanonicalFileName: (fileName) => fileName,
    getCurrentDirectory: () => desktopRoot,
    getNewLine: () => "\n",
  };
  return ts.formatDiagnosticsWithColorAndContext(diags, formatHost);
}

