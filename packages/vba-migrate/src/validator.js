import { compareWorkbooks, diffWorkbooks } from "./workbook.js";
import { executePythonMigrationScript } from "./runtime/python.js";
import { executeTypeScriptMigrationScript } from "./runtime/typescript.js";
import { RustCliOracle } from "./vba/oracle.js";

export async function validateMigration({
  workbook,
  module,
  entryPoint,
  target,
  code,
  oracle = new RustCliOracle(),
  compareOptions = {}
}) {
  const before = workbook.clone();
  const scriptWorkbook = workbook.clone();

  const vbaPayloadBytes = workbook.toBytes({ vbaModules: [module] });
  const vbaRun = await oracle.runMacro({
    workbookBytes: vbaPayloadBytes,
    macroName: entryPoint,
    inputs: []
  });
  const vbaWorkbook = workbook.constructor.fromBytes(vbaRun.workbookAfter);

  if (target === "python") {
    executePythonMigrationScript({ workbook: scriptWorkbook, code });
  } else if (target === "typescript") {
    executeTypeScriptMigrationScript({ workbook: scriptWorkbook, code });
  } else {
    throw new Error(`Unknown target: ${target}`);
  }

  const vbaDiff = diffWorkbooks(before, vbaWorkbook, compareOptions);
  const scriptDiff = diffWorkbooks(before, scriptWorkbook, compareOptions);
  const mismatches = compareWorkbooks(vbaWorkbook, scriptWorkbook, compareOptions);

  return {
    ok: mismatches.length === 0,
    vbaDiff,
    scriptDiff,
    mismatches,
    oracle: {
      ok: vbaRun.ok,
      logs: vbaRun.logs,
      errors: vbaRun.errors,
      report: vbaRun.report
    }
  };
}
