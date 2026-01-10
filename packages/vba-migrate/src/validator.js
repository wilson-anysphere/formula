import { compareWorkbooks, diffWorkbooks } from "./workbook.js";
import { executeVbaModuleSub } from "./vba/execute.js";
import { executePythonMigrationScript } from "./runtime/python.js";
import { executeTypeScriptMigrationScript } from "./runtime/typescript.js";

export function validateMigration({ workbook, module, entryPoint, target, code }) {
  const before = workbook.clone();
  const vbaWorkbook = workbook.clone();
  const scriptWorkbook = workbook.clone();

  executeVbaModuleSub({ workbook: vbaWorkbook, module, entryPoint });

  if (target === "python") {
    executePythonMigrationScript({ workbook: scriptWorkbook, code });
  } else if (target === "typescript") {
    executeTypeScriptMigrationScript({ workbook: scriptWorkbook, code });
  } else {
    throw new Error(`Unknown target: ${target}`);
  }

  const vbaDiff = diffWorkbooks(before, vbaWorkbook);
  const scriptDiff = diffWorkbooks(before, scriptWorkbook);
  const mismatches = compareWorkbooks(vbaWorkbook, scriptWorkbook);

  return {
    ok: mismatches.length === 0,
    vbaDiff,
    scriptDiff,
    mismatches
  };
}

