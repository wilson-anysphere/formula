import { rowColToA1 } from "../a1.js";

function parseJsStringLiteral(expr) {
  const trimmed = expr.trim();
  if (!/^(['"]).*\1$/.test(trimmed)) return null;
  const quote = trimmed[0];
  const inner = trimmed.slice(1, -1);
  return inner.replace(new RegExp(`\\\\${quote}`, "g"), quote);
}

function parseJsLiteral(expr) {
  const trimmed = expr.trim().replace(/;$/, "");
  const str = parseJsStringLiteral(trimmed);
  if (str !== null) return str;
  if (/^(true|false)$/i.test(trimmed)) return /^true$/i.test(trimmed);
  if (/^[+-]?\d+(\.\d+)?$/.test(trimmed)) return Number(trimmed);
  return null;
}

function stripLineComment(line) {
  const idx = line.indexOf("//");
  if (idx === -1) return line;
  return line.slice(0, idx);
}

export function executeTypeScriptMigrationScript({ workbook, code }) {
  const lines = String(code || "").split(/\r?\n/);
  let currentSheet = workbook.activeSheet;

  for (const rawLine of lines) {
    const noComment = stripLineComment(rawLine);
    const line = noComment.trim();
    if (!line) continue;
    if (/^export\b/.test(line)) continue;
    if (/^import\b/.test(line)) continue;
    if (/^const\b/.test(line) || /^let\b/.test(line)) {
      // Handle: const sheet = ctx.activeSheet;
      if (/=\s*ctx\.activeSheet\s*;?$/.test(line)) {
        currentSheet = workbook.activeSheet;
      }
      continue;
    }
    if (line === "{" || line === "}") continue;

    const setValueRange = /^sheet\.range\(\s*(['"])(?<addr>[^'"]+)\1\s*\)\.value\s*=\s*(?<expr>.+)$/.exec(line);
    if (setValueRange) {
      const value = parseJsLiteral(setValueRange.groups.expr);
      if (value === null) throw new Error(`Unsupported TS literal: ${setValueRange.groups.expr}`);
      currentSheet.setCellValue(setValueRange.groups.addr, value);
      continue;
    }

    const setFormulaRange = /^sheet\.range\(\s*(['"])(?<addr>[^'"]+)\1\s*\)\.formula\s*=\s*(?<expr>.+)$/.exec(line);
    if (setFormulaRange) {
      const value = parseJsLiteral(setFormulaRange.groups.expr);
      if (value === null) throw new Error(`Unsupported TS literal: ${setFormulaRange.groups.expr}`);
      currentSheet.setCellFormula(setFormulaRange.groups.addr, value);
      continue;
    }

    const setValueCell = /^sheet\.cell\(\s*(?<row>[0-9]+)\s*,\s*(?<col>[0-9]+)\s*\)\.value\s*=\s*(?<expr>.+)$/.exec(line);
    if (setValueCell) {
      const value = parseJsLiteral(setValueCell.groups.expr);
      if (value === null) throw new Error(`Unsupported TS literal: ${setValueCell.groups.expr}`);
      const address = rowColToA1(Number(setValueCell.groups.row), Number(setValueCell.groups.col));
      currentSheet.setCellValue(address, value);
      continue;
    }

    const setFormulaCell = /^sheet\.cell\(\s*(?<row>[0-9]+)\s*,\s*(?<col>[0-9]+)\s*\)\.formula\s*=\s*(?<expr>.+)$/.exec(line);
    if (setFormulaCell) {
      const value = parseJsLiteral(setFormulaCell.groups.expr);
      if (value === null) throw new Error(`Unsupported TS literal: ${setFormulaCell.groups.expr}`);
      const address = rowColToA1(Number(setFormulaCell.groups.row), Number(setFormulaCell.groups.col));
      currentSheet.setCellFormula(address, value);
      continue;
    }
  }
}
