import { a1ToRowCol, rowColToA1 } from "../a1.js";

const UNSUPPORTED = Symbol("unsupported");

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
  if (/^null$/i.test(trimmed) || /^undefined$/i.test(trimmed)) return null;
  if (/^(true|false)$/i.test(trimmed)) return /^true$/i.test(trimmed);
  if (/^[+-]?\d+(\.\d+)?$/.test(trimmed)) return Number(trimmed);
  return UNSUPPORTED;
}

function stripLineComment(line) {
  const idx = line.indexOf("//");
  if (idx === -1) return line;
  return line.slice(0, idx);
}

function parseJsStringLiteralPrefix(expr) {
  const trimmed = expr.trimStart();
  const quote = trimmed[0];
  if (quote !== '"' && quote !== "'") return null;

  let out = "";
  let escape = false;
  for (let i = 1; i < trimmed.length; i += 1) {
    const ch = trimmed[i];
    if (escape) {
      out += ch;
      escape = false;
      continue;
    }
    if (ch === "\\") {
      escape = true;
      continue;
    }
    if (ch === quote) {
      return { value: out, rest: trimmed.slice(i + 1) };
    }
    out += ch;
  }
  return null;
}

function parseJsLiteralPrefix(expr) {
  const trimmed = expr.trimStart();
  const str = parseJsStringLiteralPrefix(trimmed);
  if (str) return str;
  const nullMatch = /^(null|undefined)\b/i.exec(trimmed);
  if (nullMatch) {
    return { value: null, rest: trimmed.slice(nullMatch[0].length) };
  }
  const boolMatch = /^(true|false)\b/i.exec(trimmed);
  if (boolMatch) {
    return { value: /^true$/i.test(boolMatch[1]), rest: trimmed.slice(boolMatch[0].length) };
  }
  const numMatch = /^[+-]?\d+(?:\.\d+)?\b/.exec(trimmed);
  if (numMatch) {
    return { value: Number(numMatch[0]), rest: trimmed.slice(numMatch[0].length) };
  }
  return null;
}

function parseSingletonMatrix(expr) {
  const trimmed = String(expr || "").trim().replace(/;$/, "");
  let rest = trimmed;
  if (!rest.startsWith("[")) return null;
  rest = rest.slice(1).trimStart();
  if (!rest.startsWith("[")) return null;
  rest = rest.slice(1).trimStart();

  const parsed = parseJsLiteralPrefix(rest);
  if (!parsed) return null;
  const value = parsed.value;
  rest = parsed.rest.trimStart();
  if (!rest.startsWith("]")) return null;
  rest = rest.slice(1).trimStart();
  if (!rest.startsWith("]")) return null;
  rest = rest.slice(1).trimStart();
  if (rest !== "") return null;

  return value;
}

function parseJsMatrix(expr) {
  const trimmed = String(expr || "").trim().replace(/;$/, "");
  let rest = trimmed;
  if (!rest.startsWith("[")) return null;
  rest = rest.slice(1).trimStart();

  const rows = [];
  while (true) {
    rest = rest.trimStart();
    if (rest.startsWith("]")) {
      rest = rest.slice(1).trimStart();
      break;
    }
    if (!rest.startsWith("[")) return null;
    rest = rest.slice(1);

    const row = [];
    while (true) {
      rest = rest.trimStart();
      if (rest.startsWith("]")) {
        rest = rest.slice(1);
        break;
      }
      const parsed = parseJsLiteralPrefix(rest);
      if (!parsed) return null;
      row.push(parsed.value);
      rest = parsed.rest.trimStart();
      if (rest.startsWith(",")) {
        rest = rest.slice(1);
        continue;
      }
      if (rest.startsWith("]")) {
        rest = rest.slice(1);
        break;
      }
      return null;
    }

    rows.push(row);
    rest = rest.trimStart();
    if (rest.startsWith(",")) {
      rest = rest.slice(1);
      continue;
    }
    if (rest.startsWith("]")) {
      rest = rest.slice(1).trimStart();
      break;
    }
    return null;
  }

  if (rest.trim() !== "") return null;
  return rows;
}

function parseA1Range(address) {
  const parts = String(address || "").trim().toUpperCase().split(":");
  const start = a1ToRowCol(parts[0]);
  const end = parts.length === 2 ? a1ToRowCol(parts[1]) : start;

  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startCol = Math.min(start.col, end.col);
  const endCol = Math.max(start.col, end.col);

  return { startRow, endRow, startCol, endCol };
}

function isIdentifier(expr) {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(String(expr || "").trim());
}

function resolveScalar(expr, env) {
  const parsed = parseJsLiteral(expr);
  if (parsed !== UNSUPPORTED) return parsed;
  const ident = String(expr || "").trim();
  if (!isIdentifier(ident)) return UNSUPPORTED;
  const binding = env.get(ident);
  if (!binding || binding.kind !== "scalar") return UNSUPPORTED;
  return binding.value;
}

function parseArrayFromFill(expr, env) {
  const trimmed = String(expr || "").trim().replace(/;$/, "");
  const match =
    /^Array\.from\(\s*\{\s*length\s*:\s*(?<rows>[0-9]+)\s*\}\s*,\s*\(\s*\)\s*=>\s*Array\(\s*(?<cols>[0-9]+)\s*\)\.fill\(\s*(?<fill>.+)\s*\)\s*\)\s*$/.exec(
      trimmed,
    );
  if (!match?.groups) return null;
  const rows = Number(match.groups.rows);
  const cols = Number(match.groups.cols);
  if (!Number.isFinite(rows) || rows <= 0 || !Number.isFinite(cols) || cols <= 0) return null;

  const fillValue = resolveScalar(match.groups.fill, env);
  if (fillValue === UNSUPPORTED) return null;

  return Array.from({ length: rows }, () => Array.from({ length: cols }, () => fillValue));
}

function resolveMatrix(expr, env) {
  const trimmed = String(expr || "").trim().replace(/;$/, "");
  const literal = parseJsMatrix(trimmed);
  if (literal) return literal;

  if (isIdentifier(trimmed)) {
    const binding = env.get(trimmed);
    if (binding?.kind === "matrix") return binding.value;
  }

  return parseArrayFromFill(trimmed, env);
}

export function executeTypeScriptMigrationScript({ workbook, code }) {
  const lines = String(code || "").split(/\r?\n/);
  let currentSheet = workbook.activeSheet;
  const env = new Map();

  for (const rawLine of lines) {
    const noComment = stripLineComment(rawLine);
    const line = noComment.trim();
    if (!line) continue;
    if (/^export\b/.test(line)) continue;
    if (/^import\b/.test(line)) continue;
    if (/^(const|let|var)\b/.test(line)) {
      // Handle: const sheet = ctx.activeSheet;
      if (/=\s*ctx\.activeSheet\s*;?$/.test(line)) {
        currentSheet = workbook.activeSheet;
        continue;
      }

      const match = /^(?:const|let|var)\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\s*=\s*(?<expr>.+?)\s*;?$/.exec(line);
      if (match?.groups?.name) {
        const name = match.groups.name;
        const expr = match.groups.expr;

        const scalar = resolveScalar(expr, env);
        if (scalar !== UNSUPPORTED) {
          env.set(name, { kind: "scalar", value: scalar });
        } else {
          const matrix = resolveMatrix(expr, env);
          if (matrix) env.set(name, { kind: "matrix", value: matrix });
        }
      }

      continue;
    }
    if (line === "{" || line === "}") continue;

    const setValueCall = /\.getRange\(\s*(['"])(?<addr>[^'"]+)\1\s*\)\.setValue\(\s*(?<expr>.+)\)\s*;?$/.exec(line);
    if (setValueCall) {
      const value = resolveScalar(setValueCall.groups.expr, env);
      if (value === UNSUPPORTED) throw new Error(`Unsupported TS literal: ${setValueCall.groups.expr}`);
      currentSheet.setCellValue(setValueCall.groups.addr, value);
      continue;
    }

    const setValuesCall = /\.getRange\(\s*(['"])(?<addr>[^'"]+)\1\s*\)\.setValues\(\s*(?<expr>.+)\)\s*;?$/.exec(line);
    if (setValuesCall) {
      const range = parseA1Range(setValuesCall.groups.addr);
      const matrix = resolveMatrix(setValuesCall.groups.expr, env);
      if (!matrix) throw new Error(`Unsupported TS values matrix: ${setValuesCall.groups.expr}`);
      const rowCount = range.endRow - range.startRow + 1;
      const colCount = range.endCol - range.startCol + 1;
      if (matrix.length !== rowCount || matrix.some((row) => row.length !== colCount)) {
        throw new Error(
          `setValues expected ${rowCount}x${colCount} matrix for range ${setValuesCall.groups.addr}, got ${matrix.length}x${matrix[0]?.length ?? 0}`,
        );
      }

      for (let r = 0; r < rowCount; r += 1) {
        for (let c = 0; c < colCount; c += 1) {
          const addr = rowColToA1(range.startRow + r, range.startCol + c);
          currentSheet.setCellValue(addr, matrix[r][c]);
        }
      }
      continue;
    }

    const setFormulasCall =
      /\.getRange\(\s*(['"])(?<addr>[^'"]+)\1\s*\)\.setFormulas\(\s*(?<expr>.+)\)\s*;?$/.exec(line);
    if (setFormulasCall) {
      const range = parseA1Range(setFormulasCall.groups.addr);
      const matrix = resolveMatrix(setFormulasCall.groups.expr, env);
      if (!matrix) throw new Error(`Unsupported TS formulas matrix: ${setFormulasCall.groups.expr}`);
      const rowCount = range.endRow - range.startRow + 1;
      const colCount = range.endCol - range.startCol + 1;
      if (matrix.length !== rowCount || matrix.some((row) => row.length !== colCount)) {
        throw new Error(
          `setFormulas expected ${rowCount}x${colCount} matrix for range ${setFormulasCall.groups.addr}, got ${matrix.length}x${matrix[0]?.length ?? 0}`,
        );
      }

      for (let r = 0; r < rowCount; r += 1) {
        for (let c = 0; c < colCount; c += 1) {
          const addr = rowColToA1(range.startRow + r, range.startCol + c);
          const formula = matrix[r][c];
          if (formula === null) {
            currentSheet.setCellValue(addr, null);
          } else if (typeof formula === "string") {
            currentSheet.setCellFormula(addr, formula);
          } else {
            throw new Error(`Expected formula string, got ${typeof formula}`);
          }
        }
      }
      continue;
    }

    const setValueRange = /^sheet\.range\(\s*(['"])(?<addr>[^'"]+)\1\s*\)\.value\s*=\s*(?<expr>.+)$/.exec(line);
    if (setValueRange) {
      const value = resolveScalar(setValueRange.groups.expr, env);
      if (value === UNSUPPORTED) throw new Error(`Unsupported TS literal: ${setValueRange.groups.expr}`);
      currentSheet.setCellValue(setValueRange.groups.addr, value);
      continue;
    }

    const setFormulaRange = /^sheet\.range\(\s*(['"])(?<addr>[^'"]+)\1\s*\)\.formula\s*=\s*(?<expr>.+)$/.exec(line);
    if (setFormulaRange) {
      const value = resolveScalar(setFormulaRange.groups.expr, env);
      if (value === UNSUPPORTED) throw new Error(`Unsupported TS literal: ${setFormulaRange.groups.expr}`);
      if (value === null) {
        currentSheet.setCellValue(setFormulaRange.groups.addr, null);
      } else if (typeof value === "string") {
        currentSheet.setCellFormula(setFormulaRange.groups.addr, value);
      } else {
        throw new Error(`Expected formula string, got ${typeof value}`);
      }
      continue;
    }

    const setValueCell = /^sheet\.cell\(\s*(?<row>[0-9]+)\s*,\s*(?<col>[0-9]+)\s*\)\.value\s*=\s*(?<expr>.+)$/.exec(line);
    if (setValueCell) {
      const value = resolveScalar(setValueCell.groups.expr, env);
      if (value === UNSUPPORTED) throw new Error(`Unsupported TS literal: ${setValueCell.groups.expr}`);
      const address = rowColToA1(Number(setValueCell.groups.row), Number(setValueCell.groups.col));
      currentSheet.setCellValue(address, value);
      continue;
    }

    const setFormulaCell = /^sheet\.cell\(\s*(?<row>[0-9]+)\s*,\s*(?<col>[0-9]+)\s*\)\.formula\s*=\s*(?<expr>.+)$/.exec(line);
    if (setFormulaCell) {
      const value = resolveScalar(setFormulaCell.groups.expr, env);
      if (value === UNSUPPORTED) throw new Error(`Unsupported TS literal: ${setFormulaCell.groups.expr}`);
      const address = rowColToA1(Number(setFormulaCell.groups.row), Number(setFormulaCell.groups.col));
      if (value === null) {
        currentSheet.setCellValue(address, null);
      } else if (typeof value === "string") {
        currentSheet.setCellFormula(address, value);
      } else {
        throw new Error(`Expected formula string, got ${typeof value}`);
      }
      continue;
    }
  }
}
