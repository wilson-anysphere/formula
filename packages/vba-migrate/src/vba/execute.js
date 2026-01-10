import { rowColToA1 } from "../a1.js";

function extractSubroutineBody(code, entryPoint) {
  const lines = String(code || "").split(/\r?\n/);
  const startRegex = new RegExp(`^\\s*Sub\\s+${entryPoint}\\b`, "i");
  let startIndex = -1;
  for (let i = 0; i < lines.length; i += 1) {
    if (startRegex.test(lines[i])) {
      startIndex = i;
      break;
    }
  }
  if (startIndex === -1) {
    throw new Error(`Could not find Sub ${entryPoint} in module`);
  }

  let endIndex = -1;
  for (let i = startIndex + 1; i < lines.length; i += 1) {
    if (/^\s*End\s+Sub\b/i.test(lines[i])) {
      endIndex = i;
      break;
    }
  }
  if (endIndex === -1) throw new Error(`Could not find End Sub for ${entryPoint}`);
  return lines.slice(startIndex + 1, endIndex);
}

function parseVbaStringLiteral(expr) {
  const trimmed = expr.trim();
  if (!/^".*"$/.test(trimmed)) return null;
  // VBA escapes quotes by doubling them.
  const inner = trimmed.slice(1, -1).replace(/""/g, '"');
  return inner;
}

function parseVbaLiteral(expr) {
  const trimmed = expr.trim();
  const str = parseVbaStringLiteral(trimmed);
  if (str !== null) return str;
  if (/^(True|False)$/i.test(trimmed)) return /^True$/i.test(trimmed);
  if (/^[+-]?\d+(\.\d+)?$/.test(trimmed)) return Number(trimmed);
  return null;
}

function parseAssignment(line) {
  // Very small subset: <target> = <expr> (no line continuation).
  const idx = line.indexOf("=");
  if (idx === -1) return null;
  const left = line.slice(0, idx).trim();
  const right = line.slice(idx + 1).trim();
  if (!left || !right) return null;
  return { left, right };
}

function parseTarget(left) {
  // Supports:
  //   Range("A1").Value
  //   Range("A1").Formula
  //   Cells(1, 2).Value
  //   Worksheets("Sheet1").Range("A1").Value
  //   Sheets("Sheet1").Cells(1,2).Formula
  const sheetMatch = /^(?<sheetObj>Worksheets|Sheets)\("(?<sheet>[^"]+)"\)\.(?<rest>.+)$/i.exec(left);
  const sheetName = sheetMatch?.groups?.sheet ?? null;
  const rest = sheetMatch?.groups?.rest ?? left;

  const rangeMatch = /^(?<obj>Range)\(\s*"(?<addr>[^"]+)"\s*\)\.(?<prop>Value|Formula)$/i.exec(rest);
  if (rangeMatch) {
    return {
      sheetName,
      address: rangeMatch.groups.addr,
      prop: rangeMatch.groups.prop.toLowerCase()
    };
  }

  const cellsMatch = /^(?<obj>Cells)\(\s*(?<row>[0-9]+)\s*,\s*(?<col>[0-9]+)\s*\)\.(?<prop>Value|Formula)$/i.exec(rest);
  if (cellsMatch) {
    const address = rowColToA1(Number(cellsMatch.groups.row), Number(cellsMatch.groups.col));
    return {
      sheetName,
      address,
      prop: cellsMatch.groups.prop.toLowerCase()
    };
  }

  return null;
}

export function executeVbaModuleSub({ workbook, module, entryPoint }) {
  const bodyLines = extractSubroutineBody(module.code, entryPoint);

  for (const rawLine of bodyLines) {
    const line = rawLine.trim();
    if (!line) continue;
    if (line.startsWith("'")) continue;
    if (/^\s*Rem\b/i.test(rawLine)) continue;

    const assignment = parseAssignment(rawLine);
    if (!assignment) continue;

    const target = parseTarget(assignment.left);
    if (!target) {
      throw new Error(`Unsupported VBA statement: ${rawLine.trim()}`);
    }

    const sheet = target.sheetName ? workbook.getSheet(target.sheetName) : workbook.activeSheet;
    const literal = parseVbaLiteral(assignment.right);
    if (literal === null) {
      throw new Error(`Unsupported VBA expression: ${assignment.right}`);
    }

    if (target.prop === "value") {
      sheet.setCellValue(target.address, literal);
      continue;
    }

    if (target.prop === "formula") {
      sheet.setCellFormula(target.address, literal);
      continue;
    }

    throw new Error(`Unsupported property: ${target.prop}`);
  }
}

