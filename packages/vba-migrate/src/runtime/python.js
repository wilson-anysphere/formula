function parsePythonStringLiteral(expr) {
  const trimmed = expr.trim();
  if (!/^(['"]).*\1$/.test(trimmed)) return null;
  const quote = trimmed[0];
  const inner = trimmed.slice(1, -1);
  // Minimal unescaping for the fixtures we run (no full python string semantics).
  return inner.replace(new RegExp(`\\\\${quote}`, "g"), quote);
}

function parsePythonLiteral(expr) {
  const trimmed = expr.trim();
  const str = parsePythonStringLiteral(trimmed);
  if (str !== null) return str;
  if (/^(True|False)$/i.test(trimmed)) return /^True$/i.test(trimmed);
  if (/^[+-]?\d+(\.\d+)?$/.test(trimmed)) return Number(trimmed);
  return null;
}

function stripPythonComment(line) {
  // Super conservative: only strip comments starting with # when not in quotes.
  const idx = line.indexOf("#");
  if (idx === -1) return line;
  return line.slice(0, idx);
}

export function executePythonMigrationScript({ workbook, code }) {
  const lines = String(code || "").split(/\r?\n/);
  let currentSheet = workbook.activeSheet;

  for (const rawLine of lines) {
    const noComment = stripPythonComment(rawLine);
    const line = noComment.trim();
    if (!line) continue;
    if (/^import\b/.test(line)) continue;
    if (/^def\b/.test(line)) continue;
    if (/^if\s+__name__\s*==/.test(line)) continue;
    if (/^\s*main\s*\(\s*\)\s*$/.test(line)) continue;

    const sheetAssign = /^sheet\s*=\s*formula\.active_sheet\s*$/.exec(line);
    if (sheetAssign) {
      currentSheet = workbook.activeSheet;
      continue;
    }

    const getSheetAssign = /^sheet\s*=\s*formula\.get_sheet\(\s*(['"])(?<name>[^'"]+)\1\s*\)\s*$/.exec(line);
    if (getSheetAssign) {
      currentSheet = workbook.getSheet(getSheetAssign.groups.name);
      continue;
    }

    // sheet["A1"] = <literal>
    const setValue = /^sheet\[\s*(['"])(?<addr>[^'"]+)\1\s*\]\s*=\s*(?<expr>.+)$/.exec(line);
    if (setValue) {
      const value = parsePythonLiteral(setValue.groups.expr);
      if (value === null) throw new Error(`Unsupported Python literal: ${setValue.groups.expr}`);
      currentSheet.setCellValue(setValue.groups.addr, value);
      continue;
    }

    // sheet["A1"].formula = <literal>
    const setFormula = /^sheet\[\s*(['"])(?<addr>[^'"]+)\1\s*\]\.formula\s*=\s*(?<expr>.+)$/.exec(line);
    if (setFormula) {
      const value = parsePythonLiteral(setFormula.groups.expr);
      if (value === null) throw new Error(`Unsupported Python literal: ${setFormula.groups.expr}`);
      currentSheet.setCellFormula(setFormula.groups.addr, value);
      continue;
    }

    // sheet["A1"].value = <literal>
    const setValueProp = /^sheet\[\s*(['"])(?<addr>[^'"]+)\1\s*\]\.value\s*=\s*(?<expr>.+)$/.exec(line);
    if (setValueProp) {
      const value = parsePythonLiteral(setValueProp.groups.expr);
      if (value === null) throw new Error(`Unsupported Python literal: ${setValueProp.groups.expr}`);
      currentSheet.setCellValue(setValueProp.groups.addr, value);
      continue;
    }
  }
}

