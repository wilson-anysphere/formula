function sheetIdForIndex(index) {
  return `sheet_${index}`;
}

function normalizeSheetNameForCaseInsensitiveCompare(name) {
  // Excel compares sheet names case-insensitively with Unicode NFKC normalization.
  // Match the semantics used by the backend and desktop DocumentController.
  try {
    return String(name ?? "").normalize("NFKC").toUpperCase();
  } catch {
    return String(name ?? "").toUpperCase();
  }
}

function cellKey(row, col) {
  return `${row},${col}`;
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function deepMerge(base, patch) {
  if (!isPlainObject(base) || !isPlainObject(patch)) return patch;
  const out = { ...base };
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) continue;
    if (isPlainObject(value) && isPlainObject(out[key])) {
      out[key] = deepMerge(out[key], value);
    } else {
      out[key] = value;
    }
  }
  return out;
}

function applyFormatPatch(base, patch) {
  if (patch == null) return {};
  return deepMerge(isPlainObject(base) ? base : {}, patch);
}

function colLettersToIndex(letters) {
  let col = 0;
  for (const ch of letters) {
    const code = ch.toUpperCase().charCodeAt(0);
    if (code < 65 || code > 90) throw new Error(`Invalid column letter "${ch}"`);
    col = col * 26 + (code - 64);
  }
  return col - 1;
}

function parseCellRef(ref) {
  const match = /^([A-Za-z]+)([0-9]+)$/.exec(ref);
  if (!match) throw new Error(`Invalid cell reference "${ref}"`);
  const [, colLetters, rowDigits] = match;
  const row = Number.parseInt(rowDigits, 10) - 1;
  const col = colLettersToIndex(colLetters);
  if (!Number.isFinite(row) || row < 0) throw new Error(`Invalid row in reference "${ref}"`);
  return { row, col };
}

function tokenize(expr) {
  const tokens = [];
  const re = /\s*([A-Za-z]+[0-9]+|\d+(?:\.\d+)?(?:[eE][+-]?\d+)?|[()+\-*/^])\s*/y;
  let pos = 0;
  while (pos < expr.length) {
    re.lastIndex = pos;
    const match = re.exec(expr);
    if (!match) {
      throw new Error(`Unexpected token at "${expr.slice(pos)}"`);
    }
    const tok = match[1];
    pos = re.lastIndex;
    if (/^[A-Za-z]+[0-9]+$/.test(tok)) tokens.push({ type: "cell", value: tok });
    else if (/^\d/.test(tok)) tokens.push({ type: "number", value: Number(tok) });
    else tokens.push({ type: "op", value: tok });
  }
  return tokens;
}

const PRECEDENCE = new Map([
  ["u-", 4],
  ["^", 3],
  ["*", 2],
  ["/", 2],
  ["+", 1],
  ["-", 1],
]);

function isRightAssociative(op) {
  return op === "^" || op === "u-";
}

function toRpn(tokens) {
  const output = [];
  const ops = [];

  let prevType = "start";
  for (const token of tokens) {
    if (token.type === "number" || token.type === "cell") {
      output.push(token);
      prevType = "value";
      continue;
    }

    const op = token.value;
    if (op === "(") {
      ops.push(op);
      prevType = "lparen";
      continue;
    }
    if (op === ")") {
      while (ops.length && ops[ops.length - 1] !== "(") output.push({ type: "op", value: ops.pop() });
      if (!ops.length) throw new Error("Mismatched parentheses");
      ops.pop();
      prevType = "value";
      continue;
    }

    let realOp = op;
    if (op === "-" && (prevType === "start" || prevType === "op" || prevType === "lparen")) {
      realOp = "u-";
    }

    while (ops.length) {
      const top = ops[ops.length - 1];
      if (top === "(") break;
      const pTop = PRECEDENCE.get(top);
      const pCur = PRECEDENCE.get(realOp);
      if (pTop === undefined || pCur === undefined) break;
      if (pTop > pCur || (pTop === pCur && !isRightAssociative(realOp))) {
        output.push({ type: "op", value: ops.pop() });
      } else {
        break;
      }
    }

    ops.push(realOp);
    prevType = "op";
  }

  while (ops.length) {
    const op = ops.pop();
    if (op === "(") throw new Error("Mismatched parentheses");
    output.push({ type: "op", value: op });
  }

  return output;
}

function evalRpn(rpn, lookupCellValue) {
  const stack = [];
  for (const token of rpn) {
    if (token.type === "number") {
      stack.push(token.value);
      continue;
    }
    if (token.type === "cell") {
      stack.push(lookupCellValue(token.value));
      continue;
    }

    const op = token.value;
    if (op === "u-") {
      const a = stack.pop();
      stack.push(-Number(a ?? 0));
      continue;
    }

    const b = Number(stack.pop() ?? 0);
    const a = Number(stack.pop() ?? 0);
    switch (op) {
      case "+":
        stack.push(a + b);
        break;
      case "-":
        stack.push(a - b);
        break;
      case "*":
        stack.push(a * b);
        break;
      case "/":
        stack.push(a / b);
        break;
      case "^":
        stack.push(a ** b);
        break;
      default:
        throw new Error(`Unsupported operator "${op}"`);
    }
  }

  if (stack.length !== 1) throw new Error("Invalid expression");
  return stack[0];
}

export class MockWorkbook {
  constructor() {
    this.sheets = new Map();
    this.activeSheetId = sheetIdForIndex(1);
    this.sheets.set(this.activeSheetId, { id: this.activeSheetId, name: "Sheet1", cells: new Map() });
    this.selection = {
      sheet_id: this.activeSheetId,
      start_row: 0,
      start_col: 0,
      end_row: 0,
      end_col: 0,
    };
  }

  get_active_sheet_id() {
    return this.activeSheetId;
  }

  get_sheet_id({ name }) {
    const desired = normalizeSheetNameForCaseInsensitiveCompare(name);
    for (const sheet of this.sheets.values()) {
      if (normalizeSheetNameForCaseInsensitiveCompare(sheet.name) === desired) return sheet.id;
    }
    return null;
  }

  create_sheet({ name, index }) {
    const desiredName = String(name ?? "").trim();
    if (!desiredName) {
      throw new Error("create_sheet expects a non-empty name");
    }
    const desiredNormalized = normalizeSheetNameForCaseInsensitiveCompare(desiredName);
    for (const sheet of this.sheets.values()) {
      if (normalizeSheetNameForCaseInsensitiveCompare(sheet.name) === desiredNormalized) {
        throw new Error("sheet name already exists");
      }
    }

    const nextIndex = this.sheets.size + 1;
    const id = sheetIdForIndex(nextIndex);
    const newSheet = { id, name: desiredName, cells: new Map() };

    const orderedIds = Array.from(this.sheets.keys());

    let insertIndex;
    if (typeof index === "number" && Number.isInteger(index) && index >= 0) {
      insertIndex = Math.min(index, orderedIds.length);
    } else {
      const activeIdx = orderedIds.indexOf(this.activeSheetId);
      insertIndex = activeIdx >= 0 ? activeIdx + 1 : orderedIds.length;
    }

    orderedIds.splice(insertIndex, 0, id);

    const nextSheets = new Map();
    for (const sheetId of orderedIds) {
      if (sheetId === id) {
        nextSheets.set(id, newSheet);
      } else {
        const existing = this.sheets.get(sheetId);
        if (existing) nextSheets.set(sheetId, existing);
      }
    }
    this.sheets = nextSheets;

    return id;
  }

  get_sheet_name({ sheet_id }) {
    const sheet = this.sheets.get(sheet_id);
    if (!sheet) throw new Error(`Unknown sheet_id "${sheet_id}"`);
    return sheet.name;
  }

  rename_sheet({ sheet_id, name }) {
    const sheet = this.sheets.get(sheet_id);
    if (!sheet) throw new Error(`Unknown sheet_id "${sheet_id}"`);
    const desiredName = String(name ?? "").trim();
    if (!desiredName) {
      throw new Error("sheet name cannot be blank");
    }
    const desiredNormalized = normalizeSheetNameForCaseInsensitiveCompare(desiredName);
    for (const other of this.sheets.values()) {
      if (other.id === sheet_id) continue;
      if (normalizeSheetNameForCaseInsensitiveCompare(other.name) === desiredNormalized) {
        throw new Error("sheet name already exists");
      }
    }
    sheet.name = desiredName;
    return null;
  }

  get_selection() {
    return { ...this.selection };
  }

  set_selection({ selection }) {
    if (!selection || typeof selection.sheet_id !== "string") {
      throw new Error("set_selection expects { selection: { sheet_id, start_row, start_col, end_row, end_col } }");
    }
    this._requireSheet(selection.sheet_id);
    this.selection = { ...selection };
    this.activeSheetId = selection.sheet_id;
    return null;
  }

  get_cell_value({ sheet_id, row, col }) {
    const sheet = this._requireSheet(sheet_id);
    const cell = sheet.cells.get(cellKey(row, col));
    return cell?.value ?? null;
  }

  get_cell_formula({ range }) {
    const { sheet_id, start_row, start_col, end_row, end_col } = range;
    if (start_row !== end_row || start_col !== end_col) throw new Error("get_cell_formula expects a single cell range");
    const sheet = this._requireSheet(sheet_id);
    const cell = sheet.cells.get(cellKey(start_row, start_col));
    return cell?.formula ?? null;
  }

  set_cell_value({ range, value }) {
    const { sheet_id, start_row, start_col, end_row, end_col } = range;
    if (start_row !== end_row || start_col !== end_col) throw new Error("set_cell_value expects a single cell range");
    const sheet = this._requireSheet(sheet_id);
    const key = cellKey(start_row, start_col);
    const existing = sheet.cells.get(key) ?? {};
    sheet.cells.set(key, { ...existing, value, formula: null, format: existing.format ?? {} });
    this._recalculateSheet(sheet);
    return null;
  }

  set_cell_formula({ range, formula }) {
    const { sheet_id, start_row, start_col, end_row, end_col } = range;
    if (start_row !== end_row || start_col !== end_col) throw new Error("set_cell_formula expects a single cell range");
    const sheet = this._requireSheet(sheet_id);
    const key = cellKey(start_row, start_col);
    const existing = sheet.cells.get(key) ?? {};
    sheet.cells.set(key, { ...existing, formula, format: existing.format ?? {} });
    this._recalculateSheet(sheet);
    return null;
  }

  get_range_values({ range }) {
    const { sheet_id, start_row, start_col, end_row, end_col } = range;
    const sheet = this._requireSheet(sheet_id);
    const rows = [];
    for (let r = start_row; r <= end_row; r++) {
      const rowVals = [];
      for (let c = start_col; c <= end_col; c++) {
        const cell = sheet.cells.get(cellKey(r, c));
        rowVals.push(cell?.value ?? null);
      }
      rows.push(rowVals);
    }
    return rows;
  }

  set_range_values({ range, values }) {
    const { sheet_id, start_row, start_col, end_row, end_col } = range;
    const sheet = this._requireSheet(sheet_id);
    const rowCount = end_row - start_row + 1;
    const colCount = end_col - start_col + 1;

    if (!Array.isArray(values) || !Array.isArray(values[0])) {
      // Scalar fill.
      for (let r = 0; r < rowCount; r++) {
        for (let c = 0; c < colCount; c++) {
          const key = cellKey(start_row + r, start_col + c);
          const existing = sheet.cells.get(key) ?? {};
          sheet.cells.set(key, { ...existing, value: values, formula: null, format: existing.format ?? {} });
        }
      }
    } else {
      for (let r = 0; r < rowCount; r++) {
        for (let c = 0; c < colCount; c++) {
          const key = cellKey(start_row + r, start_col + c);
          const existing = sheet.cells.get(key) ?? {};
          const value = values[r]?.[c] ?? null;
          sheet.cells.set(key, { ...existing, value, formula: null, format: existing.format ?? {} });
        }
      }
    }

    this._recalculateSheet(sheet);
    return null;
  }

  set_range_format({ range, format }) {
    const { sheet_id, start_row, start_col, end_row, end_col } = range;
    const sheet = this._requireSheet(sheet_id);
    for (let r = start_row; r <= end_row; r++) {
      for (let c = start_col; c <= end_col; c++) {
        const key = cellKey(r, c);
        const existing = sheet.cells.get(key) ?? { value: null, formula: null, format: {} };
        const baseFormat = existing.format ?? {};
        const nextFormat = applyFormatPatch(baseFormat, format);
        sheet.cells.set(key, { ...existing, format: nextFormat });
      }
    }
    return null;
  }

  get_range_format({ range }) {
    const { sheet_id, start_row, start_col } = range;
    const sheet = this._requireSheet(sheet_id);
    const cell = sheet.cells.get(cellKey(start_row, start_col));
    return cell?.format ?? {};
  }

  clear_range({ range }) {
    const { sheet_id, start_row, start_col, end_row, end_col } = range;
    const sheet = this._requireSheet(sheet_id);
    for (let r = start_row; r <= end_row; r++) {
      for (let c = start_col; c <= end_col; c++) {
        sheet.cells.delete(cellKey(r, c));
      }
    }
    this._recalculateSheet(sheet);
    return null;
  }

  _requireSheet(sheetId) {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) throw new Error(`Unknown sheet_id "${sheetId}"`);
    return sheet;
  }

  _recalculateSheet(sheet) {
    const visited = new Set();
    const evaluating = new Set();

    const evaluateCell = (row, col) => {
      const key = cellKey(row, col);
      if (evaluating.has(key)) throw new Error("Circular reference detected");
      const cell = sheet.cells.get(key);
      if (!cell?.formula) return cell?.value ?? null;
      if (visited.has(key)) return cell.value;

      evaluating.add(key);

      const formula = String(cell.formula);
      const expr = formula.startsWith("=") ? formula.slice(1) : formula;
      const rpn = toRpn(tokenize(expr));
      const value = evalRpn(rpn, (ref) => {
        const { row: rr, col: cc } = parseCellRef(ref);
        const v = evaluateCell(rr, cc);
        return typeof v === "number" ? v : Number(v ?? 0);
      });

      evaluating.delete(key);
      visited.add(key);

      cell.value = value;
      return value;
    };

    for (const [key, cell] of sheet.cells) {
      if (!cell.formula) continue;
      const [rowStr, colStr] = key.split(",");
      evaluateCell(Number(rowStr), Number(colStr));
    }
  }
}
