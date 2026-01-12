export type CellScalar = number | string | boolean | null;

export type FillMode = "copy" | "series" | "formulas";

export interface CellRange {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

export interface FillSourceCell {
  /**
   * The canonical user input for the cell.
   *
   * - For formulas, this should be the formula text (e.g. `"=A1+B1"`).
   * - For literals, this is the scalar value (`number | string | boolean | null`).
   */
  input: CellScalar;
  /**
   * The computed value for the cell (used when filling formulas as values).
   */
  value: CellScalar;
}

export interface FillEdit {
  row: number;
  col: number;
  value: CellScalar;
}

export interface ComputeFillEditsOptions {
  sourceRange: CellRange;
  targetRange: CellRange;
  /**
   * Matrix of source cells, sized to `sourceRange` (rows first).
   */
  sourceCells: FillSourceCell[][];
  mode: FillMode;
  /**
   * Optional hook to prevent overwriting protected/locked cells.
   *
   * Return `false` to skip emitting an edit for the cell.
   */
  canWriteCell?: (row: number, col: number) => boolean;
}

export interface ComputeFillEditsResult {
  edits: FillEdit[];
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function mod(value: number, modulus: number): number {
  const rem = value % modulus;
  return rem < 0 ? rem + modulus : rem;
}

function normalizeRange(range: CellRange): CellRange {
  let { startRow, endRow, startCol, endCol } = range;
  if (!Number.isFinite(startRow) || !Number.isFinite(endRow) || !Number.isFinite(startCol) || !Number.isFinite(endCol)) {
    throw new Error("Invalid range: non-finite coordinates");
  }
  startRow = Math.trunc(startRow);
  endRow = Math.trunc(endRow);
  startCol = Math.trunc(startCol);
  endCol = Math.trunc(endCol);
  if (startRow > endRow) [startRow, endRow] = [endRow, startRow];
  if (startCol > endCol) [startCol, endCol] = [endCol, startCol];
  return { startRow, endRow, startCol, endCol };
}

function rangeHeight(range: CellRange): number {
  return Math.max(0, range.endRow - range.startRow);
}

function rangeWidth(range: CellRange): number {
  return Math.max(0, range.endCol - range.startCol);
}

function isFormulaInput(value: CellScalar): value is string {
  return typeof value === "string" && value.trimStart().startsWith("=");
}

function colNameToIndex(name: string): number {
  let acc = 0;
  for (let i = 0; i < name.length; i++) {
    const code = name.charCodeAt(i);
    const upper = code >= 97 && code <= 122 ? code - 32 : code;
    if (upper < 65 || upper > 90) return -1;
    acc = acc * 26 + (upper - 65 + 1);
  }
  return acc - 1;
}

function indexToColName(index0: number): string {
  let value = index0 + 1;
  if (!Number.isFinite(value) || value <= 0) return "";
  let out = "";
  while (value > 0) {
    const rem = (value - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    value = Math.floor((value - 1) / 26);
  }
  return out;
}

type ParsedRef = {
  sheetPrefix: string | null;
  colAbs: boolean;
  rowAbs: boolean;
  col: string;
  row: number; // 1-based
};

const SHEET_PREFIX_RE = /(?:'[^']*(?:''[^']*)*'|[A-Za-z0-9_]+)!/y;
const CELL_REF_RE = /\$?[A-Za-z]{1,3}\$?\d+/y;

function parseCellRef(text: string): ParsedRef | null {
  // Note: this parses a single token like `Sheet1!$A$1` or `$B12`.
  let i = 0;
  let sheetPrefix: string | null = null;

  SHEET_PREFIX_RE.lastIndex = 0;
  const sheetMatch = SHEET_PREFIX_RE.exec(text);
  if (sheetMatch && sheetMatch.index === 0) {
    sheetPrefix = sheetMatch[0];
    i = sheetPrefix.length;
  }

  const rest = text.slice(i);
  CELL_REF_RE.lastIndex = 0;
  const cellMatch = CELL_REF_RE.exec(rest);
  if (!cellMatch || cellMatch.index !== 0) return null;

  const token = cellMatch[0];
  const colAbs = token.startsWith("$");
  const parts = token.split("$").filter(Boolean);
  // parts could be ["A", "1"] or ["A1"] depending on `$` placement; parse manually.
  const colPartMatch = /^\$?([A-Za-z]{1,3})/.exec(token);
  const rowPartMatch = /(\$?)(\d+)$/.exec(token);
  if (!colPartMatch || !rowPartMatch) return null;
  const col = colPartMatch[1];
  const rowAbs = rowPartMatch[1] === "$";
  const row = Number.parseInt(rowPartMatch[2], 10);
  if (!Number.isFinite(row) || row <= 0) return null;

  return { sheetPrefix, colAbs, rowAbs, col, row };
}

function formatCellRef(ref: ParsedRef): string {
  const colName = ref.colAbs ? `$${ref.col}` : ref.col;
  const rowName = ref.rowAbs ? `$${ref.row}` : String(ref.row);
  return `${ref.sheetPrefix ?? ""}${colName}${rowName}`;
}

function shiftCellRef(ref: ParsedRef, deltaRow: number, deltaCol: number): string {
  const colIndex0 = colNameToIndex(ref.col);
  if (colIndex0 < 0) return formatCellRef(ref);

  const nextCol0 = ref.colAbs ? colIndex0 : colIndex0 + deltaCol;
  const nextRow1 = ref.rowAbs ? ref.row : ref.row + deltaRow;

  if (nextCol0 < 0 || nextRow1 <= 0) {
    return `${ref.sheetPrefix ?? ""}#REF!`;
  }

  const col = indexToColName(nextCol0);
  if (!col) return `${ref.sheetPrefix ?? ""}#REF!`;

  return formatCellRef({
    sheetPrefix: ref.sheetPrefix,
    colAbs: ref.colAbs,
    rowAbs: ref.rowAbs,
    col,
    row: nextRow1
  });
}

function shouldTreatAsCellRefBoundary(prev: string | undefined, next: string | undefined): boolean {
  const prevOk = !prev || !/[A-Za-z0-9_.]/.test(prev);
  if (!prevOk) return false;
  if (!next) return true;
  if (next === "(") return false;
  return !/[A-Za-z0-9_.]/.test(next);
}

/**
 * Best-effort A1-style reference shifter.
 *
 * Supports:
 * - `$A$1`, `A$1`, `$A1`, `A1`
 * - Optional sheet prefixes: `Sheet1!A1`, `'My Sheet'!A1`
 *
 * Not a full formula parser; it intentionally avoids touching content inside string literals.
 */
export function shiftFormulaA1(formula: string, deltaRow: number, deltaCol: number): string {
  if (!isFormulaInput(formula)) return formula;
  if (deltaRow === 0 && deltaCol === 0) return formula;

  let out = "";
  let i = 0;
  let inString = false;

  while (i < formula.length) {
    const ch = formula[i];
    if (ch === '"') {
      out += ch;
      if (inString && formula[i + 1] === '"') {
        // Escaped quote inside string literal.
        out += '"';
        i += 2;
        continue;
      }
      inString = !inString;
      i++;
      continue;
    }

    if (inString) {
      out += ch;
      i++;
      continue;
    }

    // Attempt to match a sheet prefix + cell ref at the current index.
    // We do this by slicing because sticky regex with alternations is much easier on the engine;
    // formulas are typically short so this is fine.
    const slice = formula.slice(i);
    const parsed = parseCellRef(slice);
    if (parsed) {
      const token = formatCellRef(parsed);
      const prev = i > 0 ? formula[i - 1] : undefined;
      const next = formula[i + token.length];

      if (shouldTreatAsCellRefBoundary(prev, next)) {
        const shifted = shiftCellRef(parsed, deltaRow, deltaCol);
        out += shifted;
        i += token.length;
        // Excel drops the spill-range operator (`#`) once the base reference becomes invalid.
        // If shifting turns a reference into `#REF!`, consume any following `#` characters so
        // `A1#` does not become `#REF!#`.
        if (shifted.endsWith("#REF!") && formula[i] === "#") {
          while (formula[i] === "#") {
            i++;
          }
        }
        continue;
      }
    }

    out += ch;
    i++;
  }

  return out;
}

type DateLike = { kind: "iso"; date: Date } | { kind: "mdy"; date: Date };

function parseDateLike(value: CellScalar): DateLike | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;

  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(trimmed);
  if (iso) {
    const year = Number.parseInt(iso[1], 10);
    const month = Number.parseInt(iso[2], 10);
    const day = Number.parseInt(iso[3], 10);
    if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
    const date = new Date(Date.UTC(year, month - 1, day));
    if (Number.isNaN(date.getTime())) return null;
    return { kind: "iso", date };
  }

  const mdy = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(trimmed);
  if (mdy) {
    const month = Number.parseInt(mdy[1], 10);
    const day = Number.parseInt(mdy[2], 10);
    let year = Number.parseInt(mdy[3], 10);
    if (mdy[3].length === 2) year += year >= 70 ? 1900 : 2000;
    const date = new Date(Date.UTC(year, month - 1, day));
    if (Number.isNaN(date.getTime())) return null;
    return { kind: "mdy", date };
  }

  return null;
}

function formatDateLike(kind: DateLike["kind"], date: Date): string {
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth() + 1;
  const day = date.getUTCDate();
  if (kind === "iso") {
    return `${year.toString().padStart(4, "0")}-${month.toString().padStart(2, "0")}-${day.toString().padStart(2, "0")}`;
  }
  // m/d/yyyy
  return `${month}/${day}/${year}`;
}

const MONTHS_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"] as const;
const MONTHS_LONG = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December"
] as const;
const DAYS_SHORT = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"] as const;
const DAYS_LONG = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"] as const;

type NameSeriesFormat = "short" | "long";
type NameSeriesCase = "upper" | "lower" | "title";

type NameSeries = {
  kind: "month" | "day";
  format: NameSeriesFormat;
  casing: NameSeriesCase;
  startIndex: number;
  step: number;
};

function detectCasing(example: string): NameSeriesCase {
  if (example.toUpperCase() === example) return "upper";
  if (example.toLowerCase() === example) return "lower";
  return "title";
}

function formatNamed(kind: NameSeries["kind"], format: NameSeriesFormat, casing: NameSeriesCase, index: number): string {
  const values = kind === "month" ? (format === "short" ? MONTHS_SHORT : MONTHS_LONG) : format === "short" ? DAYS_SHORT : DAYS_LONG;
  const raw = values[mod(index, values.length)]!;
  if (casing === "upper") return raw.toUpperCase();
  if (casing === "lower") return raw.toLowerCase();
  return raw;
}

function detectNameSeries(inputs: CellScalar[]): NameSeries | null {
  if (inputs.length === 0) return null;
  const first = inputs[0];
  if (typeof first !== "string") return null;

  const detect = (
    kind: NameSeries["kind"],
    format: NameSeriesFormat,
    values: readonly string[],
    period: number
  ): NameSeries | null => {
    const normalize = (s: string) => s.trim().toLowerCase();
    const indices: number[] = [];
    for (const input of inputs) {
      if (typeof input !== "string") return null;
      const idx = values.findIndex((v) => normalize(v) === normalize(input));
      if (idx < 0) return null;
      indices.push(idx);
    }

    const casing = detectCasing(first);
    const step = indices.length >= 2 ? indices[1]! - indices[0]! : 1;

    for (let i = 0; i < indices.length; i++) {
      const expected = mod(indices[0]! + step * i, period);
      if (indices[i] !== expected) return null;
    }

    return { kind, format, casing, startIndex: indices[0]!, step };
  };

  return (
    detect("month", "short", MONTHS_SHORT, 12) ??
    detect("month", "long", MONTHS_LONG, 12) ??
    detect("day", "short", DAYS_SHORT, 7) ??
    detect("day", "long", DAYS_LONG, 7)
  );
}

type TextNumberSeries = {
  kind: "textNumber";
  prefix: string;
  suffix: string;
  start: number;
  step: number;
  padWidth: number;
};

function detectTextNumberSeries(inputs: CellScalar[]): TextNumberSeries | null {
  if (inputs.length < 2) return null;

  const parsed = inputs.map((value) => (typeof value === "string" ? /^(.*?)(\d+)([^0-9]*)$/.exec(value) : null));
  if (parsed.some((match) => !match)) return null;

  const [first] = parsed as RegExpExecArray[];
  const prefix = first![1] ?? "";
  const suffix = first![3] ?? "";
  const padWidth = (first![2] ?? "").length;
  if (padWidth === 0) return null;

  const nums: number[] = [];
  for (const match of parsed as RegExpExecArray[]) {
    if ((match[1] ?? "") !== prefix) return null;
    if ((match[3] ?? "") !== suffix) return null;
    if ((match[2] ?? "").length !== padWidth) return null;
    nums.push(Number.parseInt(match[2]!, 10));
  }

  const step = nums[1]! - nums[0]!;
  for (let i = 2; i < nums.length; i++) {
    if (nums[i]! - nums[i - 1]! !== step) return null;
  }

  return { kind: "textNumber", prefix, suffix, start: nums[0]!, step, padWidth };
}

type SeriesPlan =
  | { kind: "number"; start: number; step: number }
  | { kind: "date"; start: Date; stepDays: number; format: DateLike["kind"] }
  | TextNumberSeries
  | NameSeries;

function detectNumberSeries(inputs: CellScalar[]): SeriesPlan | null {
  if (inputs.length < 2) return null;
  const nums: number[] = [];
  for (const input of inputs) {
    if (typeof input !== "number" || !Number.isFinite(input)) return null;
    nums.push(input);
  }

  const step = nums[1]! - nums[0]!;
  for (let i = 2; i < nums.length; i++) {
    if (nums[i]! - nums[i - 1]! !== step) return null;
  }

  return { kind: "number", start: nums[0]!, step };
}

function detectDateSeries(inputs: CellScalar[]): SeriesPlan | null {
  const parsed: DateLike[] = [];
  for (const input of inputs) {
    const d = parseDateLike(input);
    if (!d) return null;
    parsed.push(d);
  }

  if (parsed.length === 0) return null;

  const format = parsed[0]!.kind;
  if (parsed.some((d) => d.kind !== format)) return null;

  const times = parsed.map((d) => d.date.getTime());
  const msPerDay = 86_400_000;

  const stepDays = (() => {
    if (times.length >= 2) {
      return Math.round((times[1]! - times[0]!) / msPerDay);
    }
    return 1;
  })();

  // Validate a consistent progression.
  for (let i = 0; i < times.length; i++) {
    const expected = times[0]! + stepDays * msPerDay * i;
    if (times[i] !== expected) return null;
  }

  return { kind: "date", start: parsed[0]!.date, stepDays, format };
}

function detectSeriesPlan(inputs: CellScalar[]): SeriesPlan | null {
  const textNumber = detectTextNumberSeries(inputs);
  const named = detectNameSeries(inputs);
  return detectNumberSeries(inputs) ?? detectDateSeries(inputs) ?? textNumber ?? named;
}

function seriesValueAt(plan: SeriesPlan, index: number): CellScalar {
  if (plan.kind === "number") return plan.start + plan.step * index;

  if (plan.kind === "date") {
    const msPerDay = 86_400_000;
    const t = plan.start.getTime() + plan.stepDays * msPerDay * index;
    return formatDateLike(plan.format, new Date(t));
  }

  if (plan.kind === "textNumber") {
    const n = plan.start + plan.step * index;
    const digits = Math.abs(n).toString().padStart(plan.padWidth, "0");
    return `${plan.prefix}${n < 0 ? "-" : ""}${digits}${plan.suffix}`;
  }

  return formatNamed(plan.kind, plan.format, plan.casing, plan.startIndex + plan.step * index);
}

function fillValueFromSourceCell(source: FillSourceCell, deltaRow: number, deltaCol: number, mode: FillMode): CellScalar {
  if (isFormulaInput(source.input)) {
    if (mode === "copy") return source.value;
    return shiftFormulaA1(source.input, deltaRow, deltaCol);
  }
  return source.input;
}

type FillAxis = "vertical" | "horizontal";

function detectFillAxis(sourceRange: CellRange, targetRange: CellRange): FillAxis {
  const sameCols = targetRange.startCol === sourceRange.startCol && targetRange.endCol === sourceRange.endCol;
  const sameRows = targetRange.startRow === sourceRange.startRow && targetRange.endRow === sourceRange.endRow;

  const targetOutsideRows = targetRange.endRow <= sourceRange.startRow || targetRange.startRow >= sourceRange.endRow;
  const targetOutsideCols = targetRange.endCol <= sourceRange.startCol || targetRange.startCol >= sourceRange.endCol;

  if (sameCols && targetOutsideRows) return "vertical";
  if (sameRows && targetOutsideCols) return "horizontal";

  throw new Error(
    `Unsupported fill axis: source=${JSON.stringify(sourceRange)} target=${JSON.stringify(targetRange)}`
  );
}

export function computeFillEdits(options: ComputeFillEditsOptions): ComputeFillEditsResult {
  const sourceRange = normalizeRange(options.sourceRange);
  const targetRange = normalizeRange(options.targetRange);

  const sourceHeight = rangeHeight(sourceRange);
  const sourceWidth = rangeWidth(sourceRange);
  if (sourceHeight === 0 || sourceWidth === 0) return { edits: [] };

  if (options.sourceCells.length !== sourceHeight || options.sourceCells.some((row) => row.length !== sourceWidth)) {
    throw new Error(
      `sourceCells must match sourceRange (${sourceHeight}x${sourceWidth}); got ${options.sourceCells.length}x${
        options.sourceCells[0]?.length ?? 0
      }`
    );
  }

  const targetHeight = rangeHeight(targetRange);
  const targetWidth = rangeWidth(targetRange);
  if (targetHeight === 0 || targetWidth === 0) return { edits: [] };

  const axis = detectFillAxis(sourceRange, targetRange);
  const canWrite = options.canWriteCell;
  const edits: FillEdit[] = [];

  if (axis === "vertical") {
    if (targetWidth !== sourceWidth) {
      throw new Error(`Vertical fill requires matching widths; sourceWidth=${sourceWidth}, targetWidth=${targetWidth}`);
    }

    const seriesPlans: Array<SeriesPlan | null> = new Array(sourceWidth).fill(null);
    if (options.mode === "series") {
      for (let c = 0; c < sourceWidth; c++) {
        const inputs = options.sourceCells.map((row) => row[c]!.input);
        if (inputs.some((v) => isFormulaInput(v))) continue;
        seriesPlans[c] = detectSeriesPlan(inputs);
      }
    }

    for (let row = targetRange.startRow; row < targetRange.endRow; row++) {
      const index = row - sourceRange.startRow;
      for (let col = targetRange.startCol; col < targetRange.endCol; col++) {
        if (canWrite && !canWrite(row, col)) continue;
        const colOffset = col - sourceRange.startCol;
        if (colOffset < 0 || colOffset >= sourceWidth) continue;

        const plan = options.mode === "series" ? seriesPlans[colOffset] : null;
        if (plan) {
          edits.push({ row, col, value: seriesValueAt(plan, index) });
          continue;
        }

        const sourceRow = sourceRange.startRow + mod(row - sourceRange.startRow, sourceHeight);
        const sourceCol = col;
        const sourceCell = options.sourceCells[sourceRow - sourceRange.startRow]![sourceCol - sourceRange.startCol]!;
        edits.push({
          row,
          col,
          value: fillValueFromSourceCell(sourceCell, row - sourceRow, col - sourceCol, options.mode)
        });
      }
    }

    return { edits };
  }

  if (targetHeight !== sourceHeight) {
    throw new Error(
      `Horizontal fill requires matching heights; sourceHeight=${sourceHeight}, targetHeight=${targetHeight}`
    );
  }

  const seriesPlans: Array<SeriesPlan | null> = new Array(sourceHeight).fill(null);
  if (options.mode === "series") {
    for (let r = 0; r < sourceHeight; r++) {
      const inputs = options.sourceCells[r]!.map((cell) => cell.input);
      if (inputs.some((v) => isFormulaInput(v))) continue;
      seriesPlans[r] = detectSeriesPlan(inputs);
    }
  }

  for (let row = targetRange.startRow; row < targetRange.endRow; row++) {
    const rowOffset = row - sourceRange.startRow;
    if (rowOffset < 0 || rowOffset >= sourceHeight) continue;
    const plan = options.mode === "series" ? seriesPlans[rowOffset] : null;
    for (let col = targetRange.startCol; col < targetRange.endCol; col++) {
      if (canWrite && !canWrite(row, col)) continue;

      const index = col - sourceRange.startCol;
      if (plan) {
        edits.push({ row, col, value: seriesValueAt(plan, index) });
        continue;
      }

      const sourceRow = row;
      const sourceCol = sourceRange.startCol + mod(col - sourceRange.startCol, sourceWidth);
      const sourceCell = options.sourceCells[sourceRow - sourceRange.startRow]![sourceCol - sourceRange.startCol]!;
      edits.push({
        row,
        col,
        value: fillValueFromSourceCell(sourceCell, row - sourceRow, col - sourceCol, options.mode)
      });
    }
  }

  return { edits };
}
