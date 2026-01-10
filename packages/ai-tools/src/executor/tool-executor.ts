import { ZodError } from "zod";
import { columnLabelToIndex, formatA1Cell, formatA1Range, parseA1Cell, parseA1Range } from "../spreadsheet/a1.js";
import type { SpreadsheetApi } from "../spreadsheet/api.js";
import type { CellData, CellScalar } from "../spreadsheet/types.js";
import type { ToolCall, ToolName, UnknownToolCall } from "../tool-schema.js";
import { TOOL_REGISTRY, validateToolCall } from "../tool-schema.js";

export interface ToolExecutionError {
  code: "validation_error" | "not_implemented" | "permission_denied" | "runtime_error";
  message: string;
  details?: unknown;
}

export interface ToolExecutionTiming {
  started_at_ms: number;
  duration_ms: number;
}

export type ToolResultDataByName = {
  read_range: {
    range: string;
    values: CellScalar[][];
    formulas?: Array<Array<string | null>>;
  };
  write_cell: {
    cell: string;
    changed: boolean;
  };
  set_range: {
    range: string;
    updated_cells: number;
  };
  apply_formula_column: {
    sheet: string;
    column: string;
    start_row: number;
    end_row: number;
    updated_cells: number;
  };
  create_pivot_table: {
    status: "stub";
    message: string;
  };
  create_chart: {
    status: "stub";
    message: string;
  };
  sort_range: {
    range: string;
    sorted_rows: number;
  };
  filter_range: {
    range: string;
    matching_rows: number[];
    count: number;
  };
  apply_formatting: {
    range: string;
    formatted_cells: number;
  };
  detect_anomalies: {
    range: string;
    method: string;
    anomalies: Array<{ cell: string; value: number; score?: number }>;
  };
  compute_statistics: {
    range: string;
    statistics: Record<string, number | null>;
  };
  fetch_external_data: {
    url: string;
    destination: string;
    written_cells: number;
    shape: { rows: number; cols: number };
  };
};

export interface ToolExecutionResultBase<TName extends ToolName> {
  tool: TName;
  ok: boolean;
  timing: ToolExecutionTiming;
  data?: ToolResultDataByName[TName];
  warnings?: string[];
  error?: ToolExecutionError;
}

export type ToolExecutionResult = { [K in ToolName]: ToolExecutionResultBase<K> }[ToolName];

export interface ToolExecutorOptions {
  default_sheet?: string;
  allow_external_data?: boolean;
  allowed_external_hosts?: string[];
  max_external_bytes?: number;
}

export class ToolExecutor {
  readonly spreadsheet: SpreadsheetApi;
  readonly options: Required<ToolExecutorOptions>;

  constructor(spreadsheet: SpreadsheetApi, options: ToolExecutorOptions = {}) {
    this.spreadsheet = spreadsheet;
    this.options = {
      default_sheet: options.default_sheet ?? "Sheet1",
      allow_external_data: options.allow_external_data ?? false,
      allowed_external_hosts: options.allowed_external_hosts ?? [],
      max_external_bytes: options.max_external_bytes ?? 1_000_000
    };
  }

  async execute(call: UnknownToolCall): Promise<ToolExecutionResult> {
    const startedAt = nowMs();
    try {
      const validated = validateToolCall(call);
      const data = await this.executeValidated(validated);
      return {
        tool: validated.name,
        ok: true,
        timing: { started_at_ms: startedAt, duration_ms: nowMs() - startedAt },
        ...(data ? { data } : {})
      } as ToolExecutionResult;
    } catch (error) {
      const tool = ToolNameOrUnknown(call.name);
      return {
        tool,
        ok: false,
        timing: { started_at_ms: startedAt, duration_ms: nowMs() - startedAt },
        error: normalizeToolError(error)
      } as ToolExecutionResult;
    }
  }

  async executePlan(calls: UnknownToolCall[]): Promise<ToolExecutionResult[]> {
    const results: ToolExecutionResult[] = [];
    for (const call of calls) {
      results.push(await this.execute(call));
    }
    return results;
  }

  private async executeValidated(call: ToolCall): Promise<ToolResultDataByName[ToolName]> {
    switch (call.name) {
      case "read_range":
        return this.readRange(call.parameters);
      case "write_cell":
        return this.writeCell(call.parameters);
      case "set_range":
        return this.setRange(call.parameters);
      case "apply_formula_column":
        return this.applyFormulaColumn(call.parameters);
      case "create_pivot_table":
        return { status: "stub", message: "Pivot table creation is not implemented yet." };
      case "create_chart":
        return { status: "stub", message: "Chart creation is not implemented yet." };
      case "sort_range":
        return this.sortRange(call.parameters);
      case "filter_range":
        return this.filterRange(call.parameters);
      case "apply_formatting":
        return this.applyFormatting(call.parameters);
      case "detect_anomalies":
        return this.detectAnomalies(call.parameters);
      case "compute_statistics":
        return this.computeStatistics(call.parameters);
      case "fetch_external_data":
        return this.fetchExternalData(call.parameters);
      default: {
        const exhaustive: never = call.name;
        throw new Error(`Unhandled tool: ${exhaustive}`);
      }
    }
  }

  private readRange(params: any): ToolResultDataByName["read_range"] {
    const range = parseA1Range(params.range, this.options.default_sheet);
    const cells = this.spreadsheet.readRange(range);
    const values = cells.map((row) => row.map((cell) => (cell.formula ? null : cell.value)));
    const formulas = params.include_formulas
      ? cells.map((row) => row.map((cell) => cell.formula ?? null))
      : undefined;
    return { range: formatA1Range(range), values, ...(formulas ? { formulas } : {}) };
  }

  private writeCell(params: any): ToolResultDataByName["write_cell"] {
    const address = parseA1Cell(params.cell, this.options.default_sheet);
    const before = this.spreadsheet.getCell(address);

    const rest = params as { value: CellScalar; is_formula?: boolean };
    const isFormula =
      rest.is_formula === true || (typeof rest.value === "string" && rest.value.trim().startsWith("="));

    const next: CellData = isFormula
      ? { value: null, formula: String(rest.value) }
      : { value: rest.value };

    this.spreadsheet.setCell(address, next);
    const after = this.spreadsheet.getCell(address);
    return { cell: formatA1Cell(address), changed: !cellsEqual(before, after) };
  }

  private setRange(params: any): ToolResultDataByName["set_range"] {
    const range = parseA1Range(params.range, this.options.default_sheet);
    const interpretAs: "auto" | "value" | "formula" = params.interpret_as ?? "auto";

    const cells: CellData[][] = params.values.map((row: CellScalar[]) =>
      row.map((value) => {
        const formulaCandidate = typeof value === "string" ? value.trim() : "";
        const shouldTreatAsFormula =
          interpretAs === "formula" || (interpretAs === "auto" && typeof value === "string" && formulaCandidate.startsWith("="));

        if (shouldTreatAsFormula) {
          return { value: null, formula: String(value) };
        }

        return { value };
      })
    );

    this.spreadsheet.writeRange(range, cells);
    const sizeRows = range.endRow - range.startRow + 1;
    const sizeCols = range.endCol - range.startCol + 1;
    return { range: formatA1Range(range), updated_cells: sizeRows * sizeCols };
  }

  private applyFormulaColumn(params: any): ToolResultDataByName["apply_formula_column"] {
    const sheet = this.options.default_sheet;
    const column = String(params.column).trim().toUpperCase();
    const colIndex = columnLabelToIndex(column);

    const startRow = Number(params.start_row);
    const endRowRaw = Number(params.end_row ?? -1);
    const lastUsedRow = this.spreadsheet.getLastUsedRow(sheet);
    const endRow = endRowRaw === -1 ? Math.max(startRow, lastUsedRow || 0) : endRowRaw;
    if (endRow < startRow) {
      throw new Error(`apply_formula_column end_row (${endRow}) must be >= start_row (${startRow})`);
    }

    const template = String(params.formula_template);
    let updated = 0;
    for (let row = startRow; row <= endRow; row++) {
      const formula = template.replaceAll("{row}", String(row));
      this.spreadsheet.setCell({ sheet, row, col: colIndex }, { value: null, formula });
      updated++;
    }

    return { sheet, column, start_row: startRow, end_row: endRow, updated_cells: updated };
  }

  private sortRange(params: any): ToolResultDataByName["sort_range"] {
    const range = parseA1Range(params.range, this.options.default_sheet);
    const hasHeader = Boolean(params.has_header);

    const data = this.spreadsheet.readRange(range);
    const header = hasHeader ? data.slice(0, 1) : [];
    const body = hasHeader ? data.slice(1) : data.slice();

    const sortCriteria: Array<{ offset: number; order: "asc" | "desc" }> = params.sort_by.map(
      (criterion: { column: string; order?: "asc" | "desc" }) => {
        const colIndex = columnLabelToIndex(criterion.column);
        const offset = colIndex - range.startCol;
        if (offset < 0 || offset >= data[0]!.length) {
          throw new Error(`sort_range column ${criterion.column} is outside the target range`);
        }
        return { offset, order: criterion.order ?? "asc" };
      }
    );

    body.sort((left, right) => {
      for (const criterion of sortCriteria) {
        const orderMultiplier = criterion.order === "asc" ? 1 : -1;
        const result = compareCellForSort(left[criterion.offset]!, right[criterion.offset]!);
        if (result !== 0) return result * orderMultiplier;
      }
      return 0;
    });

    const sorted = [...header, ...body];
    this.spreadsheet.writeRange(range, sorted);

    return { range: formatA1Range(range), sorted_rows: body.length };
  }

  private filterRange(params: any): ToolResultDataByName["filter_range"] {
    const range = parseA1Range(params.range, this.options.default_sheet);
    const hasHeader = Boolean(params.has_header);
    const rows = this.spreadsheet.readRange(range);
    const bodyOffset = hasHeader ? 1 : 0;

    const criteria: Array<{ offset: number; operator: string; value: string | number; value2?: string | number }> =
      params.criteria.map((criterion: any) => {
        const colIndex = columnLabelToIndex(criterion.column);
        const offset = colIndex - range.startCol;
        if (offset < 0 || offset >= rows[0]!.length) {
          throw new Error(`filter_range column ${criterion.column} is outside the target range`);
        }
        return { offset, operator: criterion.operator, value: criterion.value, value2: criterion.value2 };
      });

    const matchingRows: number[] = [];
    for (let i = bodyOffset; i < rows.length; i++) {
      const row = rows[i]!;
      const matches = criteria.every((criterion) => matchesCriterion(row[criterion.offset]!, criterion));
      if (matches) {
        matchingRows.push(range.startRow + i);
      }
    }

    return { range: formatA1Range(range), matching_rows: matchingRows, count: matchingRows.length };
  }

  private applyFormatting(params: any): ToolResultDataByName["apply_formatting"] {
    const range = parseA1Range(params.range, this.options.default_sheet);
    const formatted = this.spreadsheet.applyFormatting(range, params.format);
    return { range: formatA1Range(range), formatted_cells: formatted };
  }

  private detectAnomalies(params: any): ToolResultDataByName["detect_anomalies"] {
    const range = parseA1Range(params.range, this.options.default_sheet);
    const method: string = params.method ?? "zscore";
    const cells = this.spreadsheet.readRange(range);
    const entries: Array<{ cell: string; value: number }> = [];
    for (let r = 0; r < cells.length; r++) {
      for (let c = 0; c < cells[r]!.length; c++) {
        const cell = cells[r]![c]!;
        const numeric = toNumber(cell);
        if (numeric === null) continue;
        entries.push({
          cell: formatA1Cell({ sheet: range.sheet, row: range.startRow + r, col: range.startCol + c }),
          value: numeric
        });
      }
    }

    if (entries.length === 0) {
      return { range: formatA1Range(range), method, anomalies: [] };
    }

    switch (method) {
      case "zscore": {
        const threshold = params.threshold ?? 3;
        const mean = entries.reduce((sum, e) => sum + e.value, 0) / entries.length;
        const variance =
          entries.length > 1
            ? entries.reduce((sum, e) => sum + (e.value - mean) ** 2, 0) / (entries.length - 1)
            : 0;
        const stdev = Math.sqrt(variance);
        if (stdev === 0) return { range: formatA1Range(range), method, anomalies: [] };
        const anomalies = entries
          .map((e) => ({ ...e, score: (e.value - mean) / stdev }))
          .filter((e) => Math.abs(e.score) >= threshold)
          .map((e) => ({ cell: e.cell, value: e.value, score: e.score }));
        return { range: formatA1Range(range), method, anomalies };
      }
      case "iqr": {
        const multiplier = params.threshold ?? 1.5;
        const sorted = [...entries].sort((a, b) => a.value - b.value);
        const q1 = quantile(sorted.map((e) => e.value), 0.25);
        const q3 = quantile(sorted.map((e) => e.value), 0.75);
        const iqr = q3 - q1;
        const low = q1 - multiplier * iqr;
        const high = q3 + multiplier * iqr;
        const anomalies = entries
          .filter((e) => e.value < low || e.value > high)
          .map((e) => ({ cell: e.cell, value: e.value }));
        return { range: formatA1Range(range), method, anomalies };
      }
      case "isolation_forest":
        throw toolError("not_implemented", "detect_anomalies method isolation_forest is not implemented yet.");
      default:
        throw new Error(`Unsupported detect_anomalies method: ${method}`);
    }
  }

  private computeStatistics(params: any): ToolResultDataByName["compute_statistics"] {
    const range = parseA1Range(params.range, this.options.default_sheet);
    const measures: string[] = params.measures ?? [];
    const cells = this.spreadsheet.readRange(range);
    const values: number[] = [];
    for (const row of cells) {
      for (const cell of row) {
        const numeric = toNumber(cell);
        if (numeric === null) continue;
        values.push(numeric);
      }
    }

    const stats: Record<string, number | null> = {};
    for (const measure of measures) {
      switch (measure) {
        case "mean":
          stats.mean = values.length ? values.reduce((sum, v) => sum + v, 0) / values.length : null;
          break;
        case "median":
          stats.median = values.length ? median(values) : null;
          break;
        case "mode":
          stats.mode = values.length ? mode(values) : null;
          break;
        case "stdev":
          stats.stdev = values.length ? stdev(values) : null;
          break;
        case "variance":
          stats.variance = values.length ? variance(values) : null;
          break;
        case "min":
          stats.min = values.length ? Math.min(...values) : null;
          break;
        case "max":
          stats.max = values.length ? Math.max(...values) : null;
          break;
        case "quartiles": {
          if (!values.length) {
            stats.q1 = null;
            stats.q2 = null;
            stats.q3 = null;
            break;
          }
          const sorted = [...values].sort((a, b) => a - b);
          stats.q1 = quantile(sorted, 0.25);
          stats.q2 = quantile(sorted, 0.5);
          stats.q3 = quantile(sorted, 0.75);
          break;
        }
        case "correlation": {
          const cols = range.endCol - range.startCol + 1;
          if (cols !== 2) {
            stats.correlation = null;
            break;
          }
          const pairs: Array<[number, number]> = [];
          for (const row of cells) {
            const left = toNumber(row[0]!);
            const right = toNumber(row[1]!);
            if (left === null || right === null) continue;
            pairs.push([left, right]);
          }
          stats.correlation = pairs.length ? correlation(pairs) : null;
          break;
        }
        default:
          stats[measure] = null;
      }
    }

    return { range: formatA1Range(range), statistics: stats };
  }

  private async fetchExternalData(params: any): Promise<ToolResultDataByName["fetch_external_data"]> {
    if (!this.options.allow_external_data) {
      throw toolError("permission_denied", "fetch_external_data is disabled by default.");
    }

    const url = new URL(params.url);
    if (this.options.allowed_external_hosts.length > 0 && !this.options.allowed_external_hosts.includes(url.host)) {
      throw toolError(
        "permission_denied",
        `External host "${url.host}" is not in the allowlist for fetch_external_data.`
      );
    }

    const response = await fetch(url.toString(), {
      headers: params.headers ?? undefined
    });

    const contentLength = response.headers.get("content-length");
    if (contentLength && Number(contentLength) > this.options.max_external_bytes) {
      throw toolError(
        "permission_denied",
        `External response too large (${contentLength} bytes). Increase max_external_bytes to allow.`
      );
    }

    if (!response.ok) {
      throw toolError("runtime_error", `External fetch failed with HTTP ${response.status}`);
    }

    const destination = parseA1Cell(params.destination, this.options.default_sheet);

    if (params.transform === "raw_text") {
      const text = await response.text();
      this.spreadsheet.setCell(destination, { value: text });
      return { url: url.toString(), destination: formatA1Cell(destination), written_cells: 1, shape: { rows: 1, cols: 1 } };
    }

    const json = await response.json();
    const table = jsonToTable(json);
    const range = {
      sheet: destination.sheet,
      startRow: destination.row,
      startCol: destination.col,
      endRow: destination.row + table.length - 1,
      endCol: destination.col + (table[0]?.length ?? 1) - 1
    };

    const cells: CellData[][] = table.map((row) => row.map((value) => ({ value })));
    this.spreadsheet.writeRange(range, cells);

    return {
      url: url.toString(),
      destination: formatA1Cell(destination),
      written_cells: table.length * (table[0]?.length ?? 0),
      shape: { rows: table.length, cols: table[0]?.length ?? 0 }
    };
  }
}

function nowMs(): number {
  if (typeof performance !== "undefined" && typeof performance.now === "function") return performance.now();
  return Date.now();
}

function normalizeToolError(error: unknown): ToolExecutionError {
  if (isToolError(error)) return error;

  if (error instanceof ZodError) {
    return { code: "validation_error", message: "Tool parameters failed validation.", details: error.flatten() };
  }

  if (error instanceof Error) {
    return { code: "runtime_error", message: error.message };
  }

  return { code: "runtime_error", message: "Unknown tool execution error." };
}

function isToolError(value: unknown): value is ToolExecutionError {
  return (
    typeof value === "object" &&
    value !== null &&
    "code" in value &&
    "message" in value &&
    typeof (value as any).code === "string"
  );
}

function toolError(code: ToolExecutionError["code"], message: string, details?: unknown): ToolExecutionError {
  return { code, message, ...(details ? { details } : {}) };
}

function ToolNameOrUnknown(name: string): ToolName {
  return ToolNameSchemaSafe(name) ?? "read_range";
}

function ToolNameSchemaSafe(name: string): ToolName | null {
  if (!name) return null;
  return Object.prototype.hasOwnProperty.call(TOOL_REGISTRY, name) ? (name as ToolName) : null;
}

function compareCellForSort(left: CellData, right: CellData): number {
  const leftValue = cellComparableValue(left);
  const rightValue = cellComparableValue(right);
  return compareScalars(leftValue, rightValue);
}

function cellComparableValue(cell: CellData): string | number | boolean | null {
  if (cell.formula) return cell.formula;
  return cell.value;
}

function compareScalars(left: CellScalar | string, right: CellScalar | string): number {
  if (left === right) return 0;
  if (left === null) return -1;
  if (right === null) return 1;

  if (typeof left === "number" && typeof right === "number") return left - right;
  return String(left).localeCompare(String(right));
}

function matchesCriterion(cell: CellData, criterion: { operator: string; value: string | number; value2?: string | number }): boolean {
  const comparable = cellComparableValue(cell);
  switch (criterion.operator) {
    case "equals":
      return String(comparable ?? "") === String(criterion.value);
    case "contains":
      return String(comparable ?? "").includes(String(criterion.value));
    case "greater": {
      const a = Number(comparable);
      const b = Number(criterion.value);
      return Number.isFinite(a) && Number.isFinite(b) && a > b;
    }
    case "less": {
      const a = Number(comparable);
      const b = Number(criterion.value);
      return Number.isFinite(a) && Number.isFinite(b) && a < b;
    }
    case "between": {
      if (criterion.value2 === undefined) return false;
      const a = Number(comparable);
      const low = Number(criterion.value);
      const high = Number(criterion.value2);
      return Number.isFinite(a) && Number.isFinite(low) && Number.isFinite(high) && a >= low && a <= high;
    }
    default:
      return false;
  }
}

function toNumber(cell: CellData): number | null {
  if (cell.formula) return null;
  if (typeof cell.value === "number") return cell.value;
  if (typeof cell.value === "string") {
    const num = Number(cell.value);
    return Number.isFinite(num) ? num : null;
  }
  return null;
}

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  return quantile(sorted, 0.5);
}

function quantile(sortedValues: number[], q: number): number {
  if (sortedValues.length === 0) return NaN;
  const sorted = [...sortedValues].sort((a, b) => a - b);
  const pos = (sorted.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  if (sorted[base + 1] === undefined) return sorted[base]!;
  return sorted[base]! + rest * (sorted[base + 1]! - sorted[base]!);
}

function mode(values: number[]): number | null {
  const counts = new Map<number, number>();
  for (const value of values) {
    counts.set(value, (counts.get(value) ?? 0) + 1);
  }
  let maxCount = 0;
  let modeValue: number | null = null;
  for (const [value, count] of counts.entries()) {
    if (count > maxCount) {
      maxCount = count;
      modeValue = value;
    }
  }
  return maxCount > 1 ? modeValue : null;
}

function variance(values: number[]): number {
  if (values.length < 2) return 0;
  const mean = values.reduce((sum, v) => sum + v, 0) / values.length;
  return values.reduce((sum, v) => sum + (v - mean) ** 2, 0) / (values.length - 1);
}

function stdev(values: number[]): number {
  return Math.sqrt(variance(values));
}

function correlation(pairs: Array<[number, number]>): number {
  const xs = pairs.map(([x]) => x);
  const ys = pairs.map(([, y]) => y);
  const meanX = xs.reduce((sum, x) => sum + x, 0) / xs.length;
  const meanY = ys.reduce((sum, y) => sum + y, 0) / ys.length;
  let numerator = 0;
  let denomX = 0;
  let denomY = 0;
  for (let i = 0; i < pairs.length; i++) {
    const dx = xs[i]! - meanX;
    const dy = ys[i]! - meanY;
    numerator += dx * dy;
    denomX += dx ** 2;
    denomY += dy ** 2;
  }
  const denominator = Math.sqrt(denomX * denomY);
  return denominator === 0 ? 0 : numerator / denominator;
}

function cellsEqual(left: CellData, right: CellData): boolean {
  if (left.value !== right.value) return false;
  if ((left.formula ?? null) !== (right.formula ?? null)) return false;
  const leftFormat = left.format ?? {};
  const rightFormat = right.format ?? {};
  const leftKeys = Object.keys(leftFormat);
  const rightKeys = Object.keys(rightFormat);
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every((key) => (leftFormat as any)[key] === (rightFormat as any)[key]);
}

function jsonToTable(payload: unknown): CellScalar[][] {
  if (Array.isArray(payload)) {
    if (payload.length === 0) return [[null]];
    if (payload.every((row) => Array.isArray(row))) {
      const rows = (payload as unknown[]).map((row) => (row as unknown[]).map((value) => normalizeJsonScalar(value)));
      const maxCols = rows.reduce((max, row) => Math.max(max, row.length), 0);
      const normalizedCols = Math.max(maxCols, 1);
      return rows.map((row) => [...row, ...new Array(normalizedCols - row.length).fill(null)]);
    }
    if (payload.every((row) => row && typeof row === "object" && !Array.isArray(row))) {
      const objects = payload as Array<Record<string, unknown>>;
      const headers = Array.from(new Set(objects.flatMap((obj) => Object.keys(obj))));
      const rows = objects.map((obj) => headers.map((header) => normalizeJsonScalar(obj[header])));
      if (headers.length === 0) return [[null]];
      return [headers, ...rows];
    }
    return [(payload as unknown[]).map((value) => normalizeJsonScalar(value))];
  }

  if (payload && typeof payload === "object") {
    const obj = payload as Record<string, unknown>;
    const headers = Object.keys(obj);
    const row = headers.map((header) => normalizeJsonScalar(obj[header]));
    if (headers.length === 0) return [[null]];
    return [headers, row];
  }

  return [[normalizeJsonScalar(payload)]];
}

function normalizeJsonScalar(value: unknown): CellScalar {
  if (value === null || value === undefined) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  return JSON.stringify(value);
}
