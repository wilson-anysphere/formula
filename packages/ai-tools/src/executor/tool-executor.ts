import { ZodError } from "zod";
import { columnLabelToIndex, formatA1Cell, formatA1Range, parseA1Cell, parseA1Range } from "../spreadsheet/a1.ts";
import type { ChartType, CreateChartResult, CreateChartSpec, SpreadsheetApi } from "../spreadsheet/api.ts";
import type { CellData, CellScalar } from "../spreadsheet/types.ts";
import type { PivotAggregationType, ToolCall, ToolName, UnknownToolCall } from "../tool-schema.ts";
import { TOOL_REGISTRY, validateToolCall } from "../tool-schema.ts";

import { redactUrlSecrets } from "../utils/urlRedaction.ts";

import { DLP_ACTION } from "../../../security/dlp/src/actions.js";
import { DLP_DECISION, evaluatePolicy } from "../../../security/dlp/src/policyEngine.js";
import { CLASSIFICATION_LEVEL, classificationRank, maxClassification } from "../../../security/dlp/src/classification.js";
import { effectiveCellClassification, effectiveRangeClassification, normalizeRange } from "../../../security/dlp/src/selectors.js";

import { parseSpreadsheetNumber } from "./number-parsing.ts";

const DEFAULT_CLASSIFICATION_RANK = classificationRank(CLASSIFICATION_LEVEL.PUBLIC);
const RESTRICTED_CLASSIFICATION_RANK = classificationRank(CLASSIFICATION_LEVEL.RESTRICTED);

function normalizeFormulaTextOpt(value: unknown): string | null {
  if (value == null) return null;
  const trimmed = String(value).trim();
  const strippedLeading = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const stripped = strippedLeading.trim();
  if (stripped === "") return null;
  return `=${stripped}`;
}

const DEFAULT_READ_RANGE_MAX_CELL_CHARS = 10_000;
const UNSERIALIZABLE_CELL_VALUE_PLACEHOLDER = "[Unserializable cell value]";

function truncateCellString(value: string, maxChars: number): string {
  const limit = Number.isFinite(maxChars) ? Math.max(0, Math.floor(maxChars)) : DEFAULT_READ_RANGE_MAX_CELL_CHARS;
  if (limit === 0) return "";
  if (value.length <= limit) return value;
  const truncated = value.length - limit;
  return `${value.slice(0, limit)}…[truncated ${truncated} chars]`;
}

function stringifyCellValue(value: unknown): string {
  try {
    const json = JSON.stringify(value);
    if (typeof json === "string") return json;
  } catch {
    // ignore
  }
  try {
    return String(value);
  } catch {
    return UNSERIALIZABLE_CELL_VALUE_PLACEHOLDER;
  }
}

function normalizeCellOutput(value: unknown, opts: { maxChars?: number } = {}): CellScalar {
  const maxChars = opts.maxChars ?? DEFAULT_READ_RANGE_MAX_CELL_CHARS;

  // Normalize `undefined` to null for schema safety.
  if (value === null || value === undefined) return null;

  if (typeof value === "string") return truncateCellString(value, maxChars);
  if (typeof value === "number" || typeof value === "boolean") return value;

  return truncateCellString(stringifyCellValue(value), maxChars);
}

function normalizeFormulaOutput(value: unknown): string | null {
  if (value === null || value === undefined) return null;
  if (typeof value === "string") return value;
  const normalized = normalizeCellOutput(value);
  return normalized === null ? null : String(normalized);
}

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
    status: "ok";
    source_range: string;
    destination_range: string;
    written_cells: number;
    shape: { rows: number; cols: number };
  };
  create_chart: {
    status: "ok";
    chart_id: string;
    chart_type: ChartType;
    data_range: string;
    position?: string;
    title?: string;
  };
  sort_range: {
    range: string;
    sorted_rows: number;
  };
  filter_range: {
    range: string;
    matching_rows: number[];
    count: number;
    truncated?: boolean;
  };
  apply_formatting: {
    range: string;
    formatted_cells: number;
  };
  detect_anomalies:
    | {
        range: string;
        method: "iqr";
        anomalies: Array<{ cell: string; value: number | null }>;
        truncated?: boolean;
        total_anomalies?: number;
      }
    | {
        range: string;
        method: "zscore";
        anomalies: Array<{ cell: string; value: number | null; score: number | null }>;
        truncated?: boolean;
        total_anomalies?: number;
      }
    | {
        range: string;
        method: "isolation_forest";
        anomalies: Array<{ cell: string; value: number | null; score: number | null }>;
        truncated?: boolean;
        total_anomalies?: number;
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
    fetched_at_ms: number;
    content_type?: string;
    content_length_bytes?: number;
    status_code: number;
    truncated?: boolean;
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
  /**
   * Optional sheet name resolver.
   *
   * Some host applications keep a stable internal sheet id even after a user renames the
   * sheet (id != display name). In those cases, tools may receive A1 references that use
   * the *display name* (e.g. "Budget!A1") while internal systems (SpreadsheetApi/DLP)
   * expect the stable id (e.g. "Sheet2").
   *
   * When provided, ToolExecutor will:
   * - canonicalize parsed sheet tokens to stable ids before calling SpreadsheetApi and evaluating DLP
   * - format tool result A1 references using the display name for readability
   */
  sheet_name_resolver?: {
    getSheetIdByName(name: string): string | null;
    getSheetNameById(id: string): string | null;
  } | null;
  allow_external_data?: boolean;
  /**
   * When true, ToolExecutor should behave as if it is running in a side-effect-free
   * preview/simulation environment.
   *
   * Preview mode MUST NOT perform network access or mutate the provided SpreadsheetApi.
   * Instead, tools should return deterministic "skipped" results where appropriate.
   */
  preview_mode?: boolean;
  /**
   * Explicit allowlist for `fetch_external_data`.
   *
   * Entries can be either:
   * - `example.com` (hostname-only): matches `url.hostname` regardless of port.
   * - `example.com:8443` (host:port): matches `url.hostname` + port. For default ports (80/443),
   *   URLs that omit an explicit port are treated as using the scheme default.
   *
   * Matching is case-insensitive and whitespace is trimmed.
   */
  allowed_external_hosts?: string[];
  max_external_bytes?: number;
  /**
   * When enabled, tools treat formula cells as having a computed value (via `cell.value`)
   * instead of always treating them as `null`.
   *
   * This is opt-in because many backends (including the in-memory workbook) do not evaluate
   * formulas and therefore store `value:null` for formula cells.
   *
   * DLP-safe default: when DLP is configured, formula values are only surfaced when the
   * range-level decision is `ALLOW`. Under `REDACT`, formula values remain `null` to avoid
   * leaking restricted content via computed results.
   *
   * Note: DLP decisions are evaluated over the selected range only. ToolExecutor does not
   * attempt to trace formula dependencies; hosts that compute formula values should ensure
   * `cell.value` does not reflect restricted content that is out-of-scope for the current DLP
   * evaluation.
   */
  include_formula_values?: boolean;
  /**
   * Hard cap on the number of cells `read_range` is allowed to return.
   *
   * This prevents accidental/looping tool calls from returning massive matrices
   * that can overflow LLM context windows and/or audit log storage.
   *
   * When exceeded, `read_range` returns `ok:false` with `permission_denied` and
   * suggests requesting a smaller range (or raising this limit explicitly).
   */
  max_read_range_cells?: number;
  /**
   * Approximate cap on the size of `read_range` results (in characters).
   *
   * This is a safeguard for cases where a "small" range contains very large
   * strings (e.g. pasted documents in cells). The cap is enforced on the sum of
   * returned scalar lengths and is intentionally conservative.
   */
  max_read_range_chars?: number;
  /**
   * Hard cap on the number of cells a tool is allowed to materialize into JS memory
   * when operating over a full rectangular range.
   *
   * Unlike `max_read_range_cells` (which specifically limits `read_range` tool output so
   * LLM context isn't flooded), this limit applies to *other* tools that internally
   * call `SpreadsheetApi.readRange` (e.g. `sort_range`, `filter_range`, `detect_anomalies`,
   * `compute_statistics`, `create_pivot_table`).
   *
   * This is a safety guard for Excel-scale sheets: without it, a single tool call could
   * attempt to allocate millions of cells worth of JS objects, leading to renderer OOMs.
   *
   * When exceeded, the tool call returns `ok:false` with `permission_denied` and suggests
   * requesting a smaller range (or raising this limit explicitly).
   */
  max_tool_range_cells?: number;
  /**
   * Cap the number of matching row indices returned by `filter_range`.
   *
   * The tool still reports the full match count via `count`, but truncates the
   * `matching_rows` list when necessary.
   */
  max_filter_range_matching_rows?: number;
  /**
   * Cap the number of anomalies returned by `detect_anomalies`.
   */
  max_detect_anomalies?: number;
  /**
   * Optional DLP enforcement for tool results.
   *
   * IMPORTANT: Tool results are fed back into the LLM context as `role:"tool"`
   * messages by `runChatWithTools`. If you use cloud LLMs, sensitive data must be
   * blocked/redacted here (not only when building prompt context).
   */
  dlp?: {
    document_id: string;
    sheet_id?: string; // default_sheet if omitted
    policy: any;
    classification_records?: Array<{ selector: any; classification: any }>;
    classification_store?: { list(documentId: string): Array<{ selector: any; classification: any }> };
    /**
     * Optional resolver for table-based column selectors.
     *
     * DLP records may express column scopes using `(tableId, columnId)` pairs instead of
     * absolute sheet column indices. ToolExecutor operates on sheet coordinates, so hosts
     * can optionally provide a resolver that maps a table column to a 0-based sheet
     * `columnIndex`.
     *
     * When provided, ToolExecutor will enforce DLP policies for table-based column
     * selectors during both policy evaluation and per-cell redaction.
     */
    table_column_resolver?: {
      getColumnIndex(sheetId: string, tableId: string, columnId: string): number | null;
    };
    include_restricted_content?: boolean;
    audit_logger?: { log(event: any): void };
  };
}

const DLP_REDACTION_PLACEHOLDER = "[REDACTED]";

type ResolvedToolExecutorOptions = Required<Omit<ToolExecutorOptions, "dlp">> & { dlp?: ToolExecutorOptions["dlp"] };

type DlpNormalizedRange = ReturnType<typeof normalizeRange>;

type DlpRangeIndex = {
  /**
   * Max document-level classification rank for the document in scope.
   */
  docRankMax: number;
  /**
   * Max sheet-level classification rank for the sheet in scope.
   */
  sheetRankMax: number;
  /**
   * Cached `max(docRankMax, sheetRankMax)` used as the starting rank for per-cell checks.
   */
  baseRank: number;
  /**
   * Cached selection bounds (0-based) used for fast column/cell array indexing.
   */
  startRow: number;
  startCol: number;
  rowCount: number;
  colCount: number;
  /**
   * Max classification rank for each 0-based column offset in the selection range.
   */
  columnRankByOffset: Uint8Array;
  /**
   * Max classification rank for each cell in the selection range, stored row-major.
   *
   * Null when there are no non-Public cell-scoped classification records intersecting the selection.
   */
  cellRankByOffset: Uint8Array | null;
  /**
   * Range-scoped records for the sheet (normalized to ensure start <= end).
   */
  rangeRecords: Array<{ startRow: number; endRow: number; startCol: number; endCol: number; rank: number }>;
  /**
   * Max rank across all range-scoped selectors intersecting the selection.
   *
   * Used to skip per-cell range scanning when the current effective rank is already >= the
   * maximum possible contribution from any range record.
   */
  rangeRankMax: number;
  /**
   * Records that cannot be indexed by (row,col)/(columnIndex).
   *
   * ToolExecutor currently operates on sheet coordinates only (no table metadata). When
   * a `table_column_resolver` is provided, table-based column selectors (tableId/columnId)
   * can be resolved to sheet `columnIndex` values and will be indexed normally. Records
   * that still cannot be indexed are kept here for future extensibility.
   */
  fallbackRecords: Array<{ selector: any; classification: any }>;
};

export class ToolExecutor {
  readonly spreadsheet: SpreadsheetApi;
  readonly options: ResolvedToolExecutorOptions;
  private readonly pivots: PivotRegistration[] = [];

  constructor(spreadsheet: SpreadsheetApi, options: ToolExecutorOptions = {}) {
    this.spreadsheet = spreadsheet;
    const sheetNameResolver = options.sheet_name_resolver ?? null;
    const canonicalizeSheetId = (sheet: string) => {
      const raw = String(sheet ?? "").trim();
      if (!raw) return raw;
      if (!sheetNameResolver) return raw;
      return sheetNameResolver.getSheetIdByName(raw) ?? raw;
    };
    const canonicalDefaultSheet = canonicalizeSheetId(options.default_sheet ?? "Sheet1");
    const canonicalDlpSheetId =
      options.dlp && typeof options.dlp.sheet_id === "string" && options.dlp.sheet_id.trim()
        ? canonicalizeSheetId(options.dlp.sheet_id)
        : undefined;
    this.options = {
      default_sheet: canonicalDefaultSheet || "Sheet1",
      sheet_name_resolver: sheetNameResolver,
      allow_external_data: options.allow_external_data ?? false,
      preview_mode: options.preview_mode ?? false,
      allowed_external_hosts: (options.allowed_external_hosts ?? [])
        .map((host) => String(host).trim().toLowerCase())
        .filter((host) => host.length > 0),
      max_external_bytes: options.max_external_bytes ?? 1_000_000,
      include_formula_values: options.include_formula_values ?? false,
      max_read_range_cells: options.max_read_range_cells ?? 5_000,
      max_read_range_chars: options.max_read_range_chars ?? 200_000,
      // Many tool implementations materialize a full `CellData[][]` grid in JS (e.g. sort/filter).
      // Keep this bounded so Excel-scale grid limits can't trigger catastrophic allocations.
      max_tool_range_cells: options.max_tool_range_cells ?? 200_000,
      max_filter_range_matching_rows: options.max_filter_range_matching_rows ?? 1_000,
      max_detect_anomalies: options.max_detect_anomalies ?? 1_000,
      dlp:
        options.dlp && canonicalDlpSheetId
          ? {
              ...options.dlp,
              sheet_id: canonicalDlpSheetId
            }
          : options.dlp
    };
  }

  private assertRangeWithinMaxToolCells(
    tool: ToolName,
    range: { sheet: string; startRow: number; endRow: number; startCol: number; endCol: number },
    opts: { label?: string } = {},
  ): void {
    const maxCells = this.options.max_tool_range_cells;
    if (!Number.isFinite(maxCells)) return;
    if (maxCells <= 0) return;
    const requestedCells = rangeCellCount(range);
    if (requestedCells <= maxCells) return;
    const label = opts.label ? `${opts.label} ` : "";
    throw toolError(
      "permission_denied",
      `${tool} ${label}requested ${requestedCells} cells (${this.formatRangeForUser(range)}), which exceeds max_tool_range_cells (${maxCells}). Request a smaller range or increase max_tool_range_cells.`
    );
  }

  private resolveSheetId(sheetNameOrId: string): string {
    const trimmed = String(sheetNameOrId ?? "").trim();
    if (!trimmed) return trimmed;
    const resolver = this.options.sheet_name_resolver;
    if (!resolver) return trimmed;
    return resolver.getSheetIdByName(trimmed) ?? trimmed;
  }

  private displaySheetName(sheetId: string): string {
    const resolver = this.options.sheet_name_resolver;
    if (!resolver) return sheetId;
    return resolver.getSheetNameById(sheetId) ?? sheetId;
  }

  private parseRange(ref: unknown, defaultSheet: string): ReturnType<typeof parseA1Range> {
    const parsed = parseA1Range(ref as any, defaultSheet);
    const sheet = this.resolveSheetId(parsed.sheet);
    return sheet === parsed.sheet ? parsed : { ...parsed, sheet };
  }

  private parseCell(ref: unknown, defaultSheet: string): ReturnType<typeof parseA1Cell> {
    const parsed = parseA1Cell(ref as any, defaultSheet);
    const sheet = this.resolveSheetId(parsed.sheet);
    return sheet === parsed.sheet ? parsed : { ...parsed, sheet };
  }

  private formatRangeForUser(range: ReturnType<typeof parseA1Range>): string {
    return formatA1Range({ ...range, sheet: this.displaySheetName(range.sheet) });
  }

  private formatCellForUser(cell: ReturnType<typeof parseA1Cell>): string {
    return formatA1Cell({ ...cell, sheet: this.displaySheetName(cell.sheet) });
  }

  async execute(call: UnknownToolCall): Promise<ToolExecutionResult> {
    const startedAt = nowMs();
    try {
      const validated = validateToolCall(call);
      const { data, warnings } = await this.executeValidated(validated);
      return {
        tool: validated.name,
        ok: true,
        timing: { started_at_ms: startedAt, duration_ms: nowMs() - startedAt },
        ...(data !== undefined ? { data } : {}),
        ...(warnings && warnings.length > 0 ? { warnings } : {})
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

  private async executeValidated(
    call: ToolCall
  ): Promise<{ data?: ToolResultDataByName[ToolName]; warnings?: string[] }> {
    switch (call.name) {
      case "read_range":
        return { data: this.readRange(call.parameters) };
      case "write_cell":
        return { data: this.writeCell(call.parameters) };
      case "set_range":
        return { data: this.setRange(call.parameters) };
      case "apply_formula_column":
        return { data: this.applyFormulaColumn(call.parameters) };
      case "create_pivot_table":
        return { data: this.createPivotTable(call.parameters) };
      case "create_chart":
        return { data: this.createChart(call.parameters) };
      case "sort_range":
        return { data: this.sortRange(call.parameters) };
      case "filter_range":
        return { data: this.filterRange(call.parameters) };
      case "apply_formatting":
        return { data: this.applyFormatting(call.parameters) };
      case "detect_anomalies":
        return { data: this.detectAnomalies(call.parameters) };
      case "compute_statistics":
        return { data: this.computeStatistics(call.parameters) };
      case "fetch_external_data":
        return this.fetchExternalData(call.parameters);
      default: {
        const exhaustive: never = call.name;
        throw new Error(`Unhandled tool: ${exhaustive}`);
      }
    }
  }

  private readRange(params: any): ToolResultDataByName["read_range"] {
    const range = this.parseRange(params.range, this.options.default_sheet);
    const requestedCells = rangeCellCount(range);
    if (requestedCells > this.options.max_read_range_cells) {
      throw toolError(
        "permission_denied",
        `read_range requested ${requestedCells} cells (${this.formatRangeForUser(range)}), which exceeds max_read_range_cells (${this.options.max_read_range_cells}). Request a smaller range or increase max_read_range_cells.`
      );
    }

    const dlp = this.evaluateDlpForRange("read_range", range);
    if (dlp && dlp.decision.decision === DLP_DECISION.BLOCK) {
      this.logToolDlpDecision({
        tool: "read_range",
        range,
        dlp,
        redactedCellCount: 0
      });
      throw toolError(
        "permission_denied",
        `DLP policy blocks reading ${this.formatRangeForUser(range)} via read_range (ai.cloudProcessing).`
      );
    }

    const cells = this.spreadsheet.readRange(range);
    if (!dlp || dlp.decision.decision === DLP_DECISION.ALLOW) {
      const includeFormulaValues = Boolean(this.options.include_formula_values);
      const values = cells.map((row) =>
        row.map((cell) => normalizeCellOutput(cell.formula ? (includeFormulaValues ? cell.value : null) : cell.value))
      );
      const formulas = params.include_formulas
        ? cells.map((row) => row.map((cell) => normalizeFormulaOutput(cell.formula)))
        : undefined;
      if (dlp) {
        this.logToolDlpDecision({ tool: "read_range", range, dlp, redactedCellCount: 0 });
      }
      const rangeForUser = { ...range, sheet: this.displaySheetName(range.sheet) };
      enforceReadRangeCharLimit({ range: rangeForUser, values, formulas, maxChars: this.options.max_read_range_chars });
      return { range: this.formatRangeForUser(range), values, ...(formulas ? { formulas } : {}) };
    }

    let redactedCellCount = 0;

    // When formulas are requested, compute values + formulas in a single pass so we only
    // evaluate DLP once per cell (avoids per-cell Map caching overhead).
    if (params.include_formulas) {
      const values: CellScalar[][] = [];
      const formulas: Array<Array<string | null>> = [];
      for (let r = 0, rowIndex = range.startRow; r < cells.length; r += 1, rowIndex += 1) {
        const row = cells[r] ?? [];
        const valuesRow: CellScalar[] = new Array(row.length);
        const formulasRow: Array<string | null> = new Array(row.length);
        for (let c = 0, colIndex = range.startCol; c < row.length; c += 1, colIndex += 1) {
          const cell = row[c]!;
          if (!this.isDlpCellAllowed(dlp, rowIndex, colIndex)) {
            redactedCellCount += 1;
            valuesRow[c] = DLP_REDACTION_PLACEHOLDER;
            formulasRow[c] = DLP_REDACTION_PLACEHOLDER;
            continue;
          }
          valuesRow[c] = normalizeCellOutput(cell.formula ? null : cell.value);
          formulasRow[c] = normalizeFormulaOutput(cell.formula);
        }
        values.push(valuesRow);
        formulas.push(formulasRow);
      }

      this.logToolDlpDecision({ tool: "read_range", range, dlp, redactedCellCount });
      const rangeForUser = { ...range, sheet: this.displaySheetName(range.sheet) };
      enforceReadRangeCharLimit({ range: rangeForUser, values, formulas, maxChars: this.options.max_read_range_chars });
      return { range: this.formatRangeForUser(range), values, formulas };
    }

    const values: CellScalar[][] = [];
    for (let r = 0, rowIndex = range.startRow; r < cells.length; r += 1, rowIndex += 1) {
      const row = cells[r] ?? [];
      const valuesRow: CellScalar[] = new Array(row.length);
      for (let c = 0, colIndex = range.startCol; c < row.length; c += 1, colIndex += 1) {
        const cell = row[c]!;
        if (!this.isDlpCellAllowed(dlp, rowIndex, colIndex)) {
          redactedCellCount += 1;
          valuesRow[c] = DLP_REDACTION_PLACEHOLDER;
          continue;
        }
        valuesRow[c] = normalizeCellOutput(cell.formula ? null : cell.value);
      }
      values.push(valuesRow);
    }

    const formulas = undefined;

    this.logToolDlpDecision({ tool: "read_range", range, dlp, redactedCellCount });
    const rangeForUser = { ...range, sheet: this.displaySheetName(range.sheet) };
    enforceReadRangeCharLimit({ range: rangeForUser, values, formulas, maxChars: this.options.max_read_range_chars });
    return { range: this.formatRangeForUser(range), values, ...(formulas ? { formulas } : {}) };
  }

  private writeCell(params: any): ToolResultDataByName["write_cell"] {
    const address = this.parseCell(params.cell, this.options.default_sheet);
    const range = { sheet: address.sheet, startRow: address.row, endRow: address.row, startCol: address.col, endCol: address.col };
    const dlp = this.evaluateDlpForRange("write_cell", range);
    const shouldMaskChanged = Boolean(dlp && dlp.decision.decision !== DLP_DECISION.ALLOW);
    const before = shouldMaskChanged ? null : this.spreadsheet.getCell(address);

    const rest = params as { value: CellScalar; is_formula?: boolean };
    const shouldTreatAsFormula =
      rest.is_formula === true || (typeof rest.value === "string" && rest.value.trimStart().startsWith("="));
    const normalizedFormula = shouldTreatAsFormula ? normalizeFormulaTextOpt(rest.value) : null;

    const next: CellData =
      normalizedFormula != null ? { value: null, formula: normalizedFormula } : { value: shouldTreatAsFormula ? null : rest.value };

    this.spreadsheet.setCell(address, next);
    this.refreshPivotsForRange({
      sheet: address.sheet,
      startRow: address.row,
      endRow: address.row,
      startCol: address.col,
      endCol: address.col
    });
    const changed = shouldMaskChanged ? true : !cellsEqual(before!, this.spreadsheet.getCell(address));
    if (dlp) this.logToolDlpDecision({ tool: "write_cell", range, dlp, redactedCellCount: shouldMaskChanged ? 1 : 0 });
    return { cell: this.formatCellForUser(address), changed };
  }

  private setRange(params: any): ToolResultDataByName["set_range"] {
    const range = this.parseRange(params.range, this.options.default_sheet);
    const interpretAs: "auto" | "value" | "formula" = params.interpret_as ?? "auto";

    const rowCount = Array.isArray(params.values) ? params.values.length : 0;
    // Avoid `Math.max(...rows.map(...))` spread: large pastes can include tens of thousands of
    // rows, which would exceed JS engines' argument limits.
    let colCount = 0;
    if (rowCount > 0) {
      for (const row of params.values as unknown[]) {
        const len = Array.isArray(row) ? row.length : 0;
        if (len > colCount) colCount = len;
      }
    }

    const expanded =
      range.startRow === range.endRow && range.startCol === range.endCol && (rowCount !== 1 || colCount !== 1);

    const targetRange = expanded
      ? {
          sheet: range.sheet,
          startRow: range.startRow,
          startCol: range.startCol,
          endRow: range.startRow + rowCount - 1,
          endCol: range.startCol + colCount - 1
        }
      : range;

    this.assertRangeWithinMaxToolCells("set_range", targetRange);

    const normalizedValues: CellScalar[][] = expanded
      ? params.values.map((row: CellScalar[]) => {
          const next = Array.isArray(row) ? row.slice() : [];
          while (next.length < colCount) next.push(null);
          return next;
        })
      : params.values;

    const cells: CellData[][] = normalizedValues.map((row: CellScalar[]) =>
      row.map((value) => {
        const shouldTreatAsFormula =
          interpretAs === "formula" || (interpretAs === "auto" && typeof value === "string" && value.trimStart().startsWith("="));

        if (shouldTreatAsFormula) {
          const formula = normalizeFormulaTextOpt(value);
          if (formula == null) return { value: null };
          return { value: null, formula };
        }

        return { value };
      })
    );

    this.spreadsheet.writeRange(targetRange, cells);
    this.refreshPivotsForRange(targetRange);
    const sizeRows = targetRange.endRow - targetRange.startRow + 1;
    const sizeCols = targetRange.endCol - targetRange.startCol + 1;
    return { range: this.formatRangeForUser(targetRange), updated_cells: sizeRows * sizeCols };
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

    const range = { sheet, startRow, endRow, startCol: colIndex, endCol: colIndex };
    this.assertRangeWithinMaxToolCells("apply_formula_column", range);
    const dlp = this.evaluateDlpForRange("apply_formula_column", range);
    if (dlp && dlp.decision.decision === DLP_DECISION.BLOCK) {
      this.logToolDlpDecision({ tool: "apply_formula_column", range, dlp, redactedCellCount: 0 });
      throw toolError(
        "permission_denied",
        `DLP policy blocks applying formulas to ${this.formatRangeForUser(range)} via apply_formula_column (ai.cloudProcessing).`
      );
    }

    const template = String(params.formula_template);
    let updated = 0;
    for (let row = startRow; row <= endRow; row++) {
      const formula = normalizeFormulaTextOpt(template.replaceAll("{row}", String(row)));
      this.spreadsheet.setCell({ sheet, row, col: colIndex }, formula == null ? { value: null } : { value: null, formula });
      updated++;
    }

    this.refreshPivotsForRange({
      sheet,
      startRow,
      endRow,
      startCol: colIndex,
      endCol: colIndex
    });

    if (dlp) this.logToolDlpDecision({ tool: "apply_formula_column", range, dlp, redactedCellCount: 0 });
    return { sheet: this.displaySheetName(sheet), column, start_row: startRow, end_row: endRow, updated_cells: updated };
  }

  private createPivotTable(params: any): ToolResultDataByName["create_pivot_table"] {
    const source = this.parseRange(params.source_range, this.options.default_sheet);
    const destination = this.parseCell(params.destination, this.options.default_sheet);

    this.assertRangeWithinMaxToolCells("create_pivot_table", source, { label: "source_range" });
    const sourceCells = this.spreadsheet.readRange(source);
    const dlp = this.evaluateDlpForRange("create_pivot_table", source);
    if (dlp && dlp.decision.decision === DLP_DECISION.BLOCK) {
      this.logToolDlpDecision({ tool: "create_pivot_table", range: source, dlp, redactedCellCount: 0 });
      throw toolError(
        "permission_denied",
        `DLP policy blocks creating a pivot table from ${this.formatRangeForUser(source)} (ai.cloudProcessing).`
      );
    }

    const includeFormulaValues = Boolean(this.options.include_formula_values && (!dlp || dlp.decision.decision === DLP_DECISION.ALLOW));
    let redactedCellCount = 0;
    const sourceValues: CellScalar[][] = sourceCells.map((row, r) =>
      row.map((cell, c) => {
        if (dlp && dlp.decision.decision === DLP_DECISION.REDACT) {
          const rowIndex = source.startRow + r;
          const colIndex = source.startCol + c;
          if (!this.isDlpCellAllowed(dlp, rowIndex, colIndex)) {
            redactedCellCount++;
            return null;
          }
        }
        if (cell.formula && !includeFormulaValues) return null;
        return cell.value ?? null;
      })
    );

    const output = buildPivotTableOutput({
      sourceValues,
      rowFields: params.rows ?? [],
      columnFields: params.columns ?? [],
      values: params.values ?? []
    });

    const rowCount = output.length;
    let colCount = 1;
    for (const row of output) {
      if (row.length > colCount) colCount = row.length;
    }
    const normalized: CellScalar[][] = output.map((row) => {
      const next = row.slice();
      while (next.length < colCount) next.push(null);
      return next;
    });

    const outRange = {
      sheet: destination.sheet,
      startRow: destination.row,
      startCol: destination.col,
      endRow: destination.row + rowCount - 1,
      endCol: destination.col + colCount - 1
    };

    this.assertRangeWithinMaxToolCells("create_pivot_table", outRange, { label: "destination_range" });

    const cells: CellData[][] = normalized.map((row) => row.map((value) => ({ value })));
    this.spreadsheet.writeRange(outRange, cells);

    // Register for automatic refresh when source data changes.
    const registration: PivotRegistration = {
      source,
      destination,
      rowFields: params.rows ?? [],
      columnFields: params.columns ?? [],
      values: (params.values ?? []) as PivotValueSpec[],
      lastDestinationRange: outRange
    };
    this.pivots.push(registration);

    if (dlp) this.logToolDlpDecision({ tool: "create_pivot_table", range: source, dlp, redactedCellCount });
    return {
      status: "ok",
      source_range: this.formatRangeForUser(source),
      destination_range: this.formatRangeForUser(outRange),
      written_cells: rowCount * colCount,
      shape: { rows: rowCount, cols: colCount }
    };
  }

  private createChart(params: any): ToolResultDataByName["create_chart"] {
    if (!this.spreadsheet.createChart) {
      throw toolError("not_implemented", "create_chart requires chart support in SpreadsheetApi");
    }

    const chartType = params.chart_type as ChartType;
    const dataRangeParsed = this.parseRange(params.data_range, this.options.default_sheet);
    const dataRangeForHost = formatA1Range(dataRangeParsed);
    const dataRangeForUser = this.formatRangeForUser(dataRangeParsed);

    let positionForHost: string | undefined;
    let positionForUser: string | undefined;
    if (params.position != null && String(params.position).trim() !== "") {
      try {
        const positionParsed = this.parseRange(String(params.position), dataRangeParsed.sheet);
        positionForHost = formatA1Range(positionParsed);
        positionForUser = this.formatRangeForUser(positionParsed);
      } catch (error) {
        throw toolError(
          "validation_error",
          `create_chart position must be an A1 cell or range reference (got "${params.position}")`,
          error instanceof Error ? { message: error.message } : undefined
        );
      }
    }

    const title = params.title != null && String(params.title).trim() !== "" ? String(params.title) : undefined;

    const spec: CreateChartSpec = {
      chart_type: chartType,
      data_range: dataRangeForHost,
      ...(title ? { title } : {}),
      ...(positionForHost ? { position: positionForHost } : {})
    };

    const result = this.spreadsheet.createChart(spec) as CreateChartResult;
    if (!result || typeof result.chart_id !== "string" || result.chart_id.trim() === "") {
      throw toolError("runtime_error", "create_chart host returned an invalid chart_id", result);
    }

    return {
      status: "ok",
      chart_id: result.chart_id,
      chart_type: chartType,
      data_range: dataRangeForUser,
      ...(title ? { title } : {}),
      ...(positionForUser ? { position: positionForUser } : {})
    };
  }

  private sortRange(params: any): ToolResultDataByName["sort_range"] {
    const range = this.parseRange(params.range, this.options.default_sheet);
    this.assertRangeWithinMaxToolCells("sort_range", range);
    const dlp = this.evaluateDlpForRange("sort_range", range);
    if (dlp && dlp.decision.decision !== DLP_DECISION.ALLOW) {
      this.logToolDlpDecision({ tool: "sort_range", range, dlp, redactedCellCount: 0 });
      throw toolError(
        "permission_denied",
        `DLP policy blocks sorting ${this.formatRangeForUser(range)} via sort_range (ai.cloudProcessing).`
      );
    }
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

    const includeFormulaValues = Boolean(this.options.include_formula_values);
    body.sort((left, right) => {
      for (const criterion of sortCriteria) {
        const orderMultiplier = criterion.order === "asc" ? 1 : -1;
        const result = compareCellForSort(left[criterion.offset]!, right[criterion.offset]!, { includeFormulaValues });
        if (result !== 0) return result * orderMultiplier;
      }
      return 0;
    });

    const sorted = [...header, ...body];
    this.spreadsheet.writeRange(range, sorted);
    this.refreshPivotsForRange(range);

    if (dlp) this.logToolDlpDecision({ tool: "sort_range", range, dlp, redactedCellCount: 0 });
    return { range: this.formatRangeForUser(range), sorted_rows: body.length };
  }

  private filterRange(params: any): ToolResultDataByName["filter_range"] {
    const range = this.parseRange(params.range, this.options.default_sheet);
    this.assertRangeWithinMaxToolCells("filter_range", range);
    const dlp = this.evaluateDlpForRange("filter_range", range);
    if (dlp && dlp.decision.decision === DLP_DECISION.BLOCK) {
      this.logToolDlpDecision({ tool: "filter_range", range, dlp, redactedCellCount: 0 });
      throw toolError(
        "permission_denied",
        `DLP policy blocks filtering ${this.formatRangeForUser(range)} via filter_range (ai.cloudProcessing).`
      );
    }
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

    // Only surface formula values when there is no DLP configured, or DLP is in pure ALLOW mode.
    // Under REDACT, formula values are treated as unsafe (may depend on restricted cells).
    const includeFormulaValues = Boolean(this.options.include_formula_values && (!dlp || dlp.decision.decision === DLP_DECISION.ALLOW));
    const matchingRows: number[] = [];
    let matchCount = 0;
    for (let i = bodyOffset; i < rows.length; i++) {
      const row = rows[i]!;
      const matches = criteria.every((criterion) => {
        const cell = row[criterion.offset]!;
        if (dlp && dlp.decision.decision === DLP_DECISION.REDACT) {
          const rowIndex = range.startRow + i;
          const colIndex = range.startCol + criterion.offset;
          if (!this.isDlpCellAllowed(dlp, rowIndex, colIndex)) {
            return matchesCriterion({ value: DLP_REDACTION_PLACEHOLDER }, criterion, { includeFormulaValues: false });
          }
        }
        return matchesCriterion(cell, criterion, { includeFormulaValues });
      });
      if (matches) {
        matchCount++;
        if (matchingRows.length < this.options.max_filter_range_matching_rows) {
          matchingRows.push(range.startRow + i);
        }
      }
    }

    if (dlp) this.logToolDlpDecision({ tool: "filter_range", range, dlp, redactedCellCount: 0 });
    const truncated = matchCount > matchingRows.length;
    return {
      range: this.formatRangeForUser(range),
      matching_rows: matchingRows,
      count: matchCount,
      ...(truncated ? { truncated } : {})
    };
  }

  private applyFormatting(params: any): ToolResultDataByName["apply_formatting"] {
    const range = this.parseRange(params.range, this.options.default_sheet);
    // NOTE: Unlike tools like sort/filter/statistics, `apply_formatting` does not need to
    // materialize a full `CellData[][]` grid in JS memory. Host spreadsheet backends are
    // expected to implement their own safety caps / scalable formatting paths (e.g.
    // layered formats, compressed range runs). Do not apply `max_tool_range_cells` here.
    const formatted = this.spreadsheet.applyFormatting(range, params.format);
    return { range: this.formatRangeForUser(range), formatted_cells: formatted };
  }

  private detectAnomalies(params: any): ToolResultDataByName["detect_anomalies"] {
    const range = this.parseRange(params.range, this.options.default_sheet);
    this.assertRangeWithinMaxToolCells("detect_anomalies", range);
    const formattedRange = this.formatRangeForUser(range);
    const method = (params.method ?? "zscore") as "zscore" | "iqr" | "isolation_forest";
    const dlp = this.evaluateDlpForRange("detect_anomalies", range);
    if (dlp && dlp.decision.decision === DLP_DECISION.BLOCK) {
      this.logToolDlpDecision({ tool: "detect_anomalies", range, dlp, redactedCellCount: 0 });
      throw toolError(
        "permission_denied",
        `DLP policy blocks analyzing ${formattedRange} via detect_anomalies (ai.cloudProcessing).`
      );
    }
    const cells = this.spreadsheet.readRange(range);
    // Only surface formula values when there is no DLP configured, or DLP is in pure ALLOW mode.
    // Under REDACT, formula values are treated as unsafe (may depend on restricted cells).
    const includeFormulaValues = Boolean(
      this.options.include_formula_values && (!dlp || dlp.decision.decision === DLP_DECISION.ALLOW)
    );
    let redactedCellCount = 0;

    switch (method) {
      case "zscore": {
        // Avoid materializing `{cell,value}` objects for every numeric cell. This tool can be
        // called on ranges near `max_tool_range_cells`; for those inputs, creating and
        // retaining per-cell A1 strings causes unnecessary memory pressure.
        //
        // Instead, compute mean/stdev in a first streaming pass (Welford), then do a second
        // pass to collect only the anomaly records (capped to `max_detect_anomalies`).
        const threshold = params.threshold ?? 3;
        let count = 0;
        let mean = 0;
        let m2 = 0;
        for (let r = 0; r < cells.length; r++) {
          const row = cells[r]!;
          for (let c = 0; c < row.length; c++) {
            const rowIndex = range.startRow + r;
            const colIndex = range.startCol + c;
            // Under DLP REDACT we exclude disallowed cells from anomaly computations entirely.
            // Otherwise, even "safe" outputs (e.g. z-score) can become an inference channel for
            // restricted values via the returned scores.
            if (dlp && dlp.decision.decision === DLP_DECISION.REDACT) {
              if (!this.isDlpCellAllowed(dlp, rowIndex, colIndex)) {
                redactedCellCount++;
                continue;
              }
            }
            const numeric = toNumber(row[c]!, { includeFormulaValues });
            if (numeric === null) continue;
            count += 1;
            const delta = numeric - mean;
            mean += delta / count;
            const delta2 = numeric - mean;
            m2 += delta * delta2;
          }
        }

        if (count === 0) {
          if (dlp) this.logToolDlpDecision({ tool: "detect_anomalies", range, dlp, redactedCellCount });
          return { range: formattedRange, method, anomalies: [] };
        }

        const variance = count > 1 ? m2 / (count - 1) : 0;
        const stdev = Math.sqrt(variance);
        if (stdev === 0) {
          if (dlp) this.logToolDlpDecision({ tool: "detect_anomalies", range, dlp, redactedCellCount });
          return { range: formattedRange, method, anomalies: [] };
        }

        const max = Math.max(0, this.options.max_detect_anomalies);
        const anomalies: Array<{ cell: string; value: number | null; score: number | null }> = [];
        let total = 0;
        for (let r = 0; r < cells.length; r++) {
          const row = cells[r]!;
          for (let c = 0; c < row.length; c++) {
            const rowIndex = range.startRow + r;
            const colIndex = range.startCol + c;
            if (dlp && dlp.decision.decision === DLP_DECISION.REDACT) {
              if (!this.isDlpCellAllowed(dlp, rowIndex, colIndex)) {
                continue;
              }
            }
            const value = toNumber(row[c]!, { includeFormulaValues });
            if (value === null) continue;
            const score = (value - mean) / stdev;
            if (Math.abs(score) < threshold) continue;
            total += 1;
            if (anomalies.length < max) {
              anomalies.push({
                cell: this.formatCellForUser({ sheet: range.sheet, row: rowIndex, col: colIndex }),
                value,
                score
              });
            }
          }
        }

        if (dlp) this.logToolDlpDecision({ tool: "detect_anomalies", range, dlp, redactedCellCount });
        const truncated = total > anomalies.length;
        return {
          range: formattedRange,
          method,
          anomalies,
          ...(truncated ? { truncated: true, total_anomalies: total } : {})
        };
      }
      case "iqr": {
        const multiplier = params.threshold ?? 1.5;
        // We still need to materialize numeric values to compute quantiles, but we avoid
        // allocating per-cell A1 strings until after we determine which indices are
        // anomalous.
        const values: number[] = [];
        for (let r = 0; r < cells.length; r++) {
          const row = cells[r]!;
          for (let c = 0; c < row.length; c++) {
            const rowIndex = range.startRow + r;
            const colIndex = range.startCol + c;
            if (dlp && dlp.decision.decision === DLP_DECISION.REDACT) {
              if (!this.isDlpCellAllowed(dlp, rowIndex, colIndex)) {
                redactedCellCount++;
                continue;
              }
            }
            const numeric = toNumber(row[c]!, { includeFormulaValues });
            if (numeric === null) continue;
            values.push(numeric);
          }
        }

        if (values.length === 0) {
          if (dlp) this.logToolDlpDecision({ tool: "detect_anomalies", range, dlp, redactedCellCount });
          return { range: formattedRange, method, anomalies: [] };
        }

        values.sort((a, b) => a - b);
        const q1 = quantileSorted(values, 0.25);
        const q3 = quantileSorted(values, 0.75);
        const iqr = q3 - q1;
        const low = q1 - multiplier * iqr;
        const high = q3 + multiplier * iqr;

        const max = Math.max(0, this.options.max_detect_anomalies);
        const anomalies: Array<{ cell: string; value: number | null }> = [];
        let total = 0;
        for (let r = 0; r < cells.length; r++) {
          const row = cells[r]!;
          for (let c = 0; c < row.length; c++) {
            const rowIndex = range.startRow + r;
            const colIndex = range.startCol + c;
            if (dlp && dlp.decision.decision === DLP_DECISION.REDACT) {
              if (!this.isDlpCellAllowed(dlp, rowIndex, colIndex)) {
                continue;
              }
            }
            const value = toNumber(row[c]!, { includeFormulaValues });
            if (value === null) continue;
            if (value >= low && value <= high) continue;
            total += 1;
            if (anomalies.length < max) {
              anomalies.push({
                cell: this.formatCellForUser({ sheet: range.sheet, row: rowIndex, col: colIndex }),
                value
              });
            }
          }
        }

        if (dlp) this.logToolDlpDecision({ tool: "detect_anomalies", range, dlp, redactedCellCount });
        const truncated = total > anomalies.length;
        return {
          range: formattedRange,
          method,
          anomalies,
          ...(truncated ? { truncated: true, total_anomalies: total } : {})
        };
      }
      case "isolation_forest": {
        const entries: Array<{ cell: string; value: number }> = [];
        for (let r = 0; r < cells.length; r++) {
          const row = cells[r]!;
          for (let c = 0; c < row.length; c++) {
            const rowIndex = range.startRow + r;
            const colIndex = range.startCol + c;
            if (dlp && dlp.decision.decision === DLP_DECISION.REDACT) {
              if (!this.isDlpCellAllowed(dlp, rowIndex, colIndex)) {
                redactedCellCount++;
                continue;
              }
            }
            const numeric = toNumber(row[c]!, { includeFormulaValues });
            if (numeric === null) continue;
            entries.push({
              cell: this.formatCellForUser({ sheet: range.sheet, row: rowIndex, col: colIndex }),
              value: numeric
            });
          }
        }

        if (entries.length === 0) {
          if (dlp) this.logToolDlpDecision({ tool: "detect_anomalies", range, dlp, redactedCellCount });
          return { range: formattedRange, method, anomalies: [] };
        }

        const values = entries.map((entry) => entry.value);
        const seed = fnv1a32(`${formattedRange}|isolation_forest`);
        const scores = isolationForestScores(values, { seed });
        const scored = entries
          .map((entry, index) => ({ ...entry, score: scores[index]! }))
          .sort((a, b) => b.score - a.score || a.cell.localeCompare(b.cell));

        /**
         * Isolation forest `threshold` semantics:
         * - If omitted, we use a default score cutoff (`score >= 0.65`).
         * - If `0 < threshold <= 1`, treat it as a score cutoff (`score >= threshold`).
         * - If `threshold > 1`, treat it as a "top N" selector (rounded + clamped).
         */
        const threshold = params.threshold as number | undefined;
        if (threshold === undefined || threshold <= 1) {
          const cutoff = threshold ?? 0.65;
          const anomalies = scored
            .filter((entry) => entry.score >= cutoff)
            .map((entry) => ({ cell: entry.cell, value: entry.value, score: entry.score }));
          if (dlp) this.logToolDlpDecision({ tool: "detect_anomalies", range, dlp, redactedCellCount });
          const capped = capList(anomalies, this.options.max_detect_anomalies);
          return {
            range: formattedRange,
            method,
            anomalies: capped.items,
            ...(capped.truncated ? { truncated: true, total_anomalies: capped.total } : {})
          };
        }

        const topN = Math.min(scored.length, Math.max(0, Math.round(threshold)));
        const anomalies = scored
          .slice(0, topN)
          .map((entry) => ({ cell: entry.cell, value: entry.value, score: entry.score }));
        if (dlp) this.logToolDlpDecision({ tool: "detect_anomalies", range, dlp, redactedCellCount });
        const capped = capList(anomalies, this.options.max_detect_anomalies);
        return {
          range: formattedRange,
          method,
          anomalies: capped.items,
          ...(capped.truncated ? { truncated: true, total_anomalies: capped.total } : {})
        };
      }
      default:
        throw new Error(`Unsupported detect_anomalies method: ${method}`);
    }
  }

  private computeStatistics(params: any): ToolResultDataByName["compute_statistics"] {
    const range = this.parseRange(params.range, this.options.default_sheet);
    this.assertRangeWithinMaxToolCells("compute_statistics", range);
    const dlp = this.evaluateDlpForRange("compute_statistics", range);
    if (dlp && dlp.decision.decision === DLP_DECISION.BLOCK) {
      this.logToolDlpDecision({ tool: "compute_statistics", range, dlp, redactedCellCount: 0 });
      throw toolError(
        "permission_denied",
        `DLP policy blocks analyzing ${this.formatRangeForUser(range)} via compute_statistics (ai.cloudProcessing).`
      );
    }
    const measures: string[] = params.measures ?? [];
    const requested = new Set(measures);

    const needsDistributionValues = requested.has("median") || requested.has("mode") || requested.has("quartiles");

    const needsStreamingAgg =
      requested.has("count") ||
      requested.has("sum") ||
      requested.has("mean") ||
      requested.has("stdev") ||
      requested.has("variance") ||
      requested.has("min") ||
      requested.has("max");

    const wantsCorrelation = requested.has("correlation");
    const cols = range.endCol - range.startCol + 1;
    const correlationSupported = cols === 2;
    const needsStreamingCorrelation = wantsCorrelation && correlationSupported;

    const cells = this.spreadsheet.readRange(range);
    const includeFormulaValues = Boolean(
      this.options.include_formula_values && (!dlp || dlp.decision.decision === DLP_DECISION.ALLOW)
    );
    const values: number[] | null = needsDistributionValues ? [] : null;
    let redactedCellCount = 0;

    // Basic streaming aggregates (Welford).
    let count = 0;
    let sum = 0;
    let mean = 0;
    let m2 = 0;
    let min = 0;
    let max = 0;
    let hasMinMax = false;

    // Streaming correlation (online covariance / Pearson r).
    let correlationCount = 0;
    let correlationMeanX = 0;
    let correlationMeanY = 0;
    let correlationC = 0;
    let correlationM2X = 0;
    let correlationM2Y = 0;

    for (let r = 0; r < cells.length; r++) {
      const row = cells[r]!;
      // Track correlation pair values for this row without re-reading cells (important for Proxy-based rows).
      let leftAllowed = true;
      let rightAllowed = true;
      let leftValue: number | null = null;
      let rightValue: number | null = null;

      for (let c = 0; c < row.length; c++) {
        const cell = row[c]!;
        let allowed = true;
        if (dlp && dlp.decision.decision === DLP_DECISION.REDACT) {
          const rowIndex = range.startRow + r;
          const colIndex = range.startCol + c;
          allowed = this.isDlpCellAllowed(dlp, rowIndex, colIndex);
          if (!allowed) {
            redactedCellCount++;
            if (needsStreamingCorrelation) {
              if (c === 0) leftAllowed = false;
              else if (c === 1) rightAllowed = false;
            }
            continue;
          }
        }
        const numeric = toNumber(cell, { includeFormulaValues });
        if (needsStreamingCorrelation) {
          if (c === 0) leftValue = numeric;
          else if (c === 1) rightValue = numeric;
        }
        if (numeric === null) continue;

        if (values) values.push(numeric);

        if (needsStreamingAgg) {
          if (!hasMinMax) {
            min = numeric;
            max = numeric;
            hasMinMax = true;
          } else {
            // Avoid Math.min/Math.max so NaN values behave consistently with the previous
            // "scan an array" implementation (NaN only poisons min/max if it is the first value).
            if (numeric < min) min = numeric;
            if (numeric > max) max = numeric;
          }

          count += 1;
          sum += numeric;
          const delta = numeric - mean;
          mean += delta / count;
          const delta2 = numeric - mean;
          m2 += delta * delta2;
        }
      }

      if (needsStreamingCorrelation) {
        if (leftAllowed && rightAllowed && leftValue !== null && rightValue !== null) {
          correlationCount += 1;
          const n = correlationCount;
          const deltaX = leftValue - correlationMeanX;
          correlationMeanX += deltaX / n;
          const deltaY = rightValue - correlationMeanY;
          correlationMeanY += deltaY / n;
          correlationC += deltaX * (rightValue - correlationMeanY);
          correlationM2X += deltaX * (leftValue - correlationMeanX);
          correlationM2Y += deltaY * (rightValue - correlationMeanY);
        }
      }
    }

    const stats: Record<string, number | null> = {};
    for (const measure of measures) {
      switch (measure) {
        case "mean":
          stats.mean = count ? sum / count : null;
          break;
        case "sum":
          stats.sum = count ? sum : null;
          break;
        case "count":
          stats.count = count;
          break;
        case "median":
          stats.median = values && values.length ? median(values) : null;
          break;
        case "mode":
          stats.mode = values && values.length ? mode(values) : null;
          break;
        case "stdev":
          if (!count) {
            stats.stdev = null;
            break;
          }
          stats.stdev = Math.sqrt(count < 2 ? 0 : m2 / (count - 1));
          break;
        case "variance":
          if (!count) {
            stats.variance = null;
            break;
          }
          stats.variance = count < 2 ? 0 : m2 / (count - 1);
          break;
        case "min":
          stats.min = hasMinMax ? min : null;
          break;
        case "max":
          stats.max = hasMinMax ? max : null;
          break;
        case "quartiles": {
          if (!values || !values.length) {
            stats.q1 = null;
            stats.q2 = null;
            stats.q3 = null;
            break;
          }
          const sorted = [...values].sort((a, b) => a - b);
          stats.q1 = quantileSorted(sorted, 0.25);
          stats.q2 = quantileSorted(sorted, 0.5);
          stats.q3 = quantileSorted(sorted, 0.75);
          break;
        }
        case "correlation": {
          if (!correlationSupported) {
            stats.correlation = null;
            break;
          }
          if (!correlationCount) {
            stats.correlation = null;
            break;
          }
          const denominator = Math.sqrt(correlationM2X * correlationM2Y);
          stats.correlation = denominator === 0 ? 0 : correlationC / denominator;
          break;
        }
        default:
          stats[measure] = null;
      }
    }

    if (dlp) this.logToolDlpDecision({ tool: "compute_statistics", range, dlp, redactedCellCount });
    return { range: this.formatRangeForUser(range), statistics: stats };
  }

  private evaluateDlpForRange(tool: ToolName, range: ReturnType<typeof parseA1Range>): null | {
    documentId: string;
    sheetId: string;
    records: Array<{ selector: any; classification: any }>;
    /**
     * Normalized (0-based) range for the current tool selection. Stored so we can lazily
     * build the per-range selector index only when per-cell enforcement is needed.
     */
    selectionRange: DlpNormalizedRange;
    /**
     * Lazily populated. Most tools only need per-cell enforcement when the policy decision
     * is REDACT; in ALLOW/BLOCK cases building a full index can be unnecessary overhead.
     */
    index: DlpRangeIndex | null;
    includeRestrictedContent: boolean;
    /**
     * Precomputed policy details used by per-cell enforcement (`isDlpCellAllowed`).
     */
    maxAllowedRank: number | null;
    policyAllowsRestrictedContent: boolean;
    restrictedOverrideAllowed: boolean;
    restrictedAllowed: boolean;
    canShortCircuitOverThreshold: boolean;
    policy: any;
    decision: any;
    selectionClassification: any;
    auditLogger?: { log(event: any): void };
  } {
    const dlp = this.options.dlp;
    if (!dlp) return null;

    const documentId = dlp.document_id;
    const sheetId =
      range.sheet === this.options.default_sheet ? (dlp.sheet_id ?? range.sheet) : range.sheet;
    const records = dlp.classification_records ?? dlp.classification_store?.list(documentId) ?? [];
    const includeRestrictedContent = dlp.include_restricted_content ?? false;

    const normalizedSelectionRange = normalizeRange({
      start: { row: range.startRow - 1, col: range.startCol - 1 },
      end: { row: range.endRow - 1, col: range.endCol - 1 }
    });

    let selectionClassification = effectiveRangeClassification(
      {
        documentId,
        sheetId,
        range: normalizedSelectionRange
      },
      records
    );

    const tableColumnResolver = dlp.table_column_resolver;
    if (tableColumnResolver) {
      for (const record of records || []) {
        if (!record || !record.selector || typeof record.selector !== "object") continue;
        const selector = record.selector;
        if (selector.documentId !== documentId) continue;
        if (selector.scope !== "column") continue;
        if (selector.sheetId !== sheetId) continue;
        // Column selectors expressed directly in sheet coordinates are already handled by
        // `effectiveRangeClassification`. Only consider table-based selectors.
        if (typeof selector.columnIndex === "number") continue;
        const tableId = selector.tableId;
        const columnId = selector.columnId;
        if (!tableId || !columnId) continue;
        const resolvedColIndex = tableColumnResolver.getColumnIndex(sheetId, tableId, columnId);
        if (typeof resolvedColIndex !== "number" || !Number.isInteger(resolvedColIndex) || resolvedColIndex < 0) continue;
        if (resolvedColIndex < normalizedSelectionRange.start.col || resolvedColIndex > normalizedSelectionRange.end.col) continue;
        selectionClassification = maxClassification(selectionClassification, record.classification);
        if (selectionClassification.level === CLASSIFICATION_LEVEL.RESTRICTED) break;
      }
    }

    const decision = evaluatePolicy({
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      classification: selectionClassification,
      policy: dlp.policy,
      options: { includeRestrictedContent }
    });

    const maxAllowed = decision?.maxAllowed ?? null;
    const maxAllowedRank = maxAllowed === null ? null : classificationRank(maxAllowed);

    const policyAllowsRestrictedContent = Boolean(
      dlp.policy?.rules?.[DLP_ACTION.AI_CLOUD_PROCESSING]?.allowRestrictedContent
    );

    const restrictedOverrideAllowed = includeRestrictedContent && policyAllowsRestrictedContent;
    const restrictedAllowed =
      maxAllowedRank === null
        ? false
        : includeRestrictedContent
          ? policyAllowsRestrictedContent
          : maxAllowedRank >= RESTRICTED_CLASSIFICATION_RANK;
    const canShortCircuitOverThreshold = !restrictedOverrideAllowed;

    return {
      documentId,
      sheetId,
      records,
      selectionRange: normalizedSelectionRange,
      index: null,
      includeRestrictedContent,
      maxAllowedRank,
      policyAllowsRestrictedContent,
      restrictedOverrideAllowed,
      restrictedAllowed,
      canShortCircuitOverThreshold,
      policy: dlp.policy,
      decision,
      selectionClassification,
      auditLogger: dlp.audit_logger
    };
  }

  private buildDlpRangeIndex(
    ref: { documentId: string; sheetId: string; range: DlpNormalizedRange },
    records: Array<{ selector: any; classification: any }>,
    opts: { maxAllowedRank: number }
  ): DlpRangeIndex {
    const selectionRange = ref.range;
    const startRow = selectionRange.start.row;
    const startCol = selectionRange.start.col;
    const rowCount = selectionRange.end.row - selectionRange.start.row + 1;
    const colCount = selectionRange.end.col - selectionRange.start.col + 1;
    const { maxAllowedRank } = opts;
    const tableColumnResolver = this.options.dlp?.table_column_resolver;

    const rankFromClassification = (classification: any): number => {
      if (!classification) return DEFAULT_CLASSIFICATION_RANK;
      if (typeof classification !== "object") {
        throw new Error("Classification must be an object");
      }
      return classificationRank((classification as any).level);
    };

    let docRankMax = DEFAULT_CLASSIFICATION_RANK;
    let sheetRankMax = DEFAULT_CLASSIFICATION_RANK;
    const columnRankByOffset = new Uint8Array(colCount);
    let cellRankByOffset: Uint8Array | null = null;
    const rangeRecords: Array<{ startRow: number; endRow: number; startCol: number; endCol: number; rank: number }> = [];
    let rangeRankMax = DEFAULT_CLASSIFICATION_RANK;
    const fallbackRecords: Array<{ selector: any; classification: any }> = [];

    for (const record of records || []) {
      if (!record || !record.selector || typeof record.selector !== "object") continue;
      const selector = record.selector;
      if (selector.documentId !== ref.documentId) continue;

      // The index only needs to track restrictions above the baseline (Public). Public-scoped
      // records do not affect max-classification enforcement and would just add Map churn.
      const recordRank = rankFromClassification(record.classification);
      // Ignore classifications that cannot influence the per-cell allow/redact decision
      // (anything at or below the policy maxAllowed threshold).
      if (recordRank <= maxAllowedRank) continue;

      switch (selector.scope) {
        case "document": {
          docRankMax = Math.max(docRankMax, recordRank);
          break;
        }
        case "sheet": {
          if (selector.sheetId === ref.sheetId) {
            sheetRankMax = Math.max(sheetRankMax, recordRank);
          }
          break;
        }
        case "column": {
          if (selector.sheetId !== ref.sheetId) break;
          let colIndex: number | null = null;
          if (typeof selector.columnIndex === "number") {
            colIndex = selector.columnIndex;
          } else if (tableColumnResolver && selector.tableId && selector.columnId) {
            colIndex = tableColumnResolver.getColumnIndex(ref.sheetId, selector.tableId, selector.columnId);
          }
          if (typeof colIndex !== "number" || !Number.isInteger(colIndex) || colIndex < 0) break;
          if (colIndex < selectionRange.start.col || colIndex > selectionRange.end.col) break;
          const offset = colIndex - startCol;
          if (recordRank > columnRankByOffset[offset]!) columnRankByOffset[offset] = recordRank;
          break;
        }
        case "cell": {
          if (selector.sheetId !== ref.sheetId) break;
          if (typeof selector.row !== "number" || typeof selector.col !== "number") break;
          // Only cells that could apply to this selection need to be indexed.
          if (
            selector.row < selectionRange.start.row ||
            selector.row > selectionRange.end.row ||
            selector.col < selectionRange.start.col ||
            selector.col > selectionRange.end.col
          ) {
            break;
          }
          const rowOffset = selector.row - startRow;
          const colOffset = selector.col - startCol;
          if (rowOffset < 0 || colOffset < 0 || rowOffset >= rowCount || colOffset >= colCount) break;
          if (cellRankByOffset === null) {
            cellRankByOffset = new Uint8Array(rowCount * colCount);
          }
          const offset = rowOffset * colCount + colOffset;
          if (recordRank > cellRankByOffset[offset]!) cellRankByOffset[offset] = recordRank;
          break;
        }
        case "range": {
          if (selector.sheetId !== ref.sheetId) break;
          if (!selector.range) break;
          const normalized = normalizeRange(selector.range);
          if (!rangesIntersectNormalized(normalized, selectionRange)) break;
          if (recordRank > rangeRankMax) rangeRankMax = recordRank;
          rangeRecords.push({
            startRow: normalized.start.row,
            endRow: normalized.end.row,
            startCol: normalized.start.col,
            endCol: normalized.end.col,
            rank: recordRank
          });
          break;
        }
        default: {
          // Unknown selector scope: ignore (selectorAppliesToCell would treat it as non-matching).
          break;
        }
      }
    }

    // Sort range selectors by rank descending so per-cell evaluation can break early once the
    // remaining records can no longer increase the effective classification.
    if (rangeRecords.length > 1) {
      rangeRecords.sort((a, b) => b.rank - a.rank);
    }

    const baseRank = Math.max(docRankMax, sheetRankMax);

    return {
      docRankMax,
      sheetRankMax,
      baseRank,
      startRow,
      startCol,
      rowCount,
      colCount,
      columnRankByOffset,
      cellRankByOffset,
      rangeRecords,
      rangeRankMax,
      fallbackRecords
    };
  }

  private isDlpCellAllowed(
    dlp: NonNullable<ReturnType<ToolExecutor["evaluateDlpForRange"]>>,
    row: number,
    col: number
  ): boolean {
    if (dlp.maxAllowedRank === null) {
      return false;
    }

    const index =
      dlp.index ??
      (dlp.index = this.buildDlpRangeIndex(
        { documentId: dlp.documentId, sheetId: dlp.sheetId, range: dlp.selectionRange },
        dlp.records,
        { maxAllowedRank: dlp.maxAllowedRank }
      ));

    const row0 = row - 1;
    const col0 = col - 1;

    const maxAllowedRank = dlp.maxAllowedRank;
    const restrictedAllowed = dlp.restrictedAllowed;
    const canShortCircuitOverThreshold = dlp.canShortCircuitOverThreshold;

    let rank = index.baseRank;

    if (rank === RESTRICTED_CLASSIFICATION_RANK) {
      return restrictedAllowed;
    }
    if (canShortCircuitOverThreshold && rank > maxAllowedRank) {
      return false;
    }

    const colOffset = col0 - index.startCol;
    const colRank = index.columnRankByOffset[colOffset] ?? DEFAULT_CLASSIFICATION_RANK;
    if (colRank > rank) rank = colRank;
    if (rank === RESTRICTED_CLASSIFICATION_RANK) {
      return restrictedAllowed;
    }
    if (canShortCircuitOverThreshold && rank > maxAllowedRank) {
      return false;
    }

    if (index.cellRankByOffset !== null) {
      const rowOffset = row0 - index.startRow;
      if (rowOffset >= 0 && rowOffset < index.rowCount && colOffset >= 0 && colOffset < index.colCount) {
        const offset = rowOffset * index.colCount + colOffset;
        const cellRank = index.cellRankByOffset[offset] ?? DEFAULT_CLASSIFICATION_RANK;
        if (cellRank > rank) rank = cellRank;
      }
    }
    if (rank === RESTRICTED_CLASSIFICATION_RANK) {
      return restrictedAllowed;
    }
    if (canShortCircuitOverThreshold && rank > maxAllowedRank) {
      return false;
    }

    const rangeCanAffectDecision =
      index.rangeRankMax > maxAllowedRank || (!restrictedAllowed && index.rangeRankMax === RESTRICTED_CLASSIFICATION_RANK);
    if (rangeCanAffectDecision && index.rangeRankMax > rank) {
      for (const record of index.rangeRecords) {
        // Records are sorted by rank desc in the index builder: once we reach a record that
        // cannot increase the effective rank, we can stop scanning.
        if (record.rank <= rank) break;
        if (row0 < record.startRow || row0 > record.endRow || col0 < record.startCol || col0 > record.endCol) continue;
        rank = record.rank;
        if (rank === RESTRICTED_CLASSIFICATION_RANK) {
          return restrictedAllowed;
        }
        if (canShortCircuitOverThreshold && rank > maxAllowedRank) {
          return false;
        }
        if (rank === index.rangeRankMax) break;
      }
    }

    if (index.fallbackRecords.length > 0 && rank !== RESTRICTED_CLASSIFICATION_RANK) {
      const fallbackClassification = effectiveCellClassification(
        { documentId: dlp.documentId, sheetId: dlp.sheetId, row: row0, col: col0 },
        index.fallbackRecords
      );
      const fallbackRank = classificationRank(fallbackClassification.level);
      if (fallbackRank > rank) rank = fallbackRank;
    }

    if (rank === RESTRICTED_CLASSIFICATION_RANK) return restrictedAllowed;
    return rank <= maxAllowedRank;
  }

  private logToolDlpDecision(params: {
    tool: ToolName;
    range: ReturnType<typeof parseA1Range>;
    dlp: NonNullable<ReturnType<ToolExecutor["evaluateDlpForRange"]>>;
    redactedCellCount: number;
  }): void {
    const { tool, range, dlp, redactedCellCount } = params;
    dlp.auditLogger?.log({
      type: "ai.tool.dlp",
      tool,
      documentId: dlp.documentId,
      sheetId: dlp.sheetId,
      range: formatA1Range(range),
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      decision: dlp.decision,
      selectionClassification: dlp.selectionClassification,
      redactedCellCount,
      // Alias for downstream consumers expecting the spec wording.
      redacted_counts: redactedCellCount
    });
  }

  private async fetchExternalData(
    params: any
  ): Promise<{ data: ToolResultDataByName["fetch_external_data"]; warnings?: string[] }> {
    // PreviewEngine must never perform network access. Instead, return a deterministic
    // stub result so previews reflect "external data requested" (approval gating) rather
    // than a misleading "tool disabled" permission error.
    if (this.options.preview_mode) {
      const requestedUrl = new URL(params.url);
      const destination = this.parseCell(params.destination, this.options.default_sheet);
      return {
        data: {
          url: safeUrlForProvenance(requestedUrl),
          destination: this.formatCellForUser(destination),
          written_cells: 0,
          shape: { rows: 0, cols: 0 },
          fetched_at_ms: Date.now(),
          status_code: 0
        },
        warnings: ["fetch_external_data skipped during preview"]
      };
    }

    if (!this.options.allow_external_data) {
      throw toolError("permission_denied", "fetch_external_data is disabled by default.");
    }

    const requestedUrl = new URL(params.url);
    ensureExternalUrlAllowed(requestedUrl, this.options.allowed_external_hosts);

    // Prevent allowlist bypass via redirects by manually following redirects and
    // validating each hop.
    const maxRedirects = 5;
    let currentUrl = requestedUrl;
    let requestHeaders: Record<string, string> | undefined = params.headers ? { ...params.headers } : undefined;
    let response: Response | null = null;

    for (let redirectCount = 0; redirectCount <= maxRedirects; redirectCount++) {
      ensureExternalUrlAllowed(currentUrl, this.options.allowed_external_hosts);
      response = await fetch(currentUrl.toString(), {
        headers: requestHeaders ?? undefined,
        credentials: "omit",
        cache: "no-store",
        referrerPolicy: "no-referrer",
        redirect: "manual"
      });

      // In browser environments, `redirect: "manual"` produces an opaque redirect response
      // that does not expose redirect location details. Fall back to automatic redirects and
      // validate the final resolved URL.
      if (response.type === "opaqueredirect") {
        // We can't inspect intermediate redirect hops here, so drop any user-supplied
        // headers before following redirects to avoid leaking secrets across hosts.
        requestHeaders = undefined;
        await cancelResponseBody(response);
        response = await fetch(currentUrl.toString(), {
          headers: undefined,
          credentials: "omit",
          cache: "no-store",
          referrerPolicy: "no-referrer",
          redirect: "follow"
        });
        const resolved = response.url ? new URL(response.url) : currentUrl;
        if (currentUrl.protocol === "https:" && resolved.protocol === "http:") {
          throw toolError("permission_denied", "Redirect from https to http is not permitted for fetch_external_data.");
        }
        ensureExternalUrlAllowed(resolved, this.options.allowed_external_hosts);
        currentUrl = resolved;
        break;
      }

      if (!isRedirectStatus(response.status)) break;

      const location = response.headers.get("location");
      if (!location) {
        await cancelResponseBody(response);
        throw toolError("runtime_error", `External fetch failed with HTTP ${response.status} (missing Location header)`);
      }
      const nextUrl = new URL(location, currentUrl);
      if (currentUrl.protocol === "https:" && nextUrl.protocol === "http:") {
        await cancelResponseBody(response);
        throw toolError("permission_denied", "Redirect from https to http is not permitted for fetch_external_data.");
      }
      // Avoid leaking user-supplied headers (e.g. API keys) across redirect hops to a
      // different host.
      if (nextUrl.host !== currentUrl.host) {
        requestHeaders = undefined;
      }
      await cancelResponseBody(response);
      currentUrl = nextUrl;
    }

    if (!response) {
      throw toolError("runtime_error", "External fetch failed to produce a response.");
    }
    if (isRedirectStatus(response.status)) {
      await cancelResponseBody(response);
      throw toolError("runtime_error", `External fetch exceeded maximum redirects (${maxRedirects}).`);
    }

    const statusCode = response.status;
    const contentType = response.headers.get("content-type") ?? undefined;
    const contentLengthHeader = response.headers.get("content-length");
    const declaredLength = contentLengthHeader ? Number(contentLengthHeader) : NaN;
    if (Number.isFinite(declaredLength) && declaredLength > this.options.max_external_bytes) {
      await cancelResponseBody(response);
      throw toolError(
        "permission_denied",
        `External response too large (${declaredLength} bytes). Increase max_external_bytes to allow.`
      );
    }

    if (!response.ok) {
      await cancelResponseBody(response);
      throw toolError("runtime_error", `External fetch failed with HTTP ${statusCode}`);
    }

    const destination = this.parseCell(params.destination, this.options.default_sheet);
    const bodyBytes = await readResponseBytes(response, this.options.max_external_bytes);
    const fetchedAtMs = Date.now();
    const contentLengthBytes = bodyBytes.byteLength;

    if (params.transform === "raw_text") {
      const text = decodeUtf8(bodyBytes);
      this.spreadsheet.setCell(destination, { value: text });
      this.refreshPivotsForRange({
        sheet: destination.sheet,
        startRow: destination.row,
        endRow: destination.row,
        startCol: destination.col,
        endCol: destination.col
      });
      return {
        data: {
          url: safeUrlForProvenance(currentUrl),
          destination: this.formatCellForUser(destination),
          written_cells: 1,
          shape: { rows: 1, cols: 1 },
          fetched_at_ms: fetchedAtMs,
          content_type: contentType,
          content_length_bytes: contentLengthBytes,
          status_code: statusCode
        }
      };
    }

    const json = JSON.parse(decodeUtf8(bodyBytes));
    const table = jsonToTable(json, { maxCells: this.options.max_tool_range_cells });
    const range = {
      sheet: destination.sheet,
      startRow: destination.row,
      startCol: destination.col,
      endRow: destination.row + table.length - 1,
      endCol: destination.col + (table[0]?.length ?? 1) - 1
    };
    this.assertRangeWithinMaxToolCells("fetch_external_data", range);

    const cells: CellData[][] = table.map((row) => row.map((value) => ({ value })));
    this.spreadsheet.writeRange(range, cells);
    this.refreshPivotsForRange(range);

    return {
      data: {
        url: safeUrlForProvenance(currentUrl),
        destination: this.formatCellForUser(destination),
        written_cells: table.length * (table[0]?.length ?? 0),
        shape: { rows: table.length, cols: table[0]?.length ?? 0 },
        fetched_at_ms: fetchedAtMs,
        content_type: contentType,
        content_length_bytes: contentLengthBytes,
        status_code: statusCode
      }
    };
  }

  private refreshPivotsForRange(changed: { sheet: string; startRow: number; endRow: number; startCol: number; endCol: number }): void {
    if (this.pivots.length === 0) return;

    for (const pivot of this.pivots) {
      if (!rangesIntersect(changed, pivot.source)) continue;
      this.refreshPivot(pivot);
    }
  }

  private refreshPivot(pivot: PivotRegistration): void {
    const maxCells = this.options.max_tool_range_cells;
    if (Number.isFinite(maxCells) && maxCells > 0) {
      // Defensive: pivots are created in-process via `create_pivot_table`, which is
      // already guarded by `max_tool_range_cells`. Still, this protects us from
      // unexpected/legacy registrations and prevents a background refresh from
      // allocating a massive `CellData[][]`.
      if (rangeCellCount(pivot.source) > maxCells) return;
    }
    const sourceCells = this.spreadsheet.readRange(pivot.source);
    const dlp = this.evaluateDlpForRange("create_pivot_table", pivot.source);
    if (dlp && dlp.decision.decision === DLP_DECISION.BLOCK) {
      return;
    }

    const includeFormulaValues = Boolean(this.options.include_formula_values && (!dlp || dlp.decision.decision === DLP_DECISION.ALLOW));
    const sourceValues: CellScalar[][] = sourceCells.map((row, r) =>
      row.map((cell, c) => {
        if (dlp && dlp.decision.decision === DLP_DECISION.REDACT) {
          const rowIndex = pivot.source.startRow + r;
          const colIndex = pivot.source.startCol + c;
          if (!this.isDlpCellAllowed(dlp, rowIndex, colIndex)) return null;
        }
        if (cell.formula && !includeFormulaValues) return null;
        return cell.value ?? null;
      })
    );

    const output = buildPivotTableOutput({
      sourceValues,
      rowFields: pivot.rowFields,
      columnFields: pivot.columnFields,
      values: pivot.values
    });

    const rowCount = output.length;
    let colCount = 1;
    for (const row of output) {
      if (row.length > colCount) colCount = row.length;
    }
    const normalized: CellScalar[][] = output.map((row) => {
      const next = row.slice();
      while (next.length < colCount) next.push(null);
      return next;
    });

    const nextRange = {
      sheet: pivot.destination.sheet,
      startRow: pivot.destination.row,
      startCol: pivot.destination.col,
      endRow: pivot.destination.row + rowCount - 1,
      endCol: pivot.destination.col + colCount - 1
    };

    if (Number.isFinite(maxCells) && maxCells > 0) {
      // Skip refresh if the next output would exceed our configured safety cap.
      // This avoids building massive intermediate arrays during background refresh.
      if (rangeCellCount(nextRange) > maxCells) return;
    }

    // Pivot refresh clears the previous output range, then writes the new output range.
    //
    // We intentionally avoid writing a "union rectangle" of prev+next: if the pivot
    // changes shape significantly (e.g. wide -> tall), the union rectangle can be
    // dramatically larger than either range and lead to huge allocations.
    const prevRange = pivot.lastDestinationRange;
    const emptyCell: CellData = { value: null };

    if (
      prevRange.sheet !== nextRange.sheet ||
      prevRange.startRow !== nextRange.startRow ||
      prevRange.startCol !== nextRange.startCol ||
      prevRange.endRow !== nextRange.endRow ||
      prevRange.endCol !== nextRange.endCol
    ) {
      if (Number.isFinite(maxCells) && maxCells > 0) {
        if (rangeCellCount(prevRange) > maxCells) return;
      }
      const prevRows = prevRange.endRow - prevRange.startRow + 1;
      const prevCols = prevRange.endCol - prevRange.startCol + 1;
      const clearCells: CellData[][] = Array.from({ length: prevRows }, () =>
        Array.from({ length: prevCols }, () => emptyCell)
      );
      this.spreadsheet.writeRange(prevRange, clearCells);
    }

    const nextCells: CellData[][] = normalized.map((row) => row.map((value) => ({ value })));
    this.spreadsheet.writeRange(nextRange, nextCells);
    pivot.lastDestinationRange = nextRange;
  }
}

interface PivotRegistration {
  source: ReturnType<typeof parseA1Range>;
  destination: ReturnType<typeof parseA1Cell>;
  rowFields: string[];
  columnFields: string[];
  values: PivotValueSpec[];
  lastDestinationRange: ReturnType<typeof parseA1Range>;
}

function rangesIntersect(
  a: { sheet: string; startRow: number; endRow: number; startCol: number; endCol: number },
  b: { sheet: string; startRow: number; endRow: number; startCol: number; endCol: number }
): boolean {
  if (a.sheet !== b.sheet) return false;
  return !(a.endRow < b.startRow || a.startRow > b.endRow || a.endCol < b.startCol || a.startCol > b.endCol);
}

type PivotAggregation = PivotAggregationType;

interface PivotValueSpec {
  field: string;
  aggregation: PivotAggregation;
}

interface PivotBuildRequest {
  sourceValues: CellScalar[][];
  rowFields: string[];
  columnFields: string[];
  values: PivotValueSpec[];
}

interface AggState {
  count: number;
  countNumbers: number;
  sum: number;
  product: number;
  min: number;
  max: number;
  mean: number;
  m2: number;
}

function initAggState(): AggState {
  return {
    count: 0,
    countNumbers: 0,
    sum: 0,
    product: 1,
    min: Infinity,
    max: -Infinity,
    mean: 0,
    m2: 0
  };
}

function updateAggState(state: AggState, value: CellScalar) {
  if (value == null) return;
  state.count += 1;
  const numeric = parseSpreadsheetNumber(value);
  if (numeric === null) return;
  const nextCount = state.countNumbers + 1;
  state.countNumbers = nextCount;
  state.sum += numeric;
  state.product *= numeric;
  state.min = Math.min(state.min, numeric);
  state.max = Math.max(state.max, numeric);

  const delta = numeric - state.mean;
  state.mean += delta / nextCount;
  const delta2 = numeric - state.mean;
  state.m2 += delta * delta2;
}

function mergeAggState(into: AggState, other: AggState) {
  into.count += other.count;
  if (other.countNumbers === 0) return;
  if (into.countNumbers === 0) {
    into.countNumbers = other.countNumbers;
    into.sum = other.sum;
    into.product = other.product;
    into.min = other.min;
    into.max = other.max;
    into.mean = other.mean;
    into.m2 = other.m2;
    return;
  }

  const n1 = into.countNumbers;
  const n2 = other.countNumbers;
  const n = n1 + n2;
  const delta = other.mean - into.mean;

  into.countNumbers = n;
  into.sum += other.sum;
  into.product *= other.product;
  into.min = Math.min(into.min, other.min);
  into.max = Math.max(into.max, other.max);
  into.mean = (n1 * into.mean + n2 * other.mean) / n;
  into.m2 += other.m2 + (delta * delta * n1 * n2) / n;
}

function finalizeAgg(state: AggState, agg: PivotAggregation): CellScalar {
  switch (agg) {
    case "count":
      return state.count;
    case "countNumbers":
      return state.countNumbers;
    case "sum":
      return state.countNumbers > 0 ? state.sum : null;
    case "average":
      return state.countNumbers > 0 ? state.sum / state.countNumbers : null;
    case "product":
      return state.countNumbers > 0 ? state.product : null;
    case "min":
      return state.countNumbers > 0 ? state.min : null;
    case "max":
      return state.countNumbers > 0 ? state.max : null;
    case "var":
      return state.countNumbers >= 2 ? state.m2 / (state.countNumbers - 1) : null;
    case "varP":
      return state.countNumbers > 0 ? state.m2 / state.countNumbers : null;
    case "stdDev": {
      const variance = state.countNumbers >= 2 ? state.m2 / (state.countNumbers - 1) : null;
      return variance == null ? null : Math.sqrt(variance);
    }
    case "stdDevP": {
      const variance = state.countNumbers > 0 ? state.m2 / state.countNumbers : null;
      return variance == null ? null : Math.sqrt(variance);
    }
    default: {
      const exhaustive: never = agg;
      throw new Error(`Unhandled aggregation: ${exhaustive}`);
    }
  }
}

function aggLabel(agg: PivotAggregation): string {
  switch (agg) {
    case "sum":
      return "Sum";
    case "count":
      return "Count";
    case "average":
      return "Average";
    case "min":
      return "Min";
    case "max":
      return "Max";
    case "product":
      return "Product";
    case "countNumbers":
      return "CountNumbers";
    case "stdDev":
      return "StdDev";
    case "stdDevP":
      return "StdDevP";
    case "var":
      return "Var";
    case "varP":
      return "VarP";
    default: {
      const exhaustive: never = agg;
      return exhaustive;
    }
  }
}

function normalizeKeyPart(value: CellScalar): string {
  return value == null ? "" : String(value);
}

function buildPivotTableOutput(request: PivotBuildRequest): CellScalar[][] {
  const { sourceValues, rowFields, columnFields, values } = request;
  if (!Array.isArray(sourceValues) || sourceValues.length === 0) {
    throw new Error("create_pivot_table: source_range is empty");
  }

  const headerRow = sourceValues[0] ?? [];
  const headers = headerRow.map((cell) => normalizeKeyPart(cell).trim());
  const indexByHeader = new Map<string, number>();
  for (const [idx, name] of headers.entries()) {
    if (!name) continue;
    if (!indexByHeader.has(name)) indexByHeader.set(name, idx);
  }

  const rowIndices = rowFields.map((name) => {
    const idx = indexByHeader.get(name);
    if (idx == null) throw new Error(`create_pivot_table: missing row field \"${name}\" in header row`);
    return idx;
  });

  const colIndices = columnFields.map((name) => {
    const idx = indexByHeader.get(name);
    if (idx == null) throw new Error(`create_pivot_table: missing column field \"${name}\" in header row`);
    return idx;
  });

  const valueSpecs: PivotValueSpec[] = values.map((v) => ({
    field: v.field,
    aggregation: v.aggregation
  }));

  const valueIndices = valueSpecs.map((spec) => {
    const idx = indexByHeader.get(spec.field);
    if (idx == null) throw new Error(`create_pivot_table: missing value field \"${spec.field}\" in header row`);
    return idx;
  });

  const hasColumns = colIndices.length > 0;

  const cube = new Map<string, Map<string, AggState[]>>();
  const rowKeyParts = new Map<string, CellScalar[]>();
  const colKeyParts = new Map<string, CellScalar[]>();
  const rowKeys = new Set<string>();
  const colKeys = new Set<string>();

  for (const record of sourceValues.slice(1)) {
    const rowParts = rowIndices.map((idx) => record[idx] ?? null);
    const rowKey = JSON.stringify(rowParts.map(normalizeKeyPart));
    rowKeys.add(rowKey);
    if (!rowKeyParts.has(rowKey)) rowKeyParts.set(rowKey, rowParts);

    const colParts = colIndices.map((idx) => record[idx] ?? null);
    const colKey = hasColumns ? JSON.stringify(colParts.map(normalizeKeyPart)) : JSON.stringify([]);
    colKeys.add(colKey);
    if (!colKeyParts.has(colKey)) colKeyParts.set(colKey, colParts);

    let rowMap = cube.get(rowKey);
    if (!rowMap) {
      rowMap = new Map();
      cube.set(rowKey, rowMap);
    }

    let cellStates = rowMap.get(colKey);
    if (!cellStates) {
      cellStates = valueSpecs.map(() => initAggState());
      rowMap.set(colKey, cellStates);
    }

    for (const [idx, state] of cellStates.entries()) {
      updateAggState(state, record[valueIndices[idx]] ?? null);
    }
  }

  const sortedRowKeys = [...rowKeys].sort((a, b) => a.localeCompare(b));
  const sortedColKeys = [...colKeys].sort((a, b) => a.localeCompare(b));

  const output: CellScalar[][] = [];

  const header: CellScalar[] = [];
  for (const name of rowFields) header.push(name);

  if (hasColumns) {
    for (const colKey of sortedColKeys) {
      const parts = colKeyParts.get(colKey) ?? [];
      const label = parts.map(normalizeKeyPart).filter(Boolean).join(" / ") || "(blank)";
      for (const spec of valueSpecs) {
        header.push(`${label} - ${aggLabel(spec.aggregation)} of ${spec.field}`);
      }
    }
    for (const spec of valueSpecs) {
      header.push(`Grand Total - ${aggLabel(spec.aggregation)} of ${spec.field}`);
    }
  } else {
    for (const spec of valueSpecs) {
      header.push(`${aggLabel(spec.aggregation)} of ${spec.field}`);
    }
  }

  output.push(header);

  for (const rowKey of sortedRowKeys) {
    const parts = rowKeyParts.get(rowKey) ?? [];
    const row: CellScalar[] = [...parts];
    const rowMap = cube.get(rowKey);
    const rowTotals = valueSpecs.map(() => initAggState());

    for (const colKey of sortedColKeys) {
      const cellStates = rowMap?.get(colKey);
      if (cellStates) {
        for (const [idx, state] of cellStates.entries()) {
          row.push(finalizeAgg(state, valueSpecs[idx].aggregation));
          mergeAggState(rowTotals[idx], state);
        }
      } else {
        for (const spec of valueSpecs) row.push(finalizeAgg(initAggState(), spec.aggregation));
      }
    }

    if (hasColumns) {
      for (const [idx, total] of rowTotals.entries()) {
        row.push(finalizeAgg(total, valueSpecs[idx].aggregation));
      }
    }

    output.push(row);
  }

  if (sortedRowKeys.length > 0) {
    const grandTotalsByCol = new Map<string, AggState[]>();
    const grandTotalsAll = valueSpecs.map(() => initAggState());
    for (const colKey of sortedColKeys) {
      grandTotalsByCol.set(colKey, valueSpecs.map(() => initAggState()));
    }

    for (const rowKey of sortedRowKeys) {
      const rowMap = cube.get(rowKey);
      if (!rowMap) continue;
      for (const colKey of sortedColKeys) {
        const cellStates = rowMap.get(colKey);
        if (!cellStates) continue;
        const colTotals = grandTotalsByCol.get(colKey);
        if (!colTotals) continue;
        for (const [idx, state] of cellStates.entries()) {
          mergeAggState(colTotals[idx], state);
          mergeAggState(grandTotalsAll[idx], state);
        }
      }
    }

    const grandRow: CellScalar[] = [];
    if (rowFields.length > 0) {
      grandRow.push("Grand Total");
      for (let i = 1; i < rowFields.length; i++) grandRow.push(null);
    }

    for (const colKey of sortedColKeys) {
      const totals = grandTotalsByCol.get(colKey) ?? valueSpecs.map(() => initAggState());
      for (const [idx, state] of totals.entries()) {
        grandRow.push(finalizeAgg(state, valueSpecs[idx].aggregation));
      }
    }
    if (hasColumns) {
      for (const [idx, total] of grandTotalsAll.entries()) {
        grandRow.push(finalizeAgg(total, valueSpecs[idx].aggregation));
      }
    }

    output.push(grandRow);
  }

  return output;
}

function nowMs(): number {
  if (typeof performance !== "undefined" && typeof performance.now === "function") return performance.now();
  return Date.now();
}

function rangesIntersectNormalized(a: DlpNormalizedRange, b: DlpNormalizedRange): boolean {
  return a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;
}

function rangeCellCount(range: { startRow: number; endRow: number; startCol: number; endCol: number }): number {
  const rows = range.endRow - range.startRow + 1;
  const cols = range.endCol - range.startCol + 1;
  return Math.max(0, rows) * Math.max(0, cols);
}

function enforceReadRangeCharLimit(params: {
  range: { sheet: string; startRow: number; endRow: number; startCol: number; endCol: number };
  values: CellScalar[][];
  formulas?: Array<Array<string | null>>;
  maxChars: number;
}): void {
  const estimated = estimateReadRangeChars(params.values, params.formulas);
  if (estimated <= params.maxChars) return;
  throw toolError(
    "permission_denied",
    `read_range result for ${formatA1Range(params.range)} is too large (~${estimated} chars), exceeding max_read_range_chars (${params.maxChars}). Request a smaller range or increase max_read_range_chars.`
  );
}

function estimateReadRangeChars(values: CellScalar[][], formulas?: Array<Array<string | null>>): number {
  let chars = 0;
  for (const row of values) {
    if (!Array.isArray(row)) continue;
    for (const cell of row) {
      chars += estimateJsonScalarChars(cell) + 2;
      if (chars > Number.MAX_SAFE_INTEGER) return chars;
    }
  }
  if (formulas) {
    for (const row of formulas) {
      if (!Array.isArray(row)) continue;
      for (const cell of row) {
        chars += estimateJsonScalarChars(cell) + 2;
        if (chars > Number.MAX_SAFE_INTEGER) return chars;
      }
    }
  }
  return chars;
}

function estimateJsonScalarChars(value: unknown): number {
  if (value === null || value === undefined) return 4; // "null"
  if (typeof value === "string") return value.length + 2; // quotes
  if (typeof value === "number") return String(value).length;
  if (typeof value === "boolean") return value ? 4 : 5;
  // Defensive (CellScalar should not include objects).
  try {
    return JSON.stringify(value).length;
  } catch {
    return String(value).length;
  }
}

function capList<T>(items: T[], maxItems: number): { items: T[]; truncated: boolean; total: number } {
  const max = Math.max(0, maxItems);
  if (items.length <= max) return { items, truncated: false, total: items.length };
  return { items: items.slice(0, max), truncated: true, total: items.length };
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

function compareCellForSort(left: CellData, right: CellData, opts: { includeFormulaValues?: boolean } = {}): number {
  const leftValue = cellComparableValue(left, opts);
  const rightValue = cellComparableValue(right, opts);
  return compareScalars(leftValue, rightValue);
}

function cellComparableValue(cell: CellData, opts: { includeFormulaValues?: boolean } = {}): string | number | boolean | null {
  if (cell.formula) {
    if (opts.includeFormulaValues) return cell.value ?? null;
    return cell.formula;
  }
  return cell.value ?? null;
}

function compareScalars(left: CellScalar | string, right: CellScalar | string): number {
  if (left === right) return 0;
  if (left === null) return -1;
  if (right === null) return 1;

  const leftNum = parseSpreadsheetNumber(left);
  const rightNum = parseSpreadsheetNumber(right);
  if (leftNum !== null && rightNum !== null) return leftNum - rightNum;
  return String(left).localeCompare(String(right));
}

function matchesCriterion(
  cell: CellData,
  criterion: { operator: string; value: string | number; value2?: string | number },
  opts: { includeFormulaValues?: boolean } = {}
): boolean {
  const comparable = cellComparableValue(cell, opts);
  switch (criterion.operator) {
    case "equals":
      return String(comparable ?? "") === String(criterion.value);
    case "contains":
      return String(comparable ?? "").includes(String(criterion.value));
    case "greater": {
      const a = parseSpreadsheetNumber(comparable);
      const b = parseSpreadsheetNumber(criterion.value);
      return a !== null && b !== null && a > b;
    }
    case "less": {
      const a = parseSpreadsheetNumber(comparable);
      const b = parseSpreadsheetNumber(criterion.value);
      return a !== null && b !== null && a < b;
    }
    case "between": {
      if (criterion.value2 === undefined) return false;
      const a = parseSpreadsheetNumber(comparable);
      const low = parseSpreadsheetNumber(criterion.value);
      const high = parseSpreadsheetNumber(criterion.value2);
      return a !== null && low !== null && high !== null && a >= low && a <= high;
    }
    default:
      return false;
  }
}

function toNumber(cell: CellData, opts: { includeFormulaValues?: boolean } = {}): number | null {
  if (cell.formula && !opts.includeFormulaValues) return null;
  return parseSpreadsheetNumber(cell.value);
}

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  return quantileSorted(sorted, 0.5);
}

function quantileSorted(sorted: number[], q: number): number {
  if (sorted.length === 0) return NaN;
  const pos = (sorted.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  if (sorted[base + 1] === undefined) return sorted[base]!;
  return sorted[base]! + rest * (sorted[base + 1]! - sorted[base]!);
}

function quantile(values: number[], q: number): number {
  if (values.length === 0) return NaN;
  const sorted = [...values].sort((a, b) => a - b);
  return quantileSorted(sorted, q);
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

interface IsolationTreeNode {
  size: number;
  split?: number;
  left?: IsolationTreeNode;
  right?: IsolationTreeNode;
}

function fnv1a32(value: string): number {
  // 32-bit FNV-1a hash.
  let hash = 0x811c9dc5;
  for (let i = 0; i < value.length; i++) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}

function mulberry32(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function sampleIndicesWithoutReplacement(length: number, sampleSize: number, rng: () => number): number[] {
  const count = Math.min(sampleSize, length);
  if (count <= 0) return [];

  // Partial Fisher-Yates shuffle without materializing a full `[0..length)` array.
  // We represent the shuffled index array lazily using a sparse swap map.
  const swapByIndex = new Map<number, number>();
  const result = new Array<number>(count);

  const get = (index: number): number => swapByIndex.get(index) ?? index;

  for (let i = 0; i < count; i++) {
    const j = i + Math.floor(rng() * (length - i));
    const valueAtI = get(i);
    const valueAtJ = get(j);
    result[i] = valueAtJ;

    // After swapping positions i and j, only position j remains in the pool.
    if (valueAtI === j) swapByIndex.delete(j);
    else swapByIndex.set(j, valueAtI);

    // Position i will never be read again (future j's are always >= next i), so drop it.
    swapByIndex.delete(i);
  }
  return result;
}

const harmonicNumberCache: number[] = [0];

function harmonicNumber(n: number): number {
  for (let i = harmonicNumberCache.length; i <= n; i++) {
    harmonicNumberCache[i] = harmonicNumberCache[i - 1]! + 1 / i;
  }
  return harmonicNumberCache[n]!;
}

const isolationForestAveragePathLengthCache: number[] = [];

function isolationForestAveragePathLength(n: number): number {
  // c(n) in the isolation forest paper: average path length of unsuccessful search in a BST.
  const cached = isolationForestAveragePathLengthCache[n];
  if (cached !== undefined) return cached;

  let next: number;
  if (n <= 1) next = 0;
  else if (n === 2) next = 1;
  else next = 2 * harmonicNumber(n - 1) - (2 * (n - 1)) / n;

  isolationForestAveragePathLengthCache[n] = next;
  return next;
}

function buildIsolationTree(values: number[], depth: number, maxDepth: number, rng: () => number): IsolationTreeNode {
  const size = values.length;
  if (size <= 1 || depth >= maxDepth) return { size };

  let min = Infinity;
  let max = -Infinity;
  for (const value of values) {
    min = Math.min(min, value);
    max = Math.max(max, value);
  }

  // All values are identical -> cannot split.
  if (min === max) return { size };

  const split = min + rng() * (max - min);
  const leftValues: number[] = [];
  const rightValues: number[] = [];
  for (const value of values) {
    if (value <= split) leftValues.push(value);
    else rightValues.push(value);
  }

  // Defensive: if a degenerate split happens (should be extremely rare), stop growing this branch.
  if (leftValues.length === 0 || rightValues.length === 0) return { size };

  return {
    size,
    split,
    left: buildIsolationTree(leftValues, depth + 1, maxDepth, rng),
    right: buildIsolationTree(rightValues, depth + 1, maxDepth, rng)
  };
}

function isolationTreePathLength(node: IsolationTreeNode, value: number, depth: number): number {
  if (!node.left || !node.right || node.split === undefined) {
    return depth + isolationForestAveragePathLength(node.size);
  }
  if (value <= node.split) return isolationTreePathLength(node.left, value, depth + 1);
  return isolationTreePathLength(node.right, value, depth + 1);
}

function isolationForestScores(values: number[], options: { seed: number; trees?: number; sampleSize?: number }): number[] {
  const trees = options.trees ?? 100;
  const sampleSize = Math.min(options.sampleSize ?? 256, values.length);
  const cSample = isolationForestAveragePathLength(sampleSize);
  if (values.length === 0 || trees <= 0 || sampleSize <= 1 || cSample === 0) {
    return values.map(() => 0);
  }

  const rng = mulberry32(options.seed);
  const maxDepth = Math.ceil(Math.log2(sampleSize));
  const pathLengthSums = new Array<number>(values.length).fill(0);

  for (let t = 0; t < trees; t++) {
    const sampleIndices = sampleIndicesWithoutReplacement(values.length, sampleSize, rng);
    const sampleValues = sampleIndices.map((idx) => values[idx]!);
    const tree = buildIsolationTree(sampleValues, 0, maxDepth, rng);

    for (let i = 0; i < values.length; i++) {
      pathLengthSums[i]! += isolationTreePathLength(tree, values[i]!, 0);
    }
  }

  return pathLengthSums.map((sum) => {
    const avgPath = sum / trees;
    return Math.pow(2, -avgPath / cSample);
  });
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
  if (!cellValuesEqual(left.value, right.value)) return false;
  if ((left.formula ?? null) !== (right.formula ?? null)) return false;
  const leftFormat = left.format ?? {};
  const rightFormat = right.format ?? {};
  const leftKeys = Object.keys(leftFormat);
  const rightKeys = Object.keys(rightFormat);
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every((key) => (leftFormat as any)[key] === (rightFormat as any)[key]);
}

function cellValuesEqual(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (typeof left !== typeof right) return false;
  if (left === null || right === null) return left === right;

  if (typeof left === "object") {
    try {
      return JSON.stringify(left) === JSON.stringify(right);
    } catch {
      return false;
    }
  }

  return false;
}

function jsonToTable(payload: unknown, options: { maxCells?: number } = {}): CellScalar[][] {
  const rawMaxCells = options.maxCells;
  const maxCells = (() => {
    if (!Number.isFinite(rawMaxCells)) return null;
    const n = Number(rawMaxCells);
    if (!Number.isFinite(n) || n <= 0) return null;
    return Math.floor(n);
  })();

  const assertWithinMaxCells = (rows: number, cols: number): void => {
    if (maxCells == null) return;
    const cellCount = rows * cols;
    if (!Number.isFinite(cellCount) || cellCount < 0) {
      throw toolError(
        "permission_denied",
        `fetch_external_data would materialize an unsafe table size (rows=${rows}, cols=${cols}). Reduce the response size or increase max_tool_range_cells (${maxCells}).`
      );
    }
    if (cellCount > maxCells) {
      throw toolError(
        "permission_denied",
        `fetch_external_data would write ${cellCount} cells (rows=${rows}, cols=${cols}), which exceeds max_tool_range_cells (${maxCells}). Reduce the response size or increase max_tool_range_cells.`
      );
    }
  };

  if (Array.isArray(payload)) {
    if (payload.length === 0) return [[null]];
    if (payload.every((row) => Array.isArray(row))) {
      const rowCount = payload.length;
      if (rowCount > 0) {
        // Fast-path rejection: even 1 column per row would exceed the limit.
        assertWithinMaxCells(rowCount, 1);
      }

      let maxCols = 0;
      for (const row of payload as unknown[]) {
        const cols = Array.isArray(row) ? row.length : 0;
        if (cols > maxCols) maxCols = cols;
        if (maxCols > 0) {
          // Check incrementally so wide rows short-circuit before we allocate matrices.
          assertWithinMaxCells(rowCount, Math.max(maxCols, 1));
        }
      }

      const normalizedCols = Math.max(maxCols, 1);
      assertWithinMaxCells(rowCount, normalizedCols);

      return (payload as unknown[]).map((rawRow) => {
        const rowValues = (rawRow as unknown[]).map((value) => normalizeJsonScalar(value));
        if (rowValues.length < normalizedCols) {
          while (rowValues.length < normalizedCols) rowValues.push(null);
        }
        return rowValues;
      });
    }
    if (payload.every((row) => row && typeof row === "object" && !Array.isArray(row))) {
      const objects = payload as Array<Record<string, unknown>>;
      const headersSet = new Set<string>();
      const rowCount = objects.length + 1;
      for (const obj of objects) {
        for (const key of Object.keys(obj)) {
          if (headersSet.has(key)) continue;
          headersSet.add(key);
          assertWithinMaxCells(rowCount, headersSet.size);
        }
      }
      const headers = Array.from(headersSet);
      const rows = objects.map((obj) => headers.map((header) => normalizeJsonScalar(obj[header])));
      if (headers.length === 0) return [[null]];
      assertWithinMaxCells(rowCount, headers.length);
      return [headers, ...rows];
    }
    assertWithinMaxCells(1, Math.max((payload as unknown[]).length, 1));
    return [(payload as unknown[]).map((value) => normalizeJsonScalar(value))];
  }

  if (payload && typeof payload === "object") {
    const obj = payload as Record<string, unknown>;
    const headers = Object.keys(obj);
    const row = headers.map((header) => normalizeJsonScalar(obj[header]));
    if (headers.length === 0) return [[null]];
    assertWithinMaxCells(2, headers.length);
    return [headers, row];
  }

  return [[normalizeJsonScalar(payload)]];
}

function normalizeJsonScalar(value: unknown): CellScalar {
  if (value === null || value === undefined) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  return JSON.stringify(value);
}

function ensureExternalUrlAllowed(url: URL, allowedHosts: string[]): void {
  if (url.protocol !== "http:" && url.protocol !== "https:") {
    throw toolError("permission_denied", `External protocol "${url.protocol}" is not supported for fetch_external_data.`);
  }
  if (url.username || url.password) {
    throw toolError(
      "permission_denied",
      "External URLs with embedded credentials are not supported for fetch_external_data. Pass credentials via headers instead."
    );
  }
  if (allowedHosts.length === 0) {
    throw toolError(
      "permission_denied",
      "fetch_external_data requires an explicit host allowlist (allowed_external_hosts)."
    );
  }

  type AllowedHostEntry = { hostname: string; port?: string };
  const normalizeHostname = (value: string): string => {
    const normalized = String(value ?? "")
      .trim()
      .toLowerCase();
    // WHATWG URL includes brackets in `.hostname` for IPv6 literals (e.g. "[::1]").
    // Normalize by stripping those brackets so allowlist entries can be provided either
    // bracketed (`[::1]`) or unbracketed (`::1`).
    if (normalized.startsWith("[") && normalized.endsWith("]")) {
      return normalized.slice(1, -1);
    }
    return normalized;
  };
  const parseAllowedHostEntry = (value: string): AllowedHostEntry | null => {
    const trimmed = String(value ?? "")
      .trim()
      .toLowerCase();
    if (!trimmed) return null;

    // IPv6 hosts use brackets when combined with ports (e.g. "[::1]:443").
    const ipv6Match = trimmed.match(/^\[(?<hostname>[^\]]+)\](?::(?<port>\d+))?$/);
    if (ipv6Match?.groups?.hostname) {
      const hostname = normalizeHostname(ipv6Match.groups.hostname);
      const port = ipv6Match.groups.port?.trim();
      return port ? { hostname, port } : { hostname };
    }

    // Hostnames with optional explicit port (e.g. "api.example.com:8443").
    const hostPortMatch = trimmed.match(/^(?<hostname>[^:]+)(?::(?<port>\d+))?$/);
    if (hostPortMatch?.groups?.hostname) {
      const hostname = normalizeHostname(hostPortMatch.groups.hostname);
      const port = hostPortMatch.groups.port?.trim();
      return port ? { hostname, port } : { hostname };
    }

    // Fall back to the full value as a hostname-only entry. This preserves previous strictness:
    // malformed allowlist entries won't accidentally match a broader set of URLs.
    return { hostname: normalizeHostname(trimmed) };
  };

  const urlHostname = normalizeHostname(url.hostname);
  // `URL.port` is empty when the URL uses the scheme default (80/443), even if the
  // caller explicitly wrote `:443`/`:80`. Use the effective port so allowlist entries
  // like `example.com:443` behave as expected.
  const urlPort = url.port || (url.protocol === "http:" ? "80" : "443");

  const allowlist = allowedHosts.map(parseAllowedHostEntry).filter((entry): entry is AllowedHostEntry => entry !== null);
  const isAllowed = allowlist.some((entry) => {
    if (entry.port != null) {
      return entry.hostname === urlHostname && entry.port === urlPort;
    }
    return entry.hostname === urlHostname;
  });

  if (!isAllowed) {
    throw toolError("permission_denied", `External host "${url.host}" is not in the allowlist for fetch_external_data.`);
  }
}

function isRedirectStatus(status: number): boolean {
  return status === 301 || status === 302 || status === 303 || status === 307 || status === 308;
}

function safeUrlForProvenance(url: URL): string {
  return redactUrlSecrets(url);
}

async function readResponseBytes(response: Response, maxBytes: number): Promise<Uint8Array> {
  if (!response.body) return new Uint8Array();

  const bodyAny = response.body as any;
  if (typeof bodyAny.getReader === "function") {
    const reader = bodyAny.getReader();
    const chunks: Uint8Array[] = [];
    let total = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      if (!value) continue;
      total += value.byteLength;
      if (total > maxBytes) {
        try {
          await reader.cancel();
        } catch {
          // ignore
        }
        throw toolError("permission_denied", `External response too large (>${maxBytes} bytes). Increase max_external_bytes to allow.`);
      }
      chunks.push(value);
    }
    const combined = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
      combined.set(chunk, offset);
      offset += chunk.byteLength;
    }
    return combined;
  }

  const buffer = new Uint8Array(await response.arrayBuffer());
  if (buffer.byteLength > maxBytes) {
    throw toolError("permission_denied", `External response too large (>${maxBytes} bytes). Increase max_external_bytes to allow.`);
  }
  return buffer;
}

function decodeUtf8(bytes: Uint8Array): string {
  if (bytes.byteLength === 0) return "";
  if (typeof TextDecoder !== "undefined") return new TextDecoder().decode(bytes);
  // Fallback for environments without TextDecoder.
  return Buffer.from(bytes).toString("utf8");
}

async function cancelResponseBody(response: Response): Promise<void> {
  try {
    await response.body?.cancel();
  } catch {
    // ignore
  }
}
