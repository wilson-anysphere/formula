import type { SpreadsheetApi } from "../spreadsheet/api.ts";
import { formatA1Cell } from "../spreadsheet/a1.ts";
import { isCellEmpty, type CellData } from "../spreadsheet/types.ts";
import type { UnknownToolCall } from "../tool-schema.ts";
import { ToolExecutor, type ToolExecutionResult, type ToolExecutorOptions } from "../executor/tool-executor.ts";

export type CellChangeType = "create" | "modify" | "delete";

export interface CellChangePreview {
  cell: string;
  type: CellChangeType;
  before: CellData;
  after: CellData;
}

export interface PreviewSummary {
  total_changes: number;
  creates: number;
  modifies: number;
  deletes: number;
}

export interface ToolPlanPreview {
  timing_ms: number;
  tool_results: ToolExecutionResult[];
  changes: CellChangePreview[];
  summary: PreviewSummary;
  warnings: string[];
  requires_approval: boolean;
  approval_reasons: string[];
}

export interface PreviewEngineOptions {
  max_preview_changes?: number;
  approval_cell_threshold?: number;
}

export class PreviewEngine {
  readonly options: Required<PreviewEngineOptions>;

  constructor(options: PreviewEngineOptions = {}) {
    this.options = {
      max_preview_changes: options.max_preview_changes ?? 20,
      approval_cell_threshold: options.approval_cell_threshold ?? 100
    };
  }

  /**
   * Simulate a tool plan without mutating the provided spreadsheet.
   */
  async generatePreview(
    toolCalls: UnknownToolCall[],
    spreadsheet: SpreadsheetApi,
    executorOptions: ToolExecutorOptions = {}
  ): Promise<ToolPlanPreview> {
    const started = nowMs();
    const before = spreadsheet.clone();
    const simulated = spreadsheet.clone();

    // Hard-disable external data fetches during preview to avoid side effects. When a plan
    // requests external data, ToolExecutor runs in `preview_mode` so `fetch_external_data`
    // returns a "skipped" result instead of surfacing a misleading permission error.
    const executor = new ToolExecutor(simulated, { ...executorOptions, allow_external_data: false, preview_mode: true });
    const toolResults = await executor.executePlan(toolCalls);

    const changes = diffSpreadsheets(before, simulated, executorOptions.sheet_name_resolver);
    const summary = summarizeChanges(changes);

    const warnings: string[] = [];
    let toolReportedCellsTouched = 0;
    let hasToolExecutionWarnings = false;
    for (const result of toolResults) {
      if (!result.ok) {
        hasToolExecutionWarnings = true;
        warnings.push(`${result.tool}: ${result.error?.message ?? "Tool failed"}`);
        continue;
      }

      // PreviewEngine diffs enumerate non-empty cells only. Some spreadsheet
      // backends store formatting in layered defaults / range runs without
      // materializing per-cell entries, so use tool-reported cell counts as a
      // conservative signal for approval gating.
      switch (result.tool) {
        case "apply_formatting":
          toolReportedCellsTouched += result.data?.formatted_cells ?? 0;
          break;
        case "set_range":
        case "apply_formula_column":
          toolReportedCellsTouched += result.data?.updated_cells ?? 0;
          break;
        case "create_pivot_table":
        case "fetch_external_data":
          toolReportedCellsTouched += result.data?.written_cells ?? 0;
          break;
      }
    }

    if (toolReportedCellsTouched > summary.total_changes) {
      warnings.push(
        `Preview diff may be incomplete: tools reported touching ${toolReportedCellsTouched} cells, but the cell-level diff captured ${summary.total_changes}. Formatting-only edits on empty cells may not materialize as per-cell changes.`
      );
    }

    const effectiveChangeCount = Math.max(summary.total_changes, toolReportedCellsTouched);

    const approvalReasons: string[] = [];
    if (effectiveChangeCount > this.options.approval_cell_threshold) {
      approvalReasons.push(`Large edit (${effectiveChangeCount} cells)`);
    }
    if (summary.deletes > 0) {
      approvalReasons.push(`Deletes detected (${summary.deletes} cells cleared)`);
    }
    if (toolCalls.some((call) => call.name === "fetch_external_data")) {
      approvalReasons.push("External data access requested");
    }
    if (hasToolExecutionWarnings) {
      approvalReasons.push("Tool execution warnings");
    }

    const requiresApproval = approvalReasons.length > 0;
    const previewChanges = changes.slice(0, this.options.max_preview_changes);

    return {
      timing_ms: nowMs() - started,
      tool_results: toolResults,
      changes: previewChanges,
      summary,
      warnings,
      requires_approval: requiresApproval,
      approval_reasons: approvalReasons
    };
  }
}

function nowMs(): number {
  if (typeof performance !== "undefined" && typeof performance.now === "function") return performance.now();
  return Date.now();
}

function diffSpreadsheets(
  before: SpreadsheetApi,
  after: SpreadsheetApi,
  sheetNameResolver?: ToolExecutorOptions["sheet_name_resolver"]
): CellChangePreview[] {
  type DiffCellEntry = {
    cell: CellData;
    cellRef: string;
    sheetId: string;
    sheetName: string;
    row: number;
    col: number;
  };

  const beforeMap = new Map<string, DiffCellEntry>();
  for (const entry of before.listNonEmptyCells()) {
    const key = diffKey(entry.address.sheet, entry.address.row, entry.address.col);
    const sheetId = entry.address.sheet;
    const sheetName = sheetNameResolver?.getSheetNameById(sheetId) ?? sheetId;
    const cellRef = formatA1Cell(sheetName === sheetId ? entry.address : { ...entry.address, sheet: sheetName });
    beforeMap.set(key, { cell: entry.cell, cellRef, sheetId, sheetName, row: entry.address.row, col: entry.address.col });
  }

  const afterMap = new Map<string, DiffCellEntry>();
  for (const entry of after.listNonEmptyCells()) {
    const key = diffKey(entry.address.sheet, entry.address.row, entry.address.col);
    const sheetId = entry.address.sheet;
    const sheetName = sheetNameResolver?.getSheetNameById(sheetId) ?? sheetId;
    const cellRef = formatA1Cell(sheetName === sheetId ? entry.address : { ...entry.address, sheet: sheetName });
    afterMap.set(key, { cell: entry.cell, cellRef, sheetId, sheetName, row: entry.address.row, col: entry.address.col });
  }

  const keys = new Set([...beforeMap.keys(), ...afterMap.keys()]);
  const changes: Array<CellChangePreview & { sheetSort: string; sheetIdSort: string; rowSort: number; colSort: number }> = [];
  for (const key of keys) {
    const beforeEntry = beforeMap.get(key);
    const afterEntry = afterMap.get(key);
    const beforeCell = beforeEntry?.cell ?? { value: null };
    const afterCell = afterEntry?.cell ?? { value: null };

    if (cellsEqual(beforeCell, afterCell)) continue;

    const beforeEmpty = isCellEmpty(beforeCell);
    const afterEmpty = isCellEmpty(afterCell);
    const type: CellChangeType = beforeEmpty && !afterEmpty ? "create" : !beforeEmpty && afterEmpty ? "delete" : "modify";
    const cellRef = afterEntry?.cellRef ?? beforeEntry?.cellRef ?? key;

    const sheetSort = afterEntry?.sheetName ?? beforeEntry?.sheetName ?? "";
    const sheetIdSort = afterEntry?.sheetId ?? beforeEntry?.sheetId ?? "";
    const rowSort = afterEntry?.row ?? beforeEntry?.row ?? 0;
    const colSort = afterEntry?.col ?? beforeEntry?.col ?? 0;
    changes.push({ cell: cellRef, type, before: beforeCell, after: afterCell, sheetSort, sheetIdSort, rowSort, colSort });
  }

  changes.sort((a, b) => {
    const sheetCmp = a.sheetSort.localeCompare(b.sheetSort);
    if (sheetCmp !== 0) return sheetCmp;
    const sheetIdCmp = a.sheetIdSort.localeCompare(b.sheetIdSort);
    if (sheetIdCmp !== 0) return sheetIdCmp;
    if (a.rowSort !== b.rowSort) return a.rowSort - b.rowSort;
    if (a.colSort !== b.colSort) return a.colSort - b.colSort;
    return a.cell.localeCompare(b.cell);
  });

  return changes.map(({ sheetSort: _sheetSort, sheetIdSort: _sheetIdSort, rowSort: _rowSort, colSort: _colSort, ...rest }) => rest);
}

function summarizeChanges(changes: CellChangePreview[]): PreviewSummary {
  let creates = 0;
  let deletes = 0;
  let modifies = 0;
  for (const change of changes) {
    switch (change.type) {
      case "create":
        creates++;
        break;
      case "delete":
        deletes++;
        break;
      case "modify":
        modifies++;
        break;
    }
  }
  return { total_changes: changes.length, creates, modifies, deletes };
}

function diffKey(sheet: string, row: number, col: number): string {
  return `${sheet}:${row}:${col}`;
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
    // Support DocumentController rich values (objects) without producing noisy diffs
    // between cloned workbooks.
    try {
      return JSON.stringify(left) === JSON.stringify(right);
    } catch {
      return false;
    }
  }

  return false;
}
