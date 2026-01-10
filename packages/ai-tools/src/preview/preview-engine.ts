import type { SpreadsheetApi } from "../spreadsheet/api.js";
import { formatA1Cell } from "../spreadsheet/a1.js";
import { isCellEmpty, type CellData } from "../spreadsheet/types.js";
import type { UnknownToolCall } from "../tool-schema.js";
import { ToolExecutor, type ToolExecutionResult, type ToolExecutorOptions } from "../executor/tool-executor.js";

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

    // Hard-disable external data fetches during preview to avoid side effects.
    const executor = new ToolExecutor(simulated, { ...executorOptions, allow_external_data: false });
    const toolResults = await executor.executePlan(toolCalls);

    const changes = diffSpreadsheets(before, simulated);
    const summary = summarizeChanges(changes);

    const warnings: string[] = [];
    for (const result of toolResults) {
      if (!result.ok) {
        warnings.push(`${result.tool}: ${result.error?.message ?? "Tool failed"}`);
      }
    }

    const approvalReasons: string[] = [];
    if (summary.total_changes > this.options.approval_cell_threshold) {
      approvalReasons.push(`Large edit (${summary.total_changes} cells)`);
    }
    if (summary.deletes > 0) {
      approvalReasons.push(`Deletes detected (${summary.deletes} cells cleared)`);
    }
    if (toolCalls.some((call) => call.name === "fetch_external_data")) {
      approvalReasons.push("External data access requested");
    }
    if (warnings.length > 0) {
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

function diffSpreadsheets(before: SpreadsheetApi, after: SpreadsheetApi): CellChangePreview[] {
  const beforeMap = new Map<string, { cell: CellData; cellRef: string }>();
  for (const entry of before.listNonEmptyCells()) {
    const key = diffKey(entry.address.sheet, entry.address.row, entry.address.col);
    beforeMap.set(key, { cell: entry.cell, cellRef: formatA1Cell(entry.address) });
  }

  const afterMap = new Map<string, { cell: CellData; cellRef: string }>();
  for (const entry of after.listNonEmptyCells()) {
    const key = diffKey(entry.address.sheet, entry.address.row, entry.address.col);
    afterMap.set(key, { cell: entry.cell, cellRef: formatA1Cell(entry.address) });
  }

  const keys = new Set([...beforeMap.keys(), ...afterMap.keys()]);
  const changes: CellChangePreview[] = [];
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
    changes.push({ cell: cellRef, type, before: beforeCell, after: afterCell });
  }

  changes.sort((a, b) => a.cell.localeCompare(b.cell));
  return changes;
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
  if (left.value !== right.value) return false;
  if ((left.formula ?? null) !== (right.formula ?? null)) return false;
  const leftFormat = left.format ?? {};
  const rightFormat = right.format ?? {};
  const leftKeys = Object.keys(leftFormat);
  const rightKeys = Object.keys(rightFormat);
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every((key) => (leftFormat as any)[key] === (rightFormat as any)[key]);
}
