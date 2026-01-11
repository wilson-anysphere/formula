import { parseA1Range, type RangeAddress } from "../spreadsheet/a1.js";
import type { ToolExecutionResult, ToolExecutionResultBase } from "../executor/tool-executor.js";

import { DLP_ACTION } from "../../../security/dlp/src/actions.js";
import { DLP_DECISION, evaluatePolicy } from "../../../security/dlp/src/policyEngine.js";
import { effectiveCellClassification, effectiveRangeClassification } from "../../../security/dlp/src/selectors.js";
import { DlpViolationError } from "../../../security/dlp/src/errors.js";

export interface ToolOutputDlpOptions {
  /**
   * Workbook/document identifier used by the classification store.
   */
  documentId: string;
  /**
   * Policy object consumed by `packages/security/dlp`.
   */
  policy: any;
  /**
   * Optional classification records for this document (preferred in tests).
   * If omitted, `classificationStore.list(documentId)` is used.
   */
  classificationRecords?: Array<{ selector: any; classification: any }>;
  classificationStore?: { list(documentId: string): Array<{ selector: any; classification: any }> };
  /**
   * Allows Restricted content to be sent to the cloud only when the policy explicitly
   * permits it. Default: false.
   */
  includeRestrictedContent?: boolean;
  /**
   * Optional audit logger hook (e.g. enterprise audit log pipeline).
   */
  auditLogger?: { log(event: any): void };
  /**
   * Deterministic placeholder used to replace disallowed cell content.
   * Default: "[REDACTED]".
   */
  redactionPlaceholder?: string;
}

export interface ToolOutputDlpEnforcementParams {
  call: { id?: string; name: string; arguments: unknown };
  result: ToolExecutionResult;
  dlp: ToolOutputDlpOptions;
  defaultSheet?: string;
}

/**
 * Enforce DLP policy for spreadsheet tool outputs that might be sent to a cloud LLM.
 *
 * IMPORTANT: This function must not leak restricted cell contents via any field in the
 * tool result object. When DLP is enabled, callers should use the returned result (not
 * the raw executor result) as the payload stored in the tool message history.
 */
export function enforceToolOutputDlp(params: ToolOutputDlpEnforcementParams): ToolExecutionResult {
  const { call, result, dlp } = params;
  // Never modify error results: they should not contain cell contents, and downstream code
  // expects validation/runtime errors to pass through verbatim.
  if (!result.ok) return result;

  switch (call.name) {
    case "read_range":
      return enforceReadRange({ ...params, result: asOkResult(result) });
    case "compute_statistics":
      return enforceComputeStatistics({ ...params, result: asOkResult(result) });
    case "detect_anomalies":
      return enforceDetectAnomalies({ ...params, result: asOkResult(result) });
    case "filter_range":
      return enforceFilterRange({ ...params, result: asOkResult(result) });
    default:
      // Most tools do not return workbook contents. For safety we only apply DLP to known
      // exfiltration surfaces (read_range + tools derived from input ranges).
      return result;
  }
}

type OkToolResult = ToolExecutionResult & { ok: true };

function asOkResult(result: ToolExecutionResult): OkToolResult {
  if (!result.ok) throw new Error("Expected ok result");
  return result as OkToolResult;
}

function getRecords(dlp: ToolOutputDlpOptions): Array<{ selector: any; classification: any }> {
  return dlp.classificationRecords ?? dlp.classificationStore?.list(dlp.documentId) ?? [];
}

function getRedactionPlaceholder(dlp: ToolOutputDlpOptions): string {
  return dlp.redactionPlaceholder ?? "[REDACTED]";
}

function evaluateRangeDecision(params: {
  dlp: ToolOutputDlpOptions;
  range: string;
  defaultSheet?: string;
}): {
  parsedRange: RangeAddress;
  selectionClassification: any;
  decision: any;
} {
  const { dlp, range, defaultSheet } = params;
  const parsedRange = parseA1Range(range, defaultSheet);
  const rangeRef = {
    documentId: dlp.documentId,
    sheetId: parsedRange.sheet,
    range: {
      start: { row: parsedRange.startRow - 1, col: parsedRange.startCol - 1 },
      end: { row: parsedRange.endRow - 1, col: parsedRange.endCol - 1 },
    },
  };

  const records = getRecords(dlp);
  const includeRestrictedContent = dlp.includeRestrictedContent ?? false;

  const selectionClassification = effectiveRangeClassification(rangeRef, records);
  const decision = evaluatePolicy({
    action: DLP_ACTION.AI_CLOUD_PROCESSING,
    classification: selectionClassification,
    policy: dlp.policy,
    options: { includeRestrictedContent },
  });

  return { parsedRange, selectionClassification, decision };
}

function blockedResult<TName extends string>(
  result: ToolExecutionResultBase<any>,
  tool: TName,
  decision: any,
  messageOverride?: string,
): ToolExecutionResult {
  const message = messageOverride ?? new DlpViolationError(decision).message;
  return {
    tool: tool as any,
    ok: false,
    timing: result.timing,
    error: { code: "permission_denied", message },
  } as ToolExecutionResult;
}

function withWarning(result: OkToolResult, warning: string): OkToolResult {
  const warnings = [...(result.warnings ?? [])];
  warnings.push(warning);
  return { ...result, warnings };
}

function enforceReadRange(params: ToolOutputDlpEnforcementParams & { result: OkToolResult }): ToolExecutionResult {
  const { call, result, dlp, defaultSheet } = params;
  if (result.tool !== "read_range" || !result.data) return result;

  const records = getRecords(dlp);
  const includeRestrictedContent = dlp.includeRestrictedContent ?? false;
  const placeholder = getRedactionPlaceholder(dlp);

  const { parsedRange, selectionClassification, decision } = evaluateRangeDecision({
    dlp,
    range: result.data.range,
    defaultSheet,
  });

  if (decision.decision === DLP_DECISION.BLOCK) {
    dlp.auditLogger?.log({
      type: "ai.tool_dlp",
      documentId: dlp.documentId,
      tool: call.name,
      toolCallId: call.id,
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: result.data.range,
      selectionClassification,
      decision,
      redactedCellCount: 0,
    });
    return blockedResult(result, call.name, decision);
  }

  let redactedCellCount = 0;

  const values = result.data.values.map((row, rIndex) =>
    row.map((value, cIndex) => {
      const row0 = parsedRange.startRow - 1 + rIndex;
      const col0 = parsedRange.startCol - 1 + cIndex;
      const classification = effectiveCellClassification(
        { documentId: dlp.documentId, sheetId: parsedRange.sheet, row: row0, col: col0 },
        records,
      );
      const cellDecision = evaluatePolicy({
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        classification,
        policy: dlp.policy,
        options: { includeRestrictedContent },
      });
      if (cellDecision.decision === DLP_DECISION.ALLOW) return value;
      redactedCellCount += 1;
      return placeholder;
    }),
  );

  const formulas =
    result.data.formulas?.map((row, rIndex) =>
      row.map((formula, cIndex) => {
        const row0 = parsedRange.startRow - 1 + rIndex;
        const col0 = parsedRange.startCol - 1 + cIndex;
        const classification = effectiveCellClassification(
          { documentId: dlp.documentId, sheetId: parsedRange.sheet, row: row0, col: col0 },
          records,
        );
        const cellDecision = evaluatePolicy({
          action: DLP_ACTION.AI_CLOUD_PROCESSING,
          classification,
          policy: dlp.policy,
          options: { includeRestrictedContent },
        });
        if (cellDecision.decision === DLP_DECISION.ALLOW) return formula;
        // Keep formulas in sync with values: a redacted cell should not reveal its formula.
        return placeholder;
      }),
    ) ?? undefined;

  const nextData = { ...result.data, values, ...(formulas ? { formulas } : {}) };
  const next = redactedCellCount > 0 ? withWarning({ ...result, data: nextData }, `DLP: ${redactedCellCount} cells redacted.`) : { ...result, data: nextData };

  dlp.auditLogger?.log({
    type: "ai.tool_dlp",
    documentId: dlp.documentId,
    tool: call.name,
    toolCallId: call.id,
    action: DLP_ACTION.AI_CLOUD_PROCESSING,
    range: result.data.range,
    selectionClassification,
    decision,
    redactedCellCount,
  });

  return next;
}

/**
 * Derived-range tools do not directly return source cell values, but their outputs are still
 * a form of data exfiltration (e.g. statistics can reveal confidential numbers).
 *
 * Decision: Treat derived tool outputs as *fully derived from the input selection*.
 * - If the selection is BLOCKed, return a permission_denied tool result.
 * - If the selection is REDACTed, return a deterministic redacted payload that contains no
 *   derived values.
 */
function enforceComputeStatistics(
  params: ToolOutputDlpEnforcementParams & { result: OkToolResult },
): ToolExecutionResult {
  const { call, result, dlp, defaultSheet } = params;
  if (result.tool !== "compute_statistics" || !result.data) return result;

  const { selectionClassification, decision } = evaluateRangeDecision({ dlp, range: result.data.range, defaultSheet });
  if (decision.decision === DLP_DECISION.BLOCK) {
    dlp.auditLogger?.log({
      type: "ai.tool_dlp",
      documentId: dlp.documentId,
      tool: call.name,
      toolCallId: call.id,
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: result.data.range,
      selectionClassification,
      decision,
      redactedDerived: true,
    });
    return blockedResult(result, call.name, decision);
  }

  if (decision.decision !== DLP_DECISION.REDACT) {
    dlp.auditLogger?.log({
      type: "ai.tool_dlp",
      documentId: dlp.documentId,
      tool: call.name,
      toolCallId: call.id,
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: result.data.range,
      selectionClassification,
      decision,
      redactedDerived: false,
    });
    return result;
  }

  const redactedStats: Record<string, number | null> = {};
  for (const key of Object.keys(result.data.statistics ?? {})) {
    redactedStats[key] = null;
  }

  const next = withWarning(
    { ...result, data: { ...result.data, statistics: redactedStats } },
    "DLP: statistics redacted.",
  );

  dlp.auditLogger?.log({
    type: "ai.tool_dlp",
    documentId: dlp.documentId,
    tool: call.name,
    toolCallId: call.id,
    action: DLP_ACTION.AI_CLOUD_PROCESSING,
    range: result.data.range,
    selectionClassification,
    decision,
    redactedDerived: true,
  });

  return next;
}

function enforceDetectAnomalies(params: ToolOutputDlpEnforcementParams & { result: OkToolResult }): ToolExecutionResult {
  const { call, result, dlp, defaultSheet } = params;
  if (result.tool !== "detect_anomalies" || !result.data) return result;

  const { selectionClassification, decision } = evaluateRangeDecision({ dlp, range: result.data.range, defaultSheet });

  if (decision.decision === DLP_DECISION.BLOCK) {
    dlp.auditLogger?.log({
      type: "ai.tool_dlp",
      documentId: dlp.documentId,
      tool: call.name,
      toolCallId: call.id,
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: result.data.range,
      selectionClassification,
      decision,
      redactedDerived: true,
    });
    return blockedResult(result, call.name, decision);
  }

  if (decision.decision !== DLP_DECISION.REDACT) {
    dlp.auditLogger?.log({
      type: "ai.tool_dlp",
      documentId: dlp.documentId,
      tool: call.name,
      toolCallId: call.id,
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: result.data.range,
      selectionClassification,
      decision,
      redactedDerived: false,
    });
    return result;
  }

  // Redaction strategy: remove anomalies entirely. Even cell coordinates/counts can leak
  // information about the underlying confidential distribution.
  const next = withWarning({ ...result, data: { ...result.data, anomalies: [] } as any }, "DLP: anomalies redacted.");

  dlp.auditLogger?.log({
    type: "ai.tool_dlp",
    documentId: dlp.documentId,
    tool: call.name,
    toolCallId: call.id,
    action: DLP_ACTION.AI_CLOUD_PROCESSING,
    range: result.data.range,
    selectionClassification,
    decision,
    redactedDerived: true,
  });

  return next;
}

function enforceFilterRange(params: ToolOutputDlpEnforcementParams & { result: OkToolResult }): ToolExecutionResult {
  const { call, result, dlp, defaultSheet } = params;
  if (result.tool !== "filter_range" || !result.data) return result;

  const { selectionClassification, decision } = evaluateRangeDecision({ dlp, range: result.data.range, defaultSheet });
  if (decision.decision === DLP_DECISION.BLOCK) {
    dlp.auditLogger?.log({
      type: "ai.tool_dlp",
      documentId: dlp.documentId,
      tool: call.name,
      toolCallId: call.id,
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: result.data.range,
      selectionClassification,
      decision,
      redactedDerived: true,
    });
    return blockedResult(result, call.name, decision);
  }

  if (decision.decision !== DLP_DECISION.REDACT) {
    dlp.auditLogger?.log({
      type: "ai.tool_dlp",
      documentId: dlp.documentId,
      tool: call.name,
      toolCallId: call.id,
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      range: result.data.range,
      selectionClassification,
      decision,
      redactedDerived: false,
    });
    return result;
  }

  const next = withWarning(
    { ...result, data: { ...result.data, matching_rows: [], count: 0 } },
    "DLP: filter results redacted.",
  );

  dlp.auditLogger?.log({
    type: "ai.tool_dlp",
    documentId: dlp.documentId,
    tool: call.name,
    toolCallId: call.id,
    action: DLP_ACTION.AI_CLOUD_PROCESSING,
    range: result.data.range,
    selectionClassification,
    decision,
    redactedDerived: true,
  });

  return next;
}

