import type { PreviewApprovalRequest } from "../../../../../packages/ai-tools/src/llm/integration.js";
import type { CellData } from "../../../../../packages/ai-tools/src/spreadsheet/types.js";

function safeStringify(value: unknown): string {
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function formatCellData(cell: CellData): string {
  if (!cell) return "null";
  if (typeof cell.formula === "string" && cell.formula.length > 0) return `formula=${cell.formula}`;
  return `value=${safeStringify(cell.value)}`;
}

export interface PreviewApprovalPromptOptions {
  max_changes?: number;
}

export function formatPreviewApprovalPrompt(request: PreviewApprovalRequest, options: PreviewApprovalPromptOptions = {}): string {
  const maxChanges = options.max_changes ?? 10;
  const { call, preview } = request;

  const lines: string[] = [];
  lines.push(`AI wants to run: ${call.name}`);
  lines.push(`Arguments: ${safeStringify(call.arguments)}`);
  lines.push("");
  lines.push(
    `Summary: ${preview.summary.total_changes} changes (creates=${preview.summary.creates}, modifies=${preview.summary.modifies}, deletes=${preview.summary.deletes})`,
  );

  if (preview.approval_reasons.length) {
    lines.push(`Reasons: ${preview.approval_reasons.join("; ")}`);
  }

  if (preview.warnings.length) {
    lines.push("");
    lines.push("Warnings:");
    for (const warning of preview.warnings) lines.push(`- ${warning}`);
  }

  if (preview.changes.length) {
    lines.push("");
    lines.push(`Changes (showing ${Math.min(maxChanges, preview.changes.length)} of ${preview.changes.length}):`);
    for (const change of preview.changes.slice(0, maxChanges)) {
      lines.push(`- ${change.cell} (${change.type}): ${formatCellData(change.before)} -> ${formatCellData(change.after)}`);
    }
  }

  lines.push("");
  lines.push("Approve?");

  return lines.join("\n");
}

export async function confirmPreviewApproval(request: PreviewApprovalRequest): Promise<boolean> {
  const prompt = formatPreviewApprovalPrompt(request);
  if (typeof globalThis.confirm !== "function") return false;
  return globalThis.confirm(prompt);
}
