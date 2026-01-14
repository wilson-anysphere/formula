import type { QueryOperation } from "@formula/power-query";

import { t } from "../../i18n/index.js";

function humanizeOperationType(type: string): string {
  return type
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/^./, (ch) => ch.toUpperCase());
}

function formatDetails(details: string | null): string {
  const trimmed = details?.trim() ?? "";
  return trimmed ? ` (${trimmed})` : "";
}

function joinNames(names: string[] | null | undefined): string | null {
  if (!Array.isArray(names)) return null;
  const cleaned = names.map((name) => String(name).trim()).filter(Boolean);
  return cleaned.length > 0 ? cleaned.join(", ") : null;
}

/**
 * Human-readable label for a `QueryOperation`.
 *
 * This is used in multiple places (AI suggestions, applied steps) so we keep it
 * centralized to avoid drifting labels.
 */
export function formatQueryOperationLabel(op: QueryOperation): string {
  switch (op.type) {
    case "addColumn": {
      const name = typeof op.name === "string" ? op.name : null;
      return `${t("queryEditor.addStep.op.addColumn")}${formatDetails(name)}`;
    }
    case "take": {
      const count = typeof op.count === "number" && Number.isFinite(op.count) ? String(op.count) : null;
      return `${t("queryEditor.addStep.op.keepTopRows")}${formatDetails(count)}`;
    }
    case "filterRows": {
      const column =
        op.predicate?.type === "comparison" && typeof op.predicate.column === "string" ? op.predicate.column : null;
      return `${t("queryEditor.addStep.op.filterRows")}${formatDetails(column)}`;
    }
    case "sortRows": {
      const col = op.sortBy?.[0]?.column ?? null;
      return `${t("queryEditor.addStep.op.sort")}${formatDetails(col)}`;
    }
    case "removeColumns": {
      const cols = joinNames(op.columns);
      return `${t("queryEditor.addStep.op.removeColumns")}${formatDetails(cols)}`;
    }
    case "selectColumns": {
      const cols = joinNames(op.columns);
      return `${t("queryEditor.addStep.op.keepColumns")}${formatDetails(cols)}`;
    }
    case "renameColumn": {
      const oldName = typeof op.oldName === "string" ? op.oldName : "";
      const newName = typeof op.newName === "string" ? op.newName : "";
      const details = oldName && newName ? `${oldName} → ${newName}` : oldName || newName || null;
      return `${t("queryEditor.addStep.op.renameColumns")}${formatDetails(details)}`;
    }
    case "changeType": {
      const col = typeof op.column === "string" ? op.column : "";
      const type = typeof op.newType === "string" ? op.newType : "";
      const details = col && type ? `${col} → ${type}` : col || type || null;
      return `${t("queryEditor.addStep.op.changeType")}${formatDetails(details)}`;
    }
    case "splitColumn": {
      const col = typeof op.column === "string" ? op.column : null;
      return `${t("queryEditor.addStep.op.splitColumn")}${formatDetails(col)}`;
    }
    case "fillDown": {
      const cols = joinNames(op.columns);
      return `${t("queryEditor.addStep.op.fillDown")}${formatDetails(cols)}`;
    }
    case "replaceValues": {
      const col = typeof op.column === "string" ? op.column : null;
      return `${t("queryEditor.addStep.op.replaceValues")}${formatDetails(col)}`;
    }
    case "distinctRows": {
      const cols = joinNames(op.columns ?? null);
      return `${t("queryEditor.addStep.op.removeDuplicates")}${formatDetails(cols)}`;
    }
    case "removeRowsWithErrors": {
      const cols = joinNames(op.columns ?? null);
      return `${t("queryEditor.addStep.op.removeRowsWithErrors")}${formatDetails(cols)}`;
    }
    case "groupBy": {
      const groupCols = joinNames(op.groupColumns);
      return `${t("queryEditor.addStep.op.groupBy")}${formatDetails(groupCols)}`;
    }
    default:
      return humanizeOperationType(op.type);
  }
}
