import { DLP_DECISION } from "./policyEngine.js";
import { DLP_ACTION } from "./actions.js";

export class DlpViolationError extends Error {
  /**
   * @param {any} decision
   */
  constructor(decision) {
    super(formatDlpDecisionMessage(decision));
    this.name = "DlpViolationError";
    this.decision = decision;
  }
}

/**
 * @param {any} decision
 */
export function formatDlpDecisionMessage(decision) {
  if (!decision) return "Operation blocked by data loss prevention policy.";
  if (decision.decision !== DLP_DECISION.BLOCK && decision.decision !== DLP_DECISION.REDACT) return "";

  const classification = decision.classification?.level || "Unknown";
  const maxAllowed = decision.maxAllowed ?? "None";

  switch (decision.action) {
    case DLP_ACTION.CLIPBOARD_COPY:
      return `Clipboard copy is blocked by your organization's data loss prevention policy because the selected data is classified as ${classification} (max allowed: ${maxAllowed}).`;
    case DLP_ACTION.EXPORT_CSV:
    case DLP_ACTION.EXPORT_PDF:
    case DLP_ACTION.EXPORT_XLSX:
      return `Export is blocked by your organization's data loss prevention policy because the selected data is classified as ${classification} (max allowed: ${maxAllowed}).`;
    case DLP_ACTION.SHARE_EXTERNAL_LINK:
      return `Creating an external sharing link is blocked by your organization's data loss prevention policy because the data is classified as ${classification} (max allowed: ${maxAllowed}).`;
    case DLP_ACTION.EXTERNAL_CONNECTOR:
      return `Sending data to external connectors is blocked by your organization's data loss prevention policy because the data is classified as ${classification} (max allowed: ${maxAllowed}).`;
    case DLP_ACTION.AI_CLOUD_PROCESSING:
      return `Sending data to cloud AI is restricted by your organization's data loss prevention policy because the data is classified as ${classification} (max allowed: ${maxAllowed}).`;
    default:
      return `Operation blocked by your organization's data loss prevention policy (classification: ${classification}, max allowed: ${maxAllowed}).`;
  }
}

