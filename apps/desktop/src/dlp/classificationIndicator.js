import { CLASSIFICATION_LEVEL } from "../../../../packages/security/dlp/src/classification.js";

/**
 * Returns a small view model for a status bar indicator.
 *
 * In the real desktop app this would be rendered as a React component; here we keep it
 * framework-agnostic so it can be used by tests and non-React shells.
 */
export function getClassificationIndicator(classification) {
  const level = classification?.level || CLASSIFICATION_LEVEL.PUBLIC;
  switch (level) {
    case CLASSIFICATION_LEVEL.PUBLIC:
      return { text: "Public", tone: "neutral" };
    case CLASSIFICATION_LEVEL.INTERNAL:
      return { text: "Internal", tone: "info" };
    case CLASSIFICATION_LEVEL.CONFIDENTIAL:
      return { text: "Confidential", tone: "warning" };
    case CLASSIFICATION_LEVEL.RESTRICTED:
      return { text: "Restricted", tone: "danger" };
    default:
      return { text: String(level), tone: "neutral" };
  }
}
