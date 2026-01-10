/**
 * Canonical action identifiers used by the DLP policy engine.
 *
 * These are intentionally stable strings so policy documents can be stored in
 * local storage and the cloud backend without migrations caused by code
 * refactors.
 */
export const DLP_ACTION = Object.freeze({
  SHARE_EXTERNAL_LINK: "sharing.externalLink",
  EXPORT_CSV: "export.csv",
  EXPORT_PDF: "export.pdf",
  EXPORT_XLSX: "export.xlsx",
  CLIPBOARD_COPY: "clipboard.copy",
  AI_CLOUD_PROCESSING: "ai.cloudProcessing",
  EXTERNAL_CONNECTOR: "connector.external",
});

