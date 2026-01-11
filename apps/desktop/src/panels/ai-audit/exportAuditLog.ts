import type { AIAuditEntry } from "../../../../../packages/ai-audit/src/types.js";

export interface AuditLogExport {
  blob: Blob;
  fileName: string;
}

export function createAuditLogExport(entries: AIAuditEntry[], options: { fileName?: string } = {}): AuditLogExport {
  const json = JSON.stringify(entries, null, 2);
  const blob = new Blob([json], { type: "application/json" });
  const fileName = options.fileName ?? `ai-audit-log-${new Date().toISOString().replaceAll(":", "-")}.json`;
  return { blob, fileName };
}

export function downloadAuditLogExport(exp: AuditLogExport): void {
  // Download behavior is browser-specific. Keep the core export as a pure Blob,
  // and make the download best-effort so this can run in non-browser contexts.
  if (typeof document === "undefined") return;
  if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") return;

  const url = URL.createObjectURL(exp.blob);
  try {
    const a = document.createElement("a");
    a.href = url;
    a.download = exp.fileName;
    a.rel = "noopener";
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    a.remove();
  } finally {
    URL.revokeObjectURL(url);
  }
}

