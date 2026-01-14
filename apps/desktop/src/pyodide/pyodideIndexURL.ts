import { getTauriEventApiOrNull } from "../tauri/api.js";

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

export type PyodideDownloadProgress = {
  kind: "checking" | "downloadStart" | "downloadProgress" | "downloadComplete" | "ready";
  fileName?: string | null;
  completedFiles: number;
  totalFiles: number;
  bytesDownloaded?: number | null;
  bytesTotal?: number | null;
  message?: string | null;
};

const PYODIDE_DOWNLOAD_PROGRESS_EVENT = "pyodide-download-progress";

function safeGetGlobal(name: string): unknown {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (globalThis as any)[name];
  } catch {
    return undefined;
  }
}

function getTauriInvokeOrNull(): TauriInvoke | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
    return typeof invoke === "function" ? invoke : null;
  } catch {
    return null;
  }
}

export function normalizePyodideIndexURL(raw: unknown): string | undefined {
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (!trimmed) return undefined;
  return trimmed.endsWith("/") ? trimmed : `${trimmed}/`;
}

export function pickPyodideIndexURL(params: {
  explicitIndexURL?: string | undefined;
  cachedIndexURL?: string | undefined;
}): string | undefined {
  return params.explicitIndexURL ?? params.cachedIndexURL;
}

export function getExplicitPyodideIndexURL(explicitIndexURL?: string | undefined): string | undefined {
  const normalizedExplicit = normalizePyodideIndexURL(explicitIndexURL);
  if (normalizedExplicit) return normalizedExplicit;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return normalizePyodideIndexURL((globalThis as any).__pyodideIndexURL);
}

function normalizeProgressPayload(payload: unknown): PyodideDownloadProgress | null {
  if (!payload || typeof payload !== "object") return null;
  const obj = payload as Record<string, unknown>;
  const kind = obj.kind;
  if (typeof kind !== "string") return null;
  if (
    kind !== "checking" &&
    kind !== "downloadStart" &&
    kind !== "downloadProgress" &&
    kind !== "downloadComplete" &&
    kind !== "ready"
  ) {
    return null;
  }

  const completedFiles = typeof obj.completedFiles === "number" ? obj.completedFiles : NaN;
  const totalFiles = typeof obj.totalFiles === "number" ? obj.totalFiles : NaN;
  if (!Number.isFinite(completedFiles) || !Number.isFinite(totalFiles)) return null;

  return {
    kind: kind as PyodideDownloadProgress["kind"],
    fileName: typeof obj.fileName === "string" ? obj.fileName : null,
    completedFiles,
    totalFiles,
    bytesDownloaded: typeof obj.bytesDownloaded === "number" ? obj.bytesDownloaded : null,
    bytesTotal: typeof obj.bytesTotal === "number" ? obj.bytesTotal : null,
    message: typeof obj.message === "string" ? obj.message : null,
  };
}

async function desktopPyodideIndexURL(options: {
  downloadIfMissing: boolean;
  onProgress?: ((progress: PyodideDownloadProgress) => void) | undefined;
}): Promise<string | undefined> {
  const invoke = getTauriInvokeOrNull();
  if (!invoke) return undefined;

  const tauriEvent = getTauriEventApiOrNull();
  let unlisten: (() => void) | null = null;

  if (tauriEvent && typeof options.onProgress === "function") {
    unlisten = await tauriEvent.listen(PYODIDE_DOWNLOAD_PROGRESS_EVENT, (event: any) => {
      const progress = normalizeProgressPayload(event?.payload);
      if (progress) {
        try {
          options.onProgress?.(progress);
        } catch {
          // ignore progress handler errors
        }
      }
    });
  }

  try {
    const result = await invoke("pyodide_index_url", { download: options.downloadIfMissing });
    return normalizePyodideIndexURL(result);
  } finally {
    try {
      unlisten?.();
    } catch {
      // ignore
    }
  }
}

export async function getCachedPyodideIndexURL(options: {
  explicitIndexURL?: string | undefined;
} = {}): Promise<string | undefined> {
  const explicit = getExplicitPyodideIndexURL(options.explicitIndexURL);
  if (explicit) return explicit;

  // Desktop-only: return the local protocol URL only if the cache is already present.
  return await desktopPyodideIndexURL({ downloadIfMissing: false });
}

export async function ensurePyodideIndexURL(options: {
  explicitIndexURL?: string | undefined;
  onProgress?: ((progress: PyodideDownloadProgress) => void) | undefined;
} = {}): Promise<string | undefined> {
  const explicit = getExplicitPyodideIndexURL(options.explicitIndexURL);
  if (explicit) return explicit;

  // Desktop-only: download + cache on demand.
  const resolved = await desktopPyodideIndexURL({
    downloadIfMissing: true,
    onProgress: options.onProgress,
  });

  // Web builds: fall back to PyodideRuntime's default CDN behavior.
  return resolved ?? normalizePyodideIndexURL(safeGetGlobal("__pyodideIndexURL"));
}

