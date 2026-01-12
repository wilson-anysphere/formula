export type StoredCollabConnection = {
  wsUrl: string;
  docId: string;
  updatedAtMs: number;
};

const CONNECTION_PREFIX = "formula:collab:connection:";

function storageKeyForWorkbook(workbookKey: string): string {
  return `${CONNECTION_PREFIX}${workbookKey}`;
}

function getLocalStorage(): Storage | null {
  try {
    if (typeof window === "undefined") return null;
    return window.localStorage ?? null;
  } catch {
    return null;
  }
}

export function saveCollabConnectionForWorkbook(opts: { workbookKey: string; wsUrl: string; docId: string }): void {
  const workbookKey = String(opts.workbookKey ?? "").trim();
  const wsUrl = String(opts.wsUrl ?? "").trim();
  const docId = String(opts.docId ?? "").trim();
  if (!workbookKey || !wsUrl || !docId) return;

  const storage = getLocalStorage();
  if (!storage) return;

  const payload: StoredCollabConnection = { wsUrl, docId, updatedAtMs: Date.now() };
  try {
    storage.setItem(storageKeyForWorkbook(workbookKey), JSON.stringify(payload));
  } catch {
    // ignore storage failures
  }
}

export function loadCollabConnectionForWorkbook(opts: { workbookKey: string }): StoredCollabConnection | null {
  const workbookKey = String(opts.workbookKey ?? "").trim();
  if (!workbookKey) return null;

  const storage = getLocalStorage();
  if (!storage) return null;

  let raw: string | null = null;
  try {
    raw = storage.getItem(storageKeyForWorkbook(workbookKey));
  } catch {
    return null;
  }
  if (!raw) return null;

  try {
    const parsed = JSON.parse(raw) as unknown;
    if (!parsed || typeof parsed !== "object") return null;
    const wsUrl = typeof (parsed as any).wsUrl === "string" ? String((parsed as any).wsUrl).trim() : "";
    const docId = typeof (parsed as any).docId === "string" ? String((parsed as any).docId).trim() : "";
    const updatedAtMs = typeof (parsed as any).updatedAtMs === "number" ? (parsed as any).updatedAtMs : NaN;
    if (!wsUrl || !docId) return null;
    return { wsUrl, docId, updatedAtMs: Number.isFinite(updatedAtMs) ? updatedAtMs : 0 };
  } catch {
    return null;
  }
}

export function clearCollabConnectionForWorkbook(opts: { workbookKey: string }): void {
  const workbookKey = String(opts.workbookKey ?? "").trim();
  if (!workbookKey) return;

  const storage = getLocalStorage();
  if (!storage) return;

  try {
    storage.removeItem(storageKeyForWorkbook(workbookKey));
  } catch {
    // ignore
  }
}

