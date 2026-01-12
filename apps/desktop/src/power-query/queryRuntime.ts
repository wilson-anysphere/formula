import type { Query } from "@formula/power-query";
import { formatDlpDecisionMessage } from "../../../../packages/security/dlp/src/errors.js";
import type { SheetNameResolver } from "../sheet/sheetNameResolver";
import { formatSheetNameForA1 } from "../sheet/formatSheetNameForA1.js";

export type QueryRunStatus = "idle" | "queued" | "refreshing" | "applying" | "success" | "error" | "cancelled";

export type QueryRuntime = {
  status: QueryRunStatus;
  jobId?: string;
  lastRefreshAtMs?: number;
  lastError?: string | null;
  rowsWritten?: number;
};

export type QueryRuntimeState = Record<string, QueryRuntime>;

function isAbortError(error: unknown): boolean {
  return (error as any)?.name === "AbortError";
}

export function summarizeQueryError(error: unknown): string {
  if (!error) return "Unknown error";
  if (typeof error === "string") return error;
  if (typeof (error as any)?.message === "string") return (error as any).message;
  if ((error as any)?.name === "DlpViolationError") {
    const decision = (error as any)?.decision;
    const msg = formatDlpDecisionMessage(decision);
    if (msg) return msg;
  }
  try {
    return JSON.stringify(error);
  } catch {
    return String(error);
  }
}

function runtimeFor(state: QueryRuntimeState, queryId: string): QueryRuntime {
  return state[queryId] ?? { status: "idle" };
}

export function reduceQueryRuntimeState(state: QueryRuntimeState, event: any): QueryRuntimeState {
  const queryId = (event?.job?.queryId ?? event?.queryId) as string | undefined;
  if (!queryId || typeof queryId !== "string") return state;

  const prev = runtimeFor(state, queryId);
  const next: QueryRuntime = { ...prev };

  switch (event?.type) {
    case "queued": {
      const jobId = event?.job?.id;
      if (typeof jobId === "string") next.jobId = jobId;
      next.status = "queued";
      break;
    }
    case "started": {
      const jobId = event?.job?.id;
      if (typeof jobId === "string") next.jobId = jobId;
      next.status = "refreshing";
      next.rowsWritten = undefined;
      break;
    }
    case "progress": {
      const jobId = event?.job?.id;
      if (typeof jobId === "string") next.jobId = jobId;
      // Keep the status as refreshing unless we're already applying.
      if (prev.status !== "applying") next.status = "refreshing";
      break;
    }
    case "completed": {
      const jobId = event?.job?.id;
      if (typeof jobId === "string") next.jobId = jobId;
      next.status = "success";
      const refreshedAt = event?.result?.meta?.refreshedAt;
      if (refreshedAt instanceof Date && !Number.isNaN(refreshedAt.getTime())) {
        next.lastRefreshAtMs = refreshedAt.getTime();
      } else if (event?.job?.completedAt instanceof Date && !Number.isNaN(event.job.completedAt.getTime())) {
        next.lastRefreshAtMs = event.job.completedAt.getTime();
      }
      next.lastError = null;
      break;
    }
    case "error": {
      const jobId = event?.job?.id;
      if (typeof jobId === "string") next.jobId = jobId;
      next.status = "error";
      next.lastError = summarizeQueryError(event?.error);
      break;
    }
    case "cancelled": {
      const jobId = event?.job?.id;
      if (typeof jobId === "string") next.jobId = jobId;
      next.status = "cancelled";
      next.lastError = null;
      break;
    }
    case "apply:started": {
      const jobId = event?.jobId;
      if (typeof jobId === "string") next.jobId = jobId;
      next.status = "applying";
      next.rowsWritten = 0;
      break;
    }
    case "apply:progress": {
      const jobId = event?.jobId;
      if (typeof jobId === "string") next.jobId = jobId;
      next.status = "applying";
      const rowsWritten = event?.rowsWritten;
      if (typeof rowsWritten === "number" && Number.isFinite(rowsWritten)) next.rowsWritten = rowsWritten;
      break;
    }
    case "apply:completed": {
      const jobId = event?.jobId;
      if (typeof jobId === "string") next.jobId = jobId;
      next.status = "success";
      next.lastError = null;
      next.rowsWritten = undefined;
      break;
    }
    case "apply:error": {
      const jobId = event?.jobId;
      if (typeof jobId === "string") next.jobId = jobId;
      next.status = isAbortError(event?.error) ? "cancelled" : "error";
      next.lastError = isAbortError(event?.error) ? null : summarizeQueryError(event?.error);
      next.rowsWritten = undefined;
      break;
    }
    case "apply:cancelled": {
      const jobId = event?.jobId;
      if (typeof jobId === "string") next.jobId = jobId;
      next.status = "cancelled";
      next.lastError = null;
      next.rowsWritten = undefined;
      break;
    }
    default:
      return state;
  }

  return { ...state, [queryId]: next };
}

function isQuerySheetDestination(dest: unknown): dest is { sheetId: string; start: { row: number; col: number } } {
  if (!dest || typeof dest !== "object") return false;
  const obj = dest as any;
  if (typeof obj.sheetId !== "string") return false;
  if (!obj.start || typeof obj.start !== "object") return false;
  if (typeof obj.start.row !== "number" || typeof obj.start.col !== "number") return false;
  return true;
}

function sheetDisplayName(sheetId: string, sheetNameResolver?: SheetNameResolver | null): string {
  const id = String(sheetId ?? "").trim();
  if (!id) return "";
  return sheetNameResolver?.getSheetNameById(id) ?? id;
}

function describeDestination(query: Query, sheetNameResolver?: SheetNameResolver | null): string {
  const dest = isQuerySheetDestination(query.destination) ? query.destination : null;
  if (!dest) return "Not loaded";
  const sheetName = sheetDisplayName(dest.sheetId, sheetNameResolver);
  const sheetPrefix = formatSheetNameForA1(sheetName || dest.sheetId);
  const lastOutputSize = (dest as any)?.lastOutputSize;
  const rows = typeof lastOutputSize?.rows === "number" ? lastOutputSize.rows : null;
  const cols = typeof lastOutputSize?.cols === "number" ? lastOutputSize.cols : null;
  if (typeof rows === "number" && typeof cols === "number" && Number.isFinite(rows) && Number.isFinite(cols) && rows > 0 && cols > 0) {
    const start = coordToA1(dest.start.row, dest.start.col);
    const end = coordToA1(dest.start.row + rows - 1, dest.start.col + cols - 1);
    return start === end ? `${sheetPrefix}!${start}` : `${sheetPrefix}!${start}:${end}`;
  }
  return `${sheetPrefix}!${coordToA1(dest.start.row, dest.start.col)}`;
}

function coordToA1(row: number, col: number): string {
  const r = Math.max(0, Math.floor(row));
  let c = Math.max(0, Math.floor(col));
  let letters = "";
  do {
    const rem = c % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    c = Math.floor(c / 26) - 1;
  } while (c >= 0);
  return `${letters}${r + 1}`;
}

type AuthInfo = { required: boolean; label: string | null };

function describeAuth(query: Query): AuthInfo {
  const source = query.source as any;
  if (source?.type === "api" && source?.auth?.type === "oauth2") {
    const providerId = typeof source.auth.providerId === "string" ? source.auth.providerId : "oauth2";
    return { required: true, label: `OAuth2: ${providerId}` };
  }
  if (source?.type === "database") {
    return { required: true, label: "Database" };
  }
  return { required: false, label: null };
}

export type QueryListRow = {
  id: string;
  name: string;
  destination: string;
  lastRefreshAtMs: number | null;
  status: QueryRunStatus;
  errorSummary: string | null;
  rowsWritten?: number;
  authRequired: boolean;
  authLabel: string | null;
};

export function deriveQueryListRows(
  queries: Query[],
  runtime: QueryRuntimeState,
  lastRunAtMsByQueryId: Record<string, number> = {},
  options: { sheetNameResolver?: SheetNameResolver | null } = {},
): QueryListRow[] {
  const sheetNameResolver = options.sheetNameResolver ?? null;
  return queries.map((query) => {
    const run = runtimeFor(runtime, query.id);
    const auth = describeAuth(query);
    const lastRefreshAtMs = run.lastRefreshAtMs ?? lastRunAtMsByQueryId[query.id] ?? null;
    const errorSummary = run.lastError ?? null;
    return {
      id: query.id,
      name: query.name,
      destination: describeDestination(query, sheetNameResolver),
      lastRefreshAtMs,
      status: run.status,
      errorSummary,
      rowsWritten: run.rowsWritten,
      authRequired: auth.required,
      authLabel: auth.label,
    };
  });
}
