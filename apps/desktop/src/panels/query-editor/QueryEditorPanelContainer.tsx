import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

import type { Query } from "../../../../../packages/power-query/src/model.js";
import { QueryEngine } from "../../../../../packages/power-query/src/engine.js";
import { parseCronExpression } from "../../../../../packages/power-query/src/cron.js";

import { parseA1 } from "../../document/coords.js";

import { applyQueryToDocument, type QuerySheetDestination } from "../../power-query/applyToDocument.js";
import { maybeGetPowerQueryDlpContext } from "../../power-query/dlpContext.js";
import { createDesktopQueryEngine } from "../../power-query/engine.js";
import { DesktopPowerQueryRefreshManager } from "../../power-query/refresh.js";
import { createPowerQueryRefreshStateStore } from "../../power-query/refreshStateStore.js";
import { DesktopPowerQueryRefreshOrchestrator } from "../../power-query/refreshAll.js";

import { QueryEditorPanel } from "./QueryEditorPanel.js";

type Props = {
  getDocumentController: () => any;
  getActiveSheetId?: () => string;
  workbookId?: string;
};

function safeParseJson(text: string): any | null {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function storageKey(workbookId: string | undefined): string {
  return `formula.desktop.powerQuery.query:${workbookId ?? "default"}`;
}

function isQuerySheetDestination(dest: unknown): dest is QuerySheetDestination {
  if (!dest || typeof dest !== "object") return false;
  const obj = dest as any;
  if (typeof obj.sheetId !== "string") return false;
  if (!obj.start || typeof obj.start !== "object") return false;
  if (typeof obj.start.row !== "number" || typeof obj.start.col !== "number") return false;
  if (typeof obj.includeHeader !== "boolean") return false;
  return true;
}

function defaultQuery(): Query {
  return {
    id: "q1",
    name: "Query 1",
    source: { type: "range", range: { values: [["Value"], [1], [2], [3]], hasHeaders: true } },
    steps: [],
    refreshPolicy: { type: "manual" },
  };
}

type TauriDialogOpen = (options?: Record<string, unknown>) => Promise<string | string[] | null>;

function getTauriDialogOpen(): TauriDialogOpen | null {
  const open = (globalThis as any).__TAURI__?.dialog?.open as TauriDialogOpen | undefined;
  return typeof open === "function" ? open : null;
}

async function pickFile(extensions: string[]): Promise<string | null> {
  const open = getTauriDialogOpen();
  if (open) {
    const result = await open({
      multiple: false,
      filters: [{ name: extensions.join(", ").toUpperCase(), extensions }],
    });
    if (Array.isArray(result)) return result[0] ?? null;
    return result ?? null;
  }

  if (typeof window !== "undefined" && typeof window.prompt === "function") {
    const path = window.prompt(`Enter path to ${extensions.join("/").toUpperCase()} file`, "");
    return path && path.trim() ? path.trim() : null;
  }

  return null;
}

function describeSource(source: any): string {
  if (!source || typeof source !== "object") return "Unknown";
  switch (source.type) {
    case "range":
      return "Range";
    case "csv":
      return `CSV: ${source.path}`;
    case "json":
      return `JSON: ${source.path}${source.jsonPath ? ` (${source.jsonPath})` : ""}`;
    case "parquet":
      return `Parquet: ${source.path}`;
    case "api":
      return `Web: ${source.url}`;
    case "database":
      return "Database";
    case "table":
      return `Table: ${source.table}`;
    case "query":
      return `Query: ${source.queryId}`;
    default:
      return String(source.type ?? "Unknown");
  }
}

export function QueryEditorPanelContainer(props: Props) {
  const storageId = useMemo(() => storageKey(props.workbookId), [props.workbookId]);

  const [{ engine, engineError }] = useState(() => {
    try {
      const dlp = maybeGetPowerQueryDlpContext({ documentId: props.workbookId ?? "default" });
      return { engine: createDesktopQueryEngine({ dlp: dlp ?? undefined }), engineError: null as string | null };
    } catch (err: any) {
      // Fall back to an in-memory engine in non-Tauri contexts (e.g. browser previews).
      return { engine: new QueryEngine(), engineError: err?.message ?? String(err) };
    }
  });

  const doc = props.getDocumentController();

  const refreshStateStore = useMemo(() => {
    return createPowerQueryRefreshStateStore({ workbookId: props.workbookId });
  }, [props.workbookId]);

  const refreshManager = useMemo(() => {
    return new DesktopPowerQueryRefreshManager({
      engine,
      document: doc,
      getContext: () => ({}),
      concurrency: 1,
      batchSize: 1024,
      stateStore: refreshStateStore,
    });
  }, [doc, engine, refreshStateStore]);

  const refreshOrchestrator = useMemo(() => {
    return new DesktopPowerQueryRefreshOrchestrator({
      engine,
      document: doc,
      getContext: () => ({}),
      concurrency: 2,
      batchSize: 1024,
    });
  }, [doc, engine]);

  const [query, setQuery] = useState<Query>(() => {
    if (typeof localStorage === "undefined") return defaultQuery();
    const stored = localStorage.getItem(storageId);
    if (!stored) return defaultQuery();
    const parsed = safeParseJson(stored);
    if (!parsed) return defaultQuery();
    return parsed as Query;
  });

  const [refreshEvent, setRefreshEvent] = useState<unknown>(null);
  const [actionError, setActionError] = useState<string | null>(null);
  const [activeLoad, setActiveLoad] = useState<{ jobId: string; controller: AbortController } | null>(null);
  const [activeRefresh, setActiveRefresh] = useState<{ jobId: string; cancel: () => void; applying: boolean } | null>(null);
  const [activeRefreshAll, setActiveRefreshAll] = useState<{ sessionId: string; cancel: () => void } | null>(null);
  const triggeredOnOpenForQueryId = useRef<string | null>(null);

  useEffect(() => {
    try {
      refreshManager.registerQuery(query);
      // Trigger "on-open" policies once per query id while the panel is mounted.
      if (query.refreshPolicy?.type === "on-open" && triggeredOnOpenForQueryId.current !== query.id) {
        refreshManager.triggerOnOpen(query.id);
        triggeredOnOpenForQueryId.current = query.id;
      }
    } catch (err: any) {
      // Invalid refresh policies (e.g. malformed cron expression) should not crash the panel.
      setActionError(err?.message ?? String(err));
      const fallback: Query = { ...query, refreshPolicy: { type: "manual" } };
      try {
        refreshManager.registerQuery(fallback);
      } catch {
        // ignore
      }
      // Persist the fallback so the user doesn't get stuck in a crash loop.
      setQuery(fallback);
    }

    return () => refreshManager.unregisterQuery(query.id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refreshManager, query]);

  useEffect(() => {
    refreshOrchestrator.registerQuery(query);
    return () => refreshOrchestrator.unregisterQuery(query.id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refreshOrchestrator, query]);

  useEffect(() => {
    if (typeof localStorage === "undefined") return;
    try {
      localStorage.setItem(storageId, JSON.stringify(query));
    } catch {
      // Ignore storage failures (e.g. quota, disabled storage).
    }
  }, [query, storageId]);

  const handleRefreshEvent = useCallback((evt: any): void => {
    setRefreshEvent(evt);

    if (evt?.type === "apply:completed" && typeof evt?.queryId === "string") {
      const rows = evt?.result?.rows;
      const cols = evt?.result?.cols;
      if (typeof rows === "number" && typeof cols === "number") {
        // Persist the last output size in local state so `clearExisting` can clear
        // the prior output range across subsequent query edits / refreshes.
        setQuery((prev) => {
          if (prev.id !== evt.queryId) return prev;
          const dest = isQuerySheetDestination(prev.destination) ? prev.destination : null;
          if (!dest) return prev;
          const existing = dest.lastOutputSize;
          if (existing?.rows === rows && existing?.cols === cols) return prev;
          return { ...prev, destination: { ...dest, lastOutputSize: { rows, cols } } };
        });
      }
    }

    // The per-query refresh manager and dependency-aware refresh orchestrator can
    // emit overlapping job id ranges (`refresh_1`, etc). Avoid letting graph refresh
    // events stomp the per-query refresh UI state by only tracking job ids on the
    // legacy manager (which doesn't include `sessionId`).
    if (typeof evt?.sessionId === "string") return;

    const refreshJobId = evt?.job?.id;
    if (typeof refreshJobId === "string" && (evt.type === "error" || evt.type === "cancelled")) {
      setActiveRefresh((prev) => (prev?.jobId === refreshJobId ? null : prev));
    }

    if (typeof refreshJobId === "string" && evt.type === "completed") {
      // If the refresh completes but no "apply" phase starts (e.g. no destination),
      // clear the active refresh state on the next tick.
      queueMicrotask(() => {
        setActiveRefresh((prev) => (prev?.jobId === refreshJobId && !prev.applying ? null : prev));
      });
    }

    const applyJobId = evt?.jobId;
    if (typeof applyJobId === "string" && evt.type === "apply:started") {
      setActiveRefresh((prev) => (prev?.jobId === applyJobId ? { ...prev, applying: true } : prev));
    }

    if (
      typeof applyJobId === "string" &&
      (evt.type === "apply:completed" || evt.type === "apply:error" || evt.type === "apply:cancelled")
    ) {
      setActiveRefresh((prev) => (prev?.jobId === applyJobId ? null : prev));
    }
  }, []);

  useEffect(() => {
    return refreshManager.onEvent(handleRefreshEvent);
  }, [handleRefreshEvent, refreshManager]);

  useEffect(() => {
    return refreshOrchestrator.onEvent(handleRefreshEvent);
  }, [handleRefreshEvent, refreshOrchestrator]);

  useEffect(() => {
    return () => refreshManager.dispose();
  }, [refreshManager]);

  useEffect(() => {
    return () => refreshOrchestrator.dispose();
  }, [refreshOrchestrator]);

  function activeSheetId(): string {
    return props.getActiveSheetId?.() ?? doc?.getSheetIds?.()?.[0] ?? "Sheet1";
  }

  function cancelActiveLoad(): void {
    activeLoad?.controller.abort();
  }

  async function setCsvSource(): Promise<void> {
    setActionError(null);
    const path = await pickFile(["csv"]);
    if (!path) return;
    setQuery((prev) => ({ ...prev, source: { type: "csv", path, options: { hasHeaders: true } } }));
  }

  async function setJsonSource(): Promise<void> {
    setActionError(null);
    const path = await pickFile(["json"]);
    if (!path) return;

    const jsonPath =
      typeof window !== "undefined" && typeof window.prompt === "function"
        ? window.prompt("Optional JSON path (e.g. data.items)", "") ?? ""
        : "";

    setQuery((prev) => ({ ...prev, source: { type: "json", path, jsonPath: jsonPath.trim() || undefined } }));
  }

  async function setParquetSource(): Promise<void> {
    setActionError(null);
    const path = await pickFile(["parquet"]);
    if (!path) return;
    setQuery((prev) => ({ ...prev, source: { type: "parquet", path } }));
  }

  async function setWebSource(): Promise<void> {
    setActionError(null);
    const url =
      typeof window !== "undefined" && typeof window.prompt === "function"
        ? window.prompt("Enter URL (GET)", "https://")?.trim()
        : null;
    if (!url) return;
    setQuery((prev) => ({ ...prev, source: { type: "api", url, method: "GET" } }));
  }

  function updateRefreshPolicy(next: any): void {
    setQuery((prev) => ({ ...prev, refreshPolicy: next }));
  }

  async function setRefreshPolicy(type: string): Promise<void> {
    setActionError(null);
    if (type === "manual") {
      updateRefreshPolicy({ type: "manual" });
      return;
    }
    if (type === "on-open") {
      updateRefreshPolicy({ type: "on-open" });
      return;
    }
    if (type === "interval") {
      const currentMs =
        query.refreshPolicy?.type === "interval" && typeof query.refreshPolicy.intervalMs === "number"
          ? query.refreshPolicy.intervalMs
          : 60_000;
      const input =
        typeof window !== "undefined" && typeof window.prompt === "function"
          ? window.prompt("Refresh interval (milliseconds)", String(currentMs))
          : String(currentMs);
      if (input == null) return;
      const ms = Number(input);
      if (!Number.isFinite(ms) || ms <= 0) {
        setActionError("Interval must be a positive number of milliseconds.");
        return;
      }
      updateRefreshPolicy({ type: "interval", intervalMs: ms });
      return;
    }
    if (type === "cron") {
      const currentCron = query.refreshPolicy?.type === "cron" ? query.refreshPolicy.cron : "* * * * *";
      const input =
        typeof window !== "undefined" && typeof window.prompt === "function"
          ? window.prompt("Cron schedule (minute hour day-of-month month day-of-week)", currentCron)
          : currentCron;
      if (input == null) return;
      const cron = String(input).trim();
      try {
        parseCronExpression(cron);
      } catch (err: any) {
        setActionError(err?.message ?? String(err));
        return;
      }
      updateRefreshPolicy({ type: "cron", cron });
    }
  }

  async function loadToSheet(current: Query): Promise<void> {
    setActionError(null);

    const sheetId = activeSheetId();
    const existingDest = isQuerySheetDestination(current.destination) ? current.destination : null;

    const startText =
      typeof window !== "undefined" && typeof window.prompt === "function"
        ? window.prompt("Load query to sheet starting cell (A1)", "A1")
        : "A1";
    if (startText == null) return;

    let start;
    try {
      start = parseA1(startText);
    } catch (err: any) {
      setActionError(err?.message ?? String(err));
      return;
    }

    const includeHeader =
      typeof window !== "undefined" && typeof window.confirm === "function"
        ? window.confirm("Include header row?")
        : true;

    const clearExisting =
      typeof window !== "undefined" && typeof window.confirm === "function"
        ? window.confirm("Clear previous output range (if known)?")
        : true;

    const destination: QuerySheetDestination = {
      sheetId,
      start,
      includeHeader,
      clearExisting,
      lastOutputSize: existingDest?.lastOutputSize,
    };

    // Cancel any existing "load to sheet" operation so we don't interleave writes.
    cancelActiveLoad();
    const controller = new AbortController();
    const jobId = `load_${crypto.randomUUID()}`;
    setActiveLoad({ jobId, controller });

    try {
      setRefreshEvent({ type: "apply:started", jobId, queryId: current.id, destination });
      const result = await applyQueryToDocument(doc, current, destination, {
        engine,
        batchSize: 1024,
        signal: controller.signal,
        onProgress: (evt) => {
          if (evt.type === "batch") {
            setRefreshEvent({ type: "apply:progress", jobId, queryId: current.id, rowsWritten: evt.totalRowsWritten });
          }
        },
      });
      setRefreshEvent({ type: "apply:completed", jobId, queryId: current.id, result });
      setQuery({ ...current, destination });
    } catch (err: any) {
      if (controller.signal.aborted || err?.name === "AbortError") {
        setRefreshEvent({ type: "apply:cancelled", jobId, queryId: current.id });
        return;
      }
      setRefreshEvent({ type: "apply:error", jobId, queryId: current.id, error: err });
      setActionError(err?.message ?? String(err));
    } finally {
      setActiveLoad((prev) => (prev?.jobId === jobId ? null : prev));
    }
  }

  function refreshNow(queryId: string): void {
    setActionError(null);
    try {
      const handle = refreshManager.refresh(queryId);
      setActiveRefresh({ jobId: handle.id, cancel: handle.cancel, applying: false });
    } catch (err: any) {
      setActionError(err?.message ?? String(err));
    }
  }

  function refreshAll(): void {
    setActionError(null);
    try {
      const handle = refreshOrchestrator.refreshAll();
      setActiveRefreshAll({ sessionId: handle.sessionId, cancel: handle.cancel });
      handle.promise.finally(() => setActiveRefreshAll(null)).catch(() => {});
    } catch (err: any) {
      setActionError(err?.message ?? String(err));
    }
  }

  return (
    <div style={{ flex: 1, minHeight: 0 }}>
      {engineError ? (
        <div style={{ padding: 12, color: "var(--text-muted)", borderBottom: "1px solid var(--border)" }}>
          Power Query engine running in fallback mode: {engineError}
        </div>
      ) : null}

      <div style={{ padding: 12, borderBottom: "1px solid var(--border)", display: "flex", flexWrap: "wrap", gap: 8 }}>
        <div style={{ color: "var(--text-muted)", fontSize: 12, marginRight: 8 }}>Source: {describeSource(query.source)}</div>
        <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12, color: "var(--text-muted)" }}>
          Refresh:
          <select
            value={query.refreshPolicy?.type ?? "manual"}
            onChange={(e) => {
              void setRefreshPolicy(e.target.value);
            }}
          >
            <option value="manual">Manual</option>
            <option value="on-open">On open</option>
            <option value="interval">Interval</option>
            <option value="cron">Cron</option>
          </select>
        </label>
        <button type="button" onClick={setCsvSource}>
          CSV…
        </button>
        <button type="button" onClick={setJsonSource}>
          JSON…
        </button>
        <button type="button" onClick={setParquetSource}>
          Parquet…
        </button>
        <button type="button" onClick={setWebSource}>
          Web…
        </button>
        {activeLoad ? (
          <button type="button" onClick={cancelActiveLoad}>
            Cancel load
          </button>
        ) : null}
        {activeRefresh ? (
          <button type="button" onClick={activeRefresh.cancel}>
            Cancel refresh
          </button>
        ) : null}
        <button type="button" onClick={refreshAll}>
          Refresh all (graph)
        </button>
        {activeRefreshAll ? (
          <button type="button" onClick={activeRefreshAll.cancel}>
            Cancel refresh all
          </button>
        ) : null}
      </div>

      {actionError ? (
        <div style={{ padding: 12, color: "var(--error)", borderBottom: "1px solid var(--border)" }}>{actionError}</div>
      ) : null}

      <QueryEditorPanel
        query={query}
        engine={engine}
        context={{}}
        refreshEvent={refreshEvent}
        onQueryChange={(next) => setQuery(next)}
        onLoadToSheet={loadToSheet}
        onRefreshNow={refreshNow}
      />
    </div>
  );
}
