import React, { useEffect, useMemo, useRef, useState } from "react";

import { parseCronExpression, type Query } from "@formula/power-query";

import { parseA1 } from "../../document/coords.js";
import * as nativeDialogs from "../../tauri/nativeDialogs.js";
import { getTauriDialogOpenOrNull, hasTauri } from "../../tauri/api";
import { showInputBox } from "../../extensions/ui.js";

import type { QuerySheetDestination } from "../../power-query/applyToDocument.js";
import { getContextForDocument } from "../../power-query/engine.js";
import {
  DesktopPowerQueryService,
  getDesktopPowerQueryService,
  onDesktopPowerQueryServiceChanged,
} from "../../power-query/service.js";

import { PanelIds } from "../panelRegistry.js";

import { QueryEditorPanel } from "./QueryEditorPanel.js";
import { suggestQueryNextSteps } from "./aiSuggestNextSteps.js";

type Props = {
  getDocumentController: () => any;
  getActiveSheetId?: () => string;
  workbookId?: string;
  /**
   * Optional SpreadsheetApp-like object for read-only detection.
   */
  app?: { isReadOnly?: () => boolean } | null;
};

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

async function pickFile(extensions: string[]): Promise<string | null> {
  const open = getTauriDialogOpenOrNull();
  if (open) {
    const result = await open({
      multiple: false,
      filters: [{ name: extensions.join(", ").toUpperCase(), extensions }],
    });
    if (Array.isArray(result)) return result[0] ?? null;
    return result ?? null;
  }

  const path = await showInputBox({ prompt: `Enter path to ${extensions.join("/").toUpperCase()} file`, value: "" });
  return path && path.trim() ? path.trim() : null;
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

type StorageLike = { getItem(key: string): string | null; setItem(key: string, value: string): void };

function getLocalStorageOrNull(): StorageLike | null {
  try {
    const storage = (globalThis as any)?.localStorage as StorageLike | undefined;
    if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") return storage;
  } catch {
    // ignore
  }
  return null;
}

function selectedQueryKey(workbookId: string): string {
  return `formula.desktop.powerQuery.selectedQuery:${workbookId}`;
}

function loadSelectedQueryId(workbookId: string): string | null {
  const storage = getLocalStorageOrNull();
  if (!storage) return null;
  try {
    const raw = storage.getItem(selectedQueryKey(workbookId));
    const trimmed = typeof raw === "string" ? raw.trim() : "";
    return trimmed ? trimmed : null;
  } catch {
    return null;
  }
}

function saveSelectedQueryId(workbookId: string, queryId: string): void {
  const storage = getLocalStorageOrNull();
  if (!storage) return;
  try {
    const trimmed = String(queryId ?? "").trim();
    storage.setItem(selectedQueryKey(workbookId), trimmed);
  } catch {
    // ignore
  }
}

export function QueryEditorPanelContainer(props: Props) {
  const workbookId = props.workbookId ?? "default";
  const doc = props.getDocumentController();
  const app = props.app ?? null;

  const [isReadOnly, setIsReadOnly] = useState<boolean>(() => {
    if (!app || typeof app.isReadOnly !== "function") return false;
    try {
      return Boolean(app.isReadOnly());
    } catch {
      return false;
    }
  });

  const [isEditing, setIsEditing] = useState<boolean>(() => {
    const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
    return globalEditing === true;
  });

  const mutationsDisabled = isReadOnly || isEditing;

  useEffect(() => {
    if (typeof window === "undefined") return;
    const onReadOnlyChanged = (evt: Event) => {
      const detail = (evt as CustomEvent)?.detail as any;
      if (detail && typeof detail.readOnly === "boolean") {
        setIsReadOnly(detail.readOnly);
        return;
      }
      if (!app || typeof app.isReadOnly !== "function") return;
      try {
        setIsReadOnly(Boolean(app.isReadOnly()));
      } catch {
        // ignore
      }
    };
    window.addEventListener("formula:read-only-changed", onReadOnlyChanged as EventListener);
    return () => window.removeEventListener("formula:read-only-changed", onReadOnlyChanged as EventListener);
  }, [app]);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const onEditingChanged = (evt: Event) => {
      const detail = (evt as CustomEvent)?.detail as any;
      if (detail && typeof detail.isEditing === "boolean") {
        setIsEditing(detail.isEditing);
        return;
      }
      const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
      setIsEditing(globalEditing === true);
    };
    window.addEventListener("formula:spreadsheet-editing-changed", onEditingChanged as EventListener);
    return () => window.removeEventListener("formula:spreadsheet-editing-changed", onEditingChanged as EventListener);
  }, []);
  const queryContext = useMemo(() => getContextForDocument(doc), [doc]);

  const [service, setService] = useState<DesktopPowerQueryService | null>(() => getDesktopPowerQueryService(workbookId));
  const serviceRef = useRef<DesktopPowerQueryService | null>(service);
  serviceRef.current = service;
  const [queries, setQueries] = useState<Query[]>(() => service?.getQueries?.() ?? []);
  const [activeQueryId, setActiveQueryId] = useState<string | null>(() => loadSelectedQueryId(workbookId));
  const activeQueryIdRef = useRef<string | null>(activeQueryId);
  activeQueryIdRef.current = activeQueryId;

  useEffect(() => {
    return onDesktopPowerQueryServiceChanged(workbookId, setService);
  }, [workbookId]);

  useEffect(() => {
    if (service) return;
    if (hasTauri()) return;

    const local = new DesktopPowerQueryService({
      workbookId,
      document: doc,
      getContext: () => queryContext,
      concurrency: 1,
      batchSize: 1024,
    });

    setService(local);
    return () => local.dispose();
  }, [doc, queryContext, service, workbookId]);

  const [query, setQuery] = useState<Query>(() => defaultQuery());
  const [refreshEvent, setRefreshEvent] = useState<unknown>(null);
  const [actionError, setActionError] = useState<string | null>(null);
  const [activeLoad, setActiveLoad] = useState<{ jobId: string; cancel: () => void } | null>(null);
  const [activeRefresh, setActiveRefresh] = useState<{ jobId: string; cancel: () => void; applying: boolean } | null>(null);
  const [activeRefreshAll, setActiveRefreshAll] = useState<{ sessionId: string; cancel: () => void } | null>(null);

  const hasSeededDefaultQuery = useRef(false);

  useEffect(() => {
    hasSeededDefaultQuery.current = false;
    setQueries(serviceRef.current?.getQueries?.() ?? []);
    setActiveQueryId(loadSelectedQueryId(workbookId));
    setActiveLoad(null);
    setActiveRefresh(null);
    setActiveRefreshAll(null);
    setActionError(null);
  }, [workbookId]);

  useEffect(() => {
    if (!service) return;
    let cancelled = false;

    void (async () => {
      try {
        await service.ready;
      } catch {
        // ignore
      }
      if (cancelled) return;

      const existing = service.getQueries();
      if (existing.length > 0) {
        setQueries(existing);
        const preferredId = loadSelectedQueryId(workbookId);
        const selected =
          (preferredId && existing.find((q) => q.id === preferredId)) ??
          existing.find((q) => q.id === activeQueryIdRef.current) ??
          existing[0];
        if (selected) {
          setActiveQueryId(selected.id);
          saveSelectedQueryId(workbookId, selected.id);
          setQuery(selected);
        }
        return;
      }

      if (hasSeededDefaultQuery.current) return;

      const seeded = defaultQuery();
      service.setQueries([seeded]);
      hasSeededDefaultQuery.current = true;
      setQueries([seeded]);
      setActiveQueryId(seeded.id);
      saveSelectedQueryId(workbookId, seeded.id);
      setQuery(seeded);
    })();

    return () => {
      cancelled = true;
    };
  }, [service, workbookId]);

  useEffect(() => {
    if (!service) return;

    return service.onEvent((evt) => {
      if (evt?.type === "queries:changed") {
        const queries = Array.isArray((evt as any).queries) ? (evt as any).queries : [];
        setQueries(queries);

        const preferredId = activeQueryIdRef.current ?? loadSelectedQueryId(workbookId);
        const selected = (preferredId && queries.find((q: any) => q?.id === preferredId)) ?? queries[0] ?? null;
        if (selected) {
          setActiveQueryId(selected.id);
          saveSelectedQueryId(workbookId, selected.id);
          setQuery(selected);
        }
        return;
      }

      setRefreshEvent(evt);

      if (evt?.type === "apply:completed" && typeof evt?.queryId === "string") {
        const rows = evt?.result?.rows;
        const cols = evt?.result?.cols;
        if (typeof rows === "number" && typeof cols === "number") {
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

      if (typeof (evt as any)?.sessionId === "string") return;

      const refreshJobId = (evt as any)?.job?.id;
      if (typeof refreshJobId === "string" && ((evt as any).type === "error" || (evt as any).type === "cancelled")) {
        setActiveRefresh((prev) => (prev?.jobId === refreshJobId ? null : prev));
      }

      if (typeof refreshJobId === "string" && (evt as any).type === "completed") {
        queueMicrotask(() => {
          setActiveRefresh((prev) => (prev?.jobId === refreshJobId && !prev.applying ? null : prev));
        });
      }

      const applyJobId = (evt as any)?.jobId;
      if (typeof applyJobId === "string" && (evt as any).type === "apply:started") {
        setActiveRefresh((prev) => (prev?.jobId === applyJobId ? { ...prev, applying: true } : prev));
      }

      if (
        typeof applyJobId === "string" &&
        ((evt as any).type === "apply:completed" || (evt as any).type === "apply:error" || (evt as any).type === "apply:cancelled")
      ) {
        setActiveRefresh((prev) => (prev?.jobId === applyJobId ? null : prev));
      }
    });
  }, [service, workbookId]);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const handler = (evt: Event) => {
      const detail = (evt as any)?.detail;
      if (detail?.panelId !== PanelIds.QUERY_EDITOR) return;
      const requested = detail?.queryId;
      if (typeof requested !== "string" || !requested.trim()) return;
      const queryId = requested.trim();
      setActiveQueryId(queryId);
      saveSelectedQueryId(workbookId, queryId);
      const next = serviceRef.current?.getQuery?.(queryId) ?? queries.find((q) => q.id === queryId);
      if (next) setQuery(next);
    };
    window.addEventListener("formula:open-panel", handler as any);
    return () => window.removeEventListener("formula:open-panel", handler as any);
  }, [queries, workbookId]);

  function persistQuery(next: Query): void {
    setQuery(next);
    service?.registerQuery(next);
  }

  function switchActiveQuery(nextQueryId: string): void {
    const next = service?.getQuery(nextQueryId) ?? queries.find((q) => q.id === nextQueryId);
    if (!next) return;
    setActiveQueryId(nextQueryId);
    saveSelectedQueryId(workbookId, nextQueryId);
    setQuery(next);
  }

  function activeSheetId(): string {
    return props.getActiveSheetId?.() ?? doc?.getSheetIds?.()?.[0] ?? "Sheet1";
  }

  function cancelActiveLoad(): void {
    activeLoad?.cancel();
  }

  async function setCsvSource(): Promise<void> {
    setActionError(null);
    const path = await pickFile(["csv"]);
    if (!path) return;
    persistQuery({ ...query, source: { type: "csv", path, options: { hasHeaders: true } } });
  }

  async function setJsonSource(): Promise<void> {
    setActionError(null);
    const path = await pickFile(["json"]);
    if (!path) return;

    const jsonPath = (await showInputBox({ prompt: "Optional JSON path (e.g. data.items)", value: "" })) ?? "";

    persistQuery({ ...query, source: { type: "json", path, jsonPath: jsonPath.trim() || undefined } });
  }

  async function setParquetSource(): Promise<void> {
    setActionError(null);
    const path = await pickFile(["parquet"]);
    if (!path) return;
    persistQuery({ ...query, source: { type: "parquet", path } });
  }

  async function setWebSource(): Promise<void> {
    setActionError(null);
    const url = (await showInputBox({ prompt: "Enter URL (GET)", value: "https://" }))?.trim();
    if (!url) return;
    persistQuery({ ...query, source: { type: "api", url, method: "GET" } });
  }

  async function setRefreshPolicy(type: string): Promise<void> {
    setActionError(null);

    if (type === "manual") {
      persistQuery({ ...query, refreshPolicy: { type: "manual" } });
      return;
    }

    if (type === "on-open") {
      persistQuery({ ...query, refreshPolicy: { type: "on-open" } });
      return;
    }

    if (type === "interval") {
      const currentMs =
        query.refreshPolicy?.type === "interval" && typeof query.refreshPolicy.intervalMs === "number" ? query.refreshPolicy.intervalMs : 60_000;

      const input = await showInputBox({ prompt: "Refresh interval (milliseconds)", value: String(currentMs) });
      if (input == null) return;

      const ms = Number(input);
      if (!Number.isFinite(ms) || ms <= 0) {
        setActionError("Interval must be a positive number of milliseconds.");
        return;
      }

      persistQuery({ ...query, refreshPolicy: { type: "interval", intervalMs: ms } });
      return;
    }

    if (type === "cron") {
      const currentCron = query.refreshPolicy?.type === "cron" ? query.refreshPolicy.cron : "* * * * *";
      const input = await showInputBox({ prompt: "Cron schedule (minute hour day-of-month month day-of-week)", value: currentCron });
      if (input == null) return;

      const cron = String(input).trim();
      try {
        parseCronExpression(cron);
      } catch (err: any) {
        setActionError(err?.message ?? String(err));
        return;
      }

      persistQuery({ ...query, refreshPolicy: { type: "cron", cron } });
    }
  }

  async function loadToSheet(current: Query): Promise<void> {
    setActionError(null);
    if (mutationsDisabled) return;
    if (!service) return;

    const sheetId = activeSheetId();
    const existingDest = isQuerySheetDestination(current.destination) ? current.destination : null;

    const startText = await showInputBox({ prompt: "Load query to sheet starting cell (A1)", value: "A1" });
    if (startText == null) return;

    let start;
    try {
      start = parseA1(startText);
    } catch (err: any) {
      setActionError(err?.message ?? String(err));
      return;
    }

    const includeHeader = await nativeDialogs.confirm("Include header row?", { fallbackValue: true });

    const clearExisting = await nativeDialogs.confirm("Clear previous output range (if known)?", { fallbackValue: true });

    const destination: QuerySheetDestination = {
      sheetId,
      start,
      includeHeader,
      clearExisting,
      lastOutputSize: existingDest?.lastOutputSize,
    };

    cancelActiveLoad();

    let handle: ReturnType<DesktopPowerQueryService["loadToSheet"]>;
    try {
      handle = service.loadToSheet(current.id, destination, { batchSize: 1024 });
    } catch (err: any) {
      setActionError(err?.message ?? String(err));
      return;
    }

    setActiveLoad({ jobId: handle.id, cancel: handle.cancel });
    try {
      await handle.promise;
      setQuery((prev) => (prev.id === current.id ? { ...prev, destination } : prev));
    } catch (err: any) {
      if (err?.name === "AbortError") return;
      setActionError(err?.message ?? String(err));
    } finally {
      setActiveLoad((prev) => (prev?.jobId === handle.id ? null : prev));
    }
  }

  function refreshNow(queryId: string): void {
    setActionError(null);
    if (mutationsDisabled) return;
    if (!service) return;

    try {
      const handle = service.refresh(queryId);
      setActiveRefresh({ jobId: handle.id, cancel: handle.cancel, applying: false });
    } catch (err: any) {
      setActionError(err?.message ?? String(err));
    }
  }

  function refreshAll(): void {
    setActionError(null);
    if (mutationsDisabled) return;
    if (!service) return;

    try {
      const handle = service.refreshAll();
      setActiveRefreshAll({ sessionId: handle.sessionId, cancel: handle.cancel });
      handle.promise.finally(() => setActiveRefreshAll(null)).catch(() => {});
    } catch (err: any) {
      setActionError(err?.message ?? String(err));
    }
  }

  const engine = service?.engine ?? null;
  const engineError = service?.engineError ?? null;

  if (!engine) {
    return <div className="query-editor-container__message">Power Query service not available.</div>;
  }

  return (
    <div className="query-editor-container">
      {engineError ? (
        <div className="query-editor-container__engine-warning">
          Power Query engine running in fallback mode: {engineError}
        </div>
      ) : null}

      <div className="query-editor-container__toolbar">
        <select
          value={activeQueryId ?? query.id}
          onChange={(e) => switchActiveQuery(e.target.value)}
          className="query-editor-container__select"
        >
          {queries.map((q) => (
            <option key={q.id} value={q.id}>
              {q.name}
            </option>
          ))}
        </select>
        <div className="query-editor-container__source">Source: {describeSource(query.source)}</div>
        <label className="query-editor-container__refresh-policy">
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
        <button type="button" onClick={refreshAll} disabled={mutationsDisabled}>
          Refresh all (graph)
        </button>
        {activeRefreshAll ? (
          <button type="button" onClick={activeRefreshAll.cancel}>
            Cancel refresh all
          </button>
        ) : null}
      </div>

      {actionError ? (
        <div className="query-editor-container__action-error">{actionError}</div>
      ) : null}

      <QueryEditorPanel
        query={query}
        engine={engine}
        context={queryContext}
        refreshEvent={refreshEvent}
        actionsDisabled={mutationsDisabled}
        onQueryChange={(next) => persistQuery(next)}
        onLoadToSheet={loadToSheet}
        onRefreshNow={refreshNow}
        onAiSuggestNextSteps={suggestQueryNextSteps}
      />
    </div>
  );
}
