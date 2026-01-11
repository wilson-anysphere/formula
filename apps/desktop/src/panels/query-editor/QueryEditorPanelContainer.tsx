import React, { useEffect, useMemo, useState } from "react";

import type { Query } from "../../../../../packages/power-query/src/model.js";
import { QueryEngine } from "../../../../../packages/power-query/src/engine.js";

import { parseA1 } from "../../document/coords.js";

import { applyQueryToDocument, type QuerySheetDestination } from "../../power-query/applyToDocument.js";
import { createDesktopQueryEngine } from "../../power-query/engine.js";
import { DesktopPowerQueryRefreshManager } from "../../power-query/refresh.js";

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

export function QueryEditorPanelContainer(props: Props) {
  const storageId = useMemo(() => storageKey(props.workbookId), [props.workbookId]);

  const [{ engine, engineError }] = useState(() => {
    try {
      return { engine: createDesktopQueryEngine(), engineError: null as string | null };
    } catch (err: any) {
      // Fall back to an in-memory engine in non-Tauri contexts (e.g. browser previews).
      return { engine: new QueryEngine(), engineError: err?.message ?? String(err) };
    }
  });

  const doc = props.getDocumentController();
  const refreshManager = useMemo(() => {
    return new DesktopPowerQueryRefreshManager({
      engine,
      document: doc,
      getContext: () => ({}),
      concurrency: 1,
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

  useEffect(() => {
    refreshManager.registerQuery(query);
    return () => refreshManager.unregisterQuery(query.id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refreshManager, query]);

  useEffect(() => {
    if (typeof localStorage === "undefined") return;
    try {
      localStorage.setItem(storageId, JSON.stringify(query));
    } catch {
      // Ignore storage failures (e.g. quota, disabled storage).
    }
  }, [query, storageId]);

  useEffect(() => {
    return refreshManager.onEvent((evt) => {
      setRefreshEvent(evt);
    });
  }, [refreshManager]);

  useEffect(() => {
    return () => refreshManager.dispose();
  }, [refreshManager]);

  function activeSheetId(): string {
    return props.getActiveSheetId?.() ?? doc?.getSheetIds?.()?.[0] ?? "Sheet1";
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

    try {
      await applyQueryToDocument(doc, current, destination, { engine, batchSize: 1024 });
      setQuery({ ...current, destination });
    } catch (err: any) {
      setActionError(err?.message ?? String(err));
    }
  }

  function refreshNow(queryId: string): void {
    setActionError(null);
    try {
      refreshManager.refresh(queryId);
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

