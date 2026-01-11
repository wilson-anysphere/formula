import type { Query, RefreshPolicy } from "../../../../packages/power-query/src/model.js";
import type { QueryExecutionContext, QueryEngine } from "../../../../packages/power-query/src/engine.js";
import { RefreshManager } from "../../../../packages/power-query/src/refresh.js";
import { RefreshOrchestrator } from "../../../../packages/power-query/src/refreshGraph.js";

import type { DocumentController } from "../document/documentController.js";

// Use `.ts` extension so Node's `--experimental-strip-types` test runner can resolve
// the module without relying on bundler-specific `.js`→`.ts` mapping.
import { applyTableToDocument, type ApplyToDocumentResult, type QuerySheetDestination } from "./applyToDocument.ts";

// `packages/power-query` is authored in JS; in the desktop layer we treat refresh
// events as an opaque payload and primarily use their `type` + `job` fields.
type RefreshEvent = any;

export type DesktopPowerQueryEvent =
  | RefreshEvent
  | { type: "apply:started"; jobId: string; queryId: string; destination: QuerySheetDestination }
  | { type: "apply:progress"; jobId: string; queryId: string; rowsWritten: number }
  | { type: "apply:completed"; jobId: string; queryId: string; result: ApplyToDocumentResult }
  | { type: "apply:error"; jobId: string; queryId: string; error: unknown }
  | { type: "apply:cancelled"; jobId: string; queryId: string };

class Emitter<T> {
  listeners: Set<(payload: T) => void> = new Set();

  on(handler: (payload: T) => void): () => void {
    this.listeners.add(handler);
    return () => this.listeners.delete(handler);
  }

  emit(payload: T): void {
    for (const handler of this.listeners) handler(payload);
  }
}

function isAbortError(error: unknown): boolean {
  return (error as any)?.name === "AbortError";
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

class RefreshingEngine {
  engine: QueryEngine;

  constructor(engine: QueryEngine) {
    this.engine = engine;
  }

  createSession(options?: { now?: () => number }) {
    // Backwards compatibility: `QueryEngine.createSession` was introduced for shared
    // refresh sessions. Fall back to local caches when an older engine is used.
    // (The desktop app normally uses a modern `QueryEngine`.)
    return typeof (this.engine as any).createSession === "function"
      ? (this.engine as any).createSession(options)
      : { credentialCache: new Map(), permissionCache: new Map(), now: options?.now };
  }

  executeQueryWithMetaInSession(query: Query, context: QueryExecutionContext, options: any, session: any) {
    const nextOptions = {
      ...(options ?? {}),
      cache: { ...(options?.cache ?? {}), mode: "refresh" as const },
    };

    if (typeof (this.engine as any).executeQueryWithMetaInSession === "function") {
      return (this.engine as any).executeQueryWithMetaInSession(query, context, nextOptions, session);
    }
    return this.engine.executeQueryWithMeta(query, context, nextOptions);
  }

  executeQueryWithMeta(query: Query, context: QueryExecutionContext, options: any) {
    const nextOptions = {
      ...(options ?? {}),
      cache: { ...(options?.cache ?? {}), mode: "refresh" as const },
    };
    return this.engine.executeQueryWithMeta(query, context, nextOptions);
  }
}

export type DesktopPowerQueryRefreshOptions = {
  engine: QueryEngine;
  document: DocumentController;
  getContext?: () => QueryExecutionContext;
  concurrency?: number;
  batchSize?: number;
  stateStore?: { load(): Promise<any>; save(state: any): Promise<void> };
};

/**
 * Desktop wrapper around `RefreshManager` that applies refreshed query outputs into the sheet.
 */
export class DesktopPowerQueryRefreshManager {
  doc: DocumentController;
  batchSize: number;
  emitter = new Emitter<DesktopPowerQueryEvent>();
  queries = new Map<string, Query>();
  applyControllers = new Map<string, AbortController>();
  activeRefreshAll = new Set<{ cancel: () => void; promise: Promise<any> }>();

  manager: RefreshManager;
  orchestrator: RefreshOrchestrator;

  constructor(options: DesktopPowerQueryRefreshOptions) {
    this.doc = options.document;
    this.batchSize = options.batchSize ?? 1024;

    // Force refreshes to bypass/overwrite cache entries, but still allow the engine
    // to use deterministic cache keys for subsequent "load to sheet" operations.
    const engine = new RefreshingEngine(options.engine);
    this.manager = new RefreshManager({
      engine: engine as any,
      getContext: options.getContext,
      concurrency: options.concurrency,
      stateStore: options.stateStore as any,
    });

    this.orchestrator = new RefreshOrchestrator({
      engine: engine as any,
      getContext: options.getContext,
      concurrency: options.concurrency,
    });

    this.manager.onEvent((evt: any) => {
      this.emitter.emit(evt);
      if (evt?.type === "completed") {
        void this.applyCompletedJob(evt);
      }
    });

    this.orchestrator.onEvent((evt: any) => {
      this.emitter.emit(evt);
      if (evt?.type === "completed") {
        // Best-effort: keep the RefreshManager state store's last-run timestamps in sync
        // with dependency-aware refreshAll sessions so interval/cron policies can restore
        // accurately on the next app launch.
        const queryId = evt?.job?.queryId;
        const completedAt = evt?.job?.completedAt;
        if (typeof queryId === "string" && completedAt instanceof Date && !Number.isNaN(completedAt.getTime())) {
          // `recordSuccessfulRun` is intentionally internal to RefreshManager, but using
          // it here avoids duplicating the state-store persistence logic.
          (this.manager as any).recordSuccessfulRun?.(queryId, completedAt.getTime());
        }
        void this.applyCompletedJob(evt);
      }
    });
  }

  onEvent(handler: (event: DesktopPowerQueryEvent) => void): () => void {
    return this.emitter.on(handler);
  }

  registerQuery(query: Query, policy: RefreshPolicy = query.refreshPolicy ?? { type: "manual" }) {
    this.queries.set(query.id, query);
    this.manager.registerQuery(query, policy);
    this.orchestrator.registerQuery(query);
  }

  unregisterQuery(queryId: string) {
    this.queries.delete(queryId);
    this.manager.unregisterQuery(queryId);
    this.orchestrator.unregisterQuery(queryId);
  }

  triggerOnOpen(queryId?: string) {
    this.manager.triggerOnOpen(queryId);
  }

  refresh(queryId: string, reason: any = "manual") {
    const handle = this.manager.refresh(queryId, reason);
    return {
      ...handle,
      cancel: () => {
        handle.cancel();
        this.applyControllers.get(handle.id)?.abort();
      },
    };
  }

  refreshAll(queryIds?: string[], reason: any = "manual") {
    const handle = this.orchestrator.refreshAll(queryIds, reason);
    const sessionPrefix = `${handle.sessionId}:`;

    this.activeRefreshAll.add(handle);
    handle.promise.finally(() => this.activeRefreshAll.delete(handle)).catch(() => {});

    return {
      ...handle,
      cancel: () => {
        handle.cancel();
        for (const [jobId, controller] of this.applyControllers) {
          if (jobId.startsWith(sessionPrefix)) controller.abort();
        }
      },
    };
  }

  dispose() {
    for (const controller of this.applyControllers.values()) controller.abort();
    this.applyControllers.clear();
    for (const handle of this.activeRefreshAll) handle.cancel();
    this.activeRefreshAll.clear();
    this.manager.dispose();
  }

  async applyCompletedJob(evt: any): Promise<void> {
    const jobId = evt?.job?.id;
    const queryId = evt?.job?.queryId;
    if (typeof jobId !== "string" || typeof queryId !== "string") return;

    const query = this.queries.get(queryId);
    if (!query) return;
    const destination = query.destination;
    if (!isQuerySheetDestination(destination)) return;

    const table = evt?.result?.table;
    if (!table) return;

    const controller = new AbortController();
    this.applyControllers.set(jobId, controller);

    this.emitter.emit({ type: "apply:started", jobId, queryId, destination });

    try {
      const result = await applyTableToDocument(this.doc, table, destination, {
        batchSize: this.batchSize,
        signal: controller.signal,
        label: `Refresh query: ${query.name}`,
        queryId,
        onProgress: async (progress) => {
          if (progress.type === "batch") {
            this.emitter.emit({
              type: "apply:progress",
              jobId,
              queryId,
              rowsWritten: progress.totalRowsWritten,
            });
          }
        },
      });
      this.emitter.emit({ type: "apply:completed", jobId, queryId, result });
    } catch (error) {
      if (controller.signal.aborted || isAbortError(error)) {
        this.emitter.emit({ type: "apply:cancelled", jobId, queryId });
      } else {
        this.emitter.emit({ type: "apply:error", jobId, queryId, error });
      }
    } finally {
      this.applyControllers.delete(jobId);
    }
  }
}
