import type { Query } from "../../../../packages/power-query/src/model.js";
import type { QueryExecutionContext, QueryEngine } from "../../../../packages/power-query/src/engine.js";
import { RefreshOrchestrator } from "../../../../packages/power-query/src/refreshGraph.js";

import type { DocumentController } from "../document/documentController.js";

// Use `.ts` extension so Node's `--experimental-strip-types` test runner can resolve
// the module without relying on bundler-specific `.js`→`.ts` mapping.
import { applyTableToDocument, type QuerySheetDestination } from "./applyToDocument.ts";
import { enqueueApplyForDocument } from "./applyQueue.ts";
import type { DesktopPowerQueryEvent } from "./refresh.ts";

// `packages/power-query` is authored in JS; in the desktop layer we treat refresh graph
// events as an opaque payload and primarily use their `type` + `job` fields.
type RefreshGraphEvent = any;

export type DesktopPowerQueryRefreshAllOptions = {
  engine: QueryEngine;
  document: DocumentController;
  /**
   * Base context for query execution.
   *
   * Note: the core RefreshOrchestrator merges `context.queries` with registered queries
   * so query references resolve during refresh.
   */
  getContext?: () => QueryExecutionContext;
  concurrency?: number;
  /** Batch size for sheet writes. */
  batchSize?: number;
  /**
   * Optional callback invoked when a query successfully refreshes.
   *
   * The query editor uses this to keep scheduled refresh persistence (`lastRunAtMs`)
   * in sync with dependency-aware refreshAll sessions.
   */
  onSuccessfulRun?: (queryId: string, completedAtMs: number) => void;
};

export type DesktopPowerQueryRefreshAllHandle = {
  sessionId: string;
  queryIds: string[];
  // Matches the core `RefreshOrchestrator` API shape: resolves with results for the
  // requested target query ids (not necessarily including dependencies).
  promise: Promise<Record<string, any>>;
  cancel: () => void;
  cancelQuery?: (queryId: string) => void;
};

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

  createSession(options: any = {}) {
    const fn = (this.engine as any).createSession;
    if (typeof fn === "function") return fn.call(this.engine, options);
    return { credentialCache: new Map(), permissionCache: new Map(), now: options?.now };
  }

  executeQueryWithMeta(query: Query, context: QueryExecutionContext, options: any) {
    const nextOptions = {
      ...(options ?? {}),
      cache: { ...(options?.cache ?? {}), mode: "refresh" as const },
    };
    return (this.engine as any).executeQueryWithMeta(query, context, nextOptions);
  }

  executeQueryWithMetaInSession(query: Query, context: QueryExecutionContext, options: any, session: any) {
    const nextOptions = {
      ...(options ?? {}),
      cache: { ...(options?.cache ?? {}), mode: "refresh" as const },
    };
    const fn = (this.engine as any).executeQueryWithMetaInSession;
    if (typeof fn === "function") return fn.call(this.engine, query, context, nextOptions, session);
    return this.executeQueryWithMeta(query, context, nextOptions);
  }
}

type ApplyControllerEntry = { sessionId: string; jobId: string; queryId: string; controller: AbortController };

/**
 * Desktop wrapper around the core dependency-aware `RefreshOrchestrator`.
 *
 * - Forwards all refresh graph events
 * - Applies completed query results to sheet destinations (when present)
 * - Cancels in-flight applies when the refresh session is cancelled
 */
export class DesktopPowerQueryRefreshOrchestrator {
  doc: DocumentController;
  batchSize: number;
  private readonly onSuccessfulRun: DesktopPowerQueryRefreshAllOptions["onSuccessfulRun"];
  emitter = new Emitter<DesktopPowerQueryEvent>();
  queries = new Map<string, Query>();
  applyControllers = new Map<string, ApplyControllerEntry>();
  cancelledSessions = new Set<string>();
  cancelledQueriesBySession = new Map<string, Set<string>>();
  activeSessions = new Map<string, () => void>();

  orchestrator: RefreshOrchestrator;

  constructor(options: DesktopPowerQueryRefreshAllOptions) {
    this.doc = options.document;
    this.batchSize = options.batchSize ?? 1024;
    this.onSuccessfulRun = options.onSuccessfulRun;

    // Force refreshes to bypass/overwrite cache entries, but still allow the engine
    // to use deterministic cache keys for subsequent "load to sheet" operations.
    const engine = new RefreshingEngine(options.engine);

    this.orchestrator = new RefreshOrchestrator({
      engine: engine as any,
      getContext: options.getContext,
      concurrency: options.concurrency,
    });

    this.orchestrator.onEvent((evt: RefreshGraphEvent) => {
      this.emitter.emit(evt);
      if (evt?.type === "completed") {
        const queryId = evt?.job?.queryId;
        const completedAt = evt?.job?.completedAt;
        if (this.onSuccessfulRun && typeof queryId === "string" && completedAt instanceof Date) {
          const completedAtMs = completedAt.getTime();
          if (!Number.isNaN(completedAtMs)) {
            try {
              this.onSuccessfulRun(queryId, completedAtMs);
            } catch {
              // Best-effort: scheduled refresh persistence should never break refreshAll.
            }
          }
        }
        void this.applyCompletedJob(evt);
      } else if (evt?.type === "cancelled") {
        this.emitApplyCancelled(evt);
      }
    });
  }

  onEvent(handler: (event: DesktopPowerQueryEvent) => void): () => void {
    return this.emitter.on(handler);
  }

  registerQuery(query: Query): void {
    this.queries.set(query.id, query);
    this.orchestrator.registerQuery(query);
  }

  unregisterQuery(queryId: string): void {
    this.queries.delete(queryId);
    this.orchestrator.unregisterQuery(queryId);
  }

  refreshAll(queryIds?: string[], reason: any = "manual"): DesktopPowerQueryRefreshAllHandle {
    const handle = this.orchestrator.refreshAll(queryIds, reason);
    const sessionId = handle.sessionId;

    const cancelledQueries = new Set<string>();
    this.cancelledQueriesBySession.set(sessionId, cancelledQueries);

    const cancel = () => {
      this.cancelledSessions.add(sessionId);
      handle.cancel();
      for (const entry of this.applyControllers.values()) {
        if (entry.sessionId === sessionId) entry.controller.abort();
      }
    };

    const cancelQuery = (queryId: string) => {
      (handle as any).cancelQuery?.(queryId);
      cancelledQueries.add(queryId);
      for (const entry of this.applyControllers.values()) {
        if (entry.sessionId !== sessionId) continue;
        if (entry.queryId !== queryId) continue;
        entry.controller.abort();
      }
    };

    this.activeSessions.set(sessionId, cancel);
    handle.promise
      .finally(() => {
        this.activeSessions.delete(sessionId);
        this.cancelledSessions.delete(sessionId);
        this.cancelledQueriesBySession.delete(sessionId);
      })
      .catch(() => {});

    return {
      ...handle,
      cancel,
      cancelQuery,
    };
  }

  dispose(): void {
    for (const cancel of this.activeSessions.values()) cancel();
    this.activeSessions.clear();
    for (const entry of this.applyControllers.values()) entry.controller.abort();
    this.applyControllers.clear();
    this.cancelledSessions.clear();
    this.cancelledQueriesBySession.clear();
  }

  emitApplyCancelled(evt: any): void {
    const sessionId = evt?.sessionId;
    const jobId = evt?.job?.id;
    const queryId = evt?.job?.queryId;
    if (typeof sessionId !== "string" || typeof jobId !== "string" || typeof queryId !== "string") return;

    const query = this.queries.get(queryId);
    if (!query) return;
    const destination = query.destination;
    if (!isQuerySheetDestination(destination)) return;

    this.emitter.emit({ type: "apply:cancelled", jobId, queryId, sessionId });
  }

  async applyCompletedJob(evt: any): Promise<void> {
    const sessionId = evt?.sessionId;
    const jobId = evt?.job?.id;
    const queryId = evt?.job?.queryId;
    if (typeof sessionId !== "string" || typeof jobId !== "string" || typeof queryId !== "string") return;

    const query = this.queries.get(queryId);
    if (!query) return;
    const destination = query.destination;
    if (!isQuerySheetDestination(destination)) return;

    const table = evt?.result?.table;
    if (!table) return;

    const cancelledQueries = this.cancelledQueriesBySession.get(sessionId);
    if (this.cancelledSessions.has(sessionId) || cancelledQueries?.has(queryId)) {
      this.emitter.emit({ type: "apply:cancelled", jobId, queryId, sessionId });
      return;
    }

    const controller = new AbortController();
    // Core `RefreshOrchestrator` namespaces `job.id` with the session id, but older
    // builds may not. Ensure uniqueness either way.
    const applyKey = jobId.startsWith(`${sessionId}:`) ? jobId : `${sessionId}:${jobId}`;
    this.applyControllers.set(applyKey, { sessionId, jobId, queryId, controller });

    this.emitter.emit({ type: "apply:started", jobId, queryId, destination, sessionId });

    // Serialize apply operations. The DocumentController batching model is global
    // (single `activeBatch`), so overlapping apply operations can corrupt undo
    // grouping and prevent cancellation from reverting partial writes.
    enqueueApplyForDocument(this.doc, async () => {
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
                sessionId,
              });
            }
          },
        });
        this.emitter.emit({ type: "apply:completed", jobId, queryId, result, sessionId });
      } catch (error) {
        if (controller.signal.aborted || isAbortError(error)) {
          this.emitter.emit({ type: "apply:cancelled", jobId, queryId, sessionId });
        } else {
          this.emitter.emit({ type: "apply:error", jobId, queryId, error, sessionId });
        }
      } finally {
        this.applyControllers.delete(applyKey);
      }
    });
  }
}
