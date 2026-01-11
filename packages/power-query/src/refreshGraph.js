/**
 * Dependency-aware refresh orchestration ("Refresh All") for Power Query queries.
 *
 * The existing `RefreshManager` is a great primitive for concurrency + per-query
 * cancellation/progress. This module builds on top of it by:
 *   - Extracting a dependency graph across registered queries.
 *   - Refreshing dependencies before dependents (Excel-like semantics).
 *   - Deduping shared dependencies so they execute at most once per session.
 *   - Allowing independent subgraphs to continue even if another branch errors;
 *     downstream dependents of the failed query are cancelled.
 *   - Sharing a `QueryExecutionSession` across all query executions to minimize
 *     repeated credential/permission prompts.
 */

import { RefreshManager } from "./refresh.js";

/**
 * @typedef {import("./model.js").Query} Query
 * @typedef {import("./model.js").QuerySource} QuerySource
 * @typedef {import("./model.js").QueryOperation} QueryOperation
 * @typedef {import("./model.js").QueryStep} QueryStep
 * @typedef {import("./refresh.js").RefreshHandle} RefreshHandle
 * @typedef {import("./engine.js").QueryEngine} QueryEngine
 * @typedef {import("./engine.js").QueryExecutionContext} QueryExecutionContext
 * @typedef {import("./engine.js").QueryExecutionResult} QueryExecutionResult
 */

/**
 * @typedef {"manual" | "interval" | "on-open" | "cron"} RefreshReason
 */

/**
 * @typedef {"dependency" | "target"} RefreshPhase
 */

/**
 * @typedef {{
 *   type: "queued";
 *   sessionId: string;
 *   phase: RefreshPhase;
 *   job: import("./refresh.js").RefreshJobInfo;
 * } | {
 *   type: "started";
 *   sessionId: string;
 *   phase: RefreshPhase;
 *   job: import("./refresh.js").RefreshJobInfo;
 * } | {
 *   type: "progress";
 *   sessionId: string;
 *   phase: RefreshPhase;
 *   job: import("./refresh.js").RefreshJobInfo;
 *   event: import("./engine.js").EngineProgressEvent;
 * } | {
 *   type: "completed";
 *   sessionId: string;
 *   phase: RefreshPhase;
 *   job: import("./refresh.js").RefreshJobInfo;
 *   result: QueryExecutionResult;
 * } | {
 *   type: "error";
 *   sessionId: string;
 *   phase: RefreshPhase;
 *   job: import("./refresh.js").RefreshJobInfo;
 *   error: unknown;
 * } | {
 *   type: "cancelled";
 *   sessionId: string;
 *   phase: RefreshPhase;
 *   job: import("./refresh.js").RefreshJobInfo;
 * }} RefreshGraphEvent
 */

class Emitter {
  constructor() {
    /** @type {Map<string, Set<(payload: any) => void>>} */
    this.listeners = new Map();
  }

  /**
   * @template T
   * @param {string} event
   * @param {(payload: T) => void} handler
   * @returns {() => void}
   */
  on(event, handler) {
    const existing = this.listeners.get(event) ?? new Set();
    existing.add(handler);
    this.listeners.set(event, existing);
    return () => this.off(event, handler);
  }

  /**
   * @param {string} event
   * @param {(payload: any) => void} handler
   */
  off(event, handler) {
    const existing = this.listeners.get(event);
    if (!existing) return;
    existing.delete(handler);
    if (existing.size === 0) this.listeners.delete(event);
  }

  /**
   * @param {string} event
   * @param {any} payload
   */
  emit(event, payload) {
    const existing = this.listeners.get(event);
    if (!existing) return;
    for (const handler of existing) handler(payload);
  }
}

/**
 * @param {string} message
 * @returns {Error}
 */
function abortError(message) {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

/**
 * Find direct dependencies (referenced query IDs) for a query.
 *
 * Dependencies are extracted from:
 *  - source.type === "query"
 *  - merge operations (`rightQuery`)
 *  - append operations (`queries[]`)
 *
 * @param {Query} query
 * @returns {string[]}
 */
export function computeQueryDependencies(query) {
  /** @type {Set<string>} */
  const deps = new Set();

  if (query.source?.type === "query") {
    deps.add(query.source.queryId);
  }

  for (const step of query.steps ?? []) {
    const op = step.operation;
    if (!op || typeof op !== "object") continue;
    if (op.type === "merge") {
      deps.add(op.rightQuery);
    } else if (op.type === "append") {
      for (const id of op.queries ?? []) deps.add(id);
    }
  }

  return Array.from(deps);
}

/**
 * @param {Map<string, string[]>} graph
 * @returns {string[] | null}
 */
function findCycle(graph) {
  /** @type {Map<string, 0 | 1 | 2>} */
  const state = new Map();
  /** @type {string[]} */
  const stack = [];

  for (const id of graph.keys()) state.set(id, 0);

  /**
   * @param {string} id
   * @returns {string[] | null}
   */
  const dfs = (id) => {
    state.set(id, 1);
    stack.push(id);

    for (const dep of graph.get(id) ?? []) {
      if (!graph.has(dep)) continue;
      const depState = state.get(dep) ?? 0;
      if (depState === 0) {
        const cycle = dfs(dep);
        if (cycle) return cycle;
      } else if (depState === 1) {
        const idx = stack.indexOf(dep);
        if (idx >= 0) return [...stack.slice(idx), dep];
        return [dep, id, dep];
      }
    }

    stack.pop();
    state.set(id, 2);
    return null;
  };

  for (const id of graph.keys()) {
    if ((state.get(id) ?? 0) !== 0) continue;
    const cycle = dfs(id);
    if (cycle) return cycle;
  }

  return null;
}

/**
 * @typedef {{
 *   engine: QueryEngine;
 *   getContext?: () => QueryExecutionContext;
 *   concurrency?: number;
 *   now?: () => number;
 * }} RefreshOrchestratorOptions
 */

/**
 * @typedef {{
 *   sessionId: string;
 *   queryIds: string[];
 *   promise: Promise<Record<string, QueryExecutionResult>>;
 *   cancel: () => void;
 *   cancelQuery?: (queryId: string) => void;
 * }} RefreshAllHandle
 */

export class RefreshOrchestrator {
  /**
   * @param {RefreshOrchestratorOptions} options
   */
  constructor(options) {
    this.engine = options.engine;
    this.getContext = options.getContext ?? (() => ({}));
    this.concurrency = Math.max(1, options.concurrency ?? 2);
    this.now = options.now ?? (() => Date.now());

    this.emitter = new Emitter();

    /** @type {Map<string, Query>} */
    this.registrations = new Map();

    this.nextSessionId = 1;
  }

  /**
   * Subscribe to orchestrator events.
   *
   * @param {(event: RefreshGraphEvent) => void} handler
   * @returns {() => void}
   */
  onEvent(handler) {
    return this.emitter.on("event", handler);
  }

  /**
   * Register a query so it can participate in dependency-aware refresh.
   * @param {Query} query
   */
  registerQuery(query) {
    this.registrations.set(query.id, query);
  }

  /**
   * @param {string} queryId
   */
  unregisterQuery(queryId) {
    this.registrations.delete(queryId);
  }

  /**
   * Refresh a set of queries (or all registered queries if none provided).
   *
   * The orchestrator will:
   *  - Compute transitive dependencies.
   *  - Detect graph cycles.
   *  - Execute dependencies before requested targets.
   *  - Share a `QueryExecutionSession` across the whole refresh.
   *
   * @param {string[]} [queryIds]
   * @param {RefreshReason} [reason]
   * @returns {RefreshAllHandle}
   */
  refreshAll(queryIds, reason = "manual") {
    const sessionId = `refresh_all_${this.nextSessionId++}`;

    const targetIds = queryIds ? Array.from(new Set(queryIds)) : Array.from(this.registrations.keys());
    const targetSet = new Set(targetIds);

    if (targetIds.length === 0) {
      return { sessionId, queryIds: targetIds, promise: Promise.resolve({}), cancel: () => {} };
    }

    /**
     * @param {Error} error
     * @param {string} queryId
     * @returns {RefreshAllHandle}
     */
    const errorHandle = (error, queryId) => {
      // Ensure consumers that don't immediately attach a handler don't trigger an
      // unhandled rejection warning, while still exposing the rejection to callers
      // that await/catch `handle.promise`.
      const promise = Promise.reject(error);
      promise.catch(() => {});

      const now = new Date(this.now());
      const job = {
        id: `${sessionId}:graph`,
        queryId,
        reason,
        queuedAt: now,
        completedAt: now,
      };
      this.emitter.emit(
        "event",
        /** @type {RefreshGraphEvent} */ ({
          type: "error",
          sessionId,
          phase: targetSet.has(queryId) ? "target" : "dependency",
          job,
          error,
        }),
      );

      return {
        sessionId,
        queryIds: targetIds,
        promise,
        cancel: () => {},
      };
    };

    /** @type {Map<string, string[]>} */
    const depsById = new Map();
    for (const [id, query] of this.registrations) depsById.set(id, computeQueryDependencies(query));

    /** @type {Set<string>} */
    const closure = new Set();
    /** @type {string[]} */
    const stack = [...targetIds];
    while (stack.length > 0) {
      const id = stack.pop();
      if (!id || closure.has(id)) continue;

      const query = this.registrations.get(id);
      if (!query) {
        return errorHandle(new Error(`Unknown query '${id}'`), id);
      }
      closure.add(id);

      for (const dep of depsById.get(id) ?? []) {
        if (!this.registrations.has(dep)) {
          return errorHandle(new Error(`Unknown query '${dep}' (dependency of '${id}')`), id);
        }
        stack.push(dep);
      }
    }

    /** @type {Map<string, string[]>} */
    const graph = new Map();
    for (const id of closure) {
      graph.set(
        id,
        (depsById.get(id) ?? []).filter((dep) => closure.has(dep)),
      );
    }

    const cycle = findCycle(graph);
    if (cycle) {
      const root = cycle[0] ?? "<unknown>";
      return errorHandle(new Error(`Query dependency cycle detected: ${cycle.join(" -> ")}`), root);
    }

    /** @type {Record<string, QueryExecutionResult>} */
    const queryResults = {};

    const baseContext = this.getContext();
    const registeredQueries = Object.fromEntries(this.registrations.entries());
    const queries = { ...(baseContext.queries ?? {}), ...registeredQueries };

    /** @type {QueryExecutionContext} */
    const context = { ...baseContext, queries, queryResults };

    const engineSession =
      typeof this.engine.createSession === "function"
        ? this.engine.createSession({ now: this.now })
        : { credentialCache: new Map(), permissionCache: new Map(), now: this.now };

    const engine = {
      /**
       * @param {Query} query
       * @param {QueryExecutionContext} ctx
       * @param {import("./engine.js").ExecuteOptions} options
       */
      executeQueryWithMeta: (query, ctx, options) => {
        if (typeof this.engine.executeQueryWithMetaInSession === "function") {
          return this.engine.executeQueryWithMetaInSession(query, ctx, options, engineSession);
        }
        return this.engine.executeQueryWithMeta(query, ctx, options);
      },
    };

    const manager = new RefreshManager({ engine, getContext: () => context, concurrency: this.concurrency, now: this.now });
    for (const id of closure) {
      const query = this.registrations.get(id);
      if (query) manager.registerQuery(query, { type: "manual" });
    }

    /** @type {Map<string, Set<string>>} */
    const dependents = new Map();
    /** @type {Map<string, number>} */
    const remainingDeps = new Map();
    for (const [id, deps] of graph) {
      remainingDeps.set(id, deps.length);
      for (const dep of deps) {
        const existing = dependents.get(dep) ?? new Set();
        existing.add(id);
        dependents.set(dep, existing);
      }
    }

    /** @type {Map<string, RefreshHandle>} */
    const handles = new Map();
    /** @type {Set<string>} */
    const scheduled = new Set();

    /** @type {Record<string, QueryExecutionResult>} */
    const targetResults = {};

    let done = false;
    let cancelled = false;
    let cancelStarted = false;
    /** @type {unknown | null} */
    let terminalError = null;

    /** @type {(value: Record<string, QueryExecutionResult>) => void} */
    let resolve;
    /** @type {(reason?: any) => void} */
    let reject;
    const promise = new Promise((res, rej) => {
      resolve = res;
      reject = rej;
    });

    let remainingJobs = closure.size;
    /** @type {Set<string>} */
    const terminalIds = new Set();
    /** @type {Map<string, number>} */
    const remainingDependents = new Map();
    for (const id of closure) {
      remainingDependents.set(id, dependents.get(id)?.size ?? 0);
    }

    /**
     * Drop a cached dependency result once nothing in the remaining refresh
     * graph can reference it anymore.
     *
     * This helps keep memory bounded when refreshing long dependency chains.
     *
     * @param {string} queryId
     */
    const releaseIfUnused = (queryId) => {
      if (targetSet.has(queryId)) return;
      if ((remainingDependents.get(queryId) ?? 0) !== 0) return;
      // Only affects orchestrator-provided dedupe; safe to delete once all direct
      // dependents have reached a terminal state.
      // @ts-ignore - runtime delete
      delete queryResults[queryId];
    };

    /**
     * @param {string} queryId
     */
    const markTerminal = (queryId) => {
      if (terminalIds.has(queryId)) return;
      terminalIds.add(queryId);
      remainingJobs -= 1;

      for (const dep of graph.get(queryId) ?? []) {
        remainingDependents.set(dep, (remainingDependents.get(dep) ?? 0) - 1);
        releaseIfUnused(dep);
      }
    };

    /**
     * Emit a synthetic cancelled event for a query that will never be scheduled
     * (e.g. because the session was cancelled, or a dependency failed).
     *
     * @param {string} id
     */
    const emitSyntheticCancelled = (id) => {
      const now = new Date(this.now());
      const job = { id: `${sessionId}:cancel_${id}`, queryId: id, reason, queuedAt: now, completedAt: now };
      const phase = targetSet.has(id) ? "target" : "dependency";
      this.emitter.emit(
        "event",
        /** @type {RefreshGraphEvent} */ ({
          type: "cancelled",
          sessionId,
          phase,
          job,
        }),
      );
    };

    const cancelUnscheduledAll = () => {
      for (const id of closure) {
        if (scheduled.has(id)) continue;
        if (terminalIds.has(id)) continue;
        emitSyntheticCancelled(id);
        markTerminal(id);
      }
    };

    /**
     * Cancel the downstream dependents of a failed/cancelled query.
     *
     * This keeps independent subgraphs running while ensuring dependents that can no
     * longer succeed are marked cancelled and never scheduled.
     *
     * @param {string} rootId
     */
    const cancelDependents = (rootId) => {
      /** @type {string[]} */
      const queue = [...(dependents.get(rootId) ?? [])];
      /** @type {Set<string>} */
      const seen = new Set();

      while (queue.length > 0) {
        const id = queue.pop();
        if (!id) continue;
        if (seen.has(id)) continue;
        seen.add(id);

        if (terminalIds.has(id)) continue;

        if (scheduled.has(id)) {
          handles.get(id)?.cancel();
        } else {
          emitSyntheticCancelled(id);
          markTerminal(id);
        }

        for (const next of dependents.get(id) ?? []) queue.push(next);
      }
    };

    /**
     * Cancel the entire refresh session (user cancellation).
     *
     * @param {unknown} error
     */
    const cancelSession = (error) => {
      if (!terminalError) terminalError = error;
      cancelled = true;
      if (cancelStarted) return;
      cancelStarted = true;
      manager.dispose();
      cancelUnscheduledAll();
    };

    const finalize = () => {
      if (done) return;
      if (remainingJobs !== 0) return;
      done = true;
      unsubscribe();
      manager.dispose();
      if (terminalError) {
        reject(terminalError);
      } else {
        resolve(targetResults);
      }
    };

    const schedule = (id) => {
      if (cancelled || done) return;
      if (scheduled.has(id)) return;
      if (terminalIds.has(id)) return;
      scheduled.add(id);

      const handle = manager.refresh(id, reason);
      handles.set(id, handle);
      // The orchestrator is primarily consumed via the aggregate `RefreshAllHandle`.
      // Ensure we don't surface unhandled rejections for individual job promises.
      handle.promise.catch(() => {});
    };

    const forward = (evt) => {
      const phase = targetSet.has(evt.job.queryId) ? "target" : "dependency";
      // Namespace job IDs with the session id so callers can safely treat `job.id`
      // as globally unique across multiple refreshAll sessions.
      const job = { ...evt.job, id: `${sessionId}:${evt.job.id}` };
      this.emitter.emit("event", /** @type {RefreshGraphEvent} */ ({ ...evt, sessionId, phase, job }));
    };

    const unsubscribe = manager.onEvent((evt) => {
      forward(evt);

      if (evt.type === "completed") {
        queryResults[evt.job.queryId] = evt.result;
        if (targetSet.has(evt.job.queryId)) targetResults[evt.job.queryId] = evt.result;
        releaseIfUnused(evt.job.queryId);

        markTerminal(evt.job.queryId);
        for (const dependent of dependents.get(evt.job.queryId) ?? []) {
          remainingDeps.set(dependent, (remainingDeps.get(dependent) ?? 0) - 1);
          if ((remainingDeps.get(dependent) ?? 0) === 0) schedule(dependent);
        }

        finalize();
        return;
      }

      if (evt.type === "error") {
        markTerminal(evt.job.queryId);
        if (!terminalError) terminalError = evt.error;
        cancelDependents(evt.job.queryId);
        finalize();
        return;
      }

      if (evt.type === "cancelled") {
        markTerminal(evt.job.queryId);
        if (!terminalError) terminalError = abortError("Aborted");
        cancelDependents(evt.job.queryId);
        finalize();
      }
    });

    for (const id of closure) {
      if ((remainingDeps.get(id) ?? 0) === 0) schedule(id);
    }

    // If there were no registered queries (or caller passed an empty list) resolve immediately.
    if (closure.size === 0) {
      done = true;
      unsubscribe();
      manager.dispose();
      resolve({});
    }

    return {
      sessionId,
      queryIds: targetIds,
      promise,
      cancel: () => {
        if (done) return;
        cancelSession(abortError("Aborted"));
        finalize();
      },
      cancelQuery: (queryId) => {
        if (done) return;
        if (!closure.has(queryId)) return;
        if (terminalIds.has(queryId)) return;

        const handle = handles.get(queryId);
        if (handle) {
          handle.cancel();
          return;
        }

        // Query hasn't been scheduled yet. Emit a synthetic cancellation and cancel
        // anything downstream that can no longer run.
        if (!terminalError) terminalError = abortError("Aborted");
        emitSyntheticCancelled(queryId);
        markTerminal(queryId);
        cancelDependents(queryId);
        finalize();
      },
    };
  }
}
