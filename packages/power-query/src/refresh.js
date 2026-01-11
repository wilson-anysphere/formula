/**
 * Refresh policy + scheduling for queries.
 *
 * `RefreshManager` is designed to be UI-agnostic. Host applications can wire
 * events into whatever UX they want (progress bars, notifications, prompts).
 */

/**
 * @typedef {import("./model.js").Query} Query
 * @typedef {import("./model.js").RefreshPolicy} RefreshPolicy
 * @typedef {import("./engine.js").QueryEngine} QueryEngine
 * @typedef {import("./engine.js").QueryExecutionContext} QueryExecutionContext
 * @typedef {import("./engine.js").EngineProgressEvent} EngineProgressEvent
 * @typedef {import("./engine.js").QueryExecutionResult} QueryExecutionResult
 */

/**
 * @typedef {"manual" | "interval" | "on-open"} RefreshReason
 */

/**
 * @typedef {{
 *   id: string;
 *   queryId: string;
 *   reason: RefreshReason;
 *   queuedAt: Date;
 *   startedAt?: Date;
 *   completedAt?: Date;
 * }} RefreshJobInfo
 */

/**
 * @typedef {{
 *   type: "queued";
 *   job: RefreshJobInfo;
 * } | {
 *   type: "started";
 *   job: RefreshJobInfo;
 * } | {
 *   type: "progress";
 *   job: RefreshJobInfo;
 *   event: EngineProgressEvent;
 * } | {
 *   type: "completed";
 *   job: RefreshJobInfo;
 *   result: QueryExecutionResult;
 * } | {
 *   type: "error";
 *   job: RefreshJobInfo;
 *   error: unknown;
 * } | {
 *   type: "cancelled";
 *   job: RefreshJobInfo;
 * }} RefreshEvent
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
    for (const handler of existing) {
      handler(payload);
    }
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
 * @typedef {{
 *   engine: QueryEngine;
 *   getContext?: () => QueryExecutionContext;
 *   concurrency?: number;
 *   timers?: { setTimeout: typeof setTimeout; clearTimeout: typeof clearTimeout };
 *   now?: () => number;
 * }} RefreshManagerOptions
 */

export class RefreshManager {
  /**
   * @param {RefreshManagerOptions} options
   */
  constructor(options) {
    this.engine = options.engine;
    this.getContext = options.getContext ?? (() => ({}));
    this.concurrency = Math.max(1, options.concurrency ?? 2);
    this.timers = options.timers ?? { setTimeout: globalThis.setTimeout, clearTimeout: globalThis.clearTimeout };
    this.now = options.now ?? (() => Date.now());

    this.emitter = new Emitter();

    /** @type {Map<string, { query: Query, policy: RefreshPolicy, timer: any }>} */
    this.registrations = new Map();
    /** @type {RefreshJob[]} */
    this.queue = [];
    /** @type {Map<string, RefreshJob>} */
    this.running = new Map();

    this.nextJobId = 1;
  }

  /**
   * Subscribe to manager events.
   * @param {(event: RefreshEvent) => void} handler
   * @returns {() => void}
   */
  onEvent(handler) {
    return this.emitter.on("event", handler);
  }

  /**
   * @param {Query} query
   * @param {RefreshPolicy} [policy]
   */
  registerQuery(query, policy = query.refreshPolicy ?? { type: "manual" }) {
    this.unregisterQuery(query.id);

    const timer = null;
    this.registrations.set(query.id, { query, policy, timer });

    if (policy.type === "interval") {
      this.scheduleInterval(query.id, policy.intervalMs);
    }
  }

  /**
   * @param {string} queryId
   */
  unregisterQuery(queryId) {
    const existing = this.registrations.get(queryId);
    if (!existing) return;
    if (existing.timer) this.timers.clearTimeout(existing.timer);
    this.registrations.delete(queryId);
  }

  /**
   * Trigger all `on-open` queries (or a single query if an ID is provided).
   * @param {string} [queryId]
   */
  triggerOnOpen(queryId) {
    if (queryId) {
      const reg = this.registrations.get(queryId);
      if (reg?.policy.type === "on-open") {
        this.enqueue(queryId, "on-open", { dedupe: true });
      }
      return;
    }

    for (const [id, reg] of this.registrations) {
      if (reg.policy.type === "on-open") {
        this.enqueue(id, "on-open", { dedupe: true });
      }
    }
  }

  /**
   * Enqueue a refresh and return a handle for awaiting/cancelling it.
   * @param {string} queryId
   * @param {RefreshReason} [reason]
   */
  refresh(queryId, reason = "manual") {
    return this.enqueue(queryId, reason, { dedupe: false });
  }

  /**
   * @private
   * @param {string} queryId
   * @param {RefreshReason} reason
   * @param {{ dedupe: boolean }} options
   * @returns {RefreshHandle}
   */
  enqueue(queryId, reason, options) {
    const reg = this.registrations.get(queryId);
    if (!reg) {
      throw new Error(`Unknown query '${queryId}'`);
    }

    if (options.dedupe) {
      if (this.runningHasQuery(queryId) || this.queueHasQuery(queryId)) {
        const existing =
          Array.from(this.running.values()).find((job) => job.info.queryId === queryId) ??
          this.queue.find((job) => job.info.queryId === queryId);
        if (existing) {
          return {
            id: existing.info.id,
            queryId,
            promise: existing.promise,
            cancel: () => this.cancel(existing.info.id),
          };
        }
      }
    }

    const id = `refresh_${this.nextJobId++}`;
    /** @type {RefreshJobInfo} */
    const info = { id, queryId, reason, queuedAt: new Date(this.now()) };

    const controller = new AbortController();

    /** @type {(value: QueryExecutionResult) => void} */
    let resolve;
    /** @type {(reason?: any) => void} */
    let reject;
    const promise = new Promise((res, rej) => {
      resolve = res;
      reject = rej;
    });

    /** @type {RefreshJob} */
    const job = {
      info,
      query: reg.query,
      controller,
      promise,
      resolve,
      reject,
    };

    this.queue.push(job);
    this.emitter.emit("event", /** @type {RefreshEvent} */ ({ type: "queued", job: { ...info } }));
    this.pump();

    return {
      id,
      queryId,
      promise,
      cancel: () => this.cancel(id),
    };
  }

  /**
   * Cancel a queued or running refresh job.
   * @param {string} jobId
   */
  cancel(jobId) {
    const running = this.running.get(jobId);
    if (running) {
      running.controller.abort();
      return;
    }

    const idx = this.queue.findIndex((j) => j.info.id === jobId);
    if (idx >= 0) {
      const [job] = this.queue.splice(idx, 1);
      job.reject(abortError("Aborted"));
      this.emitter.emit("event", /** @type {RefreshEvent} */ ({ type: "cancelled", job: { ...job.info } }));
    }
  }

  dispose() {
    for (const [id] of this.registrations) this.unregisterQuery(id);
    for (const job of this.queue.slice()) this.cancel(job.info.id);
    for (const job of this.running.values()) this.cancel(job.info.id);
  }

  /**
   * @private
   * @param {string} queryId
   */
  runningHasQuery(queryId) {
    for (const job of this.running.values()) {
      if (job.info.queryId === queryId) return true;
    }
    return false;
  }

  /**
   * @private
   * @param {string} queryId
   */
  queueHasQuery(queryId) {
    return this.queue.some((job) => job.info.queryId === queryId);
  }

  /**
   * @private
   * @param {string} queryId
   * @param {number} intervalMs
   */
  scheduleInterval(queryId, intervalMs) {
    const reg = this.registrations.get(queryId);
    if (!reg) return;
    if (reg.timer) this.timers.clearTimeout(reg.timer);

    reg.timer = this.timers.setTimeout(() => {
      if (!this.registrations.has(queryId)) return;
      this.enqueue(queryId, "interval", { dedupe: true });
      this.scheduleInterval(queryId, intervalMs);
    }, intervalMs);
  }

  /**
   * @private
   */
  pump() {
    while (this.running.size < this.concurrency && this.queue.length > 0) {
      const job = this.queue.shift();
      if (!job) break;
      void this.start(job);
    }
  }

  /**
   * @private
   * @param {RefreshJob} job
   */
  async start(job) {
    job.info.startedAt = new Date(this.now());
    this.running.set(job.info.id, job);
    this.emitter.emit("event", /** @type {RefreshEvent} */ ({ type: "started", job: { ...job.info } }));

    try {
      const context = this.getContext();
      const result = await this.engine.executeQueryWithMeta(job.query, context, {
        signal: job.controller.signal,
        onProgress: (event) => {
          this.emitter.emit("event", /** @type {RefreshEvent} */ ({ type: "progress", job: { ...job.info }, event }));
        },
      });
      job.info.completedAt = new Date(this.now());
      this.emitter.emit("event", /** @type {RefreshEvent} */ ({ type: "completed", job: { ...job.info }, result }));
      job.resolve(result);
    } catch (error) {
      job.info.completedAt = new Date(this.now());
      if (job.controller.signal.aborted || /** @type {any} */ (error)?.name === "AbortError") {
        this.emitter.emit("event", /** @type {RefreshEvent} */ ({ type: "cancelled", job: { ...job.info } }));
        job.reject(abortError("Aborted"));
      } else {
        this.emitter.emit("event", /** @type {RefreshEvent} */ ({ type: "error", job: { ...job.info }, error }));
        job.reject(error);
      }
    } finally {
      this.running.delete(job.info.id);
      this.pump();
    }
  }
}

/**
 * @typedef {{
 *   info: RefreshJobInfo;
 *   query: Query;
 *   controller: AbortController;
 *   promise: Promise<QueryExecutionResult>;
 *   resolve: (value: QueryExecutionResult) => void;
 *   reject: (reason?: any) => void;
 * }} RefreshJob
 */

/**
 * @typedef {{
 *   id: string;
 *   queryId: string;
 *   promise: Promise<QueryExecutionResult>;
 *   cancel: () => void;
 * }} RefreshHandle
 */

// Backwards-compatible shim around the old `QueryScheduler` prototype API.
export class QueryScheduler {
  /**
   * @param {{ engine: QueryEngine, getContext?: () => QueryExecutionContext, concurrency?: number }} options
   */
  constructor(options) {
    this.manager = new RefreshManager({ engine: options.engine, getContext: options.getContext, concurrency: options.concurrency ?? 1 });
  }

  /**
   * @param {Query} query
   * @param {(table: import("./table.js").DataTable, meta: any) => void} onResult
   */
  schedule(query, onResult) {
    this.manager.registerQuery(query, query.refreshPolicy ?? { type: "manual" });
    if ((query.refreshPolicy ?? { type: "manual" }).type === "interval") {
      this.manager.onEvent((evt) => {
        if (evt.type === "completed" && evt.job.queryId === query.id) onResult(evt.result.table, evt.result.meta);
      });
    }
  }

  /**
   * @param {string} queryId
   */
  unschedule(queryId) {
    this.manager.unregisterQuery(queryId);
  }

  /**
   * @param {Query} query
   */
  async refreshNow(query) {
    this.manager.registerQuery(query, query.refreshPolicy ?? { type: "manual" });
    const handle = this.manager.refresh(query.id, "manual");
    const result = await handle.promise;
    return result.table;
  }
}
