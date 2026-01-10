/**
 * Refresh policy + scheduling for queries.
 *
 * This module is deliberately small: the "real" product should plug this into
 * a persisted job scheduler (e.g. Tauri background tasks) and a document-level
 * refresh coordinator.
 */

/**
 * @typedef {import("./model.js").Query} Query
 * @typedef {import("./model.js").RefreshPolicy} RefreshPolicy
 * @typedef {import("./engine.js").QueryEngine} QueryEngine
 * @typedef {import("./engine.js").QueryExecutionContext} QueryExecutionContext
 * @typedef {import("./table.js").DataTable} DataTable
 */

/**
 * @typedef {{
 *   queryId: string;
 *   startedAt: Date;
 *   completedAt: Date;
 *   rowCount: number;
 * }} RefreshResult
 */

export class QueryScheduler {
  /**
   * @param {{
   *   engine: QueryEngine;
   *   getContext?: () => QueryExecutionContext;
   * }} options
   */
  constructor(options) {
    this.engine = options.engine;
    this.getContext = options.getContext ?? (() => ({}));
    /** @type {Map<string, { timer: any, query: Query, onResult: (table: DataTable, meta: RefreshResult) => void }>} */
    this.jobs = new Map();
  }

  /**
   * Start (or replace) a scheduled refresh job for a query.
   * @param {Query} query
   * @param {(table: DataTable, meta: RefreshResult) => void} onResult
   */
  schedule(query, onResult) {
    this.unschedule(query.id);

    const policy = query.refreshPolicy ?? { type: "manual" };
    if (policy.type === "manual") return;
    if (policy.type === "interval") {
      const timer = setInterval(() => {
        void this.refresh(query, onResult);
      }, policy.intervalMs);
      this.jobs.set(query.id, { timer, query, onResult });
      return;
    }

    if (policy.type === "cron") {
      throw new Error("Cron refresh policy is not implemented in this prototype");
    }

    /** @type {never} */
    const exhausted = policy;
    throw new Error(`Unsupported refresh policy '${exhausted.type}'`);
  }

  /**
   * Stop a scheduled job.
   * @param {string} queryId
   */
  unschedule(queryId) {
    const existing = this.jobs.get(queryId);
    if (!existing) return;
    clearInterval(existing.timer);
    this.jobs.delete(queryId);
  }

  /**
   * Trigger a refresh immediately.
   * @param {Query} query
   * @returns {Promise<DataTable>}
   */
  async refreshNow(query) {
    const context = this.getContext();
    return this.engine.executeQuery(query, context, {});
  }

  /**
   * @private
   * @param {Query} query
   * @param {(table: DataTable, meta: RefreshResult) => void} onResult
   */
  async refresh(query, onResult) {
    const startedAt = new Date();
    const table = await this.refreshNow(query);
    const completedAt = new Date();
    onResult(table, { queryId: query.id, startedAt, completedAt, rowCount: table.rows.length });
  }
}

