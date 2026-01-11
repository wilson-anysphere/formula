import type { Query, RefreshPolicy } from "../../../../packages/power-query/src/model.js";
import { QueryEngine } from "../../../../packages/power-query/src/engine.js";

import type { QueryExecutionContext } from "../../../../packages/power-query/src/engine.js";

import type { DocumentController } from "../document/documentController.js";

import { createPowerQueryCredentialManager, type PowerQueryCredentialPrompt } from "./credentialManager.ts";
import { maybeGetPowerQueryDlpContext } from "./dlpContext.ts";
import { createDesktopQueryEngine, getContextForDocument } from "./engine.ts";
import { enqueueApplyForDocument } from "./applyQueue.ts";
import {
  DesktopPowerQueryRefreshManager,
  type DesktopPowerQueryEvent,
  type DesktopPowerQueryRefreshOptions,
} from "./refresh.ts";
import { applyQueryToDocument, type ApplyToDocumentResult, type QuerySheetDestination } from "./applyToDocument.ts";
import { createPowerQueryRefreshStateStore, type RefreshStateStore } from "./refreshStateStore.ts";

type StorageLike = { getItem(key: string): string | null; setItem(key: string, value: string): void; removeItem(key: string): void };

function getLocalStorageOrNull(): StorageLike | null {
  if (typeof window !== "undefined") {
    try {
      return window.localStorage;
    } catch {
      return null;
    }
  }
  try {
    const storage = (globalThis as any)?.localStorage as any;
    if (storage && typeof storage.getItem === "function") return storage as StorageLike;
  } catch {
  }
  return null;
}

function safeParseJson(text: string): any | null {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function normalizeWorkbookId(workbookId: string | undefined): string {
  return workbookId?.trim() ? workbookId : "default";
}

function queriesStorageKey(workbookId: string): string {
  return `formula.desktop.powerQuery.queries:${normalizeWorkbookId(workbookId)}`;
}

function legacyQueryStorageKey(workbookId: string): string {
  return `formula.desktop.powerQuery.query:${normalizeWorkbookId(workbookId)}`;
}

export function loadQueriesFromStorage(workbookId: string): Query[] {
  const storage = getLocalStorageOrNull();
  if (!storage) return [];

  const stored = storage.getItem(queriesStorageKey(workbookId));
  if (stored) {
    const parsed = safeParseJson(stored);
    if (Array.isArray(parsed)) return parsed as Query[];
    if (parsed && typeof parsed === "object") return [parsed as Query];
  }

  const legacy = storage.getItem(legacyQueryStorageKey(workbookId));
  if (!legacy) return [];
  const parsed = safeParseJson(legacy);
  if (!parsed || typeof parsed !== "object") return [];
  return [parsed as Query];
}

export function saveQueriesToStorage(workbookId: string, queries: Query[]): void {
  const storage = getLocalStorageOrNull();
  if (!storage) return;
  try {
    storage.setItem(queriesStorageKey(workbookId), JSON.stringify(queries));
    storage.removeItem(legacyQueryStorageKey(workbookId));
  } catch {
  }
}

function isAbortError(error: unknown): boolean {
  return (error as any)?.name === "AbortError";
}

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

export type DesktopPowerQueryServiceEvent =
  | DesktopPowerQueryEvent
  | { type: "queries:changed"; queries: Query[] };

export type DesktopPowerQueryServiceOptions = {
  workbookId: string;
  document: DocumentController;
  getContext?: () => QueryExecutionContext;
  concurrency?: number;
  batchSize?: number;
  engine?: QueryEngine;
  refresh?: Pick<DesktopPowerQueryRefreshOptions, "timers" | "now" | "timezone"> & { stateStore?: RefreshStateStore };
  credentialPrompt?: PowerQueryCredentialPrompt;
};

export class DesktopPowerQueryService {
  readonly workbookId: string;
  readonly document: DocumentController;
  readonly engine: QueryEngine;
  readonly engineError: string | null;

  private readonly emitter = new Emitter<DesktopPowerQueryServiceEvent>();
  private readonly refreshManager: DesktopPowerQueryRefreshManager;
  private readonly getContext: () => QueryExecutionContext;
  private readonly queries = new Map<string, Query>();
  private readonly applyControllers = new Map<string, AbortController>();
  private readonly unsubscribeRefreshEvents: (() => void) | null;

  constructor(options: DesktopPowerQueryServiceOptions) {
    this.workbookId = normalizeWorkbookId(options.workbookId);
    this.document = options.document;
    this.getContext = options.getContext ?? (() => getContextForDocument(this.document));

    const creds = createPowerQueryCredentialManager({ prompt: options.credentialPrompt });

    let engine: QueryEngine;
    let engineError: string | null = null;
    if (options.engine) {
      engine = options.engine;
    } else {
      try {
        const dlp = maybeGetPowerQueryDlpContext({ documentId: this.workbookId });
        engine = createDesktopQueryEngine({ dlp: dlp ?? undefined, onCredentialRequest: creds.onCredentialRequest });
      } catch (err: any) {
        engine = new QueryEngine();
        engineError = err?.message ?? String(err);
      }
    }

    this.engine = engine;
    this.engineError = engineError;

    const refreshStateStore = options.refresh?.stateStore ?? createPowerQueryRefreshStateStore({ workbookId: this.workbookId });

    this.refreshManager = new DesktopPowerQueryRefreshManager({
      engine,
      document: this.document,
      getContext: this.getContext,
      concurrency: options.concurrency ?? 1,
      batchSize: options.batchSize ?? 1024,
      timers: options.refresh?.timers,
      now: options.refresh?.now,
      timezone: options.refresh?.timezone,
      stateStore: refreshStateStore,
    });

    this.unsubscribeRefreshEvents = this.refreshManager.onEvent((evt) => {
      this.emitter.emit(evt);
      if (evt?.type === "apply:completed" || evt?.type === "apply:cancelled" || evt?.type === "apply:error") {
        this.persistQueries();
      }
    });

    const initialQueries = loadQueriesFromStorage(this.workbookId);
    if (initialQueries.length > 0) {
      this.setQueries(initialQueries);
      this.refreshManager.triggerOnOpen();
    }
  }

  onEvent(handler: (event: DesktopPowerQueryServiceEvent) => void): () => void {
    return this.emitter.on(handler);
  }

  getQueries(): Query[] {
    return Array.from(this.queries.values());
  }

  getQuery(queryId: string): Query | null {
    return this.queries.get(queryId) ?? null;
  }

  registerQuery(query: Query, policyOverride?: RefreshPolicy): void {
    const effectivePolicy = policyOverride ?? query.refreshPolicy ?? { type: "manual" };
    const updated = { ...query, refreshPolicy: effectivePolicy };
    this.queries.set(updated.id, updated);
    try {
      this.refreshManager.registerQuery(updated, effectivePolicy);
    } catch {
      const fallback: Query = { ...updated, refreshPolicy: { type: "manual" } };
      this.queries.set(fallback.id, fallback);
      try {
        this.refreshManager.registerQuery(fallback, fallback.refreshPolicy);
      } catch {
        // ignore
      }
    }
    this.persistQueries();
    this.emitter.emit({ type: "queries:changed", queries: this.getQueries() });
  }

  unregisterQuery(queryId: string): void {
    if (!this.queries.has(queryId)) return;
    this.queries.delete(queryId);
    this.refreshManager.unregisterQuery(queryId);
    this.persistQueries();
    this.emitter.emit({ type: "queries:changed", queries: this.getQueries() });
  }

  setQueries(queries: Query[]): void {
    const nextIds = new Set(queries.map((q) => q.id));
    for (const existingId of this.queries.keys()) {
      if (!nextIds.has(existingId)) {
        this.refreshManager.unregisterQuery(existingId);
      }
    }

    this.queries.clear();
    for (const query of queries) {
      const policy = query.refreshPolicy ?? { type: "manual" };
      const updated = { ...query, refreshPolicy: policy };
      this.queries.set(updated.id, updated);
      try {
        this.refreshManager.registerQuery(updated, policy);
      } catch {
        const fallback: Query = { ...updated, refreshPolicy: { type: "manual" } };
        this.queries.set(fallback.id, fallback);
        try {
          this.refreshManager.registerQuery(fallback, fallback.refreshPolicy);
        } catch {
          // ignore
        }
      }
    }

    this.persistQueries();
    this.emitter.emit({ type: "queries:changed", queries: this.getQueries() });
  }

  refresh(queryId: string, reason: any = "manual") {
    return this.refreshManager.refresh(queryId, reason);
  }

  refreshAll(queryIds?: string[], reason: any = "manual") {
    return this.refreshManager.refreshAll(queryIds, reason);
  }

  refreshWithDependencies(queryId: string, reason: any = "manual") {
    return this.refreshManager.refreshWithDependencies(queryId, reason);
  }

  loadToSheet(queryId: string, destination: QuerySheetDestination, options?: { batchSize?: number }) {
    const query = this.queries.get(queryId);
    if (!query) throw new Error(`Unknown query '${queryId}'`);

    const controller = new AbortController();
    const jobId = `load_${typeof crypto !== "undefined" && typeof crypto.randomUUID === "function" ? crypto.randomUUID() : String(Date.now())}`;
    this.applyControllers.set(jobId, controller);

    this.emitter.emit({ type: "apply:started", jobId, queryId, destination });

    const promise: Promise<ApplyToDocumentResult> = enqueueApplyForDocument(this.document, async () => {
      try {
        if (controller.signal.aborted) {
          const err = new Error("Aborted");
          (err as any).name = "AbortError";
          throw err;
        }

        const result = await applyQueryToDocument(this.document, query, destination, {
          engine: this.engine,
          context: this.getContext(),
          batchSize: options?.batchSize ?? 1024,
          signal: controller.signal,
          label: `Load query: ${query.name}`,
          onProgress: async (evt) => {
            if (evt.type === "batch") {
              this.emitter.emit({ type: "apply:progress", jobId, queryId, rowsWritten: evt.totalRowsWritten });
            }
          },
        });

        const updated = { ...query, destination };
        this.registerQuery(updated);

        this.emitter.emit({ type: "apply:completed", jobId, queryId, result });
        return result;
      } catch (error) {
        if (controller.signal.aborted || isAbortError(error)) {
          this.emitter.emit({ type: "apply:cancelled", jobId, queryId });
        } else {
          this.emitter.emit({ type: "apply:error", jobId, queryId, error });
        }
        throw error;
      } finally {
        this.applyControllers.delete(jobId);
      }
    });

    return {
      id: jobId,
      queryId,
      promise,
      cancel: () => controller.abort(),
    };
  }

  dispose(): void {
    this.unsubscribeRefreshEvents?.();
    for (const controller of this.applyControllers.values()) controller.abort();
    this.applyControllers.clear();
    this.refreshManager.dispose();
    this.queries.clear();
  }

  private persistQueries(): void {
    saveQueriesToStorage(this.workbookId, this.getQueries());
  }
}

const services = new Map<string, DesktopPowerQueryService>();
const registryEmitter = new Emitter<{ workbookId: string; service: DesktopPowerQueryService | null }>();

export function getDesktopPowerQueryService(workbookId: string | undefined): DesktopPowerQueryService | null {
  return services.get(normalizeWorkbookId(workbookId)) ?? null;
}

export function setDesktopPowerQueryService(workbookId: string | undefined, service: DesktopPowerQueryService | null): void {
  const key = normalizeWorkbookId(workbookId);
  if (service) services.set(key, service);
  else services.delete(key);
  registryEmitter.emit({ workbookId: key, service });
}

export function onDesktopPowerQueryServiceChanged(
  workbookId: string | undefined,
  handler: (service: DesktopPowerQueryService | null) => void,
): () => void {
  const key = normalizeWorkbookId(workbookId);
  handler(services.get(key) ?? null);
  return registryEmitter.on((evt) => {
    if (evt.workbookId !== key) return;
    handler(evt.service);
  });
}
