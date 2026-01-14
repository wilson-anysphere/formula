import type { AIAuditEntry, AIAuditStore } from "@formula/ai-audit/browser";
import { serializeAuditEntries } from "@formula/ai-audit/export";

import { downloadAuditLogExport } from "./exportAuditLog";
import { getDesktopAIAuditStore } from "../../ai/audit/auditStore";

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Record<string, unknown> = {},
  children: Array<Node | string> = [],
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "className" && typeof value === "string") {
      node.className = value;
      continue;
    }
    if (key.startsWith("on") && typeof value === "function") {
      node.addEventListener(key.slice(2).toLowerCase(), value as EventListener);
      continue;
    }
    if (value === undefined) continue;
    node.setAttribute(key, String(value));
  }
  for (const child of children) node.append(typeof child === "string" ? document.createTextNode(child) : child);
  return node;
}

function formatTimestamp(timestamp_ms: number): string {
  try {
    return new Date(timestamp_ms).toISOString();
  } catch {
    return String(timestamp_ms);
  }
}

function formatTokenUsage(entry: AIAuditEntry): string | null {
  if (!entry.token_usage) return null;
  const { prompt_tokens, completion_tokens, total_tokens } = entry.token_usage;
  const parts = [`prompt ${prompt_tokens}`, `completion ${completion_tokens}`];
  if (typeof total_tokens === "number") parts.push(`total ${total_tokens}`);
  return `Tokens: ${parts.join(", ")}`;
}

function formatLatency(entry: AIAuditEntry): string | null {
  if (typeof entry.latency_ms !== "number") return null;
  const rounded = Math.round(entry.latency_ms);
  return `Latency: ${rounded}ms`;
}

function renderVerification(entry: AIAuditEntry): HTMLElement | null {
  const verification = entry.verification;
  if (!verification) return null;

  const confidence = Number(verification.confidence ?? 0);
  const confidencePct = Number.isFinite(confidence) ? Math.round(confidence * 100) : 0;
  const status = verification.verified ? "Verified" : "Unverified";
  const warnings =
    Array.isArray(verification.warnings) && verification.warnings.length > 0 ? ` • ${verification.warnings.join(" ")}` : "";
  const text = `Verification: ${status} (confidence ${confidencePct}%)${warnings}`;

  const claims = Array.isArray((verification as any).claims) ? ((verification as any).claims as any[]) : [];
  const detailsNode =
    claims.length > 0
      ? (() => {
          const verifiedCount = claims.filter((c) => c?.verified === true).length;
          const summary = `Claims: ${verifiedCount}/${claims.length} verified`;
           return el(
             "details",
            { "data-testid": "ai-audit-verification-claims", style: "margin-top: var(--space-3);" },
            [
              el("summary", { style: "cursor: pointer;" }, [summary]),
              el(
                "pre",
                { style: "white-space: pre-wrap; font-size: 11px; opacity: 0.9; margin: var(--space-3) 0 0 0;" },
                [JSON.stringify(claims, null, 2)],
              ),
            ],
          );
        })()
      : null;

  const children: Array<Node | string> = [text];
  if (detailsNode) children.push(detailsNode);

  if (verification.verified) {
    return el(
      "div",
      {
        "data-testid": "ai-audit-verification",
        style: "font-size: 12px; opacity: 0.85; margin-bottom: var(--space-3);",
      },
      children,
    );
  }

  return el(
    "div",
    {
      "data-testid": "ai-audit-verification",
      style:
        "font-size: 12px; margin-bottom: var(--space-3); padding: var(--space-3) var(--space-4); border: 1px solid var(--border); border-radius: var(--radius); background: var(--warning-bg);",
    },
    children,
  );
}

function formatToolCall(call: AIAuditEntry["tool_calls"][number]): string {
  const approved = call.approved === undefined ? "—" : String(call.approved);
  const ok = call.ok === undefined ? "—" : String(call.ok);
  const requiresApproval = call.requires_approval === undefined ? "—" : String(call.requires_approval);
  const duration = typeof call.duration_ms === "number" && Number.isFinite(call.duration_ms) ? Math.round(call.duration_ms) : null;
  const error =
    typeof call.error === "string" && call.error.trim().length > 0 ? truncate(call.error.trim(), 120) : null;

  const parts = [`approved: ${approved}`, `ok: ${ok}`];
  if (requiresApproval !== "—") parts.push(`requires_approval: ${requiresApproval}`);
  if (duration != null) parts.push(`duration: ${duration}ms`);
  if (error) parts.push(`error: ${error}`);

  return `${call.name} (${parts.join(", ")})`;
}

function extractWorkbookId(entry: AIAuditEntry): string | null {
  if (typeof entry.workbook_id === "string" && entry.workbook_id.trim()) return entry.workbook_id;
  const input = entry.input as unknown;
  if (!input || typeof input !== "object") return null;
  const obj = input as Record<string, unknown>;
  const workbookId = obj.workbook_id ?? obj.workbookId;
  return typeof workbookId === "string" ? workbookId : null;
}

function truncate(text: string, maxChars: number): string {
  if (text.length <= maxChars) return text;
  return `${text.slice(0, maxChars - 1)}…`;
}

function sanitizeFileNameComponent(value: string): string {
  // Replace characters that are problematic across filesystems (slashes, colons,
  // etc). This is only used for download names, not for storage keys.
  return value.replace(/[^a-z0-9_-]+/gi, "_").replace(/^_+|_+$/g, "").slice(0, 64) || "unknown";
}

function buildExportFileName(
  filters: { workbookId?: string; sessionId?: string },
  format: "ndjson" | "json" = "ndjson",
): string {
  const stamp = new Date().toISOString().replaceAll(":", "-");
  const parts: string[] = [];
  if (filters.workbookId) parts.push(`workbook-${sanitizeFileNameComponent(filters.workbookId)}`);
  if (filters.sessionId) parts.push(`session-${sanitizeFileNameComponent(filters.sessionId)}`);
  const suffix = parts.length ? `-${parts.join("_")}` : "";
  const ext = format === "json" ? "json" : "ndjson";
  return `ai-audit-log-${stamp}${suffix}.${ext}`;
}

export interface CreateAIAuditPanelOptions {
  container: HTMLElement;
  store?: AIAuditStore;
  initialSessionId?: string;
  initialWorkbookId?: string;
  /**
   * Auto-refresh the list on an interval while mounted (useful during demos).
   * Disabled by default.
   */
  autoRefreshMs?: number;
}

export function createAIAuditPanel(options: CreateAIAuditPanelOptions) {
  const store = options.store ?? getDesktopAIAuditStore();

  const timeRangeSelect = el(
    "select",
    {
      "data-testid": "ai-audit-filter-time-range",
      style: "min-width: 140px;",
      title: "Filter entries by recency",
    },
    [
      el("option", { value: "all" }, ["All time"]),
      el("option", { value: "1h" }, ["Last 1h"]),
      el("option", { value: "24h" }, ["Last 24h"]),
      el("option", { value: "7d" }, ["Last 7d"]),
      el("option", { value: "30d" }, ["Last 30d"]),
    ],
  );
  (timeRangeSelect as HTMLSelectElement).value = "all";

  const pageSizeInput = el("input", {
    type: "number",
    min: "1",
    max: "1000",
    step: "1",
    value: "200",
    placeholder: "page size",
    "data-testid": "ai-audit-filter-page-size",
    style: "width: 100px;",
    title: "How many entries to load per page",
  });

  const sessionInput = el("input", {
    type: "text",
    placeholder: "session_id (optional)",
    value: options.initialSessionId ?? "",
    "data-testid": "ai-audit-filter-session",
    style: "flex: 1; min-width: 180px;",
  });

  const workbookInput = el("input", {
    type: "text",
    placeholder: "workbook_id (optional)",
    value: options.initialWorkbookId ?? "",
    "data-testid": "ai-audit-filter-workbook",
    style: "flex: 1; min-width: 180px;",
  });

  const entriesMeta = el("div", { "data-testid": "ai-audit-meta", style: "font-size: 12px; opacity: 0.8;" }, [
    "Loading…",
  ]);

  const list = el("div", {
    "data-testid": "ai-audit-entries",
    style:
      "display: flex; flex-direction: column; gap: calc(var(--space-4) + var(--space-1)); overflow: auto; min-height: 0; flex: 1;",
  });

  let currentEntries: AIAuditEntry[] = [];
  let hasMore: boolean = false;

  const refreshButton = el(
    "button",
    {
      type: "button",
      "data-testid": "ai-audit-refresh",
      onClick: () =>
        void refresh().catch(() => {
          // Best-effort: avoid unhandled rejections from fire-and-forget UI handlers.
        }),
    },
    ["Refresh"],
  );

  const loadMoreButton = el(
    "button",
    {
      type: "button",
      "data-testid": "ai-audit-load-more",
      onClick: () =>
        void loadMore().catch(() => {
          // Best-effort: avoid unhandled rejections from fire-and-forget UI handlers.
        }),
      disabled: true,
      title: "Load older entries",
    },
    ["Load more"],
  );

  function renderEntry(entry: AIAuditEntry): HTMLElement {
    const toolCalls = entry.tool_calls ?? [];
    const tokenUsage = formatTokenUsage(entry);
    const latency = formatLatency(entry);
    const stats = [tokenUsage, latency].filter(Boolean).join(" • ");
    const verificationNode = renderVerification(entry);
    const workbookId = extractWorkbookId(entry);

    const toolsNode =
      toolCalls.length === 0
        ? el("div", { style: "font-size: 12px; opacity: 0.8;" }, ["Tools: none"])
        : el(
            "div",
            { style: "display: flex; flex-direction: column; gap: var(--space-2); font-size: 12px;" },
            toolCalls.map((call) => el("div", { "data-testid": "ai-audit-tool-call" }, [formatToolCall(call)])),
          );

    return el(
      "div",
      {
        "data-testid": "ai-audit-entry",
        style:
          "border: 1px solid var(--border); border-radius: var(--radius); padding: calc(var(--space-4) + var(--space-1)); color: var(--text-primary);",
      },
      [
        el("div", { style: "font-size: 12px; opacity: 0.75; margin-bottom: var(--space-2);" }, [
          `${formatTimestamp(entry.timestamp_ms)} • ${entry.mode} • ${entry.model}`,
        ]),
        el("div", { style: "font-size: 12px; opacity: 0.8; margin-bottom: var(--space-3);" }, [`session_id: ${entry.session_id}`]),
        el("div", { style: "font-size: 12px; opacity: 0.8; margin-bottom: var(--space-3);" }, [`workbook_id: ${workbookId ?? "—"}`]),
        ...(stats
          ? [el("div", { style: "font-size: 12px; opacity: 0.85; margin-bottom: var(--space-3);" }, [stats])]
          : []),
        ...(verificationNode ? [verificationNode] : []),
        toolsNode,
      ],
    );
  }

  function replaceEntries(entries: AIAuditEntry[]) {
    list.replaceChildren();
    if (entries.length === 0) {
      list.append(el("div", { style: "font-size: 12px; opacity: 0.8;" }, ["No audit entries found."]));
      return;
    }
    for (const entry of entries) list.append(renderEntry(entry));
  }

  function appendEntries(entries: AIAuditEntry[]) {
    if (entries.length === 0) return;
    // If we were previously showing the empty placeholder, replace it.
    if (currentEntries.length === 0) {
      replaceEntries(entries);
      return;
    }
    for (const entry of entries) list.append(renderEntry(entry));
  }

  function parsePageSize(): number {
    const raw = pageSizeInput.value.trim();
    const value = Number.parseInt(raw, 10);
    if (!Number.isFinite(value)) return 200;
    if (value <= 0) return 200;
    return Math.min(1000, Math.max(1, value));
  }

  function computeAfterTimestampMs(nowMs: number): number | undefined {
    const value = (timeRangeSelect as HTMLSelectElement).value;
    if (value === "1h") return nowMs - 60 * 60 * 1000;
    if (value === "24h") return nowMs - 24 * 60 * 60 * 1000;
    if (value === "7d") return nowMs - 7 * 24 * 60 * 60 * 1000;
    if (value === "30d") return nowMs - 30 * 24 * 60 * 60 * 1000;
    return undefined;
  }

  function updateMeta() {
    const count = currentEntries.length;
    const more = hasMore ? " (more available)" : "";
    entriesMeta.textContent = `Showing ${count} entr${count === 1 ? "y" : "ies"}${more}.`;
  }

  type PendingOperation = "refresh" | "load_more";
  let operationInFlight = false;
  let queuedOperation: PendingOperation | null = null;
  let exportInFlight = false;

  async function runOperation(op: PendingOperation) {
    if (exportInFlight) {
      // Queue the latest operation, but don't run concurrent fetches while exporting.
      if (op === "refresh" || queuedOperation !== "refresh") queuedOperation = op;
      return;
    }

    if (operationInFlight) {
      // Refresh always wins over load_more (since it resets cursor state).
      if (op === "refresh" || queuedOperation !== "refresh") {
        queuedOperation = op;
      }
      return;
    }

    operationInFlight = true;
    entriesMeta.textContent = "Loading…";
    loadMoreButton.disabled = true;

    const session_id = sessionInput.value.trim() || undefined;
    const workbookId = workbookInput.value.trim() || undefined;
    const pageSize = parsePageSize();
    const queryLimit = pageSize + 1;
    const nowMs = Date.now();
    const after_timestamp_ms = computeAfterTimestampMs(nowMs);

    try {
      if (op === "refresh") {
        const entries = await store.listEntries({ session_id, workbook_id: workbookId, after_timestamp_ms, limit: queryLimit });
        const page = entries.slice(0, pageSize);
        currentEntries = page;
        hasMore = entries.length > pageSize;
        replaceEntries(page);
        updateMeta();
      } else {
        const last = currentEntries[currentEntries.length - 1];
        if (!last) {
          hasMore = false;
          updateMeta();
          return;
        }
        const older = await store.listEntries({
          session_id,
          workbook_id: workbookId,
          after_timestamp_ms,
          limit: queryLimit,
          cursor: { before_timestamp_ms: last.timestamp_ms, before_id: last.id },
        });
        const page = older.slice(0, pageSize);
        const existingIds = new Set(currentEntries.map((e) => e.id));
        const deduped = page.filter((e) => !existingIds.has(e.id));
        currentEntries = currentEntries.concat(deduped);
        hasMore = older.length > pageSize && deduped.length > 0;
        appendEntries(deduped);
        updateMeta();
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      currentEntries = [];
      hasMore = false;
      entriesMeta.textContent = "Failed to load audit log.";
      list.replaceChildren(el("div", { style: "font-size: 12px; opacity: 0.8;" }, [`Error: ${message}`]));
    } finally {
      operationInFlight = false;
      const next = queuedOperation;
      queuedOperation = null;
      if (next) {
        void runOperation(next).catch(() => {
          // Best-effort: avoid unhandled rejections from internal queue chaining.
        });
      }
      else loadMoreButton.disabled = !hasMore;
    }
  }

  async function refresh() {
    await runOperation("refresh");
  }

  async function loadMore(): Promise<void> {
    await runOperation("load_more");
  }

  async function exportLog(): Promise<{ blob: Blob; fileName: string } | null> {
    if (exportInFlight) return null;
    if (operationInFlight) {
      // Avoid exporting a moving target while a refresh/load-more fetch is inflight.
      queuedOperation = "refresh";
      return null;
    }

    exportInFlight = true;
    let exportErrorMessage: string | null = null;
    const btn = exportButton as HTMLButtonElement;
    const prevDisabled = btn.disabled;
    btn.disabled = true;
    refreshButton.disabled = true;
    loadMoreButton.disabled = true;

    const sessionId = sessionInput.value.trim() || undefined;
    const workbookId = workbookInput.value.trim() || undefined;
    const after_timestamp_ms = computeAfterTimestampMs(Date.now());

    // Export can be large; fetch in pages and serialize as NDJSON to avoid holding
    // the entire log in memory as objects.
    const exportPageSize = 1000;
    const parts: Array<BlobPart> = [];
    let cursor: { before_timestamp_ms: number; before_id: string } | undefined;
    let total = 0;

    try {
      for (let pageIndex = 0; pageIndex < 10_000; pageIndex++) {
        entriesMeta.textContent = total === 0 ? "Exporting…" : `Exporting… ${total} entr${total === 1 ? "y" : "ies"}`;

        const page = await store.listEntries({
          session_id: sessionId,
          workbook_id: workbookId,
          after_timestamp_ms,
          limit: exportPageSize,
          ...(cursor ? { cursor } : {}),
        });

        if (page.length === 0) break;
        const serialized = serializeAuditEntries(page, { format: "ndjson" });
        if (parts.length > 0) parts.push("\n");
        parts.push(serialized);

        total += page.length;

        const last = page[page.length - 1];
        if (!last) break;
        if (page.length < exportPageSize) break;
        cursor = { before_timestamp_ms: last.timestamp_ms, before_id: last.id };
      }

      const blob = new Blob(parts, { type: "application/x-ndjson" });
      const fileName = buildExportFileName(
        {
          workbookId,
          sessionId,
        },
        "ndjson",
      );
      downloadAuditLogExport({ blob, fileName });
      return { blob, fileName };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      exportErrorMessage = `Failed to export audit log: ${message}`;
      entriesMeta.textContent = exportErrorMessage;
      return null;
    } finally {
      exportInFlight = false;
      btn.disabled = prevDisabled;
      refreshButton.disabled = false;
      loadMoreButton.disabled = !hasMore;
      if (!exportErrorMessage) updateMeta();

      const next = queuedOperation;
      if (next) {
        queuedOperation = null;
        void runOperation(next).catch(() => {
          // Best-effort: avoid unhandled rejections from internal queue chaining.
        });
      }
    }
  }

  const exportButton = el(
    "button",
    {
      type: "button",
      "data-testid": "ai-audit-export-json",
      onClick: () => {
        void exportLog().catch(() => {
          // Best-effort: avoid unhandled rejections from fire-and-forget UI handlers.
        });
      },
    },
    ["Export log"],
  );

  function onFilterKeyDown(event: KeyboardEvent) {
    if (event.key === "Enter") {
      event.preventDefault();
      void refresh().catch(() => {
        // Best-effort: avoid unhandled rejections from fire-and-forget UI handlers.
      });
    }
  }

  sessionInput.addEventListener("keydown", onFilterKeyDown);
  workbookInput.addEventListener("keydown", onFilterKeyDown);
  pageSizeInput.addEventListener("keydown", onFilterKeyDown);
  timeRangeSelect.addEventListener("change", () => {
    void refresh().catch(() => {
      // Best-effort: avoid unhandled rejections from fire-and-forget UI handlers.
    });
  });
  pageSizeInput.addEventListener("change", () => {
    void refresh().catch(() => {
      // Best-effort: avoid unhandled rejections from fire-and-forget UI handlers.
    });
  });

  const controls = el(
    "div",
    {
      style:
        "display: flex; flex-wrap: wrap; gap: var(--space-4); align-items: center; padding-bottom: var(--space-4); border-bottom: 1px solid var(--border); margin-bottom: calc(var(--space-4) + var(--space-1));",
    },
    [
      sessionInput,
      workbookInput,
      timeRangeSelect,
      pageSizeInput,
      refreshButton,
      loadMoreButton,
      exportButton,
      el("div", { style: "flex-basis: 100%; height: 0;" }),
      entriesMeta,
    ],
  );

  const root = el(
    "div",
    { "data-testid": "ai-audit-panel", style: "display: flex; flex-direction: column; height: 100%; min-height: 0;" },
    [controls, list],
  );

  options.container.replaceChildren(root);

  const ready = refresh();
  void ready.catch(() => {
    // Best-effort: avoid unhandled rejections if callers don't await `ready`.
  });

  const autoRefreshMs = options.autoRefreshMs ?? 0;
  const intervalId =
    autoRefreshMs > 0
      ? globalThis.setInterval(() => {
          void refresh().catch(() => {
            // Best-effort: avoid unhandled rejections from fire-and-forget auto-refresh.
          });
        }, autoRefreshMs)
      : null;

  // Avoid keeping Node-based test runners alive when auto-refresh is enabled.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (intervalId as any)?.unref?.();

  return {
    ready,
    refresh,
    loadMore,
    exportLog,
    getEntries() {
      return currentEntries.slice();
    },
    dispose() {
      if (intervalId != null) globalThis.clearInterval(intervalId);
      options.container.textContent = "";
    },
  };
}
