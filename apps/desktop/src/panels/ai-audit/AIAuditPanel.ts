import { LocalStorageAIAuditStore } from "@formula/ai-audit/browser";
import type { AIAuditEntry, AIAuditStore } from "@formula/ai-audit/browser";

import { createAuditLogExport, downloadAuditLogExport } from "./exportAuditLog";

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

function formatToolCall(call: AIAuditEntry["tool_calls"][number]): string {
  const approved = call.approved === undefined ? "—" : String(call.approved);
  const ok = call.ok === undefined ? "—" : String(call.ok);
  return `${call.name} (approved: ${approved}, ok: ${ok})`;
}

function extractWorkbookId(entry: AIAuditEntry): string | null {
  const input = entry.input as unknown;
  if (!input || typeof input !== "object") return null;
  const obj = input as Record<string, unknown>;
  const workbookId = obj.workbook_id ?? obj.workbookId;
  return typeof workbookId === "string" ? workbookId : null;
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
  const store = options.store ?? new LocalStorageAIAuditStore();

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
    style: "display: flex; flex-direction: column; gap: 10px; overflow: auto; min-height: 0; flex: 1;",
  });

  let currentEntries: AIAuditEntry[] = [];

  function renderEntries(entries: AIAuditEntry[]) {
    list.replaceChildren();
    if (entries.length === 0) {
      list.append(el("div", { style: "font-size: 12px; opacity: 0.8;" }, ["No audit entries found."]));
      return;
    }

    for (const entry of entries) {
      const toolCalls = entry.tool_calls ?? [];
      const tokenUsage = formatTokenUsage(entry);
      const latency = formatLatency(entry);
      const stats = [tokenUsage, latency].filter(Boolean).join(" • ");

      const toolsNode =
        toolCalls.length === 0
          ? el("div", { style: "font-size: 12px; opacity: 0.8;" }, ["Tools: none"])
          : el(
              "div",
              { style: "display: flex; flex-direction: column; gap: 4px; font-size: 12px;" },
              toolCalls.map((call) => el("div", { "data-testid": "ai-audit-tool-call" }, [formatToolCall(call)])),
            );

      const entryNode = el(
        "div",
        {
          "data-testid": "ai-audit-entry",
          style: "border: 1px solid var(--border); border-radius: 8px; padding: 10px; color: var(--text-primary);",
        },
        [
          el("div", { style: "font-size: 12px; opacity: 0.75; margin-bottom: 4px;" }, [
            `${formatTimestamp(entry.timestamp_ms)} • ${entry.mode} • ${entry.model}`,
          ]),
          el("div", { style: "font-size: 12px; opacity: 0.8; margin-bottom: 6px;" }, [`session_id: ${entry.session_id}`]),
          stats ? el("div", { style: "font-size: 12px; opacity: 0.85; margin-bottom: 6px;" }, [stats]) : el("div"),
          toolsNode,
        ],
      );

      list.append(entryNode);
    }
  }

  let refreshInFlight = false;
  let refreshQueued = false;

  async function refresh() {
    if (refreshInFlight) {
      refreshQueued = true;
      return;
    }

    refreshInFlight = true;
    entriesMeta.textContent = "Loading…";
    const session_id = sessionInput.value.trim() || undefined;
    const workbookId = workbookInput.value.trim() || undefined;

    try {
      const entries = await store.listEntries({ session_id });
      const filtered = workbookId
        ? entries.filter((entry) => extractWorkbookId(entry) === workbookId)
        : entries;
      currentEntries = filtered;
      entriesMeta.textContent = `Showing ${filtered.length} entr${filtered.length === 1 ? "y" : "ies"}.`;
      renderEntries(filtered);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      currentEntries = [];
      entriesMeta.textContent = "Failed to load audit log.";
      list.replaceChildren(el("div", { style: "font-size: 12px; opacity: 0.8;" }, [`Error: ${message}`]));
    } finally {
      refreshInFlight = false;
      if (refreshQueued) {
        refreshQueued = false;
        void refresh();
      }
    }
  }

  const refreshButton = el(
    "button",
    { type: "button", "data-testid": "ai-audit-refresh", onClick: () => void refresh() },
    ["Refresh"],
  );

  const exportButton = el(
    "button",
    {
      type: "button",
      "data-testid": "ai-audit-export-json",
      onClick: () => {
        const exp = createAuditLogExport(currentEntries);
        downloadAuditLogExport(exp);
      },
    },
    ["Export JSON"],
  );

  function onFilterKeyDown(event: KeyboardEvent) {
    if (event.key === "Enter") {
      event.preventDefault();
      void refresh();
    }
  }

  sessionInput.addEventListener("keydown", onFilterKeyDown);
  workbookInput.addEventListener("keydown", onFilterKeyDown);

  const controls = el(
    "div",
    {
      style:
        "display: flex; flex-wrap: wrap; gap: 8px; align-items: center; padding-bottom: 8px; border-bottom: 1px solid var(--border); margin-bottom: 10px;",
    },
    [
      sessionInput,
      workbookInput,
      refreshButton,
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

  const autoRefreshMs = options.autoRefreshMs ?? 0;
  const intervalId =
    autoRefreshMs > 0
      ? globalThis.setInterval(() => {
          void refresh();
        }, autoRefreshMs)
      : null;

  // Avoid keeping Node-based test runners alive when auto-refresh is enabled.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (intervalId as any)?.unref?.();

  return {
    ready,
    refresh,
    getEntries() {
      return currentEntries.slice();
    },
    dispose() {
      if (intervalId != null) globalThis.clearInterval(intervalId);
      options.container.textContent = "";
    },
  };
}
