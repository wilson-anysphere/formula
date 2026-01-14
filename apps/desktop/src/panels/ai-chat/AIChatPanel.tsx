import React, { useEffect, useMemo, useRef, useState } from "react";

import type { Attachment, ChatMessage } from "./types.js";
import { showQuickPick, showToast } from "../../extensions/ui.js";
import type {
  ChatStreamEvent,
  LLMClient,
  LLMMessage,
  ToolCall,
  ToolExecutor,
} from "../../../../../packages/llm/src/index.js";
import { runChatWithToolsStreaming } from "../../../../../packages/llm/src/index.js";
import { classifyQueryNeedsTools, verifyAssistantClaims, verifyToolUsage } from "../../../../../packages/ai-tools/src/llm/verification.js";
import { t, tWithVars } from "../../i18n/index.js";

const CONFIDENCE_WARNING_THRESHOLD = 0.7;

function safeStringify(value: unknown, opts: { pretty?: boolean } = {}): string {
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value, null, opts.pretty ? 2 : 0);
  } catch {
    return String(value);
  }
}

function formatAttachmentsForPrompt(attachments: Attachment[]) {
  return attachments
    // Prompts may be sent to cloud models. Never inline raw table/range attachment payloads
    // (they can contain copied spreadsheet values) even when small.
    // Keep formula/chart behavior unchanged.
    .map((a) => {
      const includeData =
        a.data !== undefined && (a.type === "formula" || a.type === "chart");
      return `- ${a.type}: ${a.reference}${includeData ? ` (${safeStringify(a.data)})` : ""}`;
    })
    .join("\n");
}

export interface AIChatPanelSendMessageArgs {
  messages: LLMMessage[];
  userText: string;
  attachments: Attachment[];
  signal?: AbortSignal;
  onStreamEvent?: (event: ChatStreamEvent) => void;
  onToolCall: (call: ToolCall, meta: { requiresApproval: boolean }) => void;
  onToolResult?: (call: ToolCall, result: unknown) => void;
}

export type AIChatPanelSendMessage = (
  args: AIChatPanelSendMessageArgs,
) => Promise<{ messages: LLMMessage[]; final: string; verification?: ChatMessage["verification"] }>;

export type AIChatPanelTableOption = {
  name: string;
  description?: string;
  detail?: string;
};

export type AIChatPanelChartOption = {
  id: string;
  label: string;
  description?: string;
  detail?: string;
};

export type AIChatPanelAttachmentProps = {
  /**
   * Optional providers for building spreadsheet context attachments.
   *
   * These callbacks keep the panel UI decoupled from the desktop SpreadsheetApp
   * implementation while still letting callers wire in selection/table/etc.
   */
  getSelectionAttachment?: () => Attachment | null;
  getFormulaAttachment?: () => Attachment | null;
  getTableOptions?: () => AIChatPanelTableOption[];
  getChartOptions?: () => AIChatPanelChartOption[];
  getChartAttachment?: () => Attachment | null;
};

export type AIChatPanelProps =
  | (AIChatPanelAttachmentProps & {
      /**
       * When `sendMessage` is provided, the panel becomes a pure UI shell and
       * delegates orchestration (context, tools, audit, approvals) to the caller.
       */
      sendMessage: AIChatPanelSendMessage;
      systemPrompt?: string;
      onRequestToolApproval?: (call: ToolCall) => Promise<boolean>;
      client?: LLMClient;
      toolExecutor?: ToolExecutor;
    })
  | (AIChatPanelAttachmentProps & {
      /**
       * If `sendMessage` is omitted, the panel runs the provider-agnostic
       * `runChatWithTools` loop directly (legacy/demo mode).
       */
      sendMessage?: undefined;
      client: LLMClient;
      toolExecutor: ToolExecutor;
      systemPrompt?: string;
      onRequestToolApproval?: (call: ToolCall) => Promise<boolean>;
    });

export function AIChatPanel(props: AIChatPanelProps) {
  const [input, setInput] = useState("");
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [sending, setSending] = useState(false);
  // The attachment toolbar depends on external providers (selection, tables, charts) that
  // may change independently of React state. We use this reducer as a lightweight
  // `forceUpdate()` when the user interacts with the composer so disabled/enabled states
  // stay in sync without wiring the panel directly to the spreadsheet app.
  const [, refreshAttachmentToolbar] = React.useReducer((n: number) => n + 1, 0);
  const attachmentToolbarRefreshTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const abortControllerRef = useRef<AbortController | null>(null);
  const messageSeqRef = useRef(0);
  const scrollRef = useRef<HTMLDivElement>(null);
  // Provider-facing history is stored separately from UI messages. UI tool
  // entries created in `onToolCall` / `onToolResult` are for display only and
  // don't carry the `toolCallId` required by the LLM protocol.
  const [llmHistory, setLlmHistory] = useState<LLMMessage[]>([]);

  function messageId(): string {
    messageSeqRef.current += 1;
    const maybeCrypto = globalThis.crypto as Crypto | undefined;
    const base =
      maybeCrypto && typeof maybeCrypto.randomUUID === "function"
        ? maybeCrypto.randomUUID()
        : `msg-${Date.now()}-${Math.round(Math.random() * 1e9)}`;
    return `${base}-${messageSeqRef.current}`;
  }

  const systemPrompt = useMemo(
    () =>
      props.systemPrompt ??
      "You are an AI assistant inside a spreadsheet app. Prefer using tools to read data before making claims.",
    [props.systemPrompt],
  );

  function safeInvoke<T>(fn: (() => T) | undefined): T | null {
    if (!fn) return null;
    try {
      return fn();
    } catch {
      return null;
    }
  }

  const selectionAttachmentPreview = safeInvoke(props.getSelectionAttachment);
  const formulaAttachmentPreview = safeInvoke(props.getFormulaAttachment);
  const chartAttachmentPreview = safeInvoke(props.getChartAttachment);
  const tableOptionsPreview = safeInvoke(props.getTableOptions) ?? [];
  const chartOptionsPreview = safeInvoke(props.getChartOptions) ?? [];
  const chartDisabledReason =
    chartOptionsPreview.length > 0
      ? null
      : props.getChartOptions
        ? t("chat.attachChart.disabled.noCharts")
        : t("chat.attachChart.disabled.noSelection");

  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    el.scrollTop = el.scrollHeight;
  }, [messages]);

  useEffect(() => {
    return () => {
      abortControllerRef.current?.abort();
      if (attachmentToolbarRefreshTimer.current != null) {
        globalThis.clearTimeout(attachmentToolbarRefreshTimer.current);
        attachmentToolbarRefreshTimer.current = null;
      }
    };
  }, []);

  function scheduleAttachmentToolbarRefresh() {
    if (attachmentToolbarRefreshTimer.current != null) return;
    // Defer refresh so we don't re-render in the middle of a click/focus sequence.
    attachmentToolbarRefreshTimer.current = globalThis.setTimeout(() => {
      attachmentToolbarRefreshTimer.current = null;
      refreshAttachmentToolbar();
    }, 0);
  }

  function isAbortError(err: unknown): boolean {
    if (!err || (typeof err !== "object" && typeof err !== "function")) return false;
    const name = (err as any).name;
    const message = (err as any).message;
    return name === "AbortError" || /aborted/i.test(String(message ?? "")) || /abort/i.test(String(message ?? ""));
  }

  function formatToolResult(result: unknown): string {
    const rendered = safeStringify(result, { pretty: true });
    const limit = 2_000;
    if (rendered.length <= limit) return rendered;
    return `${rendered.slice(0, limit)}\n…(truncated)`;
  }

  function addAttachment(next: Attachment) {
    setAttachments((prev) => {
      if (prev.some((a) => a.type === next.type && a.reference === next.reference)) return prev;
      return [...prev, next];
    });
  }

  function toastBestEffort(message: string, type: "info" | "warning" | "error" = "info") {
    try {
      showToast(message, type);
    } catch {
      // `showToast` requires a #toast-root; unit tests don't always include it.
    }
  }

  function removeAttachmentAt(index: number) {
    setAttachments((prev) => prev.filter((_a, i) => i !== index));
  }

  async function attachTable() {
    if (sending) return;
    const tables = safeInvoke(props.getTableOptions) ?? [];
    if (!tables.length) {
      toastBestEffort(t("chat.attachTable.disabled"));
      return;
    }
    const picked = await showQuickPick(
      tables.map((t) => ({
        label: t.name,
        value: t.name,
        description: t.description,
        detail: t.detail,
      })),
      { placeHolder: t("chat.attachTable.placeholder") },
    );
    if (!picked) return;
    addAttachment({ type: "table", reference: picked });
  }

  async function attachChart() {
    if (sending) return;
    // Prefer attaching an already-selected chart when the host app can provide one.
    // This matches "attach selected chart" UX, while still allowing users to pick
    // from all known charts when nothing is selected.
    const selected = safeInvoke(props.getChartAttachment);
    if (selected) {
      addAttachment(selected);
      return;
    }

    const charts = safeInvoke(props.getChartOptions) ?? [];
    if (!charts.length) {
      const reason = props.getChartOptions
        ? t("chat.attachChart.disabled.noCharts")
        : t("chat.attachChart.disabled.noSelection");
      toastBestEffort(reason);
      return;
    }

    const picked = await showQuickPick(
      charts.map((c) => ({
        label: c.label,
        value: c.id,
        description: c.description,
        detail: c.detail,
      })),
      { placeHolder: t("chat.attachChart.placeholder") },
    );
    if (!picked) return;
    addAttachment({ type: "chart", reference: picked });
  }

  async function send() {
    if (sending) return;
    const text = input.trim();
    if (!text) return;
    setSending(true);

    abortControllerRef.current?.abort();
    const abortController = new AbortController();
    abortControllerRef.current = abortController;

    const needsTools = classifyQueryNeedsTools({ userText: text, attachments });
    const executedToolCalls: Array<{ name: string; ok?: boolean }> = [];

    const userMsg: ChatMessage = {
      id: messageId(),
      role: "user",
      content: text,
      attachments,
    };

    setMessages((prev) => [...prev, userMsg]);
    setInput("");
    setAttachments([]);

    const userContent =
      attachments.length > 0 ? `${text}\n\nAttachments:\n${formatAttachmentsForPrompt(attachments)}` : text;
    const base = llmHistory.length ? llmHistory : [{ role: "system", content: systemPrompt } satisfies LLMMessage];
    const requestMessages: LLMMessage[] = [...base, { role: "user", content: userContent }];

    try {
      setMessages((prev) => [...prev, { id: messageId(), role: "assistant", content: "", pending: true }]);

      const onToolCall = (call: ToolCall, meta: { requiresApproval: boolean }) => {
        setMessages((prev) => [
          ...insertBeforePendingAssistant(prev, {
            id: messageId(),
            role: "tool",
            content: `${call.name}(${safeStringify(call.arguments)})`,
            requiresApproval: meta.requiresApproval,
          }),
        ]);
      };

      const onToolResult = (call: ToolCall, result: unknown) => {
        // Many tool executors return domain objects without an explicit `{ ok }` field.
        // If the tool loop reached `onToolResult`, we treat that as a success signal
        // unless the result explicitly reports `{ ok: false }`.
        const ok = typeof (result as any)?.ok === "boolean" ? (result as any).ok : true;
        executedToolCalls.push({ name: call.name, ok });
        setMessages((prev) => [
          ...insertBeforePendingAssistant(prev, {
            id: messageId(),
            role: "tool",
            content: `${call.name} result:\n${formatToolResult(result)}`,
          }),
        ]);
      };

      // When a streaming call triggers tool execution, some models may still
      // emit additional text blocks after the tool call (e.g. multi-block
      // streaming payloads). We only want to stream the *final* assistant answer
      // to the UI, so once a tool call begins we suppress further text deltas
      // until the current model stream ends (`done`), then resume streaming for
      // the next iteration.
      let suppressTextDeltasForCurrentStream = false;

      const onStreamEvent = (event: ChatStreamEvent) => {
        if (event.type === "done") {
          suppressTextDeltasForCurrentStream = false;
          return;
        }

        if (event.type === "tool_call_start" || event.type === "tool_call_delta") {
          suppressTextDeltasForCurrentStream = true;
          // We only display the final assistant answer. Clear any pre-tool chatter so the
          // pending assistant message doesn't briefly show "planning" text that will be
          // replaced after tool execution.
          setMessages((prev) => {
            const next = prev.slice();
            const msg = [...next].reverse().find((m) => m.role === "assistant" && m.pending);
            if (msg && msg.pending) msg.content = "";
            return next;
          });
          return;
        }

        if (event.type !== "text") return;
        if (!event.delta) return;
        if (suppressTextDeltasForCurrentStream) return;
        setMessages((prev) => {
          const next = prev.slice();
          const msg = [...next].reverse().find((m) => m.role === "assistant" && m.pending);
          if (msg && msg.pending) msg.content += event.delta;
          return next;
        });
      };

      // NOTE: `props` is a discriminated union. Keep narrowing outside of closures
      // so TypeScript doesn't lose the refinement when capturing variables.
      let runner: AIChatPanelSendMessage;
      if (props.sendMessage) {
        runner = props.sendMessage;
      } else {
        const { client, toolExecutor } = props;
        const requireApproval = props.onRequestToolApproval ?? (async () => true);
        runner = async ({ messages, onToolCall, onToolResult, onStreamEvent, signal }: AIChatPanelSendMessageArgs) =>
          runChatWithToolsStreaming({
            client,
            toolExecutor,
            messages,
            signal,
            onStreamEvent,
            onToolCall,
            onToolResult,
            requireApproval,
          });
      }

      const result = await runner({
        messages: requestMessages,
        userText: text,
        attachments,
        signal: abortController.signal,
        onStreamEvent,
        onToolCall,
        onToolResult,
      });
      setLlmHistory(result.messages);

      const baseVerification = verifyToolUsage({ needsTools, toolCalls: executedToolCalls });
      let verification = result.verification ?? baseVerification;

      // In legacy/demo mode (no external orchestrator), we can optionally run a
      // post-response verification pass to check numeric claims against
      // spreadsheet computations.
      if (!result.verification && !props.sendMessage) {
        const claims = await verifyAssistantClaims({
          assistantText: result.final,
          userText: text,
          attachments,
          toolCalls: [],
          toolExecutor: props.toolExecutor as any
        });

        if (claims) {
          const warnings = [...claims.warnings];
          if (needsTools && !baseVerification.used_tools) warnings.unshift("Model did not use tools for a data question.");
          verification = {
            needs_tools: baseVerification.needs_tools,
            used_tools: baseVerification.used_tools,
            verified: claims.verified,
            confidence: claims.confidence,
            warnings,
            claims: claims.claims
          };
        }
      }

      setMessages((prev) => {
        const next = prev.slice();
        const lastAssistant = [...next].reverse().find((m) => m.role === "assistant" && m.pending);
        if (lastAssistant) {
          lastAssistant.content = result.final;
          lastAssistant.pending = false;
          lastAssistant.verification = verification;
        } else {
          next.push({ id: messageId(), role: "assistant", content: result.final, verification });
        }
        return next;
      });
    } catch (err) {
      const message = isAbortError(err) ? t("chat.cancelled") : err instanceof Error ? err.message : String(err);
      const content = isAbortError(err) ? message : tWithVars("chat.errorWithMessage", { message });
      setMessages((prev) => {
        const next = prev.slice();
        const lastAssistant = [...next].reverse().find((m) => m.role === "assistant" && m.pending);
        if (lastAssistant) {
          lastAssistant.content = content;
          lastAssistant.pending = false;
          return next;
        }
        next.push({ id: messageId(), role: "assistant", content });
        return next;
      });
    } finally {
      setSending(false);
      if (abortControllerRef.current === abortController) {
        abortControllerRef.current = null;
      }
    }
  }

  function insertBeforePendingAssistant(prev: ChatMessage[], message: ChatMessage): ChatMessage[] {
    const next = prev.slice();
    const pendingIndex = next.findIndex((m) => m.role === "assistant" && m.pending);
    if (pendingIndex === -1) {
      next.push(message);
    } else {
      next.splice(pendingIndex, 0, message);
    }
    return next;
  }

  return (
    <div className="ai-chat-panel">
      <div className="ai-chat-panel__header">{t("chat.title")}</div>
      <div ref={scrollRef} className="ai-chat-panel__messages">
        {messages.map((m) => (
          <div key={m.id} className="ai-chat-panel__message">
            <div className="ai-chat-panel__meta">
              {m.role === "user"
                ? t("chat.role.user")
                : m.role === "assistant"
                  ? t("chat.role.assistant")
                  : t("chat.role.tool")}
              {m.pending ? t("chat.meta.thinking") : ""}
              {m.requiresApproval ? t("chat.meta.requiresApproval") : ""}
            </div>
            {m.role === "assistant" &&
            m.verification &&
            (!m.verification.verified || (m.verification.confidence ?? 0) < CONFIDENCE_WARNING_THRESHOLD) ? (
              <div className="ai-chat-panel__unverified">
                {t("chat.meta.unverifiedAnswer")}
              </div>
            ) : null}
            <div className="ai-chat-panel__message-content">{m.content}</div>
            {m.attachments?.length ? (
              <div className="ai-chat-panel__attachments">
                {t("chat.attachmentsLabel")}
                <ul>
                  {m.attachments.map((a, i) => (
                    <li key={i}>
                      {a.type}: {a.reference}
                    </li>
                  ))}
                </ul>
              </div>
            ) : null}
          </div>
        ))}
      </div>
      <div
        className="ai-chat-panel__composer"
        onMouseEnter={() => scheduleAttachmentToolbarRefresh()}
        onFocusCapture={() => scheduleAttachmentToolbarRefresh()}
      >
        <div className="ai-chat-panel__composer-toolbar">
          <div className="ai-chat-panel__attachment-toolbar" role="toolbar" aria-label="Attach context">
            <AttachmentButton
              testId="ai-chat-attach-selection"
              disabled={sending || !selectionAttachmentPreview}
              title={!selectionAttachmentPreview ? t("chat.attachSelection.disabled") : undefined}
              onClick={() => {
                const attachment = safeInvoke(props.getSelectionAttachment);
                if (!attachment) {
                  toastBestEffort(t("chat.attachSelection.disabled"));
                  return;
                }
                addAttachment(attachment);
                const clampInfo = (attachment as any)?.data?.clamped;
                if (clampInfo && typeof clampInfo === "object") {
                  toastBestEffort(
                    tWithVars("chat.attachSelection.clamped", {
                      original: (clampInfo as any).originalCellCount ?? (clampInfo as any).original ?? "",
                      attached: (clampInfo as any).attachedCellCount ?? (clampInfo as any).attached ?? "",
                      max: (clampInfo as any).maxCells ?? (clampInfo as any).max ?? "",
                    }),
                    "warning",
                  );
                }
              }}
            >
              {t("chat.attachSelection")}
            </AttachmentButton>

            <AttachmentButton
              testId="ai-chat-attach-table"
              disabled={sending || tableOptionsPreview.length === 0}
              title={tableOptionsPreview.length === 0 ? t("chat.attachTable.disabled") : undefined}
              onClick={() => void attachTable()}
            >
              {t("chat.attachTable")}
            </AttachmentButton>

            {props.getChartOptions || props.getChartAttachment ? (
              <AttachmentButton
                testId="ai-chat-attach-chart"
                disabled={sending || (chartOptionsPreview.length === 0 && !chartAttachmentPreview)}
                title={chartOptionsPreview.length === 0 && !chartAttachmentPreview ? chartDisabledReason ?? undefined : undefined}
                onClick={() => void attachChart()}
              >
                {t("chat.attachChart")}
              </AttachmentButton>
            ) : null}
            {props.getFormulaAttachment ? (
              <button
                type="button"
                className="ai-chat-panel__attachment-button"
                data-testid="ai-chat-attach-formula"
                disabled={sending}
                title={!formulaAttachmentPreview ? t("chat.attachFormula.disabled") : undefined}
                onClick={() => {
                  const attachment = safeInvoke(props.getFormulaAttachment);
                  if (!attachment) {
                    toastBestEffort(t("chat.attachFormula.disabled"));
                    return;
                  }
                  addAttachment(attachment);
                }}
              >
                {t("chat.attachFormula")}
              </button>
            ) : null}
          </div>
          {attachments.length ? (
            <div className="ai-chat-panel__pending-attachments" data-testid="ai-chat-pending-attachments">
              <span className="ai-chat-panel__pending-attachments-label">{t("chat.pendingAttachments")}</span>
              <div className="ai-chat-panel__pending-attachments-chips">
                {attachments.map((a, idx) => (
                  <span key={`${a.type}:${a.reference}:${idx}`} className="ai-chat-panel__attachment-chip" data-testid={`ai-chat-attachment-chip-${idx}`}>
                    <span className="ai-chat-panel__attachment-chip-label">
                      {a.type}: {a.reference}
                    </span>
                    <button
                      type="button"
                      className="ai-chat-panel__attachment-chip-remove"
                      aria-label={`Remove ${a.type} attachment`}
                      data-testid={`ai-chat-attachment-remove-${idx}`}
                      onClick={() => removeAttachmentAt(idx)}
                    >
                      ×
                    </button>
                  </span>
                ))}
              </div>
            </div>
          ) : null}
        </div>
        <div className="ai-chat-panel__composer-row">
          <input
            className="ai-chat-panel__input"
            placeholder={t("chat.input.placeholder")}
            value={input}
            disabled={sending}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                void send();
              }
            }}
          />
          <button onClick={() => void send()} className="ai-chat-panel__button" disabled={sending}>
            {t("chat.send")}
          </button>
          <button
            onClick={() => abortControllerRef.current?.abort()}
            className="ai-chat-panel__button"
            disabled={!sending}
            type="button"
          >
            {t("chat.cancel")}
          </button>
        </div>
      </div>
    </div>
  );
}

function AttachmentButton(props: {
  testId: string;
  disabled?: boolean;
  title?: string;
  onClick: () => void;
  children: React.ReactNode;
}) {
  const button = (
    <button
      type="button"
      className="ai-chat-panel__attachment-button"
      data-testid={props.testId}
      disabled={props.disabled}
      title={!props.disabled ? props.title : undefined}
      onClick={props.onClick}
    >
      {props.children}
    </button>
  );

  // Disabled buttons do not reliably show `title` tooltips in browsers. Wrap in a
  // non-disabled element so the tooltip is still available to mouse users.
  if (props.disabled && props.title) {
    return (
      <span className="ai-chat-panel__attachment-button-wrap" title={props.title}>
        {button}
      </span>
    );
  }

  return button;
}
