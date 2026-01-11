import React, { useEffect, useMemo, useRef, useState } from "react";

import type { Attachment, ChatMessage } from "./types.js";
import type { LLMClient, LLMMessage, ToolCall, ToolExecutor } from "../../../../../packages/llm/src/types.js";
import { runChatWithTools } from "../../../../../packages/llm/src/toolCalling.js";
import { classifyQueryNeedsTools, verifyToolUsage } from "../../../../../packages/ai-tools/src/llm/verification.js";
import { t, tWithVars } from "../../i18n/index.js";

function formatAttachmentsForPrompt(attachments: Attachment[]) {
  return attachments
    .map((a) => `- ${a.type}: ${a.reference}${a.data ? ` (${JSON.stringify(a.data)})` : ""}`)
    .join("\n");
}

export interface AIChatPanelProps {
  client: LLMClient;
  toolExecutor: ToolExecutor;
  systemPrompt?: string;
  onRequestToolApproval?: (call: ToolCall) => Promise<boolean>;
  sendMessage?: AIChatPanelSendMessage;
}

export interface AIChatPanelSendMessageArgs {
  messages: LLMMessage[];
  userText: string;
  attachments: Attachment[];
  onToolCall: (call: ToolCall, meta: { requiresApproval: boolean }) => void;
  onToolResult?: (call: ToolCall, result: unknown) => void;
}

export type AIChatPanelSendMessage = (args: AIChatPanelSendMessageArgs) => Promise<{ messages: LLMMessage[]; final: string }>;

export function AIChatPanel(props: AIChatPanelProps) {
  const [input, setInput] = useState("");
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [sending, setSending] = useState(false);
  const scrollRef = useRef<HTMLDivElement>(null);
  // Provider-facing history is stored separately from UI messages. UI tool
  // entries created in `onToolCall` / `onToolResult` are for display only and
  // don't carry the `toolCallId` required by the LLM protocol.
  const [llmHistory, setLlmHistory] = useState<LLMMessage[]>([]);

  function messageId(): string {
    const maybeCrypto = globalThis.crypto as Crypto | undefined;
    if (maybeCrypto && typeof maybeCrypto.randomUUID === "function") return maybeCrypto.randomUUID();
    return `msg-${Date.now()}-${Math.round(Math.random() * 1e9)}`;
  }

  const systemPrompt = useMemo(
    () =>
      props.systemPrompt ??
      "You are an AI assistant inside a spreadsheet app. Prefer using tools to read data before making claims.",
    [props.systemPrompt],
  );

  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    el.scrollTop = el.scrollHeight;
  }, [messages]);

  function safeStringify(value: unknown): string {
    if (typeof value === "string") return value;
    try {
      return JSON.stringify(value, null, 2);
    } catch {
      return String(value);
    }
  }

  function formatToolResult(result: unknown): string {
    const rendered = safeStringify(result);
    const limit = 2_000;
    if (rendered.length <= limit) return rendered;
    return `${rendered.slice(0, limit)}\n…(truncated)`;
  }

  async function send() {
    if (sending) return;
    const text = input.trim();
    if (!text) return;
    setSending(true);

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
      setMessages((prev) => [
        ...prev,
        { id: messageId(), role: "assistant", content: "", pending: true },
      ]);

      const onToolCall = (call: ToolCall, meta: { requiresApproval: boolean }) => {
        setMessages((prev) => [
          ...prev,
          {
            id: messageId(),
            role: "tool",
            content: `${call.name}(${JSON.stringify(call.arguments)})`,
            requiresApproval: meta.requiresApproval,
          },
        ]);
      };

      const onToolResult = (call: ToolCall, result: unknown) => {
        executedToolCalls.push({ name: call.name, ok: typeof (result as any)?.ok === "boolean" ? (result as any).ok : undefined });
        setMessages((prev) => [
          ...prev,
          {
            id: messageId(),
            role: "tool",
            content: `${call.name} result:\n${formatToolResult(result)}`,
          },
        ]);
      };

      const runner =
        props.sendMessage ??
        (async ({ messages, onToolCall, onToolResult }: AIChatPanelSendMessageArgs) =>
          runChatWithTools({
            client: props.client,
            toolExecutor: props.toolExecutor,
            messages,
            onToolCall,
            onToolResult,
            requireApproval: props.onRequestToolApproval ?? (async () => true),
          }));

      const result = await runner({
        messages: requestMessages,
        userText: text,
        attachments,
        onToolCall,
        onToolResult,
      });
      setLlmHistory(result.messages);

      const verification = verifyToolUsage({ needsTools, toolCalls: executedToolCalls });

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
      const message = err instanceof Error ? err.message : String(err);
      setMessages((prev) => [
        ...prev,
        { id: messageId(), role: "assistant", content: tWithVars("chat.errorWithMessage", { message }) },
      ]);
    } finally {
      setSending(false);
    }
  }

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100%",
        borderInlineStart: "1px solid var(--border)",
      }}
    >
      <div style={{ padding: "8px 12px", borderBottom: "1px solid var(--border)", fontWeight: 600 }}>
        {t("chat.title")}
      </div>
      <div ref={scrollRef} style={{ flex: 1, overflow: "auto", padding: 12 }}>
        {messages.map((m) => (
          <div key={m.id} style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 12, opacity: 0.7 }}>
              {m.role === "user"
                ? t("chat.role.user")
                : m.role === "assistant"
                  ? t("chat.role.assistant")
                  : t("chat.role.tool")}
              {m.pending ? t("chat.meta.thinking") : ""}
              {m.requiresApproval ? t("chat.meta.requiresApproval") : ""}
            </div>
            {m.role === "assistant" && m.verification && !m.verification.verified ? (
              <div
                style={{
                  marginTop: 6,
                  padding: "6px 8px",
                  border: "1px solid var(--border)",
                  borderRadius: 6,
                  background: "var(--warning-bg)",
                  fontSize: 12,
                }}
              >
                {t("chat.meta.unverifiedAnswer")}
              </div>
            ) : null}
            <div style={{ whiteSpace: "pre-wrap" }}>{m.content}</div>
            {m.attachments?.length ? (
              <div style={{ marginTop: 6, fontSize: 12, opacity: 0.85 }}>
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
      {attachments.length ? (
        <div style={{ padding: "6px 12px", borderTop: "1px solid var(--border)", fontSize: 12 }}>
          {t("chat.pendingAttachments")}{" "}
          {attachments.map((a) => (
            <span key={`${a.type}:${a.reference}`} style={{ marginRight: 8 }}>
              {a.type}:{a.reference}
            </span>
          ))}
        </div>
      ) : null}
      <div style={{ display: "flex", gap: 8, padding: 12, borderTop: "1px solid var(--border)" }}>
        <input
          style={{ flex: 1, padding: 8 }}
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
        <button onClick={() => void send()} style={{ padding: "8px 12px" }} disabled={sending}>
          {t("chat.send")}
        </button>
      </div>
      <div
        style={{
          padding: "6px 12px",
          borderTop: "1px solid var(--border)",
          fontSize: 12,
          opacity: 0.7,
        }}
      >
        {t("chat.attachmentsApiPlaceholder")}
        <button
          style={{ marginInlineStart: 8 }}
          onClick={() =>
            setAttachments((prev) => [
              ...prev,
              { type: "range", reference: "Sheet1!A1:D10", data: { source: "selection" } },
            ])
          }
        >
          {t("chat.addRangeDemo")}
        </button>
      </div>
    </div>
  );
}
