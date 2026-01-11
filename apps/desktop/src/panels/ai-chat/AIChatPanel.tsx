import React, { useMemo, useState } from "react";

import type { Attachment, ChatMessage } from "./types.js";
import type { LLMClient, LLMMessage, ToolCall, ToolExecutor } from "../../../../../packages/llm/src/types.js";
import { runChatWithTools } from "../../../../../packages/llm/src/toolCalling.js";
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

  const systemPrompt = useMemo(
    () =>
      props.systemPrompt ??
      "You are an AI assistant inside a spreadsheet app. Prefer using tools to read data before making claims.",
    [props.systemPrompt],
  );

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

    const userMsg: ChatMessage = {
      id: crypto.randomUUID(),
      role: "user",
      content: text,
      attachments,
    };

    setMessages((prev) => [...prev, userMsg]);
    setInput("");
    setAttachments([]);

    const llmMessages = [
      { role: "system" as const, content: systemPrompt },
      ...messages.map((m) => ({ role: m.role === "tool" ? ("tool" as const) : (m.role as any), content: m.content })),
      {
        role: "user" as const,
        content:
          attachments.length > 0 ? `${text}\n\nAttachments:\n${formatAttachmentsForPrompt(attachments)}` : text,
      },
    ];

    try {
      setMessages((prev) => [
        ...prev,
        { id: crypto.randomUUID(), role: "assistant", content: "", pending: true },
      ]);

      const onToolCall = (call: ToolCall, meta: { requiresApproval: boolean }) => {
        setMessages((prev) => [
          ...prev,
          {
            id: crypto.randomUUID(),
            role: "tool",
            content: `${call.name}(${JSON.stringify(call.arguments)})`,
            requiresApproval: meta.requiresApproval,
          },
        ]);
      };

      const onToolResult = (call: ToolCall, result: unknown) => {
        setMessages((prev) => [
          ...prev,
          {
            id: crypto.randomUUID(),
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
            messages: messages as any,
            onToolCall,
            onToolResult,
            requireApproval: props.onRequestToolApproval ?? (async () => true),
          }) as any);

      const result = await runner({
        messages: llmMessages as any,
        userText: text,
        attachments,
        onToolCall,
        onToolResult,
      });

      setMessages((prev) => {
        const next = prev.slice();
        const lastAssistant = [...next].reverse().find((m) => m.role === "assistant" && m.pending);
        if (lastAssistant) {
          lastAssistant.content = result.final;
          lastAssistant.pending = false;
        } else {
          next.push({ id: crypto.randomUUID(), role: "assistant", content: result.final });
        }
        return next;
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setMessages((prev) => [
        ...prev,
        { id: crypto.randomUUID(), role: "assistant", content: tWithVars("chat.errorWithMessage", { message }) },
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
      <div style={{ flex: 1, overflow: "auto", padding: 12 }}>
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
