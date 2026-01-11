import React, { useMemo, useRef, useState } from "react";

import type { LLMMessage } from "../../../../../packages/llm/src/types.js";
import { OpenAIClient } from "../../../../../packages/llm/src/openai.js";

import { createAiChatOrchestrator } from "../../ai/chat/orchestrator.js";
import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import { HashEmbedder, InMemoryVectorStore } from "../../../../../packages/ai-rag/src/index.js";

import { AIChatPanel, type AIChatPanelSendMessage } from "./AIChatPanel.js";
import { confirmPreviewApproval } from "./previewApproval.js";

const API_KEY_STORAGE_KEY = "formula:openaiApiKey";

function generateSessionId(): string {
  const maybeCrypto = globalThis.crypto as Crypto | undefined;
  if (maybeCrypto && typeof maybeCrypto.randomUUID === "function") return maybeCrypto.randomUUID();
  return `session-${Date.now()}-${Math.round(Math.random() * 1e9)}`;
}

function loadApiKeyFromRuntime(): string | null {
  try {
    const stored = globalThis.localStorage?.getItem(API_KEY_STORAGE_KEY);
    if (stored) return stored;
  } catch {
    // ignore
  }

  // Allow Vite devs to inject a key without touching localStorage.
  const envKey = (import.meta as any)?.env?.VITE_OPENAI_API_KEY;
  if (typeof envKey === "string" && envKey.length > 0) return envKey;

  return null;
}

export interface AIChatPanelContainerProps {
  getDocumentController: () => unknown;
  getActiveSheetId?: () => string;
  workbookId?: string;
}

export function AIChatPanelContainer(props: AIChatPanelContainerProps) {
  const [apiKey, setApiKey] = useState<string | null>(() => loadApiKeyFromRuntime());
  const [draftKey, setDraftKey] = useState("");

  const sessionId = useRef<string>(generateSessionId());
  const llmHistory = useRef<LLMMessage[] | undefined>(undefined);

  if (!apiKey) {
    return (
      <div style={{ padding: 12, display: "flex", flexDirection: "column", gap: 12 }}>
        <div style={{ fontWeight: 600 }}>AI chat setup</div>
        <div style={{ fontSize: 12, opacity: 0.8 }}>
          Enter an OpenAI API key to enable chat. It will be stored in localStorage as <code>{API_KEY_STORAGE_KEY}</code>.
        </div>
        <input
          value={draftKey}
          placeholder="sk-..."
          onChange={(e) => setDraftKey(e.target.value)}
          style={{ padding: 8, fontFamily: "monospace" }}
        />
        <div style={{ display: "flex", gap: 8 }}>
          <button
            type="button"
            style={{ padding: "8px 12px" }}
            onClick={() => {
              const next = draftKey.trim();
              if (!next) return;
              try {
                globalThis.localStorage?.setItem(API_KEY_STORAGE_KEY, next);
              } catch {
                // ignore
              }
              setApiKey(next);
            }}
          >
            Save key
          </button>
          <button
            type="button"
            style={{ padding: "8px 12px" }}
            onClick={() => {
              setDraftKey("");
              try {
                globalThis.localStorage?.removeItem(API_KEY_STORAGE_KEY);
              } catch {
                // ignore
              }
            }}
          >
            Clear
          </button>
        </div>
      </div>
    );
  }

  const client = useMemo(() => new OpenAIClient({ apiKey }), [apiKey]);

  const workbookId = props.workbookId ?? "local-workbook";

  const contextManager = useMemo(() => {
    // Keep this lightweight + dependency-free for now (deterministic hash embeddings).
    const dimension = 128;
    const embedder = new HashEmbedder({ dimension });
    const vectorStore = new InMemoryVectorStore({ dimension });
    return new ContextManager({
      tokenBudgetTokens: 8_000,
      workbookRag: { vectorStore, embedder, topK: 6 },
    });
  }, [workbookId]);

  const orchestrator = useMemo(() => {
    return createAiChatOrchestrator({
      documentController: props.getDocumentController() as any,
      workbookId,
      llmClient: client as any,
      model: (client as any).model ?? "gpt-4o-mini",
      getActiveSheetId: props.getActiveSheetId,
      onApprovalRequired: confirmPreviewApproval,
      previewOptions: { approval_cell_threshold: 0 },
      sessionId: `${workbookId}:${sessionId.current}`,
      contextManager,
    });
  }, [client, contextManager, props.getActiveSheetId, props.getDocumentController, workbookId]);

  const sendMessage: AIChatPanelSendMessage = useMemo(() => {
    return async (args) => {
      const result = await orchestrator.sendMessage({
        text: args.userText,
        attachments: args.attachments as any,
        history: llmHistory.current,
        onToolCall: args.onToolCall as any,
      });

      llmHistory.current = stripSystemPrompt(result.messages);
      return { messages: result.messages, final: result.finalText };
    };
  }, [orchestrator]);

  return (
    <AIChatPanel
      client={client as any}
      toolExecutor={{ tools: [], execute: async () => ({ ok: false, error: { message: "Tool executor not initialized" } }) } as any}
      sendMessage={sendMessage}
    />
  );
}

function stripSystemPrompt(messages: LLMMessage[]): LLMMessage[] {
  if (messages[0]?.role === "system") return messages.slice(1);
  return messages;
}
