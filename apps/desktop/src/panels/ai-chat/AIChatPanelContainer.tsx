import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

import type { LLMMessage } from "../../../../../packages/llm/src/types.js";
import { OpenAIClient } from "../../../../../packages/llm/src/openai.js";

import { LocalStorageAIAuditStore } from "../../../../../packages/ai-audit/src/local-storage-store.js";

import { createAiChatOrchestrator } from "../../ai/chat/orchestrator.js";
import { runAgentTask, type AgentProgressEvent, type AgentTaskResult } from "../../ai/agent/agentOrchestrator.js";
import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import { HashEmbedder } from "../../../../../packages/ai-rag/src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../../../../../packages/ai-rag/src/store/inMemoryVectorStore.js";
import type { LLMToolCall } from "../../../../../packages/ai-tools/src/llm/integration.js";
import type { ToolPlanPreview } from "../../../../../packages/ai-tools/src/preview/preview-engine.js";
import type { SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.js";

import { AIChatPanel, type AIChatPanelSendMessage } from "./AIChatPanel.js";
import { ApprovalModal } from "./ApprovalModal.js";
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
  createChart?: SpreadsheetApi["createChart"];
}

export function AIChatPanelContainer(props: AIChatPanelContainerProps) {
  const [apiKey, setApiKey] = useState<string | null>(() => loadApiKeyFromRuntime());
  const [draftKey, setDraftKey] = useState("");

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
              setDraftKey("");
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

  return <AIChatPanelRuntime {...props} apiKey={apiKey} />;
}

function AIChatPanelRuntime(props: AIChatPanelContainerProps & { apiKey: string }) {
  const [tab, setTab] = useState<"chat" | "agent">("chat");

  const sessionId = useRef<string>(generateSessionId());
  const llmHistory = useRef<LLMMessage[] | undefined>(undefined);

  const [approvalRequest, setApprovalRequest] = useState<{ call: LLMToolCall; preview: ToolPlanPreview } | null>(null);
  const approvalResolver = useRef<((approved: boolean) => void) | null>(null);

  const onApprovalRequired = useCallback(async (request: { call: LLMToolCall; preview: ToolPlanPreview }) => {
    // If we're not in a browser DOM environment, fall back to `window.confirm`.
    if (typeof document === "undefined") return confirmPreviewApproval(request as any);

    if (approvalResolver.current) {
      // Shouldn't happen because tool calls are sequential, but be safe.
      return false;
    }

    setApprovalRequest(request);
    return new Promise<boolean>((resolve) => {
      approvalResolver.current = resolve;
    });
  }, []);

  const resolveApproval = useCallback((approved: boolean) => {
    const resolve = approvalResolver.current;
    approvalResolver.current = null;
    setApprovalRequest(null);
    resolve?.(approved);
  }, []);

  useEffect(() => {
    return () => {
      // Ensure we never leave an awaited approval promise hanging on unmount.
      if (approvalResolver.current) {
        approvalResolver.current(false);
        approvalResolver.current = null;
      }
    };
  }, []);

  const client = useMemo(() => new OpenAIClient({ apiKey: props.apiKey }), [props.apiKey]);

  const workbookId = props.workbookId ?? "local-workbook";
  const auditStore = useMemo(() => new LocalStorageAIAuditStore(), []);

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
      createChart: props.createChart,
      auditStore,
      onApprovalRequired,
      previewOptions: { approval_cell_threshold: 0 },
      sessionId: `${workbookId}:${sessionId.current}`,
      contextManager,
    });
  }, [
    auditStore,
    client,
    contextManager,
    onApprovalRequired,
    props.createChart,
    props.getActiveSheetId,
    props.getDocumentController,
    workbookId,
  ]);

  const sendMessage: AIChatPanelSendMessage = useMemo(() => {
    return async (args) => {
      const result = await orchestrator.sendMessage({
        text: args.userText,
        attachments: args.attachments as any,
        history: llmHistory.current,
        onToolCall: args.onToolCall as any,
        onToolResult: args.onToolResult as any,
      });

      llmHistory.current = stripSystemPrompt(result.messages);
      return { messages: result.messages, final: result.finalText };
    };
  }, [orchestrator]);

  const [agentGoal, setAgentGoal] = useState("");
  const [agentConstraints, setAgentConstraints] = useState("");
  const [agentContinueOnDenied, setAgentContinueOnDenied] = useState(false);
  const [agentEvents, setAgentEvents] = useState<AgentProgressEvent[]>([]);
  const [agentResult, setAgentResult] = useState<AgentTaskResult | null>(null);
  const [agentRunning, setAgentRunning] = useState(false);
  const agentStepsRootRef = useRef<HTMLDivElement | null>(null);
  const abortControllerRef = useRef<AbortController | null>(null);

  const cancelAgent = useCallback(() => {
    abortControllerRef.current?.abort();
    // Also resolve any pending approval prompt to avoid a stuck UI.
    resolveApproval(false);
  }, [resolveApproval]);

  useEffect(() => {
    return () => cancelAgent();
  }, [cancelAgent]);

  useEffect(() => {
    const root = agentStepsRootRef.current;
    if (!root) return;
    // Auto-scroll while running, but don't yank the user if they've intentionally scrolled up.
    const distanceFromBottom = root.scrollHeight - root.scrollTop - root.clientHeight;
    const nearBottom = distanceFromBottom < 60;
    if (agentRunning && nearBottom) root.scrollTop = root.scrollHeight;
  }, [agentRunning, agentEvents.length]);

  const runAgent = useCallback(async () => {
    if (agentRunning) return;

    const goal = agentGoal.trim();
    if (!goal) return;

    abortControllerRef.current?.abort();
    const controller = new AbortController();
    abortControllerRef.current = controller;

    setAgentRunning(true);
    setAgentEvents([]);
    setAgentResult(null);

    const constraints = agentConstraints
      .split("\n")
      .map((c) => c.trim())
      .filter(Boolean);
    const defaultSheetId = props.getActiveSheetId?.() ?? "Sheet1";

    try {
      const result = await runAgentTask({
        goal,
        constraints: constraints.length ? constraints : undefined,
        workbookId,
        defaultSheetId,
        documentController: props.getDocumentController() as any,
        llmClient: client as any,
        auditStore,
        createChart: props.createChart,
        onProgress: (event) => setAgentEvents((prev) => [...prev, event]),
        onApprovalRequired: onApprovalRequired as any,
        continueOnApprovalDenied: agentContinueOnDenied,
        maxIterations: 20,
        maxDurationMs: 5 * 60 * 1000,
        signal: controller.signal,
        model: (client as any).model ?? "gpt-4o-mini"
      });
      setAgentResult(result);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setAgentResult({ status: "error", session_id: "unknown", error: message });
    } finally {
      abortControllerRef.current = null;
      setAgentRunning(false);
    }
  }, [
    agentConstraints,
    agentGoal,
    agentContinueOnDenied,
    agentRunning,
    auditStore,
    client,
    onApprovalRequired,
    props,
    workbookId
  ]);

  return (
    <div style={{ position: "relative", height: "100%", display: "flex", flexDirection: "column", minHeight: 0 }}>
      <div
        style={{
          display: "flex",
          gap: 6,
          padding: 8,
          borderBottom: "1px solid var(--border)",
          background: "var(--bg-secondary)"
        }}
      >
        <TabButton active={tab === "chat"} onClick={() => setTab("chat")} testId="ai-tab-chat">
          Chat
        </TabButton>
        <TabButton active={tab === "agent"} onClick={() => setTab("agent")} testId="ai-tab-agent">
          Agent
        </TabButton>
      </div>
      <div style={{ position: "relative", flex: 1, minHeight: 0 }}>
        {tab === "chat" ? (
          <AIChatPanel sendMessage={sendMessage} />
        ) : (
          <div style={{ padding: 12, display: "flex", flexDirection: "column", gap: 10, height: "100%", minHeight: 0 }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <label style={{ fontSize: 12, fontWeight: 600 }}>Goal</label>
              <textarea
                value={agentGoal}
                onChange={(e) => setAgentGoal(e.target.value)}
                placeholder="e.g. Summarize the data in Sheet1 and add a chart."
                rows={3}
                style={{ padding: 8, resize: "vertical" }}
                data-testid="agent-goal"
                disabled={agentRunning}
              />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <label style={{ fontSize: 12, fontWeight: 600 }}>Constraints (one per line, optional)</label>
              <textarea
                value={agentConstraints}
                onChange={(e) => setAgentConstraints(e.target.value)}
                placeholder="e.g. Don’t overwrite existing data."
                rows={2}
                style={{ padding: 8, resize: "vertical" }}
                data-testid="agent-constraints"
                disabled={agentRunning}
              />
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button type="button" onClick={() => void runAgent()} disabled={agentRunning || !agentGoal.trim()} data-testid="agent-run">
                {agentRunning ? "Running…" : "Run"}
              </button>
              <button type="button" onClick={cancelAgent} disabled={!agentRunning} data-testid="agent-cancel">
                Cancel
              </button>
            </div>
            <label style={{ display: "flex", gap: 6, alignItems: "center", fontSize: 12, opacity: agentRunning ? 0.6 : 1 }}>
              <input
                type="checkbox"
                checked={agentContinueOnDenied}
                onChange={(e) => setAgentContinueOnDenied(e.target.checked)}
                disabled={agentRunning}
                data-testid="agent-continue-on-denied"
              />
              Continue running if I deny an approval (agent will re-plan)
            </label>
            <div
              ref={agentStepsRootRef}
              style={{ borderTop: "1px solid var(--border)", paddingTop: 10, minHeight: 0, flex: 1, overflow: "auto" }}
            >
              <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 8 }}>Steps</div>
              {agentEvents.length === 0 ? (
                <div style={{ fontSize: 12, opacity: 0.8 }}>No steps yet.</div>
              ) : (
                <ol style={{ display: "flex", flexDirection: "column", gap: 8, paddingLeft: 18, margin: 0 }}>
                  {agentEvents.map((event, idx) => (
                    <li key={idx} style={{ fontSize: 12 }}>
                      <AgentEventRow event={event} />
                    </li>
                  ))}
                </ol>
              )}
              {agentResult ? (
                <div style={{ marginTop: 12, paddingTop: 10, borderTop: "1px solid var(--border)" }} data-testid="agent-result">
                  <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 4 }}>Result</div>
                  <div style={{ fontSize: 12, opacity: 0.85 }}>
                    Status: <code>{agentResult.status}</code> • session_id: <code>{agentResult.session_id}</code>
                  </div>
                  {agentResult.final ? (
                    <pre style={{ whiteSpace: "pre-wrap", fontSize: 12, marginTop: 8 }}>{agentResult.final}</pre>
                  ) : agentResult.error ? (
                    <pre style={{ whiteSpace: "pre-wrap", fontSize: 12, marginTop: 8 }}>{agentResult.error}</pre>
                  ) : null}
                </div>
              ) : null}
            </div>
          </div>
        )}
      </div>
      {approvalRequest ? (
        <ApprovalModal
          request={approvalRequest}
          onApprove={() => resolveApproval(true)}
          onReject={() => resolveApproval(false)}
        />
      ) : null}
    </div>
  );
}

function stripSystemPrompt(messages: LLMMessage[]): LLMMessage[] {
  if (messages[0]?.role === "system") return messages.slice(1);
  return messages;
}

function TabButton(props: { active: boolean; onClick: () => void; children: React.ReactNode; testId?: string }) {
  return (
    <button
      type="button"
      onClick={props.onClick}
      data-testid={props.testId}
      style={{
        padding: "6px 10px",
        borderRadius: 8,
        border: "1px solid var(--border)",
        background: props.active ? "var(--bg-tertiary)" : "transparent",
        color: "var(--text-primary)",
        fontWeight: props.active ? 600 : 500
      }}
    >
      {props.children}
    </button>
  );
}

function AgentEventRow({ event }: { event: AgentProgressEvent }) {
  switch (event.type) {
    case "planning":
      return <span>Planning (iteration {event.iteration})</span>;
    case "tool_call":
      return (
        <span>
          Tool: <code>{event.call.name}</code>
          {event.requiresApproval ? " (approval gated)" : null}
        </span>
      );
    case "tool_result":
      return (
        <span>
          Result: <code>{event.call.name}</code> •{" "}
          {event.ok === undefined ? "done" : event.ok ? "ok" : "error"}
          {event.error ? ` (${event.error})` : null}
        </span>
      );
    case "assistant_message":
      return <span>Assistant: {event.content}</span>;
    case "complete":
      return <span>Complete</span>;
    case "cancelled":
      return (
        <span>
          Cancelled ({event.reason})
          {event.message ? `: ${event.message}` : null}
        </span>
      );
    case "error":
      return <span>Error: {event.message}</span>;
    default:
      return <span>Unknown event</span>;
  }
}
