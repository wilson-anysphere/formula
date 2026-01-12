import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

import type { LLMMessage } from "../../../../../packages/llm/src/types.js";

import { createAiChatOrchestrator } from "../../ai/chat/orchestrator.js";
import { runAgentTask, type AgentProgressEvent, type AgentTaskResult } from "../../ai/agent/agentOrchestrator.js";
import { createDesktopRagService } from "../../ai/rag/ragService.js";
import { createSchemaProviderFromSearchWorkbook } from "../../ai/context/searchWorkbookSchemaProvider.js";
import type { LLMToolCall } from "../../../../../packages/ai-tools/src/llm/integration.js";
import type { ToolPlanPreview } from "../../../../../packages/ai-tools/src/preview/preview-engine.js";
import type { SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.js";
import { getDesktopAIAuditStore } from "../../ai/audit/auditStore.js";
import { getDesktopLLMClient, getDesktopModel, purgeLegacyDesktopLLMSettings } from "../../ai/llm/desktopLLMClient.js";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver.js";

import { AIChatPanel, type AIChatPanelSendMessage } from "./AIChatPanel.js";
import { ApprovalModal } from "./ApprovalModal.js";
import { confirmPreviewApproval } from "./previewApproval.js";

function generateSessionId(): string {
  const maybeCrypto = globalThis.crypto as Crypto | undefined;
  if (maybeCrypto && typeof maybeCrypto.randomUUID === "function") return maybeCrypto.randomUUID();
  return `session-${Date.now()}-${Math.round(Math.random() * 1e9)}`;
}

export interface AIChatPanelContainerProps {
  getDocumentController: () => unknown;
  getActiveSheetId?: () => string;
  /**
   * Optional UI selection provider (0-based coordinates).
   *
   * When provided, chat includes the current selection as a sampled data block in
   * workbook context, so prompts like "summarize this selection" work without an
   * explicit attachment.
   */
  getSelection?: () => { sheetId: string; range: { startRow: number; startCol: number; endRow: number; endCol: number } } | null;
  /**
   * Optional workbook metadata provider (defined names / tables) used by other
   * desktop features like the name box and formula tab completion.
   *
   * When provided, chat/agent can include this metadata in workbook context.
   */
  getSearchWorkbook?: () => unknown;
  sheetNameResolver?: SheetNameResolver | null;
  workbookId?: string;
  createChart?: SpreadsheetApi["createChart"];
}

export function AIChatPanelContainer(props: AIChatPanelContainerProps) {
  useEffect(() => {
    purgeLegacyDesktopLLMSettings();
  }, []);
  return <AIChatPanelRuntime {...props} />;
}

function AIChatPanelRuntime(props: AIChatPanelContainerProps) {
  const [tab, setTab] = useState<"chat" | "agent">("chat");

  const sessionId = useRef<string>(generateSessionId());
  const llmHistory = useRef<LLMMessage[] | undefined>(undefined);

  const [approvalRequest, setApprovalRequest] = useState<{ call: LLMToolCall; preview: ToolPlanPreview } | null>(null);
  const approvalResolver = useRef<((approved: boolean) => void) | null>(null);

  const onApprovalRequired = useCallback(async (request: { call: LLMToolCall; preview: ToolPlanPreview }) => {
    // If we're not in a browser DOM environment, fall back to a native dialog helper.
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

  const client = useMemo(() => getDesktopLLMClient(), []);
  const model = useMemo(() => getDesktopModel(), []);

  const workbookId = props.workbookId ?? "local-workbook";
  const auditStore = useMemo(() => getDesktopAIAuditStore(), []);

  const documentController = useMemo(() => props.getDocumentController() as any, [props.getDocumentController]);
  const schemaProvider = useMemo(() => {
    try {
      const wb = props.getSearchWorkbook?.();
      return wb ? createSchemaProviderFromSearchWorkbook(wb as any) : null;
    } catch {
      return null;
    }
  }, [props.getSearchWorkbook]);

  const ragService = useMemo(() => {
    return createDesktopRagService({
      documentController,
      workbookId,
      tokenBudgetTokens: 8_000,
      topK: 6,
      sampleRows: 6,
      embedder: { type: "hash", dimension: 384 },
    });
  }, [documentController, workbookId]);

  useEffect(() => {
    return () => {
      void ragService.dispose();
    };
  }, [ragService]);

  const orchestrator = useMemo(() => {
    return createAiChatOrchestrator({
      documentController,
      workbookId,
      llmClient: client as any,
      model,
      sheetNameResolver: props.sheetNameResolver ?? null,
      getActiveSheetId: props.getActiveSheetId,
      getSelectedRange: props.getSelection as any,
      schemaProvider,
      createChart: props.createChart,
      auditStore,
      onApprovalRequired,
      previewOptions: { approval_cell_threshold: 0 },
      sessionId: `${workbookId}:${sessionId.current}`,
      ragService,
    });
  }, [
    auditStore,
    client,
    documentController,
    model,
    onApprovalRequired,
    props.createChart,
    props.getActiveSheetId,
    props.getSelection,
    props.sheetNameResolver,
    schemaProvider,
    ragService,
    workbookId,
  ]);

  useEffect(() => {
    return () => {
      void orchestrator.dispose();
    };
  }, [orchestrator]);

  const sendMessage: AIChatPanelSendMessage = useMemo(() => {
    return async (args) => {
      let abortListener: (() => void) | null = null;
      if (args.signal) {
        abortListener = () => {
          // Also resolve any pending approval prompt so the UI doesn't get stuck
          // if cancellation happens while waiting for user approval.
          resolveApproval(false);
        };
        args.signal.addEventListener("abort", abortListener, { once: true });
      }

      try {
        const result = await orchestrator.sendMessage({
          text: args.userText,
          attachments: args.attachments as any,
          history: llmHistory.current,
          onToolCall: args.onToolCall as any,
          onToolResult: args.onToolResult as any,
          onStreamEvent: args.onStreamEvent as any,
          signal: args.signal,
        });

        llmHistory.current = stripSystemPrompt(result.messages);
        return { messages: result.messages, final: result.finalText, verification: result.verification as any };
      } finally {
        if (abortListener && args.signal) {
          args.signal.removeEventListener("abort", abortListener);
        }
      }
    };
  }, [orchestrator, resolveApproval]);

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
        documentController,
        sheetNameResolver: props.sheetNameResolver ?? null,
        llmClient: client as any,
        auditStore,
        createChart: props.createChart,
        schemaProvider,
        onProgress: (event) =>
          setAgentEvents((prev) => {
            const last = prev.at(-1);
            if (event.type === "assistant_message" && last?.type === "assistant_message" && last.iteration === event.iteration) {
              const next = prev.slice();
              next[next.length - 1] = event;
              return next;
            }
            return [...prev, event];
          }),
        onApprovalRequired: onApprovalRequired as any,
        ragService,
        continueOnApprovalDenied: agentContinueOnDenied,
        maxIterations: 20,
        maxDurationMs: 5 * 60 * 1000,
        signal: controller.signal,
        model,
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
    documentController,
    model,
    onApprovalRequired,
    props,
    ragService,
    schemaProvider,
    workbookId,
  ]);

  return (
    <div className="ai-chat-runtime">
      <div className="ai-chat-runtime__tabs">
        <TabButton active={tab === "chat"} onClick={() => setTab("chat")} testId="ai-tab-chat">
          Chat
        </TabButton>
        <TabButton active={tab === "agent"} onClick={() => setTab("agent")} testId="ai-tab-agent">
          Agent
        </TabButton>
        <div className="ai-chat-runtime__tabs-spacer" />
      </div>
      <div className="ai-chat-runtime__content">
        {tab === "chat" ? (
          <AIChatPanel sendMessage={sendMessage} />
        ) : (
          <div className="ai-chat-agent">
            <div className="ai-chat-agent__section">
              <label className="ai-chat-agent__label">Goal</label>
              <textarea
                value={agentGoal}
                onChange={(e) => setAgentGoal(e.target.value)}
                placeholder="e.g. Summarize the data in Sheet1 and add a chart."
                rows={3}
                className="ai-chat-agent__textarea"
                data-testid="agent-goal"
                disabled={agentRunning}
              />
            </div>
            <div className="ai-chat-agent__section">
              <label className="ai-chat-agent__label">Constraints (one per line, optional)</label>
              <textarea
                value={agentConstraints}
                onChange={(e) => setAgentConstraints(e.target.value)}
                placeholder="e.g. Don’t overwrite existing data."
                rows={2}
                className="ai-chat-agent__textarea"
                data-testid="agent-constraints"
                disabled={agentRunning}
              />
            </div>
            <div className="ai-chat-agent__buttons">
              <button type="button" onClick={() => void runAgent()} disabled={agentRunning || !agentGoal.trim()} data-testid="agent-run">
                {agentRunning ? "Running…" : "Run"}
              </button>
              <button type="button" onClick={cancelAgent} disabled={!agentRunning} data-testid="agent-cancel">
                Cancel
              </button>
            </div>
            <label className={agentRunning ? "ai-chat-agent__continue ai-chat-agent__continue--disabled" : "ai-chat-agent__continue"}>
              <input
                type="checkbox"
                checked={agentContinueOnDenied}
                onChange={(e) => setAgentContinueOnDenied(e.target.checked)}
                disabled={agentRunning}
                data-testid="agent-continue-on-denied"
              />
              Continue running if I deny an approval (agent will re-plan)
            </label>
            <div ref={agentStepsRootRef} className="ai-chat-agent__steps">
              <div className="ai-chat-agent__steps-title">Steps</div>
              {agentEvents.length === 0 ? (
                <div className="ai-chat-agent__steps-empty">No steps yet.</div>
              ) : (
                <ol className="ai-chat-agent__steps-list">
                  {agentEvents.map((event, idx) => (
                    <li key={idx} className="ai-chat-agent__steps-item">
                      <AgentEventRow event={event} />
                    </li>
                  ))}
                </ol>
              )}
              {agentResult ? (
                <div className="ai-chat-agent__result" data-testid="agent-result">
                  <div className="ai-chat-agent__result-title">Result</div>
                  <div className="ai-chat-agent__result-meta">
                    Status: <code>{agentResult.status}</code> • session_id: <code>{agentResult.session_id}</code>
                  </div>
                  {agentResult.final ? (
                    <pre className="ai-chat-agent__result-pre">{agentResult.final}</pre>
                  ) : agentResult.error ? (
                    <pre className="ai-chat-agent__result-pre">{agentResult.error}</pre>
                  ) : null}
                </div>
              ) : null}
            </div>
          </div>
        )}
      </div>
      {approvalRequest ? (
        <ApprovalModal request={approvalRequest} onApprove={() => resolveApproval(true)} onReject={() => resolveApproval(false)} />
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
      className={props.active ? "ai-chat-tab-button ai-chat-tab-button--active" : "ai-chat-tab-button"}
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
          Result: <code>{event.call.name}</code> • {event.ok === undefined ? "done" : event.ok ? "ok" : "error"}
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
