import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

import type { LLMMessage } from "../../../../../packages/llm/src/types.js";
import { createLLMClient } from "../../../../../packages/llm/src/createLLMClient.js";

import { createAiChatOrchestrator } from "../../ai/chat/orchestrator.js";
import { runAgentTask, type AgentProgressEvent, type AgentTaskResult } from "../../ai/agent/agentOrchestrator.js";
import { createDesktopRagService } from "../../ai/rag/ragService.js";
import type { LLMToolCall } from "../../../../../packages/ai-tools/src/llm/integration.js";
import type { ToolPlanPreview } from "../../../../../packages/ai-tools/src/preview/preview-engine.js";
import type { SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.js";
import { getDesktopAIAuditStore } from "../../ai/audit/auditStore.js";

import { AIChatPanel, type AIChatPanelSendMessage } from "./AIChatPanel.js";
import { ApprovalModal } from "./ApprovalModal.js";
import { confirmPreviewApproval } from "./previewApproval.js";

import {
  DEFAULT_OLLAMA_BASE_URL,
  DEFAULT_OLLAMA_MODEL,
  LEGACY_OPENAI_API_KEY_STORAGE_KEY,
  LLM_PROVIDER_STORAGE_KEY,
  OLLAMA_BASE_URL_STORAGE_KEY,
  OLLAMA_MODEL_STORAGE_KEY,
  OPENAI_API_KEY_STORAGE_KEY,
  OPENAI_BASE_URL_STORAGE_KEY,
  OPENAI_MODEL_STORAGE_KEY,
  ANTHROPIC_API_KEY_STORAGE_KEY,
  ANTHROPIC_MODEL_STORAGE_KEY,
  clearDesktopLLMConfig,
  loadDesktopLLMConfig,
  migrateLegacyOpenAIKey,
  saveDesktopLLMConfig,
  type DesktopLLMConfig,
  type LLMProvider,
} from "../../ai/llm/settings.js";

function generateSessionId(): string {
  const maybeCrypto = globalThis.crypto as Crypto | undefined;
  if (maybeCrypto && typeof maybeCrypto.randomUUID === "function") return maybeCrypto.randomUUID();
  return `session-${Date.now()}-${Math.round(Math.random() * 1e9)}`;
}

function safeGetStorageItem(key: string): string | null {
  try {
    return globalThis.localStorage?.getItem(key) ?? null;
  } catch {
    return null;
  }
}

function loadProviderPreference(): LLMProvider {
  migrateLegacyOpenAIKey();
  const raw = safeGetStorageItem(LLM_PROVIDER_STORAGE_KEY);
  if (raw === "openai" || raw === "anthropic" || raw === "ollama") return raw;
  return "openai";
}

function loadOpenAIApiKeyDraft(): string {
  // Prefer new key, fall back to legacy and Vite env injection.
  const stored = safeGetStorageItem(OPENAI_API_KEY_STORAGE_KEY) ?? safeGetStorageItem(LEGACY_OPENAI_API_KEY_STORAGE_KEY);
  if (stored) return stored;
  const envKey = (import.meta as any)?.env?.VITE_OPENAI_API_KEY;
  return typeof envKey === "string" ? envKey : "";
}

function loadOpenAIBaseUrlDraft(): string {
  const stored = safeGetStorageItem(OPENAI_BASE_URL_STORAGE_KEY);
  if (stored) return stored;
  const envUrl = (import.meta as any)?.env?.VITE_OPENAI_BASE_URL;
  return typeof envUrl === "string" ? envUrl : "";
}

function loadOpenAIModelDraft(): string {
  return safeGetStorageItem(OPENAI_MODEL_STORAGE_KEY) ?? "";
}

function loadAnthropicApiKeyDraft(): string {
  const stored = safeGetStorageItem(ANTHROPIC_API_KEY_STORAGE_KEY);
  if (stored) return stored;
  const envKey = (import.meta as any)?.env?.VITE_ANTHROPIC_API_KEY;
  return typeof envKey === "string" ? envKey : "";
}

function loadAnthropicModelDraft(): string {
  return safeGetStorageItem(ANTHROPIC_MODEL_STORAGE_KEY) ?? "";
}

function loadOllamaBaseUrlDraft(): string {
  return safeGetStorageItem(OLLAMA_BASE_URL_STORAGE_KEY) ?? DEFAULT_OLLAMA_BASE_URL;
}

function loadOllamaModelDraft(): string {
  return safeGetStorageItem(OLLAMA_MODEL_STORAGE_KEY) ?? DEFAULT_OLLAMA_MODEL;
}

function normalizeBaseUrl(baseUrl: string): string {
  return baseUrl.trim().replace(/\/$/, "");
}

export interface AIChatPanelContainerProps {
  getDocumentController: () => unknown;
  getActiveSheetId?: () => string;
  workbookId?: string;
  createChart?: SpreadsheetApi["createChart"];
}

export function AIChatPanelContainer(props: AIChatPanelContainerProps) {
  const [config, setConfig] = useState<DesktopLLMConfig | null>(() => loadDesktopLLMConfig());
  const [editing, setEditing] = useState(() => config === null);

  const [provider, setProvider] = useState<LLMProvider>(() => loadProviderPreference());
  const [openaiApiKey, setOpenaiApiKey] = useState(() => loadOpenAIApiKeyDraft());
  const [openaiBaseUrl, setOpenaiBaseUrl] = useState(() => loadOpenAIBaseUrlDraft());
  const [openaiModel, setOpenaiModel] = useState(() => loadOpenAIModelDraft());
  const [anthropicApiKey, setAnthropicApiKey] = useState(() => loadAnthropicApiKeyDraft());
  const [anthropicModel, setAnthropicModel] = useState(() => loadAnthropicModelDraft());
  const [ollamaBaseUrl, setOllamaBaseUrl] = useState(() => loadOllamaBaseUrlDraft());
  const [ollamaModel, setOllamaModel] = useState(() => loadOllamaModelDraft());

  useEffect(() => {
    if (provider === "openai") {
      setOpenaiApiKey((prev) => prev || loadOpenAIApiKeyDraft());
      setOpenaiBaseUrl((prev) => prev || loadOpenAIBaseUrlDraft());
      setOpenaiModel((prev) => prev || loadOpenAIModelDraft());
      return;
    }
    if (provider === "anthropic") {
      setAnthropicApiKey((prev) => prev || loadAnthropicApiKeyDraft());
      setAnthropicModel((prev) => prev || loadAnthropicModelDraft());
      return;
    }
    setOllamaBaseUrl((prev) => prev || DEFAULT_OLLAMA_BASE_URL);
    setOllamaModel((prev) => prev || DEFAULT_OLLAMA_MODEL);
  }, [provider]);

  const onSave = useCallback(() => {
    if (provider === "openai") {
      const apiKey = openaiApiKey.trim();
      if (!apiKey) return;
      const baseUrl = normalizeBaseUrl(openaiBaseUrl);
      const model = openaiModel.trim() || undefined;
      const next: DesktopLLMConfig = {
        provider: "openai",
        apiKey,
        ...(model ? { model } : {}),
        ...(baseUrl ? { baseUrl } : {}),
      };
      saveDesktopLLMConfig(next);
      setConfig(next);
      setEditing(false);
      return;
    }

    if (provider === "anthropic") {
      const apiKey = anthropicApiKey.trim();
      if (!apiKey) return;
      const model = anthropicModel.trim() || undefined;
      const next: DesktopLLMConfig = { provider: "anthropic", apiKey, ...(model ? { model } : {}) };
      saveDesktopLLMConfig(next);
      setConfig(next);
      setEditing(false);
      return;
    }

    const baseUrl = normalizeBaseUrl(ollamaBaseUrl) || DEFAULT_OLLAMA_BASE_URL;
    const model = ollamaModel.trim() || DEFAULT_OLLAMA_MODEL;
    const next: DesktopLLMConfig = { provider: "ollama", baseUrl, model };
    saveDesktopLLMConfig(next);
    setConfig(next);
    setEditing(false);
  }, [
    anthropicApiKey,
    anthropicModel,
    ollamaBaseUrl,
    ollamaModel,
    openaiApiKey,
    openaiBaseUrl,
    openaiModel,
    provider,
  ]);

  const onClear = useCallback(() => {
    clearDesktopLLMConfig();
    setConfig(null);
    setEditing(true);
  }, []);

  if (config === null || editing) {
    return (
      <div style={{ padding: 12, display: "flex", flexDirection: "column", gap: 12 }}>
        <div style={{ fontWeight: 600 }}>AI chat setup</div>
        <div style={{ fontSize: 12, opacity: 0.8 }}>
          Choose a provider and enter credentials. Settings are stored in localStorage under <code>formula:llm:*</code>{" "}
          (OpenAI keys are also mirrored to <code>{LEGACY_OPENAI_API_KEY_STORAGE_KEY}</code> for backward
          compatibility).
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          <label style={{ fontSize: 12, fontWeight: 600 }}>Provider</label>
          <select
            value={provider}
            onChange={(e) => setProvider(e.target.value as LLMProvider)}
            style={{ padding: 8 }}
            data-testid="ai-provider-select"
          >
            <option value="openai">OpenAI</option>
            <option value="anthropic">Anthropic</option>
            <option value="ollama">Ollama (local)</option>
          </select>
        </div>

        {provider === "openai" ? (
          <>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <label style={{ fontSize: 12, fontWeight: 600 }}>OpenAI API key</label>
              <input
                value={openaiApiKey}
                placeholder="sk-..."
                onChange={(e) => setOpenaiApiKey(e.target.value)}
                style={{ padding: 8, fontFamily: "monospace" }}
                data-testid="ai-openai-api-key"
              />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <label style={{ fontSize: 12, fontWeight: 600 }}>OpenAI base URL (optional)</label>
              <input
                value={openaiBaseUrl}
                placeholder="https://api.openai.com/v1"
                onChange={(e) => setOpenaiBaseUrl(e.target.value)}
                style={{ padding: 8, fontFamily: "monospace" }}
                data-testid="ai-openai-base-url"
              />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <label style={{ fontSize: 12, fontWeight: 600 }}>Model (optional)</label>
              <input
                value={openaiModel}
                placeholder="gpt-4o-mini"
                onChange={(e) => setOpenaiModel(e.target.value)}
                style={{ padding: 8, fontFamily: "monospace" }}
                data-testid="ai-openai-model"
              />
            </div>
          </>
        ) : null}

        {provider === "anthropic" ? (
          <>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <label style={{ fontSize: 12, fontWeight: 600 }}>Anthropic API key</label>
              <input
                value={anthropicApiKey}
                placeholder="sk-ant-..."
                onChange={(e) => setAnthropicApiKey(e.target.value)}
                style={{ padding: 8, fontFamily: "monospace" }}
                data-testid="ai-anthropic-api-key"
              />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <label style={{ fontSize: 12, fontWeight: 600 }}>Model (optional)</label>
              <input
                value={anthropicModel}
                placeholder="claude-3-5-sonnet-latest"
                onChange={(e) => setAnthropicModel(e.target.value)}
                style={{ padding: 8, fontFamily: "monospace" }}
                data-testid="ai-anthropic-model"
              />
            </div>
          </>
        ) : null}

        {provider === "ollama" ? (
          <>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <label style={{ fontSize: 12, fontWeight: 600 }}>Ollama base URL</label>
              <input
                value={ollamaBaseUrl}
                placeholder={DEFAULT_OLLAMA_BASE_URL}
                onChange={(e) => setOllamaBaseUrl(e.target.value)}
                style={{ padding: 8, fontFamily: "monospace" }}
                data-testid="ai-ollama-base-url"
              />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <label style={{ fontSize: 12, fontWeight: 600 }}>Model</label>
              <input
                value={ollamaModel}
                placeholder={DEFAULT_OLLAMA_MODEL}
                onChange={(e) => setOllamaModel(e.target.value)}
                style={{ padding: 8, fontFamily: "monospace" }}
                data-testid="ai-ollama-model"
              />
            </div>
          </>
        ) : null}

        <div style={{ display: "flex", gap: 8 }}>
          <button type="button" style={{ padding: "8px 12px" }} onClick={onSave} data-testid="ai-save-settings">
            Save
          </button>
          <button type="button" style={{ padding: "8px 12px" }} onClick={onClear} data-testid="ai-clear-settings">
            Clear
          </button>
          {config !== null && editing ? (
            <button
              type="button"
              style={{ padding: "8px 12px", marginLeft: "auto" }}
              onClick={() => setEditing(false)}
              data-testid="ai-cancel-settings"
            >
              Cancel
            </button>
          ) : null}
        </div>
      </div>
    );
  }

  return <AIChatPanelRuntime {...props} llmConfig={config} onOpenSettings={() => setEditing(true)} />;
}

function AIChatPanelRuntime(
  props: AIChatPanelContainerProps & { llmConfig: DesktopLLMConfig; onOpenSettings: () => void },
) {
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

  const client = useMemo(() => createLLMClient(props.llmConfig as any), [props.llmConfig]);

  const workbookId = props.workbookId ?? "local-workbook";
  const auditStore = useMemo(() => getDesktopAIAuditStore(), []);

  const documentController = useMemo(() => props.getDocumentController() as any, [props.getDocumentController]);

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
      model: (client as any).model ?? "gpt-4o-mini",
      getActiveSheetId: props.getActiveSheetId,
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
    onApprovalRequired,
    props.createChart,
    props.getActiveSheetId,
    ragService,
    workbookId,
  ]);

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
        llmClient: client as any,
        auditStore,
        createChart: props.createChart,
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
    documentController,
    onApprovalRequired,
    props,
    ragService,
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
        <div style={{ flex: 1 }} />
        <button
          type="button"
          onClick={props.onOpenSettings}
          style={{
            padding: "6px 10px",
            borderRadius: 8,
            border: "1px solid var(--border)",
            background: "transparent",
            color: "var(--text-primary)",
            fontWeight: 500
          }}
          data-testid="ai-open-settings"
        >
          Settings
        </button>
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
