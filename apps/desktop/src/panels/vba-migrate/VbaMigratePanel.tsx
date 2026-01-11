import React, { useCallback, useEffect, useMemo, useState } from "react";

import { OpenAIClient } from "../../../../../packages/llm/src/openai.js";
import { analyzeVbaModule } from "../../../../../packages/vba-migrate/src/analyzer.js";
import { VbaMigrator } from "../../../../../packages/vba-migrate/src/converter.js";

import { getVbaProject, type VbaModuleSummary, type VbaProjectSummary } from "../../macros/vba_project.js";

const API_KEY_STORAGE_KEY = "formula:openaiApiKey";

function loadApiKeyFromRuntime(): string | null {
  try {
    const stored = globalThis.localStorage?.getItem(API_KEY_STORAGE_KEY);
    if (stored) return stored;
  } catch {
    // ignore
  }

  const envKey = (import.meta as any)?.env?.VITE_OPENAI_API_KEY;
  if (typeof envKey === "string" && envKey.length > 0) return envKey;

  return null;
}

function moduleDisplayName(project: VbaProjectSummary, module: VbaModuleSummary) {
  const suffix = module.moduleType ? ` (${module.moduleType})` : "";
  const name = module.name || "Unnamed module";
  if (project.modules.length <= 1) return `${name}${suffix}`;
  return `${name}${suffix}`;
}

type AnalysisReport = ReturnType<typeof analyzeVbaModule>;

function sumCounts(obj: Record<string, unknown[] | undefined>) {
  const out: Record<string, number> = {};
  for (const [key, value] of Object.entries(obj)) out[key] = Array.isArray(value) ? value.length : 0;
  return out;
}

function AggregateAnalysisView(props: { report: AnalysisReport | null }) {
  const report = props.report;
  if (!report) {
    return <div style={{ opacity: 0.8 }}>Select a module to analyze.</div>;
  }

  const usage = sumCounts(report.objectModelUsage as any);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <div data-testid="vba-analysis-risk">
        <div style={{ fontWeight: 600 }}>Risk</div>
        <div>
          Score: <span style={{ fontFamily: "monospace" }}>{report.risk?.score ?? "?"}</span>{" "}
          <span style={{ fontFamily: "monospace" }}>({report.risk?.level ?? "unknown"})</span>
        </div>
      </div>

      <div>
        <div style={{ fontWeight: 600, marginBottom: 6 }}>Excel object model calls</div>
        <ul style={{ margin: 0, paddingLeft: 18 }}>
          {Object.entries(usage).map(([key, count]) => (
            <li key={key} data-testid={`vba-analysis-usage-${key}`}>
              {key}: <span style={{ fontFamily: "monospace" }}>{count}</span>
            </li>
          ))}
        </ul>
      </div>

      {report.externalReferences?.length ? (
        <div data-testid="vba-analysis-external">
          <div style={{ fontWeight: 600, marginBottom: 6 }}>External references</div>
          <ul style={{ margin: 0, paddingLeft: 18 }}>
            {report.externalReferences.map((ref: any, idx: number) => (
              <li key={`${ref.line}-${idx}`} style={{ fontFamily: "monospace", fontSize: 12 }}>
                L{ref.line}: {ref.text.trim()}
              </li>
            ))}
          </ul>
        </div>
      ) : null}

      {report.unsupportedConstructs?.length ? (
        <div data-testid="vba-analysis-unsupported">
          <div style={{ fontWeight: 600, marginBottom: 6 }}>Unsupported / risky constructs</div>
          <ul style={{ margin: 0, paddingLeft: 18 }}>
            {report.unsupportedConstructs.map((item: any, idx: number) => (
              <li key={`${item.line}-${idx}`} style={{ fontFamily: "monospace", fontSize: 12 }}>
                L{item.line}: {item.text.trim()}
              </li>
            ))}
          </ul>
        </div>
      ) : null}

      {report.unsafeConstructs?.length ? (
        <div data-testid="vba-analysis-unsafe">
          <div style={{ fontWeight: 600, marginBottom: 6 }}>Unsafe dynamic execution</div>
          <ul style={{ margin: 0, paddingLeft: 18 }}>
            {report.unsafeConstructs.map((item: any, idx: number) => (
              <li key={`${item.line}-${idx}`} style={{ fontFamily: "monospace", fontSize: 12 }}>
                L{item.line}: {item.text.trim()}
              </li>
            ))}
          </ul>
        </div>
      ) : null}

      {report.warnings?.length ? (
        <div data-testid="vba-analysis-warnings">
          <div style={{ fontWeight: 600, marginBottom: 6 }}>Warnings</div>
          <ul style={{ margin: 0, paddingLeft: 18 }}>
            {report.warnings.map((warning: any, idx: number) => (
              <li key={`${warning.line}-${idx}`} style={{ fontFamily: "monospace", fontSize: 12 }}>
                L{warning.line}: {warning.message}
              </li>
            ))}
          </ul>
        </div>
      ) : null}
    </div>
  );
}

export interface VbaMigratePanelProps {
  workbookId?: string;
  /**
   * Optional test hook to inject a migrator without requiring an API key.
   */
  createMigrator?: () => VbaMigrator;
}

export function VbaMigratePanel(props: VbaMigratePanelProps) {
  const [project, setProject] = useState<VbaProjectSummary | null>(null);
  const [loadingProject, setLoadingProject] = useState(true);
  const [projectError, setProjectError] = useState<string | null>(null);

  const [selectedModuleName, setSelectedModuleName] = useState<string | null>(null);
  const [analysisScope, setAnalysisScope] = useState<"module" | "project">("module");

  const [apiKey, setApiKey] = useState<string | null>(() => loadApiKeyFromRuntime());
  const [draftKey, setDraftKey] = useState("");

  const [conversionPrompt, setConversionPrompt] = useState("");
  const [conversionOutput, setConversionOutput] = useState("");
  const [conversionStatus, setConversionStatus] = useState<"idle" | "working" | "error">("idle");
  const [conversionError, setConversionError] = useState<string | null>(null);

  const migrator = useMemo(() => {
    if (props.createMigrator) return props.createMigrator();
    if (!apiKey) return null;
    try {
      return new VbaMigrator({ llm: new OpenAIClient({ apiKey }) });
    } catch {
      return null;
    }
  }, [apiKey, props.createMigrator]);

  const refreshProject = useCallback(async () => {
    setLoadingProject(true);
    setProjectError(null);
    try {
      const workbookId = props.workbookId ?? "local-workbook";
      const result = await getVbaProject(workbookId);
      setProject(result);
      setSelectedModuleName((prev) => {
        if (!result?.modules?.length) return null;
        if (prev && result.modules.some((m) => m.name === prev)) return prev;
        return result.modules[0]?.name ?? null;
      });
    } catch (err) {
      setProjectError(err instanceof Error ? err.message : String(err));
      setProject(null);
      setSelectedModuleName(null);
    } finally {
      setLoadingProject(false);
    }
  }, [props.workbookId]);

  useEffect(() => {
    void refreshProject();
  }, [refreshProject]);

  const selectedModule = useMemo(() => {
    if (!project || !selectedModuleName) return null;
    return project.modules.find((m) => m.name === selectedModuleName) ?? null;
  }, [project, selectedModuleName]);

  const moduleAnalysis = useMemo(() => {
    if (!selectedModule) return null;
    return analyzeVbaModule(selectedModule);
  }, [selectedModule]);

  const projectAnalysis = useMemo(() => {
    if (!project) return null;
    const modules = project.modules ?? [];
    if (modules.length === 0) return null;
    // Aggregate by summing findings across modules and using the max risk score.
    const reports = modules.map((m) => analyzeVbaModule(m));
    const aggregated: any = {
      moduleName: project.name ?? "VBA Project",
      objectModelUsage: {
        Range: [],
        Cells: [],
        Worksheets: [],
        Workbook: [],
        Application: [],
        ActiveSheet: [],
        ActiveCell: [],
        Selection: [],
      },
      externalReferences: [],
      unsafeConstructs: [],
      unsupportedConstructs: [],
      warnings: [],
      todos: [],
      risk: { score: 0, level: "low", factors: [] },
    };

    let maxScore = 0;
    for (const report of reports) {
      maxScore = Math.max(maxScore, report.risk?.score ?? 0);
      for (const key of Object.keys(aggregated.objectModelUsage)) {
        aggregated.objectModelUsage[key].push(...(report.objectModelUsage?.[key] ?? []));
      }
      aggregated.externalReferences.push(...(report.externalReferences ?? []));
      aggregated.unsafeConstructs.push(...(report.unsafeConstructs ?? []));
      aggregated.unsupportedConstructs.push(...(report.unsupportedConstructs ?? []));
      aggregated.warnings.push(...(report.warnings ?? []));
      aggregated.todos.push(...(report.todos ?? []));
      aggregated.risk.factors.push(...(report.risk?.factors ?? []));
    }
    aggregated.risk.score = maxScore;
    aggregated.risk.level = maxScore >= 70 ? "high" : maxScore >= 30 ? "medium" : "low";
    return aggregated as AnalysisReport;
  }, [project]);

  const analysis = analysisScope === "project" ? projectAnalysis : moduleAnalysis;

  const canConvert = Boolean(selectedModule && migrator);

  const onConvert = useCallback(
    async (target: "python" | "typescript") => {
      if (!selectedModule || !migrator) return;
      setConversionStatus("working");
      setConversionError(null);
      setConversionPrompt("");
      setConversionOutput("");

      try {
        const result = await migrator.convertModule(selectedModule, { target });
        setConversionPrompt(result.prompt);
        setConversionOutput(result.code);
        setConversionStatus("idle");
      } catch (err) {
        setConversionStatus("error");
        setConversionError(err instanceof Error ? err.message : String(err));
      }
    },
    [migrator, selectedModule]
  );

  async function copyToClipboard(text: string) {
    const clipboard = (globalThis.navigator as any)?.clipboard;
    if (!clipboard || typeof clipboard.writeText !== "function") return;
    try {
      await clipboard.writeText(text);
    } catch {
      // ignore
    }
  }

  const sidebar = (
    <div
      style={{
        width: 280,
        display: "flex",
        flexDirection: "column",
        borderRight: "1px solid var(--panel-border)",
        padding: 12,
        gap: 12,
        boxSizing: "border-box",
        overflow: "auto",
      }}
    >
      <div>
        <div style={{ fontWeight: 600 }}>VBA project</div>
        {loadingProject ? (
          <div style={{ opacity: 0.8 }}>Loading…</div>
        ) : project ? (
          <div data-testid="vba-project-name" style={{ fontFamily: "monospace", fontSize: 12, opacity: 0.9 }}>
            {project.name ?? "(unnamed project)"}
          </div>
        ) : (
          <div style={{ opacity: 0.8 }}>No VBA project found.</div>
        )}
      </div>

      <div style={{ display: "flex", gap: 8 }}>
        <button type="button" onClick={() => void refreshProject()} disabled={loadingProject} data-testid="vba-refresh">
          Refresh
        </button>
      </div>

      {projectError ? (
        <div style={{ color: "var(--error)", fontSize: 12 }} data-testid="vba-project-error">
          {projectError}
        </div>
      ) : null}

      {project?.constants ? (
        <details>
          <summary style={{ cursor: "pointer", fontWeight: 600 }}>Constants</summary>
          <pre style={{ whiteSpace: "pre-wrap", fontSize: 12, margin: "8px 0 0" }}>{project.constants}</pre>
        </details>
      ) : null}

      {project?.references?.length ? (
        <details>
          <summary style={{ cursor: "pointer", fontWeight: 600 }}>References</summary>
          <ul style={{ margin: "8px 0 0", paddingLeft: 18, fontSize: 12 }}>
            {project.references.map((ref, idx) => (
              <li key={`${ref.raw}-${idx}`} style={{ fontFamily: "monospace" }}>
                {ref.name ?? ref.raw}
              </li>
            ))}
          </ul>
        </details>
      ) : null}

      {project?.modules?.length ? (
        <div data-testid="vba-module-list">
          <div style={{ fontWeight: 600, marginBottom: 8 }}>Modules</div>
          <ul style={{ listStyle: "none", padding: 0, margin: 0 }}>
            {project.modules.map((mod) => {
              const selected = mod.name === selectedModuleName;
              return (
                <li key={mod.name} style={{ marginBottom: 6 }}>
                  <button
                    type="button"
                    onClick={() => setSelectedModuleName(mod.name)}
                    data-testid={`vba-module-${mod.name}`}
                    style={{
                      width: "100%",
                      textAlign: "left",
                      padding: "6px 8px",
                      background: selected ? "var(--selection-bg)" : "transparent",
                      border: "1px solid var(--panel-border)",
                      cursor: "pointer",
                      color: "inherit",
                      borderRadius: 6,
                      fontSize: 12,
                      fontFamily: "monospace",
                    }}
                    title={moduleDisplayName(project, mod)}
                  >
                    {mod.name}
                  </button>
                </li>
              );
            })}
          </ul>
        </div>
      ) : null}
    </div>
  );

  if (!loadingProject && !project) {
    const title = projectError ? "Unable to load VBA project" : "No macros to migrate";
    const detail = projectError
      ? "The VBA project could not be loaded (Tauri backend unavailable or workbook could not be parsed)."
      : "This workbook does not contain a vbaProject.bin payload.";
    return (
      <div style={{ display: "flex", height: "100%" }}>
        {sidebar}
        <div style={{ padding: 16, flex: 1 }}>
          <div style={{ fontWeight: 600, marginBottom: 8 }}>{title}</div>
          <div style={{ opacity: 0.8 }}>{detail}</div>
        </div>
      </div>
    );
  }

  const monospace = {
    fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
    fontSize: 12,
  } as const;

  return (
    <div style={{ display: "flex", height: "100%", minHeight: 0 }}>
      {sidebar}

      <div style={{ flex: 1, minHeight: 0, padding: 12, boxSizing: "border-box", display: "flex", gap: 12 }}>
        <div
          style={{
            flex: 1,
            minHeight: 0,
            border: "1px solid var(--panel-border)",
            borderRadius: 8,
            display: "flex",
            flexDirection: "column",
          }}
        >
          <div style={{ padding: 10, borderBottom: "1px solid var(--panel-border)", fontWeight: 600 }}>VBA module</div>
          <textarea
            readOnly
            value={selectedModule?.code ?? ""}
            data-testid="vba-module-code"
            style={{
              ...monospace,
              flex: 1,
              minHeight: 0,
              border: "none",
              outline: "none",
              padding: 10,
              resize: "none",
              background: "transparent",
              color: "inherit",
            }}
          />
        </div>

        <div
          style={{
            flex: 1,
            minHeight: 0,
            border: "1px solid var(--panel-border)",
            borderRadius: 8,
            display: "flex",
            flexDirection: "column",
          }}
        >
          <div
            style={{
              padding: 10,
              borderBottom: "1px solid var(--panel-border)",
              display: "flex",
              gap: 8,
              alignItems: "center",
            }}
          >
            <div style={{ fontWeight: 600, flex: 1 }}>Conversion</div>
            <button
              type="button"
              data-testid="vba-convert-python"
              onClick={() => void onConvert("python")}
              disabled={!canConvert || conversionStatus === "working"}
            >
              Convert to Python
            </button>
            <button
              type="button"
              data-testid="vba-convert-typescript"
              onClick={() => void onConvert("typescript")}
              disabled={!canConvert || conversionStatus === "working"}
            >
              Convert to TypeScript
            </button>
            <button
              type="button"
              onClick={() => void copyToClipboard(conversionOutput)}
              disabled={!conversionOutput}
              data-testid="vba-copy-converted"
            >
              Copy
            </button>
          </div>

          {!props.createMigrator && !apiKey ? (
            <div style={{ padding: 10, borderBottom: "1px solid var(--panel-border)", display: "flex", gap: 8 }}>
              <input
                value={draftKey}
                placeholder="Enter OpenAI API key to enable conversion"
                onChange={(e) => setDraftKey(e.target.value)}
                style={{ ...monospace, flex: 1, padding: 8 }}
                data-testid="vba-openai-key"
              />
              <button
                type="button"
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
                data-testid="vba-save-openai-key"
              >
                Save
              </button>
            </div>
          ) : null}

          {conversionStatus === "working" ? (
            <div style={{ padding: 10, borderBottom: "1px solid var(--panel-border)", fontSize: 12, opacity: 0.8 }}>
              Converting…
            </div>
          ) : null}

          {conversionError ? (
            <div
              style={{ padding: 10, borderBottom: "1px solid var(--panel-border)", fontSize: 12, color: "var(--error)" }}
              data-testid="vba-conversion-error"
            >
              {conversionError}
            </div>
          ) : null}

          <textarea
            readOnly
            value={conversionOutput}
            data-testid="vba-converted-code"
            style={{
              ...monospace,
              flex: 1,
              minHeight: 0,
              border: "none",
              outline: "none",
              padding: 10,
              resize: "none",
              background: "transparent",
              color: "inherit",
            }}
          />

          <details style={{ borderTop: "1px solid var(--panel-border)" }}>
            <summary style={{ cursor: "pointer", padding: 10, fontWeight: 600 }}>Prompt</summary>
            <pre
              style={{
                ...monospace,
                whiteSpace: "pre-wrap",
                margin: 0,
                padding: 10,
                borderTop: "1px solid var(--panel-border)",
              }}
              data-testid="vba-conversion-prompt"
            >
              {conversionPrompt || "(no prompt yet)"}
            </pre>
          </details>
        </div>

        <div
          style={{
            flex: 1,
            minHeight: 0,
            border: "1px solid var(--panel-border)",
            borderRadius: 8,
            display: "flex",
            flexDirection: "column",
            overflow: "hidden",
          }}
        >
          <div
            style={{
              padding: 10,
              borderBottom: "1px solid var(--panel-border)",
              display: "flex",
              gap: 8,
              alignItems: "center",
            }}
          >
            <div style={{ fontWeight: 600, flex: 1 }}>Analysis</div>
            <select
              value={analysisScope}
              onChange={(e) => setAnalysisScope(e.target.value as any)}
              style={{ fontSize: 12 }}
              data-testid="vba-analysis-scope"
            >
              <option value="module">Selected module</option>
              <option value="project">Entire project</option>
            </select>
          </div>

          <div style={{ flex: 1, minHeight: 0, overflow: "auto", padding: 10 }}>
            <AggregateAnalysisView report={analysis} />
          </div>
        </div>
      </div>
    </div>
  );
}
