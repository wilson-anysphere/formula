import React, { useCallback, useEffect, useMemo, useState } from "react";

import { analyzeVbaModule } from "../../../../../packages/vba-migrate/src/analyzer.js";
import { VbaMigrator } from "../../../../../packages/vba-migrate/src/converter.js";

import { createClipboardProvider } from "../../clipboard/index.js";
import { getVbaProject, type VbaModuleSummary, type VbaProjectSummary } from "../../macros/vba_project.js";
import { getDesktopLLMClient, purgeLegacyDesktopLLMSettings } from "../../ai/llm/desktopLLMClient.js";

type TauriInvoke = (cmd: string, args?: any) => Promise<any>;

type MacroUiContext = {
  sheetId: string;
  activeRow: number;
  activeCol: number;
  selection?: { startRow: number; startCol: number; endRow: number; endCol: number } | null;
};

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

type ProjectAnalysisSummary = {
  kind: "project";
  projectName: string;
  risk: { score: number; level: "low" | "medium" | "high" };
  objectModelUsageCounts: Record<string, number>;
  rangeShapesCounts: Record<string, number>;
  externalReferencesCount: number;
  unsafeConstructsCount: number;
  unsupportedConstructsCount: number;
  warningsCount: number;
  todosCount: number;
  modules: Array<{
    moduleName: string;
    risk: { score: number; level: "low" | "medium" | "high" };
    externalReferencesCount: number;
    unsafeConstructsCount: number;
    unsupportedConstructsCount: number;
  }>;
};

type AnalysisViewModel = { kind: "module"; report: AnalysisReport } | ProjectAnalysisSummary;

function AggregateAnalysisView(props: { report: AnalysisViewModel | null }) {
  const model = props.report;
  if (!model) {
    return <div style={{ opacity: 0.8 }}>Select a module to analyze.</div>;
  }

  if (model.kind === "project") {
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <div data-testid="vba-analysis-risk">
          <div style={{ fontWeight: 600 }}>Project risk</div>
          <div>
            Score: <span style={{ fontFamily: "monospace" }}>{model.risk.score}</span>{" "}
            <span style={{ fontFamily: "monospace" }}>({model.risk.level})</span>
          </div>
        </div>

        <div>
          <div style={{ fontWeight: 600, marginBottom: 6 }}>Totals</div>
          <ul style={{ margin: 0, paddingLeft: 18 }}>
            <li>
              External references: <span style={{ fontFamily: "monospace" }}>{model.externalReferencesCount}</span>
            </li>
            <li>
              Unsafe constructs: <span style={{ fontFamily: "monospace" }}>{model.unsafeConstructsCount}</span>
            </li>
            <li>
              Unsupported constructs: <span style={{ fontFamily: "monospace" }}>{model.unsupportedConstructsCount}</span>
            </li>
            <li>
              Warnings: <span style={{ fontFamily: "monospace" }}>{model.warningsCount}</span>
            </li>
            <li>
              TODOs: <span style={{ fontFamily: "monospace" }}>{model.todosCount}</span>
            </li>
          </ul>
        </div>

        <details>
          <summary style={{ cursor: "pointer", fontWeight: 600 }}>Excel object model calls</summary>
          <ul style={{ margin: "8px 0 0", paddingLeft: 18 }}>
            {Object.entries(model.objectModelUsageCounts).map(([key, count]) => (
              <li key={key}>
                {key}: <span style={{ fontFamily: "monospace" }}>{count}</span>
              </li>
            ))}
          </ul>
        </details>

        <details>
          <summary style={{ cursor: "pointer", fontWeight: 600 }}>Range shapes</summary>
          <ul style={{ margin: "8px 0 0", paddingLeft: 18 }}>
            {Object.entries(model.rangeShapesCounts).map(([key, count]) => (
              <li key={key}>
                {key}: <span style={{ fontFamily: "monospace" }}>{count}</span>
              </li>
            ))}
          </ul>
        </details>

        <div>
          <div style={{ fontWeight: 600, marginBottom: 6 }}>Modules</div>
          <ul style={{ margin: 0, paddingLeft: 18 }}>
            {model.modules.map((m) => (
              <li key={m.moduleName} style={{ fontFamily: "monospace", fontSize: 12 }}>
                {m.moduleName}: risk {m.risk.score} ({m.risk.level}), external {m.externalReferencesCount}, unsafe{" "}
                {m.unsafeConstructsCount}, unsupported {m.unsupportedConstructsCount}
              </li>
            ))}
          </ul>
        </div>
      </div>
    );
  }

  const report = model.report;
  const usage = sumCounts(report.objectModelUsage as any);
  const rangeShapes = sumCounts(report.rangeShapes as any);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <div data-testid="vba-analysis-risk">
        <div style={{ fontWeight: 600 }}>Risk</div>
        <div>
          Score: <span style={{ fontFamily: "monospace" }}>{report.risk?.score ?? "?"}</span>{" "}
          <span style={{ fontFamily: "monospace" }}>({report.risk?.level ?? "unknown"})</span>
        </div>
      </div>

      <details open>
        <summary style={{ cursor: "pointer", fontWeight: 600 }}>Excel object model calls</summary>
        <ul style={{ margin: "8px 0 0", paddingLeft: 18 }}>
          {Object.entries(usage).map(([key, count]) => (
            <li key={key} data-testid={`vba-analysis-usage-${key}`}>
              {key}: <span style={{ fontFamily: "monospace" }}>{count}</span>
            </li>
          ))}
        </ul>
      </details>

      <details>
        <summary style={{ cursor: "pointer", fontWeight: 600 }}>Range shapes</summary>
        <ul style={{ margin: "8px 0 0", paddingLeft: 18 }}>
          {Object.entries(rangeShapes).map(([key, count]) => (
            <li key={key}>
              {key}: <span style={{ fontFamily: "monospace" }}>{count}</span>
            </li>
          ))}
        </ul>
      </details>

      {report.externalReferences?.length ? (
        <details open data-testid="vba-analysis-external">
          <summary style={{ cursor: "pointer", fontWeight: 600 }}>External references</summary>
          <ul style={{ margin: "8px 0 0", paddingLeft: 18 }}>
            {report.externalReferences.map((ref: any, idx: number) => (
              <li key={`${ref.line}-${idx}`} style={{ fontFamily: "monospace", fontSize: 12 }}>
                L{ref.line}: {ref.text.trim()}
              </li>
            ))}
          </ul>
        </details>
      ) : null}

      {report.unsupportedConstructs?.length ? (
        <details open data-testid="vba-analysis-unsupported">
          <summary style={{ cursor: "pointer", fontWeight: 600 }}>Unsupported / risky constructs</summary>
          <ul style={{ margin: "8px 0 0", paddingLeft: 18 }}>
            {report.unsupportedConstructs.map((item: any, idx: number) => (
              <li key={`${item.line}-${idx}`} style={{ fontFamily: "monospace", fontSize: 12 }}>
                L{item.line}: {item.text.trim()}
              </li>
            ))}
          </ul>
        </details>
      ) : null}

      {report.unsafeConstructs?.length ? (
        <details open data-testid="vba-analysis-unsafe">
          <summary style={{ cursor: "pointer", fontWeight: 600 }}>Unsafe dynamic execution</summary>
          <ul style={{ margin: "8px 0 0", paddingLeft: 18 }}>
            {report.unsafeConstructs.map((item: any, idx: number) => (
              <li key={`${item.line}-${idx}`} style={{ fontFamily: "monospace", fontSize: 12 }}>
                L{item.line}: {item.text.trim()}
              </li>
            ))}
          </ul>
        </details>
      ) : null}

      {report.warnings?.length ? (
        <details data-testid="vba-analysis-warnings">
          <summary style={{ cursor: "pointer", fontWeight: 600 }}>Warnings ({report.warnings.length})</summary>
          <ul style={{ margin: "8px 0 0", paddingLeft: 18 }}>
            {report.warnings.map((warning: any, idx: number) => (
              <li key={`${warning.line}-${idx}`} style={{ fontFamily: "monospace", fontSize: 12 }}>
                L{warning.line}: {warning.message}
              </li>
            ))}
          </ul>
        </details>
      ) : null}

      {report.todos?.length ? (
        <details data-testid="vba-analysis-todos">
          <summary style={{ cursor: "pointer", fontWeight: 600 }}>TODOs ({report.todos.length})</summary>
          <ul style={{ margin: "8px 0 0", paddingLeft: 18 }}>
            {report.todos.map((todo: any, idx: number) => (
              <li key={`${todo.line}-${idx}`} style={{ fontFamily: "monospace", fontSize: 12 }}>
                L{todo.line}: {todo.message}
              </li>
            ))}
          </ul>
        </details>
      ) : null}
    </div>
  );
}

export interface VbaMigratePanelProps {
  workbookId?: string;
  /**
   * Optional test hook to inject a migrator (e.g. deterministic unit tests).
   */
  createMigrator?: () => VbaMigrator;
  /**
   * Optional invoke wrapper (e.g. the queued invoke used by the desktop shell).
   */
  invoke?: TauriInvoke;
  /**
   * Optional hook to flush pending backend workbook sync operations before running
   * expensive validation commands.
   */
  drainBackendSync?: () => Promise<void>;
  /**
   * Optional callback returning the current macro UI context so validation runs
   * under the same `ActiveSheet` / `ActiveCell` / `Selection` the user sees.
   */
  getMacroUiContext?: () => MacroUiContext;
}

export function VbaMigratePanel(props: VbaMigratePanelProps) {
  const [project, setProject] = useState<VbaProjectSummary | null>(null);
  const [loadingProject, setLoadingProject] = useState(true);
  const [projectError, setProjectError] = useState<string | null>(null);
  const [availableMacros, setAvailableMacros] = useState<Array<{ id: string; name: string; module?: string | null }>>([]);

  const [selectedModuleName, setSelectedModuleName] = useState<string | null>(null);
  const [analysisScope, setAnalysisScope] = useState<"module" | "project">("module");

  const [entryPoint, setEntryPoint] = useState("Main");

  const [conversionTarget, setConversionTarget] = useState<"python" | "typescript" | null>(null);
  const [conversionPrompt, setConversionPrompt] = useState("");
  const [conversionOutput, setConversionOutput] = useState("");
  const [conversionStatus, setConversionStatus] = useState<"idle" | "working" | "error">("idle");
  const [conversionError, setConversionError] = useState<string | null>(null);

  const [validationStatus, setValidationStatus] = useState<"idle" | "working" | "error">("idle");
  const [validationError, setValidationError] = useState<string | null>(null);
  const [validationReport, setValidationReport] = useState<any>(null);

  useEffect(() => {
    purgeLegacyDesktopLLMSettings();
  }, []);

  const llmClient = useMemo(() => getDesktopLLMClient(), []);
  const clipboardProviderPromise = useMemo(() => createClipboardProvider(), []);

  const migrator = useMemo(() => {
    if (props.createMigrator) return props.createMigrator();
    try {
      return new VbaMigrator({ llm: llmClient as any });
    } catch {
      return null;
    }
  }, [llmClient, props.createMigrator]);

  const refreshProject = useCallback(async () => {
    setLoadingProject(true);
    setProjectError(null);
    try {
      const workbookId = props.workbookId ?? "local-workbook";
      const invoke = (globalThis as any).__TAURI__?.core?.invoke as ((cmd: string, args?: any) => Promise<any>) | undefined;
      const [result, macros] = await Promise.all([
        getVbaProject(workbookId),
        invoke
          ? invoke("list_macros", { workbook_id: workbookId }).catch(() => [])
          : Promise.resolve([]),
      ]);
      setProject(result);
      setAvailableMacros(
        Array.isArray(macros)
          ? macros
              .map((m: any) => ({
                id: String(m?.id ?? ""),
                name: String(m?.name ?? ""),
                module: m?.module != null ? String(m.module) : null,
              }))
              .filter((m: any) => m.id && m.name)
          : [],
      );
      setSelectedModuleName((prev) => {
        if (!result?.modules?.length) return null;
        if (prev && result.modules.some((m) => m.name === prev)) return prev;
        return result.modules[0]?.name ?? null;
      });
      setEntryPoint((prev) => {
        if (!Array.isArray(macros) || macros.length === 0) return prev;
        const ids = new Set(macros.map((m: any) => String(m?.id ?? "")).filter(Boolean));
        if (prev && ids.has(prev)) return prev;
        return String(macros[0]?.id ?? prev);
      });
    } catch (err) {
      setProjectError(err instanceof Error ? err.message : String(err));
      setProject(null);
      setSelectedModuleName(null);
      setAvailableMacros([]);
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

  const projectSummary = useMemo<ProjectAnalysisSummary | null>(() => {
    if (!project) return null;
    const modules = project.modules ?? [];
    if (modules.length === 0) return null;

    const objectModelUsageCounts: Record<string, number> = {
      Range: 0,
      Cells: 0,
      Worksheets: 0,
      Workbook: 0,
      Application: 0,
      ActiveSheet: 0,
      ActiveCell: 0,
      Selection: 0,
    };
    const rangeShapesCounts: Record<string, number> = {
      singleCell: 0,
      multiCell: 0,
      rows: 0,
      columns: 0,
      other: 0,
    };

    const perModule = modules.map((m) => ({ module: m, report: analyzeVbaModule(m) }));

    let externalReferencesCount = 0;
    let unsafeConstructsCount = 0;
    let unsupportedConstructsCount = 0;
    let warningsCount = 0;
    let todosCount = 0;

    for (const { report } of perModule) {
      for (const key of Object.keys(objectModelUsageCounts)) {
        objectModelUsageCounts[key] += report.objectModelUsage?.[key]?.length ?? 0;
      }
      for (const key of Object.keys(rangeShapesCounts)) {
        rangeShapesCounts[key] += report.rangeShapes?.[key]?.length ?? 0;
      }
      externalReferencesCount += report.externalReferences?.length ?? 0;
      unsafeConstructsCount += report.unsafeConstructs?.length ?? 0;
      unsupportedConstructsCount += report.unsupportedConstructs?.length ?? 0;
      warningsCount += report.warnings?.length ?? 0;
      todosCount += report.todos?.length ?? 0;
    }

    const riskScore = Math.min(100, externalReferencesCount * 25 + unsafeConstructsCount * 30 + unsupportedConstructsCount * 10);
    const riskLevel = riskScore >= 70 ? "high" : riskScore >= 30 ? "medium" : "low";

    return {
      kind: "project",
      projectName: project.name ?? "VBA Project",
      risk: { score: riskScore, level: riskLevel },
      objectModelUsageCounts,
      rangeShapesCounts,
      externalReferencesCount,
      unsafeConstructsCount,
      unsupportedConstructsCount,
      warningsCount,
      todosCount,
      modules: perModule
        .map(({ module, report }) => ({
          moduleName: module.name,
          risk: {
            score: report.risk?.score ?? 0,
            level: report.risk?.level ?? "low",
          },
          externalReferencesCount: report.externalReferences?.length ?? 0,
          unsafeConstructsCount: report.unsafeConstructs?.length ?? 0,
          unsupportedConstructsCount: report.unsupportedConstructs?.length ?? 0,
        }))
        .sort((a, b) => b.risk.score - a.risk.score),
    };
  }, [project]);

  const analysis: AnalysisViewModel | null = useMemo(() => {
    if (analysisScope === "project") return projectSummary;
    if (!moduleAnalysis) return null;
    return { kind: "module", report: moduleAnalysis };
  }, [analysisScope, moduleAnalysis, projectSummary]);

  const canConvert = Boolean(selectedModule && migrator);
  const canValidate = Boolean(conversionOutput && conversionTarget);

  const onConvert = useCallback(
    async (target: "python" | "typescript") => {
      if (!selectedModule || !migrator) return;
      setConversionTarget(target);
      setConversionStatus("working");
      setConversionError(null);
      setConversionPrompt("");
      setConversionOutput("");
      setValidationStatus("idle");
      setValidationError(null);
      setValidationReport(null);

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

  const onValidate = useCallback(async () => {
    if (!conversionTarget || !conversionOutput) return;
    setValidationStatus("working");
    setValidationError(null);
    setValidationReport(null);
    try {
      const invoke =
        props.invoke ??
        ((globalThis as any).__TAURI__?.core?.invoke as ((cmd: string, args?: any) => Promise<any>) | undefined);
      if (!invoke) {
        throw new Error("Tauri invoke API not available");
      }

      // Mirror the macro runner behavior: allow microtask-batched workbook edits to enqueue
      // first, then drain the queue so validation sees the latest workbook state.
      if (props.drainBackendSync) {
        await new Promise<void>((resolve) => queueMicrotask(resolve));
        await props.drainBackendSync();
      }

      const workbookId = props.workbookId ?? "local-workbook";

      // Ensure VBA validation runs under the same UI context as real macro execution.
      // This is best-effort: older backends (or tests) may not implement the command.
      if (props.getMacroUiContext) {
        const ctx = props.getMacroUiContext();
        try {
          await invoke("set_macro_ui_context", {
            workbook_id: workbookId,
            sheet_id: ctx.sheetId,
            active_row: ctx.activeRow,
            active_col: ctx.activeCol,
            selection: ctx.selection
              ? {
                  start_row: ctx.selection.startRow,
                  start_col: ctx.selection.startCol,
                  end_row: ctx.selection.endRow,
                  end_col: ctx.selection.endCol,
                }
              : null,
          });
        } catch {
          // Ignore context sync failures; validation can still run with the last known context.
        }
      }

      const report = await invoke("validate_vba_migration", {
        workbook_id: workbookId,
        macro_id: entryPoint.trim() || "Main",
        target: conversionTarget,
        code: conversionOutput,
      });
      setValidationReport(report);
      setValidationStatus("idle");
    } catch (err) {
      setValidationStatus("error");
      setValidationError(err instanceof Error ? err.message : String(err));
    }
  }, [
    conversionOutput,
    conversionTarget,
    entryPoint,
    props.workbookId,
    props.invoke,
    props.drainBackendSync,
    props.getMacroUiContext,
  ]);

  async function copyToClipboard(text: string) {
    try {
      const provider = await clipboardProviderPromise;
      await provider.write({ text });
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
    fontFamily: "var(--font-mono)",
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
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: 600 }}>Conversion</div>
              <div style={{ fontSize: 11, opacity: 0.8 }}>
                AI backend: <span style={{ fontFamily: "monospace" }}>Cursor</span>
              </div>
            </div>
            <label style={{ display: "flex", gap: 6, alignItems: "center", fontSize: 12 }}>
              Macro:
              {availableMacros.length > 0 ? (
                <select
                  value={entryPoint}
                  onChange={(e) => setEntryPoint(e.target.value)}
                  style={{ ...monospace, padding: "4px 6px", width: 160 }}
                  data-testid="vba-entrypoint"
                >
                  {availableMacros.map((m) => (
                    <option key={m.id} value={m.id}>
                      {m.name}
                    </option>
                  ))}
                </select>
              ) : (
                <input
                  value={entryPoint}
                  onChange={(e) => setEntryPoint(e.target.value)}
                  placeholder="Main"
                  style={{ ...monospace, padding: "4px 6px", width: 120 }}
                  data-testid="vba-entrypoint"
                />
              )}
            </label>
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
              data-testid="vba-validate"
              onClick={() => void onValidate()}
              disabled={!canValidate || validationStatus === "working"}
            >
              Validate
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

          <details style={{ borderTop: "1px solid var(--panel-border)" }} open={Boolean(validationReport || validationError)}>
            <summary style={{ cursor: "pointer", padding: 10, fontWeight: 600 }}>Validation</summary>
            <div style={{ padding: 10, display: "flex", flexDirection: "column", gap: 8 }}>
              {validationStatus === "working" ? <div style={{ fontSize: 12, opacity: 0.8 }}>Validating…</div> : null}
              {validationError ? (
                <div style={{ fontSize: 12, color: "var(--error)" }} data-testid="vba-validation-error">
                  {validationError}
                </div>
              ) : null}
              {validationReport ? (
                <div data-testid="vba-validation-report">
                  <div style={{ fontSize: 12 }}>
                    Result:{" "}
                    <span style={{ fontFamily: "monospace" }}>
                      {validationReport.ok ? "ok" : "failed"}
                      {Array.isArray(validationReport.mismatches)
                        ? ` (${validationReport.mismatches.length} mismatches)`
                        : ""}
                    </span>
                  </div>
                  {validationReport.error ? (
                    <div style={{ fontSize: 12, color: "var(--error)" }}>{String(validationReport.error)}</div>
                  ) : null}
                  {Array.isArray(validationReport.mismatches) && validationReport.mismatches.length > 0 ? (
                    <ul style={{ margin: 0, paddingLeft: 18 }}>
                      {validationReport.mismatches.slice(0, 50).map((m: any, idx: number) => (
                        <li key={`${m.sheetId ?? m.sheet_id}-${m.row}-${m.col}-${idx}`} style={{ fontSize: 12 }}>
                          <span style={{ fontFamily: "monospace" }}>
                            {m.sheetId ?? m.sheet_id}:{Number(m.row) + 1},{Number(m.col) + 1}
                          </span>
                          : VBA={String(m.vba?.display_value ?? "")} / Script={String(m.script?.display_value ?? "")}
                        </li>
                      ))}
                    </ul>
                  ) : null}
                  <details>
                    <summary style={{ cursor: "pointer", fontWeight: 600, fontSize: 12 }}>Raw report</summary>
                    <pre style={{ ...monospace, whiteSpace: "pre-wrap", margin: 0, paddingTop: 8 }}>
                      {JSON.stringify(validationReport, null, 2)}
                    </pre>
                  </details>
                </div>
              ) : (
                <div style={{ fontSize: 12, opacity: 0.8 }}>No validation run.</div>
              )}
            </div>
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
