import React, { useMemo, useState } from "react";

import type { VbaMigrator } from "../../../../packages/vba-migrate/src/converter.js";
import { analyzeVbaModule, migrationReportToMarkdown, validateMigration } from "../../../../packages/vba-migrate/src/index.js";

export type VbaModuleSummary = {
  name: string;
  code: string;
};

export type VbaMigratePanelProps = {
  modules: VbaModuleSummary[];
  migrator: VbaMigrator;
  workbook: any;
  entryPointByModule?: Record<string, string>;
  onSaveConvertedScript?: (args: { moduleName: string; target: "python" | "typescript"; code: string }) => Promise<void>;
};

/**
 * Minimal UI surface for "Migrate Macros".
 *
 * The real desktop app can swap `workbook` for the actual document controller API
 * and connect `onSaveConvertedScript` to project storage.
 */
export function VbaMigratePanel(props: VbaMigratePanelProps) {
  const [selectedModuleName, setSelectedModuleName] = useState(props.modules[0]?.name ?? null);
  const [target, setTarget] = useState<"python" | "typescript">("python");
  const [conversionOutput, setConversionOutput] = useState<string>("");
  const [validationOutput, setValidationOutput] = useState<string>("");
  const selectedModule = useMemo(
    () => props.modules.find((m) => m.name === selectedModuleName) ?? null,
    [props.modules, selectedModuleName]
  );

  const analysis = useMemo(() => {
    if (!selectedModule) return null;
    return analyzeVbaModule(selectedModule);
  }, [selectedModule]);

  async function onConvert() {
    if (!selectedModule) return;
    setConversionOutput("Converting...");
    const result = await props.migrator.convertModule(selectedModule, { target });
    setConversionOutput(result.code);
  }

  async function onValidate() {
    if (!selectedModule || !conversionOutput) return;
    setValidationOutput("Validating...");
    const entryPoint = props.entryPointByModule?.[selectedModule.name] ?? "Main";
    const result = validateMigration({
      workbook: props.workbook,
      module: selectedModule,
      entryPoint,
      target,
      code: conversionOutput
    });
    setValidationOutput(JSON.stringify(result, null, 2));
  }

  async function onSave() {
    if (!selectedModule || !conversionOutput || !props.onSaveConvertedScript) return;
    await props.onSaveConvertedScript({ moduleName: selectedModule.name, target, code: conversionOutput });
  }

  return (
    <div style={{ display: "flex", height: "100%", gap: 12 }}>
      <div style={{ width: 240, borderRight: "1px solid var(--border)", paddingRight: 12 }}>
        <div style={{ fontWeight: 600, marginBottom: 8 }}>Modules</div>
        <ul style={{ listStyle: "none", padding: 0, margin: 0 }}>
          {props.modules.map((mod) => (
            <li key={mod.name}>
              <button
                onClick={() => setSelectedModuleName(mod.name)}
                style={{
                  width: "100%",
                  textAlign: "left",
                  padding: "6px 8px",
                  background: mod.name === selectedModuleName ? "var(--selection-bg)" : "transparent",
                  border: "1px solid var(--border)",
                  marginBottom: 6,
                  cursor: "pointer",
                  color: "inherit"
                }}
              >
                {mod.name}
              </button>
            </li>
          ))}
        </ul>
      </div>

      <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 12 }}>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <div style={{ fontWeight: 600, flex: 1 }}>{selectedModule?.name ?? "No module selected"}</div>
          <select value={target} onChange={(e) => setTarget(e.target.value as any)}>
            <option value="python">Python</option>
            <option value="typescript">TypeScript</option>
          </select>
          <button onClick={onConvert} disabled={!selectedModule}>
            Convert
          </button>
          <button onClick={onValidate} disabled={!selectedModule || !conversionOutput}>
            Validate
          </button>
          <button onClick={onSave} disabled={!props.onSaveConvertedScript || !conversionOutput}>
            Save
          </button>
        </div>
 
        <div style={{ display: "flex", gap: 12, flex: 1, overflow: "hidden" }}>
          <div style={{ flex: 1, overflow: "auto", border: "1px solid var(--border)", padding: 8 }}>
            <div style={{ fontWeight: 600, marginBottom: 8 }}>Analysis</div>
            <pre style={{ margin: 0, whiteSpace: "pre-wrap" }}>
              {analysis ? migrationReportToMarkdown(analysis) : "No analysis"}
            </pre>
          </div>
          <div style={{ flex: 1, overflow: "auto", border: "1px solid var(--border)", padding: 8 }}>
            <div style={{ fontWeight: 600, marginBottom: 8 }}>Conversion output</div>
            <pre style={{ margin: 0, whiteSpace: "pre-wrap" }}>{conversionOutput || "No output"}</pre>
          </div>
          <div style={{ flex: 1, overflow: "auto", border: "1px solid var(--border)", padding: 8 }}>
            <div style={{ fontWeight: 600, marginBottom: 8 }}>Validation diff</div>
            <pre style={{ margin: 0, whiteSpace: "pre-wrap" }}>{validationOutput || "No validation run"}</pre>
          </div>
        </div>
      </div>
    </div>
  );
}
