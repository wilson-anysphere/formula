import { createEngineClient } from "@formula/engine";
import { CanvasGrid, GridPlaceholder } from "@formula/grid";
import { EngineCellCache, EngineGridProvider } from "@formula/spreadsheet-frontend";
import { useEffect, useMemo, useState } from "react";

import { DEMO_WORKBOOK_JSON } from "./engine/documentControllerSync";

export function App() {
  const engine = useMemo(() => createEngineClient(), []);
  const [engineStatus, setEngineStatus] = useState("starting…");
  const [provider, setProvider] = useState<EngineGridProvider | null>(null);
  const [activeSheet, setActiveSheet] = useState("Sheet1");
  const [engineReady, setEngineReady] = useState(false);

  // +1 for frozen header row/col.
  const rowCount = 1_000_000 + 1;
  const colCount = 100 + 1;

  useEffect(() => {
    let cancelled = false;

    async function start() {
      try {
        await engine.init();
        await engine.loadWorkbookFromJson(DEMO_WORKBOOK_JSON);
        // Ensure there's a second sheet for the sheet selector demo.
        await engine.setCell("A1", "Hello from Sheet2", "Sheet2");
        await engine.recalculate();
        const b1 = await engine.getCell("B1");
        if (!cancelled) {
          setEngineStatus(`ready (B1=${b1.value === null ? "" : String(b1.value)})`);
          setEngineReady(true);
        }
      } catch (error) {
        if (!cancelled)
          setEngineStatus(
            `error: ${error instanceof Error ? error.message : String(error)}`,
          );
      }
    }

    void start();

    return () => {
      cancelled = true;
      engine.terminate();
    };
  }, [engine]);

  useEffect(() => {
    if (!engineReady) return;
    const cache = new EngineCellCache(engine);
    const nextProvider = new EngineGridProvider({ cache, rowCount, colCount, sheet: activeSheet, headers: true });
    setProvider(nextProvider);
    void nextProvider.recalculate();
  }, [engine, engineReady, activeSheet]);

  return (
    <div style={{ padding: 24, fontFamily: "system-ui, sans-serif" }}>
      <h1 style={{ margin: 0 }}>Formula (Web Preview)</h1>
      <p style={{ marginTop: 8, color: "#475569" }}>
        Engine: <strong data-testid="engine-status">{engineStatus}</strong>
      </p>
      <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8 }}>
        Sheet:
        <select
          value={activeSheet}
          onChange={(e) => setActiveSheet(e.target.value)}
          style={{ padding: "4px 6px" }}
          disabled={!engineReady}
        >
          <option value="Sheet1">Sheet1</option>
          <option value="Sheet2">Sheet2</option>
        </select>
      </label>
      <div style={{ marginTop: 16, height: 560 }}>
        {provider ? (
          <CanvasGrid provider={provider} rowCount={rowCount} colCount={colCount} frozenRows={1} frozenCols={1} />
        ) : (
          <GridPlaceholder />
        )}
      </div>
    </div>
  );
}
