import { createEngineClient, defaultWasmModuleUrl } from "@formula/engine";
import { CanvasGrid, GridPlaceholder } from "@formula/grid";
import { EngineCellCache, EngineGridProvider } from "@formula/spreadsheet-frontend";
import { useEffect, useMemo, useState } from "react";

export function App() {
  const engine = useMemo(() => createEngineClient({ wasmModuleUrl: defaultWasmModuleUrl() }), []);
  const [engineStatus, setEngineStatus] = useState("starting…");
  const [provider, setProvider] = useState<EngineGridProvider | null>(null);

  // +1 for frozen header row/col.
  const rowCount = 1_000_000 + 1;
  const colCount = 100 + 1;

  useEffect(() => {
    let cancelled = false;

    async function start() {
      try {
        await engine.init();
        const engineAny = engine as any;

        // Keep the preview deterministic by seeding a tiny workbook.
        if (typeof engineAny.newWorkbook === "function") {
          await engineAny.newWorkbook();
        }
        await engineAny.setCell("A1", 1);
        await engineAny.setCell("A2", 2);
        await engineAny.setCell("B1", "=A1+A2");
        await engineAny.setCell("B2", "=B1*2");
        await engineAny.setCell("C1", "hello");
        await engineAny.recalculate();

        const pong = typeof engineAny.ping === "function" ? await engineAny.ping() : "ok";
        if (!cancelled) {
          setEngineStatus(`ready (${pong})`);
          const cache = new EngineCellCache(engine);
          setProvider(new EngineGridProvider({ cache, rowCount, colCount, headers: true }));
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

  return (
    <div style={{ padding: 24, fontFamily: "system-ui, sans-serif" }}>
      <h1 style={{ margin: 0 }}>Formula (Web Preview)</h1>
      <p style={{ marginTop: 8, color: "#475569" }}>
        Engine: <strong>{engineStatus}</strong>
      </p>
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
