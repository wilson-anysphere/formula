import { createEngineClient } from "@formula/engine";
import { CanvasGrid, GridPlaceholder, MockCellProvider } from "@formula/grid";
import { useEffect, useMemo, useState } from "react";

export function App() {
  const engine = useMemo(() => createEngineClient(), []);
  const [engineStatus, setEngineStatus] = useState("starting…");
  const [engineReady, setEngineReady] = useState(false);

  const rowCount = 1_000_000;
  const colCount = 100;
  const provider = useMemo(() => new MockCellProvider({ rowCount, colCount }), []);

  useEffect(() => {
    let cancelled = false;

    async function start() {
      try {
        await engine.init();
        const pong = await engine.ping();
        if (!cancelled) {
          setEngineStatus(`ready (${pong})`);
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

  return (
    <div style={{ padding: 24, fontFamily: "system-ui, sans-serif" }}>
      <h1 style={{ margin: 0 }}>Formula (Web Preview)</h1>
      <p style={{ marginTop: 8, color: "#475569" }}>
        Engine: <strong>{engineStatus}</strong>
      </p>
      <div style={{ marginTop: 16, height: 560 }}>
        {engineReady ? (
          <CanvasGrid provider={provider} rowCount={rowCount} colCount={colCount} frozenRows={1} frozenCols={1} />
        ) : (
          <GridPlaceholder />
        )}
      </div>
    </div>
  );
}
