import { createEngineClient } from "@formula/engine";
import type { CellRange } from "@formula/grid";
import { CanvasGrid, GridPlaceholder } from "@formula/grid";
import { EngineCellCache, EngineGridProvider } from "@formula/spreadsheet-frontend";
import { useEffect, useLayoutEffect, useRef, useState } from "react";

import { rangeToA1 } from "./a1";
import { DEMO_WORKBOOK_JSON } from "./engine/documentControllerSync";

export function App() {
  const [engineStatus, setEngineStatus] = useState("starting…");
  const engineRef = useRef<ReturnType<typeof createEngineClient> | null>(null);
  const [provider, setProvider] = useState<EngineGridProvider | null>(null);
  const [activeSheet, setActiveSheet] = useState("Sheet1");
  const previousSheetRef = useRef<string | null>(null);

  // +1 for frozen header row/col.
  const rowCount = 1_000_000 + 1;
  const colCount = 100 + 1;
  const frozenRows = 1;
  const frozenCols = 1;

  const inputRef = useRef<HTMLInputElement | null>(null);
  const [draft, setDraft] = useState("");
  const draftRef = useRef(draft);
  const [formulaFocused, setFormulaFocused] = useState(false);

  const cursorRef = useRef<{ start: number; end: number }>({ start: 0, end: 0 });
  const rangeInsertionRef = useRef<{ start: number; end: number } | null>(null);
  const pendingSelectionRef = useRef<{ start: number; end: number } | null>(null);

  const isFormulaEditing = formulaFocused && draft.trim().startsWith("=");
  const headerRowOffset = frozenRows > 0 ? 1 : 0;
  const headerColOffset = frozenCols > 0 ? 1 : 0;

  useEffect(() => {
    draftRef.current = draft;
  }, [draft]);

  const syncCursorFromInput = () => {
    const input = inputRef.current;
    if (!input) return;
    const start = input.selectionStart ?? input.value.length;
    const end = input.selectionEnd ?? input.value.length;
    cursorRef.current = { start, end };
  };

  const cellRangeToA1 = (range: CellRange): string | null => {
    const startRow = range.startRow - headerRowOffset;
    const startCol = range.startCol - headerColOffset;
    const endRow = range.endRow - 1 - headerRowOffset;
    const endCol = range.endCol - 1 - headerColOffset;

    if (startRow < 0 || startCol < 0 || endRow < 0 || endCol < 0) return null;

    return rangeToA1({
      start: { row: startRow, col: startCol },
      end: { row: endRow, col: endCol }
    });
  };

  const insertOrReplaceRange = (rangeText: string, isBegin: boolean) => {
    const currentDraft = draftRef.current;

    if (!rangeInsertionRef.current || isBegin) {
      const start = Math.min(cursorRef.current.start, cursorRef.current.end);
      const end = Math.max(cursorRef.current.start, cursorRef.current.end);
      const nextDraft = currentDraft.slice(0, start) + rangeText + currentDraft.slice(end);

      rangeInsertionRef.current = { start, end: start + rangeText.length };
      const cursor = rangeInsertionRef.current.end;
      cursorRef.current = { start: cursor, end: cursor };
      pendingSelectionRef.current = { start: cursor, end: cursor };
      draftRef.current = nextDraft;
      setDraft(nextDraft);
      return;
    }

    const { start, end } = rangeInsertionRef.current;
    const nextDraft = currentDraft.slice(0, start) + rangeText + currentDraft.slice(end);
    rangeInsertionRef.current = { start, end: start + rangeText.length };
    const cursor = rangeInsertionRef.current.end;
    cursorRef.current = { start: cursor, end: cursor };
    pendingSelectionRef.current = { start: cursor, end: cursor };
    draftRef.current = nextDraft;
    setDraft(nextDraft);
  };

  const beginRangeSelection = (range: CellRange) => {
    if (!isFormulaEditing) return;
    const ref = cellRangeToA1(range);
    if (!ref) return;
    syncCursorFromInput();
    insertOrReplaceRange(ref, true);
  };

  const updateRangeSelection = (range: CellRange) => {
    if (!isFormulaEditing) return;
    const ref = cellRangeToA1(range);
    if (!ref) return;
    insertOrReplaceRange(ref, false);
  };

  const endRangeSelection = () => {
    rangeInsertionRef.current = null;
    const input = inputRef.current;
    if (!input) return;
    input.focus({ preventScroll: true });
    const cursor = cursorRef.current.start;
    input.setSelectionRange(cursor, cursor);
  };

  useLayoutEffect(() => {
    const input = inputRef.current;
    const pending = pendingSelectionRef.current;
    if (!input || !pending) return;
    pendingSelectionRef.current = null;
    input.setSelectionRange(pending.start, pending.end);
  }, [draft]);

  useEffect(() => {
    // Create the Engine client inside the effect. React.StrictMode re-runs effects
    // and their cleanup in dev; terminating a memoized Worker-backed client can
    // leave subsequent `init()` calls hanging.
    const engine = createEngineClient();
    engineRef.current = engine;

    setEngineStatus("starting…");
    setProvider(null);

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
          const cache = new EngineCellCache(engine);
          setProvider(new EngineGridProvider({ cache, rowCount, colCount, sheet: "Sheet1", headers: true }));
        }
      } catch (error) {
        if (!cancelled) {
          setEngineStatus(
            `error: ${error instanceof Error ? error.message : String(error)}`
          );
        }
      }
    }

    void start();

    return () => {
      cancelled = true;
      engine.terminate();
      if (engineRef.current === engine) engineRef.current = null;
    };
  }, []);

  useEffect(() => {
    if (!provider) return;

    provider.setSheet(activeSheet);

    const previousSheet = previousSheetRef.current;
    previousSheetRef.current = activeSheet;
    if (previousSheet && previousSheet !== activeSheet) {
      void provider.recalculate(activeSheet);
    }
  }, [provider, activeSheet]);

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
          disabled={!provider}
        >
          <option value="Sheet1">Sheet1</option>
          <option value="Sheet2">Sheet2</option>
        </select>
      </label>

      <div style={{ marginTop: 16 }}>
        <label style={{ display: "block", fontSize: 12, color: "#64748b" }} htmlFor="formula-input">
          Formula
        </label>
        <input
          ref={inputRef}
          id="formula-input"
          data-testid="formula-input"
          spellCheck={false}
          value={draft}
          onFocus={() => {
            setFormulaFocused(true);
            syncCursorFromInput();
          }}
          onBlur={() => {
            setFormulaFocused(false);
            rangeInsertionRef.current = null;
          }}
          onChange={(event) => {
            const value = event.currentTarget.value;
            setDraft(value);
            draftRef.current = value;
            rangeInsertionRef.current = null;
            cursorRef.current = {
              start: event.currentTarget.selectionStart ?? value.length,
              end: event.currentTarget.selectionEnd ?? value.length
            };
          }}
          onClick={syncCursorFromInput}
          onKeyUp={syncCursorFromInput}
          onSelect={syncCursorFromInput}
          style={{
            marginTop: 4,
            width: "100%",
            padding: "8px 10px",
            border: "1px solid #cbd5e1",
            borderRadius: 6,
            fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace",
            fontSize: 14
          }}
        />
      </div>

      <div data-testid="grid" style={{ marginTop: 16, height: 560 }}>
        {provider ? (
          <CanvasGrid
            provider={provider}
            rowCount={rowCount}
            colCount={colCount}
            frozenRows={frozenRows}
            frozenCols={frozenCols}
            interactionMode={isFormulaEditing ? "rangeSelection" : "default"}
            onRangeSelectionStart={beginRangeSelection}
            onRangeSelectionChange={updateRangeSelection}
            onRangeSelectionEnd={endRangeSelection}
          />
        ) : (
          <GridPlaceholder />
        )}
      </div>
    </div>
  );
}
