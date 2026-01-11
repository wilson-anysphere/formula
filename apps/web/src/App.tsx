import { createEngineClient, type CellChange, type CellScalar } from "@formula/engine";
import type { CellRange } from "@formula/grid";
import { CanvasGrid, GridPlaceholder, type GridApi } from "@formula/grid";
import { range0ToA1, toA1 } from "@formula/spreadsheet-frontend";
import { useEffect, useLayoutEffect, useRef, useState } from "react";

import { EngineCellProvider } from "./EngineCellProvider";
import { DEMO_WORKBOOK_JSON } from "./engine/documentControllerSync";

function scalarToDisplayString(value: CellScalar): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function parseFormulaBarInput(raw: string): CellScalar {
  if (raw.startsWith("=") && raw.length > 1) return raw;

  const trimmed = raw.trim();
  if (trimmed === "") return null;

  if (/^(true|false)$/i.test(trimmed)) return trimmed.toLowerCase() === "true";
  if (/^null$/i.test(trimmed)) return null;

  if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) {
    return Number(trimmed);
  }

  return raw;
}

function isFormulaInput(value: CellScalar): value is string {
  return typeof value === "string" && value.startsWith("=") && value.length > 1;
}

export function App() {
  const [engineStatus, setEngineStatus] = useState("starting…");
  const engineRef = useRef<ReturnType<typeof createEngineClient> | null>(null);
  const [provider, setProvider] = useState<EngineCellProvider | null>(null);
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
  const cellSyncTokenRef = useRef(0);

  const isFormulaEditing = formulaFocused && draft.trim().startsWith("=");
  const headerRowOffset = frozenRows > 0 ? 1 : 0;
  const headerColOffset = frozenCols > 0 ? 1 : 0;

  const gridApiRef = useRef<GridApi | null>(null);
  const [activeCell, setActiveCell] = useState<{ row: number; col: number } | null>(null);

  const activeAddress = (() => {
    if (!activeCell) return null;
    const row0 = activeCell.row - headerRowOffset;
    const col0 = activeCell.col - headerColOffset;
    if (row0 < 0 || col0 < 0) return null;
    return toA1(row0, col0);
  })();

  const [activeValue, setActiveValue] = useState<CellScalar>(null);

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
    const startRow0 = range.startRow - headerRowOffset;
    const startCol0 = range.startCol - headerColOffset;
    const endRow0Exclusive = range.endRow - headerRowOffset;
    const endCol0Exclusive = range.endCol - headerColOffset;

    if (startRow0 < 0 || startCol0 < 0) return null;
    if (endRow0Exclusive <= startRow0 || endCol0Exclusive <= startCol0) return null;

    return range0ToA1({
      startRow0,
      startCol0,
      endRow0Exclusive,
      endCol0Exclusive
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
    setActiveCell(null);

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
          setProvider(new EngineCellProvider({ engine, rowCount, colCount, sheet: "Sheet1" }));
        }
      } catch (error) {
        if (!cancelled) {
          setEngineStatus(`error: ${error instanceof Error ? error.message : String(error)}`);
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

  useEffect(() => {
    if (!provider) return;
    const id = requestAnimationFrame(() => {
      gridApiRef.current?.setSelection(headerRowOffset, headerColOffset);
    });
    return () => cancelAnimationFrame(id);
  }, [provider, headerRowOffset, headerColOffset]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;

    if (!provider || !activeAddress) {
      rangeInsertionRef.current = null;
      draftRef.current = "";
      setDraft("");
      setActiveValue(null);
      return;
    }

    if (isFormulaEditing) {
      return;
    }

    const token = ++cellSyncTokenRef.current;
    void engine
      .getCell(activeAddress, activeSheet)
      .then((cell) => {
        if (cellSyncTokenRef.current !== token) return;
        const inputText = scalarToDisplayString(cell.input as CellScalar);
        rangeInsertionRef.current = null;
        draftRef.current = inputText;
        setDraft(inputText);
        setActiveValue(cell.value as CellScalar);
      })
      .catch(() => {
        // Ignore selection reads while the engine is initializing/tearing down.
      });
  }, [provider, activeAddress, activeSheet, isFormulaEditing]);

  const commitDraft = async () => {
    const engine = engineRef.current;
    if (!engine || !provider || !activeAddress) return;

    const nextValue = parseFormulaBarInput(draftRef.current);
    await engine.setCell(activeAddress, nextValue, activeSheet);
    const changes = await engine.recalculate(activeSheet);

    const directChange: CellChange | null = isFormulaInput(nextValue)
      ? null
      : { sheet: activeSheet, address: activeAddress, value: nextValue };
    provider.applyRecalcChanges(directChange ? [...changes, directChange] : changes);

    const updated = await engine.getCell(activeAddress, activeSheet);
    const inputText = scalarToDisplayString(updated.input as CellScalar);
    rangeInsertionRef.current = null;
    draftRef.current = inputText;
    setDraft(inputText);
    setActiveValue(updated.value as CellScalar);
  };

  const onSelectionChange = (cell: { row: number; col: number } | null) => {
    if (isFormulaEditing) return;
    setActiveCell(cell);
  };

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
        <div style={{ marginTop: 4, display: "flex", alignItems: "center", gap: 8 }}>
          <div
            style={{
              width: 64,
              fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace",
              fontSize: 12,
              color: "#0f172a"
            }}
            data-testid="active-address"
          >
            {activeAddress ?? ""}
          </div>
          <input
            ref={inputRef}
            id="formula-input"
            data-testid="formula-input"
            spellCheck={false}
            value={draft}
            onFocus={(event) => {
              setFormulaFocused(true);
              cellSyncTokenRef.current++;
              const input = event.currentTarget;
              queueMicrotask(() => {
                input.select();
                syncCursorFromInput();
              });
            }}
            onBlur={() => {
              setFormulaFocused(false);
              rangeInsertionRef.current = null;
            }}
            onChange={(event) => {
              const value = event.currentTarget.value;
              setDraft(value);
              draftRef.current = value;
              cellSyncTokenRef.current++;
              rangeInsertionRef.current = null;
              cursorRef.current = {
                start: event.currentTarget.selectionStart ?? value.length,
                end: event.currentTarget.selectionEnd ?? value.length
              };
            }}
            onKeyDown={(event) => {
              if (event.key !== "Enter") return;
              event.preventDefault();
              void commitDraft();
            }}
            onClick={syncCursorFromInput}
            onKeyUp={syncCursorFromInput}
            onSelect={syncCursorFromInput}
            disabled={!provider || !activeAddress}
            style={{
              flex: 1,
              padding: "8px 10px",
              border: "1px solid #cbd5e1",
              borderRadius: 6,
              fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace",
              fontSize: 14
            }}
          />
          <div style={{ minWidth: 120, fontSize: 12, color: "#475569" }}>
            Value: <span data-testid="formula-bar-value">{scalarToDisplayString(activeValue)}</span>
          </div>
        </div>
      </div>

      <div data-testid="grid" style={{ marginTop: 16, height: 560 }}>
        {provider ? (
          <CanvasGrid
            provider={provider}
            rowCount={rowCount}
            colCount={colCount}
            frozenRows={frozenRows}
            frozenCols={frozenCols}
            apiRef={(api) => {
              gridApiRef.current = api;
              const params = new URLSearchParams(window.location.search);
              if (params.has("e2e") || params.has("perf")) {
                (window as any).__gridApi = api;
              }
            }}
            onSelectionChange={onSelectionChange}
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
