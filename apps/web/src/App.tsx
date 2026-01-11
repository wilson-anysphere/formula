import { createEngineClient, type CellChange, type CellScalar } from "@formula/engine";
import type { CellRange } from "@formula/grid";
import { CanvasGrid, GridPlaceholder, MockCellProvider, type GridApi } from "@formula/grid";
import {
  parseHtmlTableToGrid,
  parseTsvToGrid,
  range0ToA1,
  serializeGridToHtmlTable,
  serializeGridToTsv,
  toA1
} from "@formula/spreadsheet-frontend";
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ClipboardEvent } from "react";

import { CellEditorOverlay } from "./CellEditorOverlay";
import { EngineCellProvider } from "./EngineCellProvider";
import { isFormulaInput, parseCellScalarInput, scalarToDisplayString } from "./cellScalar";
import { DEMO_WORKBOOK_JSON } from "./engine/documentControllerSync";

export function App() {
  const params = typeof window !== "undefined" ? new URLSearchParams(window.location.search) : null;
  const perfMode = params?.has("perf") ?? false;
  return perfMode ? <PerfGridApp /> : <EngineDemoApp />;
}

function EngineDemoApp() {
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
  const gridContainerRef = useRef<HTMLDivElement | null>(null);
  const internalClipboardRef = useRef<{ tsv: string; html: string } | null>(null);
  const [activeCell, setActiveCell] = useState<{ row: number; col: number } | null>(null);
  const [editingCell, setEditingCell] = useState<{ row: number; col: number } | null>(null);
  const editingCellOriginalDraftRef = useRef("");

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

    if (isFormulaEditing || editingCell) {
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
  }, [provider, activeAddress, activeSheet, isFormulaEditing, editingCell]);

  const commitDraft = async () => {
    const engine = engineRef.current;
    if (!engine || !provider || !activeAddress) return;

    const nextValue = parseCellScalarInput(draftRef.current);
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

  const focusGrid = () => {
    const host = gridContainerRef.current;
    if (!host) return;
    const grid = host.querySelector<HTMLElement>('[data-testid="canvas-grid"]');
    grid?.focus({ preventScroll: true });
  };

  const beginCellEdit = (request: { row: number; col: number; initialKey?: string }) => {
    if (!provider) return;

    // Ensure React state matches the cell the grid intends to edit, even if the
    // selection hasn't changed (e.g., F2 or type-to-edit).
    setActiveCell({ row: request.row, col: request.col });
    gridApiRef.current?.scrollToCell(request.row, request.col, { align: "auto", padding: 8 });

    editingCellOriginalDraftRef.current = draftRef.current;
    cellSyncTokenRef.current++;
    rangeInsertionRef.current = null;

    if (request.initialKey !== undefined) {
      draftRef.current = request.initialKey;
      setDraft(request.initialKey);
    }

    setEditingCell({ row: request.row, col: request.col });
  };

  const cancelCellEdit = () => {
    const original = editingCellOriginalDraftRef.current;
    setEditingCell(null);
    draftRef.current = original;
    setDraft(original);
    requestAnimationFrame(() => focusGrid());
  };

  const commitCellEdit = async (nav: { deltaRow: number; deltaCol: number }) => {
    if (!editingCell) return;
    const from = editingCell;
    await commitDraft();
    setEditingCell(null);

    const nextRow = Math.max(0, Math.min(rowCount - 1, from.row + nav.deltaRow));
    const nextCol = Math.max(0, Math.min(colCount - 1, from.col + nav.deltaCol));
    requestAnimationFrame(() => {
      gridApiRef.current?.setSelection(nextRow, nextCol);
      gridApiRef.current?.scrollToCell(nextRow, nextCol, { align: "auto", padding: 8 });
      focusGrid();
    });
  };

  const onSelectionChange = (cell: { row: number; col: number } | null) => {
    if (isFormulaEditing || editingCell) return;
    setActiveCell(cell);
  };

  const getCopyRange = (): CellRange | null => {
    const api = gridApiRef.current;
    if (!api) return null;

    const range = api.getSelectionRange();
    if (!range) {
      const cell = api.getSelection();
      if (!cell) return null;
      if (cell.row < headerRowOffset || cell.col < headerColOffset) return null;
      const startRow = Math.max(headerRowOffset, cell.row);
      const startCol = Math.max(headerColOffset, cell.col);
      return { startRow, endRow: startRow + 1, startCol, endCol: startCol + 1 };
    }

    const startRow = Math.max(headerRowOffset, range.startRow);
    const startCol = Math.max(headerColOffset, range.startCol);
    const endRow = Math.max(startRow, range.endRow);
    const endCol = Math.max(startCol, range.endCol);

    if (endRow <= startRow || endCol <= startCol) return null;

    return { startRow, endRow, startCol, endCol };
  };

  const selectionRangeToStringGrid = (range: CellRange): string[][] => {
    if (!provider) return [];

    const rows: string[][] = [];

    for (let row = range.startRow; row < range.endRow; row++) {
      const outRow: string[] = [];
      for (let col = range.startCol; col < range.endCol; col++) {
        const cell = provider.getCell(row, col);
        outRow.push(scalarToDisplayString((cell?.value ?? null) as CellScalar));
      }
      rows.push(outRow);
    }

    return rows;
  };

  const syncFormulaBar = async (address: string) => {
    const engine = engineRef.current;
    if (!engine) return;
    const token = ++cellSyncTokenRef.current;
    try {
      const cell = await engine.getCell(address, activeSheet);
      if (cellSyncTokenRef.current !== token) return;
      const inputText = scalarToDisplayString(cell.input as CellScalar);
      rangeInsertionRef.current = null;
      draftRef.current = inputText;
      setDraft(inputText);
      setActiveValue(cell.value as CellScalar);
    } catch {
      // Ignore selection reads while the engine is initializing/tearing down.
    }
  };

  const handleGridCopy = (event: ClipboardEvent<HTMLDivElement>) => {
    if (editingCell) return;
    if (!provider) return;
    const range = getCopyRange();
    if (!range) return;

    // Web preview behavior: copy the displayed value grid (not formulas).
    // This matches what the canvas grid currently renders and makes TSV/HTML
    // interoperability predictable.
    const grid = selectionRangeToStringGrid(range);
    const tsv = serializeGridToTsv(grid);
    const html = serializeGridToHtmlTable(grid);

    internalClipboardRef.current = { tsv, html };

    event.clipboardData?.setData("text/plain", tsv);
    // Some consumers look for explicit TSV MIME types.
    event.clipboardData?.setData("text/tab-separated-values", tsv);
    event.clipboardData?.setData("text/tsv", tsv);
    event.clipboardData?.setData("text/html", html);
    if (!event.clipboardData) {
      const clipboard = navigator.clipboard;
      if (clipboard && typeof clipboard.write === "function" && typeof ClipboardItem !== "undefined") {
        const item = new ClipboardItem({
          "text/plain": new Blob([tsv], { type: "text/plain" }),
          "text/tab-separated-values": new Blob([tsv], { type: "text/tab-separated-values" }),
          "text/tsv": new Blob([tsv], { type: "text/tsv" }),
          "text/html": new Blob([html], { type: "text/html" })
        });
        void clipboard.write([item]).catch(() => {
          // Ignore clipboard API failures (permissions, unsupported platform, etc).
        });
      }
    }
    event.preventDefault();
  };

  const handleGridPaste = (event: ClipboardEvent<HTMLDivElement>) => {
    if (editingCell) return;
    const clipboardPlain =
      event.clipboardData?.getData("text/plain") ??
      event.clipboardData?.getData("text/tab-separated-values") ??
      event.clipboardData?.getData("text/tsv") ??
      "";
    const clipboardHtml = event.clipboardData?.getData("text/html") ?? "";
    const internal = internalClipboardRef.current;

    const grid =
      clipboardPlain !== ""
        ? parseTsvToGrid(clipboardPlain)
        : clipboardHtml !== ""
          ? parseHtmlTableToGrid(clipboardHtml)
          : internal?.tsv
            ? parseTsvToGrid(internal.tsv)
            : internal?.html
              ? parseHtmlTableToGrid(internal.html)
              : null;

    if (!grid || grid.length === 0) return;

    const engine = engineRef.current;
    const api = gridApiRef.current;
    if (!engine || !provider || !api) return;

    const selection = api.getSelection();
    if (!selection) return;

    const startRow0 = selection.row - headerRowOffset;
    const startCol0 = selection.col - headerColOffset;
    if (startRow0 < 0 || startCol0 < 0) return;

    event.preventDefault();

    void (async () => {
      const pasteRows = grid.length;
      const pasteCols = Math.max(0, ...grid.map((row) => row.length));
      if (pasteRows === 0 || pasteCols === 0) return;

      const values: CellScalar[][] = [];
      const directChanges: CellChange[] = [];

      for (let r = 0; r < pasteRows; r++) {
        const row = grid[r] ?? [];
        const outRow: CellScalar[] = [];
        for (let c = 0; c < pasteCols; c++) {
          const raw = row[c] ?? "";
          const value = parseCellScalarInput(raw);
          outRow.push(value);

          if (!isFormulaInput(value)) {
            directChanges.push({ sheet: activeSheet, address: toA1(startRow0 + r, startCol0 + c), value });
          }
        }
        values.push(outRow);
      }

      const rangeA1 = range0ToA1({
        startRow0,
        startCol0,
        endRow0Exclusive: startRow0 + pasteRows,
        endCol0Exclusive: startCol0 + pasteCols
      });

      await engine.setRange(rangeA1, values, activeSheet);
      const changes = await engine.recalculate(activeSheet);
      provider.applyRecalcChanges(directChanges.length > 0 ? [...changes, ...directChanges] : changes);

      api.setSelectionRange({
        startRow: selection.row,
        endRow: Math.min(rowCount, selection.row + pasteRows),
        startCol: selection.col,
        endCol: Math.min(colCount, selection.col + pasteCols)
      });

      await syncFormulaBar(toA1(startRow0, startCol0));
    })();
  };

  const handleGridCut = (event: ClipboardEvent<HTMLDivElement>) => {
    if (editingCell) return;
    const range = getCopyRange();
    if (!range) return;
    handleGridCopy(event);
    if (event.defaultPrevented) {
      // Clear on the next tick; clipboard population must remain synchronous.
      void (async () => {
        const engine = engineRef.current;
        if (!engine || !provider) return;

        const startRow0 = range.startRow - headerRowOffset;
        const startCol0 = range.startCol - headerColOffset;
        const endRow0Exclusive = range.endRow - headerRowOffset;
        const endCol0Exclusive = range.endCol - headerColOffset;
        if (startRow0 < 0 || startCol0 < 0) return;
        if (endRow0Exclusive <= startRow0 || endCol0Exclusive <= startCol0) return;

        const clearRows = endRow0Exclusive - startRow0;
        const clearCols = endCol0Exclusive - startCol0;

        const values = Array.from({ length: clearRows }, () => Array.from({ length: clearCols }, () => null as CellScalar));
        const rangeA1 = range0ToA1({ startRow0, startCol0, endRow0Exclusive, endCol0Exclusive });

        await engine.setRange(rangeA1, values, activeSheet);
        const changes = await engine.recalculate(activeSheet);

        const directChanges: CellChange[] = [];
        for (let r = 0; r < clearRows; r++) {
          for (let c = 0; c < clearCols; c++) {
            directChanges.push({ sheet: activeSheet, address: toA1(startRow0 + r, startCol0 + c), value: null });
          }
        }

        provider.applyRecalcChanges(directChanges.length > 0 ? [...changes, ...directChanges] : changes);
        await syncFormulaBar(toA1(startRow0, startCol0));
      })();
    }
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

      <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 12 }}>
        Import XLSX:
        <input
          type="file"
          accept=".xlsx"
          data-testid="xlsx-file-input"
          disabled={!provider}
          onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            if (!file) return;
            const engine = engineRef.current;
            if (!engine) return;

            setEngineStatus("importing xlsx…");
            setProvider(null);
            setActiveCell(null);

            void file
              .arrayBuffer()
              .then(async (buffer) => {
                const bytes = new Uint8Array(buffer);
                await engine.loadWorkbookFromXlsxBytes(bytes);
                // Ensure there's still a Sheet2 option available for the demo selector.
                await engine.setCell("A1", "Hello from Sheet2", "Sheet2");
                await engine.recalculate();
                const b1 = await engine.getCell("B1");
                setEngineStatus(`ready (imported xlsx; B1=${b1.value === null ? "" : String(b1.value)})`);
                setActiveSheet("Sheet1");
                previousSheetRef.current = null;
                setProvider(new EngineCellProvider({ engine, rowCount, colCount, sheet: "Sheet1" }));
              })
              .catch((error) => {
                setEngineStatus(`error: ${error instanceof Error ? error.message : String(error)}`);
              })
              .finally(() => {
                // Allow uploading the same fixture again.
                event.currentTarget.value = "";
              });
          }}
        />
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

      <div
        ref={gridContainerRef}
        data-testid="grid"
        onCopy={handleGridCopy}
        onCut={handleGridCut}
        onPaste={handleGridPaste}
        style={{ marginTop: 16, height: 560, position: "relative" }}
      >
        {provider ? (
          <>
            <CanvasGrid
              provider={provider}
              rowCount={rowCount}
              colCount={colCount}
              frozenRows={frozenRows}
              frozenCols={frozenCols}
              enableResize
              apiRef={(api) => {
                gridApiRef.current = api;
                const params = new URLSearchParams(window.location.search);
                if (params.has("e2e") || params.has("perf")) {
                  (window as any).__gridApi = api;
                }
              }}
              onSelectionChange={onSelectionChange}
              onRequestCellEdit={(request) => {
                if (editingCell) return;
                beginCellEdit(request);
              }}
              interactionMode={isFormulaEditing ? "rangeSelection" : "default"}
              onRangeSelectionStart={beginRangeSelection}
              onRangeSelectionChange={updateRangeSelection}
              onRangeSelectionEnd={endRangeSelection}
            />
            <CellEditorOverlay
              gridApi={gridApiRef.current}
              cell={editingCell}
              value={draft}
              onChange={(value) => {
                setDraft(value);
                draftRef.current = value;
                cellSyncTokenRef.current++;
                rangeInsertionRef.current = null;
              }}
              onCommit={(nav) => {
                void commitCellEdit(nav);
              }}
              onCancel={cancelCellEdit}
            />
          </>
        ) : (
          <GridPlaceholder />
        )}
      </div>
    </div>
  );
}

function PerfGridApp(): React.ReactElement {
  // +1 for frozen header row/col.
  const rowCount = 1_000_000 + 1;
  const colCount = 100 + 1;
  const frozenRows = 1;
  const frozenCols = 1;

  const apiRef = useRef<GridApi | null>(null);
  const provider = useMemo(() => new MockCellProvider({ rowCount, colCount }), [rowCount, colCount]);

  return (
    <div style={{ padding: 24, fontFamily: "system-ui, sans-serif" }}>
      <h1 style={{ margin: 0 }}>Formula (Web Preview)</h1>
      <p style={{ marginTop: 8, color: "#475569" }}>
        Engine: <strong data-testid="engine-status">ready (mock)</strong>
      </p>

      <div data-testid="grid" style={{ marginTop: 16, height: 560 }}>
        <CanvasGrid
          provider={provider}
          rowCount={rowCount}
          colCount={colCount}
          frozenRows={frozenRows}
          frozenCols={frozenCols}
          enableResize
          apiRef={(api) => {
            apiRef.current = api;
            (window as any).__gridApi = api;
          }}
        />
      </div>
    </div>
  );
}
