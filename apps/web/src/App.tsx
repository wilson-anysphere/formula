import {
  createEngineClient,
  isMissingGetRangeCompactError,
  type CellChange,
  type CellDataCompact,
  type CellScalar,
  type EngineClient,
} from "@formula/engine";
import { computeFillEdits, type FillSourceCell } from "@formula/fill-engine";
import type { CellRange, GridAxisSizeChange, GridViewportState } from "@formula/grid";
import { CanvasGrid, GridPlaceholder, MockCellProvider, type GridApi } from "@formula/grid";
import {
  assignFormulaReferenceColors,
  extractFormulaReferences,
  parseHtmlTableToGrid,
  parseTsvToGrid,
  range0ToA1,
  serializeGridToHtmlTable,
  serializeGridToTsv,
  toggleA1AbsoluteAtCursor,
  toA1,
  type Range0
} from "@formula/spreadsheet-frontend";
import { formatA1Range, parseGoTo, type GoToWorkbookLookup } from "../../../packages/search/index.js";
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type ClipboardEvent } from "react";

import { CellEditorOverlay } from "./CellEditorOverlay";
import { EngineCellProvider } from "./EngineCellProvider";
import { isFormulaInput, parseCellScalarInput, scalarToDisplayString } from "./cellScalar";
import { DEMO_WORKBOOK_JSON } from "./engine/documentControllerSync";

const DEMO_SHEETS = ["Sheet1", "Sheet2"] as const;

function readRootCssVar(name: string, fallback: string): string {
  if (typeof document === "undefined" || typeof getComputedStyle !== "function") return fallback;
  const value = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  return value || fallback;
}

export function App() {
  const params = typeof window !== "undefined" ? new URLSearchParams(window.location.search) : null;
  const perfMode = params?.has("perf") ?? false;
  return perfMode ? <PerfGridApp /> : <EngineDemoApp />;
}

function EngineDemoApp() {
  const [engineStatus, setEngineStatus] = useState("starting…");
  const engineRef = useRef<ReturnType<typeof createEngineClient> | null>(null);
  const [provider, setProvider] = useState<EngineCellProvider | null>(null);
  const [activeSheet, setActiveSheet] = useState<(typeof DEMO_SHEETS)[number]>(DEMO_SHEETS[0]);
  const previousSheetRef = useRef<string | null>(null);
  const activeSheetRef = useRef(activeSheet);
  activeSheetRef.current = activeSheet;
  const supportsRangeCompactRef = useRef<boolean | null>(null);

  // Persist per-sheet axis sizes (row heights / col widths). Values are stored in "base" units
  // (CSS pixels at zoom=1) so they can be reapplied consistently across zoom changes.
  const axisSizesBySheetRef = useRef(
    new Map<
      string,
      {
        cols: Map<number, number>;
        rows: Map<number, number>;
      }
    >()
  );
  const originA1BySheetRef = useRef(new Map<string, string>());

  const HEADER_ROWS = 1;
  const HEADER_COLS = 1;

  // +1 for the header row/col (implemented using frozen panes in the grid).
  const rowCount = 1_000_000 + HEADER_ROWS;
  const colCount = 100 + HEADER_COLS;

  const [frozenBySheet, setFrozenBySheet] = useState<Record<string, { frozenRows: number; frozenCols: number }>>({});
  const sheetFrozen = frozenBySheet[activeSheet] ?? { frozenRows: 0, frozenCols: 0 };
  const frozenRows = HEADER_ROWS + sheetFrozen.frozenRows;
  const frozenCols = HEADER_COLS + sheetFrozen.frozenCols;

  const inputRef = useRef<HTMLInputElement | null>(null);
  const [draft, setDraft] = useState("");
  const draftRef = useRef(draft);
  const [formulaFocused, setFormulaFocused] = useState(false);

  const cursorRef = useRef<{ start: number; end: number }>({ start: 0, end: 0 });
  const rangeInsertionRef = useRef<{ start: number; end: number } | null>(null);
  const pendingSelectionRef = useRef<{ start: number; end: number } | null>(null);
  const referenceColorByTextRef = useRef<Map<string, string>>(new Map());
  const selectedReferenceIndexRef = useRef<number | null>(null);
  const cellSyncTokenRef = useRef(0);

  const isFormulaEditing = formulaFocused && draft.trim().startsWith("=");
  const isFormulaEditingRef = useRef(isFormulaEditing);
  const headerRowOffset = HEADER_ROWS;
  const headerColOffset = HEADER_COLS;

  const gridApiRef = useRef<GridApi | null>(null);
  const [zoom, setZoom] = useState(1);
  const zoomRef = useRef(zoom);
  const gridContainerRef = useRef<HTMLDivElement | null>(null);
  const internalClipboardRef = useRef<{ tsv: string; html: string } | null>(null);
  const [activeCell, setActiveCell] = useState<{ row: number; col: number } | null>(null);
  const activeCellRef = useRef<{ row: number; col: number } | null>(null);
  const [editingCell, setEditingCell] = useState<{ row: number; col: number } | null>(null);
  const editingCellOriginalDraftRef = useRef("");
  const editingCellRef = useRef<{ row: number; col: number } | null>(null);
  activeCellRef.current = activeCell;
  editingCellRef.current = editingCell;

  const gridCellToA1Address = (cell: { row: number; col: number } | null): string | null => {
    if (!cell) return null;
    const row0 = cell.row - headerRowOffset;
    const col0 = cell.col - headerColOffset;
    if (row0 < 0 || col0 < 0) return null;
    return toA1(row0, col0);
  };

  const activeAddress = gridCellToA1Address(activeCell);

  const [activeValue, setActiveValue] = useState<CellScalar>(null);
  const defaultCellFontFamily = useMemo(() => readRootCssVar("--font-mono", "ui-monospace, monospace"), []);
  const defaultHeaderFontFamily = useMemo(() => readRootCssVar("--font-sans", "system-ui"), []);

  const focusGrid = () => {
    const host = gridContainerRef.current;
    if (!host) return;
    const grid = host.querySelector<HTMLElement>('[data-testid="canvas-grid"]');
    grid?.focus({ preventScroll: true });
  };

  const switchSheet = (delta: -1 | 1) => {
    const current = activeSheetRef.current;
    const currentIndex = DEMO_SHEETS.findIndex((sheet) => sheet === current);
    if (currentIndex < 0) {
      setActiveSheet(DEMO_SHEETS[0]);
      queueMicrotask(() => focusGrid());
      return;
    }
    const nextIndex = (currentIndex + delta + DEMO_SHEETS.length) % DEMO_SHEETS.length;
    setActiveSheet(DEMO_SHEETS[nextIndex]);
    queueMicrotask(() => focusGrid());
  };

  const setActiveSheetFrozen = (next: { frozenRows: number; frozenCols: number }) => {
    setFrozenBySheet((prev) => ({ ...prev, [activeSheet]: next }));
  };

  const handleFreezePanes = () => {
    if (!activeCell) return;
    const row0 = activeCell.row - headerRowOffset;
    const col0 = activeCell.col - headerColOffset;
    setActiveSheetFrozen({ frozenRows: Math.max(0, row0), frozenCols: Math.max(0, col0) });
  };

  const handleFreezeTopRow = () => {
    setActiveSheetFrozen({ frozenRows: 1, frozenCols: 0 });
  };

  const handleFreezeFirstColumn = () => {
    setActiveSheetFrozen({ frozenRows: 0, frozenCols: 1 });
  };

  const handleUnfreezePanes = () => {
    setActiveSheetFrozen({ frozenRows: 0, frozenCols: 0 });
  };

  const [commandPaletteOpen, setCommandPaletteOpen] = useState(false);
  const [commandPaletteQuery, setCommandPaletteQuery] = useState("");
  const [commandPaletteSelectedIndex, setCommandPaletteSelectedIndex] = useState(0);
  const commandPaletteInputRef = useRef<HTMLInputElement | null>(null);

  const closeCommandPalette = () => {
    setCommandPaletteOpen(false);
    setCommandPaletteQuery("");
    setCommandPaletteSelectedIndex(0);
  };

  const commandPaletteCommands: Array<{ id: string; title: string; run: () => void; keywords?: string[] }> = [
    {
      id: "workbook.previousSheet",
      title: "Previous Sheet",
      run: () => switchSheet(-1),
      keywords: ["sheet", "tab"],
    },
    {
      id: "workbook.nextSheet",
      title: "Next Sheet",
      run: () => switchSheet(1),
      keywords: ["sheet", "tab"],
    },
    { id: "view.freezePanes", title: "Freeze Panes", run: handleFreezePanes, keywords: ["frozen", "pane"] },
    { id: "view.freezeTopRow", title: "Freeze Top Row", run: handleFreezeTopRow, keywords: ["frozen", "row"] },
    { id: "view.freezeFirstColumn", title: "Freeze First Column", run: handleFreezeFirstColumn, keywords: ["frozen", "column"] },
    { id: "view.unfreezePanes", title: "Unfreeze Panes", run: handleUnfreezePanes, keywords: ["frozen", "pane"] },
  ];

  const filteredCommandPalette = commandPaletteCommands.filter((cmd) => {
    const query = commandPaletteQuery.trim().toLowerCase();
    if (!query) return true;
    const haystack = `${cmd.title} ${cmd.id} ${(cmd.keywords ?? []).join(" ")}`.toLowerCase();
    return haystack.includes(query);
  });

  const goToWorkbook = useMemo<GoToWorkbookLookup>(
    () => ({
      getSheet: (name: string) => {
        const trimmed = String(name ?? "").trim();
        const candidates = DEMO_SHEETS;
        const resolved = candidates.find((s) => s.toLowerCase() === trimmed.toLowerCase());
        if (!resolved) {
          throw new Error(`Unknown sheet: ${name}`);
        }
        return { name: resolved };
      },
      getTable: () => null,
      getName: () => null,
    }),
    [],
  );

  const goToSuggestion = useMemo(() => {
    const trimmed = commandPaletteQuery.trim();
    if (!trimmed) return null;
    try {
      return parseGoTo(trimmed, { workbook: goToWorkbook, currentSheetName: activeSheet });
    } catch {
      return null;
    }
  }, [commandPaletteQuery, goToWorkbook, activeSheet]);

  const commandPaletteResults = useMemo(() => {
    const trimmed = commandPaletteQuery.trim();
    const results: Array<{ id: string; title: string; secondaryText?: string; run: () => void }> = [];

    if (trimmed && goToSuggestion) {
      const resolved = `${goToSuggestion.sheetName}!${formatA1Range(goToSuggestion.range)}`;
      results.push({
        id: "goTo",
        title: `Go to ${trimmed}`,
        secondaryText: resolved,
        run: () => {
          const api = gridApiRef.current;
          if (!api) return;

          if (goToSuggestion.sheetName !== activeSheet) {
            setActiveSheet(goToSuggestion.sheetName as (typeof DEMO_SHEETS)[number]);
          }

          const { range } = goToSuggestion;
          const startRow = range.startRow + headerRowOffset;
          const startCol = range.startCol + headerColOffset;
          const endRow = range.endRow + 1 + headerRowOffset;
          const endCol = range.endCol + 1 + headerColOffset;

          if (range.startRow === range.endRow && range.startCol === range.endCol) {
            api.setSelection(startRow, startCol);
          } else {
            api.setSelectionRange({ startRow, startCol, endRow, endCol });
          }

          focusGrid();
        },
      });
    }

    for (const cmd of filteredCommandPalette) {
      results.push({ id: cmd.id, title: cmd.title, run: cmd.run });
    }

    return results;
  }, [commandPaletteQuery, goToSuggestion, filteredCommandPalette, activeSheet, headerRowOffset, headerColOffset, focusGrid]);

  const clampedCommandPaletteIndex = Math.max(
    0,
    Math.min(commandPaletteSelectedIndex, Math.max(0, commandPaletteResults.length - 1)),
  );
  const selectedCommand = commandPaletteResults[clampedCommandPaletteIndex];

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const primary = event.ctrlKey || event.metaKey;
      if (!primary || !event.shiftKey) return;
      if (event.key !== "P" && event.key !== "p") return;
      event.preventDefault();
      setCommandPaletteOpen(true);
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);

  useEffect(() => {
    const isEditableTarget = (target: EventTarget | null): boolean => {
      if (!(target instanceof Element)) return false;
      return Boolean(
        target.closest('input, textarea, select, [contenteditable=""], [contenteditable="true"]'),
      );
    };

    const onKeyDown = (event: KeyboardEvent) => {
      if (!provider) return;
      if (event.defaultPrevented) return;

      const primary = event.ctrlKey || event.metaKey;
      if (!primary || event.shiftKey || event.altKey) return;
      if (event.key !== "PageUp" && event.key !== "PageDown") return;
      if (isEditableTarget(event.target)) return;

      event.preventDefault();
      event.stopPropagation();

      switchSheet(event.key === "PageDown" ? 1 : -1);
    };

    // Capture here so sheet switching works even when the grid hasn't registered focus yet.
    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, [provider]);

  useEffect(() => {
    if (!commandPaletteOpen) return;
    queueMicrotask(() => commandPaletteInputRef.current?.focus());
  }, [commandPaletteOpen]);

  useEffect(() => {
    draftRef.current = draft;
  }, [draft]);

  useEffect(() => {
    // Reset any persisted sizing when swapping providers (e.g. importing a new workbook).
    axisSizesBySheetRef.current.clear();
  }, [provider]);

  useEffect(() => {
    zoomRef.current = zoom;
    gridApiRef.current?.setZoom(zoom);
  }, [zoom]);

  useEffect(() => {
    isFormulaEditingRef.current = isFormulaEditing;
  }, [isFormulaEditing]);

  const syncCursorFromInput = () => {
    const input = inputRef.current;
    if (!input) return;
    const start = input.selectionStart ?? input.value.length;
    const end = input.selectionEnd ?? input.value.length;
    cursorRef.current = { start, end };

    // Track when the user has explicitly selected a full reference token so we
    // can toggle click-to-select behavior (Excel UX) without getting stuck.
    if (start === end || !draftRef.current.trim().startsWith("=")) {
      selectedReferenceIndexRef.current = null;
      return;
    }
    const { references } = extractFormulaReferences(draftRef.current, start, end);
    const selected = references.find((ref) => ref.start === start && ref.end === end);
    selectedReferenceIndexRef.current = selected ? selected.index : null;
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

  const cellRangeToRange0 = (range: CellRange): Range0 | null => {
    const startRow0 = range.startRow - headerRowOffset;
    const startCol0 = range.startCol - headerColOffset;
    const endRow0Exclusive = range.endRow - headerRowOffset;
    const endCol0Exclusive = range.endCol - headerColOffset;

    if (startRow0 < 0 || startCol0 < 0) return null;
    if (endRow0Exclusive <= startRow0 || endCol0Exclusive <= startCol0) return null;

    return { startRow0, startCol0, endRow0Exclusive, endCol0Exclusive };
  };

  type FillDelta0 = { range: Range0; direction: "down" | "up" | "left" | "right" };

  const fillDeltaRange0 = (source: Range0, target: Range0): FillDelta0 | null => {
    // Fill down.
    if (
      target.startRow0 === source.startRow0 &&
      target.endRow0Exclusive > source.endRow0Exclusive &&
      target.startCol0 === source.startCol0 &&
      target.endCol0Exclusive === source.endCol0Exclusive
    ) {
      return {
        direction: "down",
        range: {
          startRow0: source.endRow0Exclusive,
          endRow0Exclusive: target.endRow0Exclusive,
          startCol0: source.startCol0,
          endCol0Exclusive: source.endCol0Exclusive
        }
      };
    }

    // Fill up.
    if (
      target.endRow0Exclusive === source.endRow0Exclusive &&
      target.startRow0 < source.startRow0 &&
      target.startCol0 === source.startCol0 &&
      target.endCol0Exclusive === source.endCol0Exclusive
    ) {
      return {
        direction: "up",
        range: {
          startRow0: target.startRow0,
          endRow0Exclusive: source.startRow0,
          startCol0: source.startCol0,
          endCol0Exclusive: source.endCol0Exclusive
        }
      };
    }

    // Fill right.
    if (
      target.startRow0 === source.startRow0 &&
      target.endRow0Exclusive === source.endRow0Exclusive &&
      target.startCol0 === source.startCol0 &&
      target.endCol0Exclusive > source.endCol0Exclusive
    ) {
      return {
        direction: "right",
        range: {
          startRow0: source.startRow0,
          endRow0Exclusive: source.endRow0Exclusive,
          startCol0: source.endCol0Exclusive,
          endCol0Exclusive: target.endCol0Exclusive
        }
      };
    }

    // Fill left.
    if (
      target.startRow0 === source.startRow0 &&
      target.endRow0Exclusive === source.endRow0Exclusive &&
      target.endCol0Exclusive === source.endCol0Exclusive &&
      target.startCol0 < source.startCol0
    ) {
      return {
        direction: "left",
        range: {
          startRow0: source.startRow0,
          endRow0Exclusive: source.endRow0Exclusive,
          startCol0: target.startCol0,
          endCol0Exclusive: source.startCol0
        }
      };
    }

    return null;
  };

  const insertOrReplaceRange = (rangeText: string, isBegin: boolean, replaceSpan?: { start: number; end: number } | null) => {
    const currentDraft = draftRef.current;

    if (!rangeInsertionRef.current || isBegin) {
      const start = replaceSpan ? replaceSpan.start : Math.min(cursorRef.current.start, cursorRef.current.end);
      const end = replaceSpan ? replaceSpan.end : Math.max(cursorRef.current.start, cursorRef.current.end);
      const nextDraft = currentDraft.slice(0, start) + rangeText + currentDraft.slice(end);

      selectedReferenceIndexRef.current = null;
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
    selectedReferenceIndexRef.current = null;
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
    const { references, activeIndex } = extractFormulaReferences(
      draftRef.current,
      cursorRef.current.start,
      cursorRef.current.end
    );
    const active = activeIndex == null ? null : references[activeIndex] ?? null;
    insertOrReplaceRange(ref, true, active ? { start: active.start, end: active.end } : null);
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
    const api = gridApiRef.current;
    if (!api) return;

    if (!isFormulaEditing) {
      api.setReferenceHighlights(null);
      referenceColorByTextRef.current = new Map();
      return;
    }

    const cursor = cursorRef.current;
    const { references, activeIndex } = extractFormulaReferences(draftRef.current, cursor.start, cursor.end);
    const { colored, nextByText } = assignFormulaReferenceColors(references, referenceColorByTextRef.current);
    referenceColorByTextRef.current = nextByText;

    const highlights = colored
      .filter((ref) => {
        if (!ref.range.sheet) return true;
        return ref.range.sheet.toLowerCase() === activeSheet.toLowerCase();
      })
      .map((ref) => ({
        range: {
          startRow: ref.range.startRow + headerRowOffset,
          endRow: ref.range.endRow + 1 + headerRowOffset,
          startCol: ref.range.startCol + headerColOffset,
          endCol: ref.range.endCol + 1 + headerColOffset
        },
        color: ref.color,
        active: activeIndex != null && ref.index === activeIndex
      }));

    api.setReferenceHighlights(highlights.length > 0 ? highlights : null);
  }, [activeSheet, draft, headerColOffset, headerRowOffset, isFormulaEditing]);

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

  const syncEngineOriginFromViewport = useCallback(
    (viewport: GridViewportState | null | undefined, sheet: string) => {
      const engine = engineRef.current;
      if (!engine) return;
      if (!viewport) return;

      const row0 = Math.max(0, viewport.main.rows.start - headerRowOffset);
      const col0 = Math.max(0, viewport.main.cols.start - headerColOffset);
      const originA1 = toA1(row0, col0);

      const prev = originA1BySheetRef.current.get(sheet);
      if (prev === originA1) return;
      originA1BySheetRef.current.set(sheet, originA1);

      // Fire-and-forget; scroll can produce high-frequency events.
      if (typeof (engine as any).setSheetOrigin === "function") {
        void (engine as any).setSheetOrigin(sheet, originA1);
      }
    },
    [headerColOffset, headerRowOffset],
  );

  useEffect(() => {
    if (!provider) return;

    provider.setSheet(activeSheet);

    const previousSheet = previousSheetRef.current;
    previousSheetRef.current = activeSheet;
    if (previousSheet && previousSheet !== activeSheet) {
      void provider.recalculate(activeSheet);
    }

    // Restore per-sheet row/col sizes. CanvasGridRenderer keeps axis sizes in-memory, so without
    // this we'd end up carrying widths/heights from the previous sheet into the new one.
    const api = gridApiRef.current;
    if (!api) return;

    const getSheetSizes = (sheet: string) => {
      let entry = axisSizesBySheetRef.current.get(sheet);
      if (!entry) {
        entry = { cols: new Map(), rows: new Map() };
        axisSizesBySheetRef.current.set(sheet, entry);
      }
      return entry;
    };

    const nextSheetSizes = getSheetSizes(activeSheet);

    const zoom = api.getZoom();

    const cols = new Map<number, number>();
    for (const [col, base] of nextSheetSizes.cols) {
      cols.set(col, base * zoom);
    }

    const rows = new Map<number, number>();
    for (const [row, base] of nextSheetSizes.rows) {
      rows.set(row, base * zoom);
    }

    api.applyAxisSizeOverrides({ rows, cols }, { resetUnspecified: true });
  }, [provider, activeSheet]);

  // CanvasGrid's `onScroll` callback intentionally doesn't fire for the initial
  // scroll baseline. Sync the starting origin so `INFO("origin")` is correct
  // immediately after the sheet/view mounts or frozen panes change.
  useEffect(() => {
    if (!provider) return;
    const api = gridApiRef.current;
    if (!api) return;
    syncEngineOriginFromViewport(api.getViewportState(), activeSheet);
  }, [provider, activeSheet, frozenRows, frozenCols, syncEngineOriginFromViewport]);

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

  const beginCellEdit = (request: { row: number; col: number; initialKey?: string }) => {
    if (!provider) return;

    // Ensure React state matches the cell the grid intends to edit, even if the
    // selection hasn't changed (e.g., F2 or type-to-edit).
    activeCellRef.current = { row: request.row, col: request.col };
    setActiveCell({ row: request.row, col: request.col });
    gridApiRef.current?.scrollToCell(request.row, request.col, { align: "auto", padding: 8 });

    editingCellRef.current = { row: request.row, col: request.col };
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
    editingCellRef.current = null;
    setEditingCell(null);
    draftRef.current = original;
    setDraft(original);
    requestAnimationFrame(() => focusGrid());
  };

  const commitCellEdit = async (nav: { deltaRow: number; deltaCol: number }) => {
    if (!editingCell) return;
    const from = editingCell;
    await commitDraft();
    editingCellRef.current = null;
    setEditingCell(null);

    const nextRow = Math.max(0, Math.min(rowCount - 1, from.row + nav.deltaRow));
    const nextCol = Math.max(0, Math.min(colCount - 1, from.col + nav.deltaCol));
    requestAnimationFrame(() => {
      gridApiRef.current?.setSelection(nextRow, nextCol);
      gridApiRef.current?.scrollToCell(nextRow, nextCol, { align: "auto", padding: 8 });
      focusGrid();
    });
  };

  const getRangeInputValues = async (
    engine: EngineClient,
    rangeA1: string,
    sheet: string
  ): Promise<CellDataCompact[][]> => {
    if (supportsRangeCompactRef.current !== false && typeof engine.getRangeCompact === "function") {
      try {
        const compact = await engine.getRangeCompact(rangeA1, sheet);
        supportsRangeCompactRef.current = true;
        return compact;
      } catch (err) {
        if (!isMissingGetRangeCompactError(err)) {
          throw err;
        }
        supportsRangeCompactRef.current = false;
      }
    }

    const legacy = await engine.getRange(rangeA1, sheet);
    return legacy.map((row) =>
      row.map((cell): CellDataCompact => [cell?.input ?? null, cell?.value ?? null]),
    );
  };

  const handleFillCommit = async (event: { sourceRange: CellRange; targetRange: CellRange; mode: "copy" | "series" | "formulas" }) => {
    const engine = engineRef.current;
    if (!engine || !provider) return;

    const source0 = cellRangeToRange0(event.sourceRange);
    const target0 = cellRangeToRange0(event.targetRange);
    if (!source0 || !target0) return;

    const sourceA1 = range0ToA1(source0);
    const targetA1 = range0ToA1(target0);

    // For formula-aware fill modes, delegate to the engine's Fill operation so reference rewriting
    // (e.g., full-row/column ranges, structured refs) matches Excel semantics.
    if (event.mode === "formulas" || event.mode === "series") {
      const seriesUpdates: Array<{ address: string; value: CellScalar; sheet: string }> = [];

      if (event.mode === "series") {
        const sourceMatrix = await getRangeInputValues(engine, sourceA1, activeSheet);
        const sourceCells: FillSourceCell[][] = sourceMatrix.map((row) =>
          row.map((cell) => ({ input: cell[0], value: cell[1] })),
        );

        const { edits } = computeFillEdits({
          sourceRange: {
            startRow: source0.startRow0,
            endRow: source0.endRow0Exclusive,
            startCol: source0.startCol0,
            endCol: source0.endCol0Exclusive
          },
          targetRange: {
            startRow: target0.startRow0,
            endRow: target0.endRow0Exclusive,
            startCol: target0.startCol0,
            endCol: target0.endCol0Exclusive
          },
          sourceCells,
          mode: "series"
        });

        for (const edit of edits) {
          // Ignore formula edits; the engine fill already copied and rewrote those formulas.
          if (typeof edit.value === "string" && edit.value.trimStart().startsWith("=")) continue;
          seriesUpdates.push({ sheet: activeSheet, address: toA1(edit.row, edit.col), value: edit.value as CellScalar });
        }
      }

      const fillResult = await engine.applyOperation({
        type: "Fill",
        sheet: activeSheet,
        src: sourceA1,
        dst: targetA1
      });

      if (seriesUpdates.length > 0) {
        await engine.setCells(seriesUpdates);
      }

      const changes = await engine.recalculate(activeSheet);

      const directChanges: CellChange[] = [
        ...fillResult.changedCells
          .filter((change) => change.after == null || change.after.formula == null)
          .map((change) => ({
            sheet: change.sheet,
            address: change.address,
            value: (change.after?.value ?? null) as CellScalar
          })),
        ...seriesUpdates.map((update) => ({ sheet: update.sheet, address: update.address, value: update.value }))
      ];

      provider.applyRecalcChanges(directChanges.length > 0 ? [...changes, ...directChanges] : changes);

      // If the user clicks a new cell while the fill operation is still
      // fetching/applying changes, the first read can race and show stale data.
      // Refresh the currently active cell after the commit completes.
      const currentActiveAddress = gridCellToA1Address(activeCellRef.current ?? activeCell);
      if (!currentActiveAddress || isFormulaEditingRef.current || editingCellRef.current) return;
      await syncFormulaBar(currentActiveAddress);
      return;
    }

    // `copy` mode fills formulas as values (no formula shifting), so the JS fill engine is fine.
    const sourceMatrix = await getRangeInputValues(engine, sourceA1, activeSheet);
    const sourceCells: FillSourceCell[][] = sourceMatrix.map((row) =>
      row.map((cell) => ({ input: cell[0], value: cell[1] })),
    );

    const { edits } = computeFillEdits({
      sourceRange: {
        startRow: source0.startRow0,
        endRow: source0.endRow0Exclusive,
        startCol: source0.startCol0,
        endCol: source0.endCol0Exclusive
      },
      targetRange: {
        startRow: target0.startRow0,
        endRow: target0.endRow0Exclusive,
        startCol: target0.startCol0,
        endCol: target0.endCol0Exclusive
      },
      sourceCells,
      mode: "copy"
    });

    if (edits.length === 0) return;

    await engine.setCells(
      edits.map((edit) => ({
        sheet: activeSheet,
        address: toA1(edit.row, edit.col),
        value: edit.value as CellScalar
      }))
    );

    const changes = await engine.recalculate(activeSheet);
    const directChanges: CellChange[] = edits.map((edit) => ({
      sheet: activeSheet,
      address: toA1(edit.row, edit.col),
      value: edit.value as CellScalar
    }));

    provider.applyRecalcChanges(directChanges.length > 0 ? [...changes, ...directChanges] : changes);

    // If the user clicks a new cell while the fill operation is still
    // fetching/applying changes, the first read can race and show stale data.
    // Refresh the currently active cell after the commit completes.
    const currentActiveAddress = gridCellToA1Address(activeCellRef.current ?? activeCell);
    if (!currentActiveAddress || isFormulaEditingRef.current || editingCellRef.current) return;
    await syncFormulaBar(currentActiveAddress);
  };

  const onSelectionChange = (cell: { row: number; col: number } | null) => {
    if (isFormulaEditing || editingCell) return;
    activeCellRef.current = cell;
    setActiveCell(cell);
  };

  const onAxisSizeChange = (change: GridAxisSizeChange) => {
    const sheet = activeSheetRef.current;
    let entry = axisSizesBySheetRef.current.get(sheet);
    if (!entry) {
      entry = { cols: new Map(), rows: new Map() };
      axisSizesBySheetRef.current.set(sheet, entry);
    }

    // Store sizes in base units (zoom=1) so we can restore regardless of current zoom.
    const baseSize = change.size / change.zoom;
    const baseDefault = change.defaultSize / change.zoom;
    const isDefault = Math.abs(baseSize - baseDefault) < 1e-6;

    if (change.kind === "col") {
      if (isDefault) entry.cols.delete(change.index);
      else entry.cols.set(change.index, baseSize);
    } else {
      if (isDefault) entry.rows.delete(change.index);
      else entry.rows.set(change.index, baseSize);
    }
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

  const onFillHandleCommit = async ({ source, target }: { source: CellRange; target: CellRange }) => {
    const engine = engineRef.current;
    if (!engine || !provider) return;

    const source0 = cellRangeToRange0(source);
    const target0 = cellRangeToRange0(target);
    if (!source0 || !target0) return;

    const fillDelta = fillDeltaRange0(source0, target0);
    if (!fillDelta) return;
    const { range: fillArea0, direction } = fillDelta;

    const sourceCells = await getRangeInputValues(engine, range0ToA1(source0), activeSheet);
    const sourceHeight = source0.endRow0Exclusive - source0.startRow0;
    const sourceWidth = source0.endCol0Exclusive - source0.startCol0;

    const fillSeries = (() => {
      const eps = 1e-9;

      type Series =
        | { kind: "number"; start: number; step: number }
        | { kind: "text"; prefix: string; suffix: string; start: number; step: number; padWidth: number };

      const seriesFromNumbers = (values: number[]): Series | null => {
        if (values.length < 2) return null;
        const step = values[1] - values[0];
        for (let i = 2; i < values.length; i++) {
          const delta = values[i] - values[i - 1]!;
          if (Math.abs(delta - step) > eps) return null;
        }
        return { kind: "number", start: values[0]!, step };
      };

      const seriesFromText = (values: string[]): Series | null => {
        if (values.length < 2) return null;
        const parsed = values.map((value) => /^(.*?)(\d+)([^0-9]*)$/.exec(value));
        if (parsed.some((match) => !match)) return null;

        const [first] = parsed as RegExpExecArray[];
        const prefix = first![1] ?? "";
        const suffix = first![3] ?? "";
        const width = (first![2] ?? "").length;
        if (width === 0) return null;

        const nums: number[] = [];
        for (const match of parsed as RegExpExecArray[]) {
          if ((match[1] ?? "") !== prefix) return null;
          if ((match[3] ?? "") !== suffix) return null;
          if ((match[2] ?? "").length !== width) return null;
          nums.push(Number.parseInt(match[2]!, 10));
        }

        const step = nums[1]! - nums[0]!;
        for (let i = 2; i < nums.length; i++) {
          const delta = nums[i]! - nums[i - 1]!;
          if (delta !== step) return null;
        }

        return { kind: "text", prefix, suffix, start: nums[0]!, step, padWidth: width };
      };

      const isVertical = direction === "down" || direction === "up";
      if (isVertical && sourceHeight >= 2) {
        const columns: Array<Series | null> = [];
        for (let c = 0; c < sourceWidth; c++) {
          let kind: "number" | "text" | null = null;
          const numberValues: number[] = [];
          const textValues: string[] = [];
          for (let r = 0; r < sourceHeight; r++) {
            const input = (sourceCells[r]?.[c]?.[0] ?? null) as CellScalar;
            if (typeof input === "number" && Number.isFinite(input)) {
              if (kind === "text") {
                kind = null;
                break;
              }
              kind = "number";
              numberValues.push(input);
              continue;
            }

            if (typeof input === "string" && !isFormulaInput(input)) {
              if (kind === "number") {
                kind = null;
                break;
              }
              kind = "text";
              textValues.push(input);
              continue;
            }

            kind = null;
            break;
          }
          columns.push(kind === "number" ? seriesFromNumbers(numberValues) : kind === "text" ? seriesFromText(textValues) : null);
        }
        return columns.some((series) => series) ? ({ axis: "vertical" as const, columns } as const) : null;
      }

      const isHorizontal = direction === "left" || direction === "right";
      if (isHorizontal && sourceWidth >= 2) {
        const rows: Array<Series | null> = [];
        for (let r = 0; r < sourceHeight; r++) {
          let kind: "number" | "text" | null = null;
          const numberValues: number[] = [];
          const textValues: string[] = [];
          for (let c = 0; c < sourceWidth; c++) {
            const input = (sourceCells[r]?.[c]?.[0] ?? null) as CellScalar;
            if (typeof input === "number" && Number.isFinite(input)) {
              if (kind === "text") {
                kind = null;
                break;
              }
              kind = "number";
              numberValues.push(input);
              continue;
            }

            if (typeof input === "string" && !isFormulaInput(input)) {
              if (kind === "number") {
                kind = null;
                break;
              }
              kind = "text";
              textValues.push(input);
              continue;
            }

            kind = null;
            break;
          }
          rows.push(kind === "number" ? seriesFromNumbers(numberValues) : kind === "text" ? seriesFromText(textValues) : null);
        }
        return rows.some((series) => series) ? ({ axis: "horizontal" as const, rows } as const) : null;
      }

      return null;
    })();

    // Always delegate to the engine's Fill operation so formula rewriting matches Excel semantics
    // (structured refs, external refs, etc). We still layer custom number/text series behavior on
    // top by overwriting those cells with computed values below.
    const fillResult = await engine.applyOperation({
      type: "Fill",
      sheet: activeSheet,
      src: range0ToA1(source0),
      dst: range0ToA1(target0)
    });

    const seriesUpdates: Array<{ address: string; value: CellScalar; sheet: string }> = [];
    for (let row0 = fillArea0.startRow0; row0 < fillArea0.endRow0Exclusive; row0++) {
      for (let col0 = fillArea0.startCol0; col0 < fillArea0.endCol0Exclusive; col0++) {
        if (fillSeries?.axis === "vertical") {
          const series = fillSeries.columns[col0 - source0.startCol0];
          if (series) {
            const k = row0 - source0.startRow0;
            let value: CellScalar;
            if (series.kind === "number") {
              value = series.start + series.step * k;
            } else {
              const n = series.start + series.step * k;
              const digits = Math.abs(n).toString().padStart(series.padWidth, "0");
              value = `${series.prefix}${n < 0 ? "-" : ""}${digits}${series.suffix}`;
            }
            seriesUpdates.push({ address: toA1(row0, col0), value, sheet: activeSheet });
            continue;
          }
        }

        if (fillSeries?.axis === "horizontal") {
          const series = fillSeries.rows[row0 - source0.startRow0];
          if (series) {
            const k = col0 - source0.startCol0;
            let value: CellScalar;
            if (series.kind === "number") {
              value = series.start + series.step * k;
            } else {
              const n = series.start + series.step * k;
              const digits = Math.abs(n).toString().padStart(series.padWidth, "0");
              value = `${series.prefix}${n < 0 ? "-" : ""}${digits}${series.suffix}`;
            }
            seriesUpdates.push({ address: toA1(row0, col0), value, sheet: activeSheet });
            continue;
          }
        }
      }
    }

    if (seriesUpdates.length > 0) {
      await engine.setCells(seriesUpdates);
    }

    const changes = await engine.recalculate(activeSheet);

    const directChanges: CellChange[] = [
      ...fillResult.changedCells
        .filter((change) => change.after == null || change.after.formula == null)
        .map((change) => ({
          sheet: change.sheet,
          address: change.address,
          value: (change.after?.value ?? null) as CellScalar
        })),
      ...seriesUpdates.map((update) => ({ sheet: update.sheet, address: update.address, value: update.value }))
    ];

    provider.applyRecalcChanges(directChanges.length > 0 ? [...changes, ...directChanges] : changes);

    // If the user clicks a new cell while the fill operation is still
    // fetching/applying changes, the first read can race and show stale data.
    // Refresh the currently active cell after the commit completes.
    const currentActiveAddress = gridCellToA1Address(activeCellRef.current ?? activeCell);
    if (!currentActiveAddress || isFormulaEditingRef.current || editingCellRef.current) return;
    await syncFormulaBar(currentActiveAddress);
  };

  const handleGridCopy = (event: ClipboardEvent<HTMLDivElement>) => {
    if (editingCell) return;
    if (!provider) return;
    const range = getCopyRange();
    if (!range) return;

    const rows = range.endRow - range.startRow;
    const cols = range.endCol - range.startCol;
    const cellCount = rows * cols;
    const maxClipboardCells = 100_000;
    if (cellCount > maxClipboardCells) return;

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
      event.clipboardData?.getData("text/plain") ||
      event.clipboardData?.getData("text/tab-separated-values") ||
      event.clipboardData?.getData("text/tsv") ||
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
      // Avoid `Math.max(...rows.map(...))` spread: a tall paste can have tens of thousands of rows,
      // which would exceed JS engines' argument limits.
      let pasteCols = 0;
      for (const row of grid) {
        if (row.length > pasteCols) pasteCols = row.length;
      }
      if (pasteRows === 0 || pasteCols === 0) return;

      const maxRows = Math.max(0, Math.min(pasteRows, rowCount - selection.row));
      const maxCols = Math.max(0, Math.min(pasteCols, colCount - selection.col));
      if (maxRows === 0 || maxCols === 0) return;

      const values: CellScalar[][] = [];
      const directChanges: CellChange[] = [];

      for (let r = 0; r < maxRows; r++) {
        const row = grid[r] ?? [];
        const outRow: CellScalar[] = [];
        for (let c = 0; c < maxCols; c++) {
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
        endRow0Exclusive: startRow0 + maxRows,
        endCol0Exclusive: startCol0 + maxCols
      });

      await engine.setRange(rangeA1, values, activeSheet);
      const changes = await engine.recalculate(activeSheet);
      provider.applyRecalcChanges(directChanges.length > 0 ? [...changes, ...directChanges] : changes);

      api.setSelectionRange({
        startRow: selection.row,
        endRow: selection.row + maxRows,
        startCol: selection.col,
        endCol: selection.col + maxCols
      });

      await syncFormulaBar(toA1(startRow0, startCol0));
    })().catch(() => {});
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
      })().catch(() => {});
    }
  };

  return (
    <div style={{ padding: 24, fontFamily: "system-ui, sans-serif" }}>
      <h1 style={{ margin: 0 }}>Formula (Web Preview)</h1>
      <p style={{ marginTop: 8, color: "var(--formula-grid-cell-text, #475569)", opacity: 0.75 }}>
        Engine: <strong data-testid="engine-status">{engineStatus}</strong>
      </p>

      <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8 }}>
        Sheet:
        <select
          data-testid="sheet-switcher"
          value={activeSheet}
          onChange={(e) => {
            // Treat sheet switching as a navigation action: after the change, restore
            // focus to the grid so keyboard workflows (F2, typing, arrows) keep working.
            const nextSheet = e.target.value;
            if (!DEMO_SHEETS.includes(nextSheet as (typeof DEMO_SHEETS)[number])) return;
            setActiveSheet(nextSheet as (typeof DEMO_SHEETS)[number]);
            queueMicrotask(() => focusGrid());
          }}
          style={{ padding: "4px 6px" }}
          disabled={!provider}
        >
          {DEMO_SHEETS.map((sheet) => (
            <option key={sheet} value={sheet}>
              {sheet}
            </option>
          ))}
        </select>
      </label>

      <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 12, flexWrap: "wrap" }}>
        <button type="button" onClick={handleFreezePanes} disabled={!provider || !activeCell}>
          Freeze Panes
        </button>
        <button type="button" onClick={handleFreezeTopRow} disabled={!provider}>
          Freeze Top Row
        </button>
        <button type="button" onClick={handleFreezeFirstColumn} disabled={!provider}>
          Freeze First Column
        </button>
        <button type="button" onClick={handleUnfreezePanes} disabled={!provider}>
          Unfreeze Panes
        </button>
        <span style={{ fontSize: 12, color: "#475569" }}>
          Frozen: {sheetFrozen.frozenRows} rows, {sheetFrozen.frozenCols} cols
        </span>
      </div>

      <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 12 }}>
        Import XLSX/XLSM:
        <input
          type="file"
          accept=".xlsx,.xlsm"
          data-testid="xlsx-file-input"
          disabled={!provider}
          onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            if (!file) return;
            const engine = engineRef.current;
            if (!engine) return;

            setEngineStatus("importing workbook…");
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
                setEngineStatus(`ready (imported workbook; B1=${b1.value === null ? "" : String(b1.value)})`);
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

      <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 12 }}>
        Zoom:
        <input
          type="range"
          min={0.25}
          max={4}
          step={0.05}
          value={zoom}
          onChange={(event) => setZoom(event.currentTarget.valueAsNumber)}
          disabled={!provider}
          style={{ width: 200 }}
        />
        <span style={{ width: 48, textAlign: "right" }}>{Math.round(zoom * 100)}%</span>
      </label>

      <div style={{ marginTop: 16 }}>
        <label style={{ display: "block", fontSize: 12, color: "var(--formula-grid-cell-text, #64748b)", opacity: 0.75 }} htmlFor="formula-input">
          Formula
        </label>
        <div style={{ marginTop: 4, display: "flex", alignItems: "center", gap: 8 }}>
          <div
            style={{
              width: 64,
              fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace",
              fontSize: 12,
              color: "var(--formula-grid-header-text, #0f172a)"
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
              selectedReferenceIndexRef.current = null;
              const input = event.currentTarget;
              queueMicrotask(() => {
                input.select();
                syncCursorFromInput();
              });
            }}
            onBlur={() => {
              setFormulaFocused(false);
              rangeInsertionRef.current = null;
              selectedReferenceIndexRef.current = null;
            }}
            onChange={(event) => {
              const value = event.currentTarget.value;
              setDraft(value);
              draftRef.current = value;
              cellSyncTokenRef.current++;
              rangeInsertionRef.current = null;
              selectedReferenceIndexRef.current = null;
              cursorRef.current = {
                start: event.currentTarget.selectionStart ?? value.length,
                end: event.currentTarget.selectionEnd ?? value.length
              };
            }}
            onKeyDown={(event) => {
              // Excel UX: F4 cycles the absolute/relative state of the reference token at the caret,
              // but only while editing a formula (draft starts with "=").
              if (
                event.key === "F4" &&
                !event.altKey &&
                !event.ctrlKey &&
                !event.metaKey &&
                event.currentTarget.value.trim().startsWith("=")
              ) {
                event.preventDefault();
                const input = event.currentTarget;
                const value = input.value;
                const start = input.selectionStart ?? value.length;
                const end = input.selectionEnd ?? value.length;
                const toggled = toggleA1AbsoluteAtCursor(value, start, end);
                if (!toggled) return;

                rangeInsertionRef.current = null;
                draftRef.current = toggled.text;
                cursorRef.current = { start: toggled.cursorStart, end: toggled.cursorEnd };
                pendingSelectionRef.current = { start: toggled.cursorStart, end: toggled.cursorEnd };

                // Track full-token selections so click-to-select behaves like Excel.
                // When the selection is collapsed (caret), treat it as "not selected".
                if (toggled.cursorStart === toggled.cursorEnd) {
                  selectedReferenceIndexRef.current = null;
                } else {
                  const min = Math.min(toggled.cursorStart, toggled.cursorEnd);
                  const max = Math.max(toggled.cursorStart, toggled.cursorEnd);
                  const { references } = extractFormulaReferences(toggled.text, min, max);
                  const selected = references.find((ref) => ref.start === min && ref.end === max);
                  selectedReferenceIndexRef.current = selected ? selected.index : null;
                }

                setDraft(toggled.text);
                return;
              }

              if (event.key !== "Enter") return;
              event.preventDefault();
              void commitDraft();
            }}
            onClick={(event) => {
              const input = event.currentTarget;
              const prevSelectedReference = selectedReferenceIndexRef.current;
              syncCursorFromInput();

              if (!isFormulaEditingRef.current) return;

              const start = input.selectionStart ?? input.value.length;
              const end = input.selectionEnd ?? input.value.length;
              if (start !== end) return;

              const { references, activeIndex } = extractFormulaReferences(draftRef.current, start, end);
              const active = activeIndex == null ? null : references[activeIndex] ?? null;
              if (!active) return;

              // Excel UX: click selects the whole reference token for easy range replacement.
              // Clicking again inside the same token toggles back to a caret so users can
              // manually edit within the reference.
              if (prevSelectedReference === activeIndex) {
                selectedReferenceIndexRef.current = null;
                return;
              }

              input.setSelectionRange(active.start, active.end);
              cursorRef.current = { start: active.start, end: active.end };
              selectedReferenceIndexRef.current = activeIndex;
            }}
            onKeyUp={syncCursorFromInput}
            onSelect={syncCursorFromInput}
            disabled={!provider || !activeAddress}
            style={{
              flex: 1,
              padding: "8px 10px",
              border: "1px solid var(--formula-grid-line, #cbd5e1)",
              borderRadius: 6,
              fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace",
              fontSize: 14
            }}
          />
          <div style={{ minWidth: 120, fontSize: 12, color: "var(--formula-grid-cell-text, #475569)", opacity: 0.75 }}>
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
                headerRows={headerRowOffset}
                headerCols={headerColOffset}
                frozenRows={frozenRows}
                frozenCols={frozenCols}
                defaultCellFontFamily={defaultCellFontFamily}
                defaultHeaderFontFamily={defaultHeaderFontFamily}
                enableResize
                onAxisSizeChange={onAxisSizeChange}
                onZoomChange={setZoom}
                onScroll={(_scroll, viewport) => {
                  syncEngineOriginFromViewport(viewport, activeSheetRef.current);
                }}
                apiRef={(api) => {
                  gridApiRef.current = api;
                  if (api) api.setZoom(zoomRef.current);
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
              onFillCommit={editingCell || isFormulaEditing ? undefined : handleFillCommit}
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

      {commandPaletteOpen ? (
        <div
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(15, 23, 42, 0.25)",
            display: "flex",
            justifyContent: "center",
            alignItems: "flex-start",
            paddingTop: 80,
            zIndex: 1000
          }}
          onMouseDown={(event) => {
            if (event.target === event.currentTarget) closeCommandPalette();
          }}
        >
          <div
            data-testid="command-palette"
            style={{
              width: 420,
              maxWidth: "92vw",
              borderRadius: 12,
              border: "1px solid #cbd5e1",
              background: "#ffffff",
              padding: 12,
              boxShadow: "0 10px 30px rgba(15, 23, 42, 0.18)"
            }}
          >
            <input
              ref={commandPaletteInputRef}
              type="text"
              value={commandPaletteQuery}
              placeholder="Type a command…"
              onChange={(event) => {
                setCommandPaletteQuery(event.currentTarget.value);
                setCommandPaletteSelectedIndex(0);
              }}
              onKeyDown={(event) => {
                if (event.key === "Escape") {
                  event.preventDefault();
                  closeCommandPalette();
                  return;
                }
                if (event.key === "ArrowDown") {
                  event.preventDefault();
                  setCommandPaletteSelectedIndex((prev) => prev + 1);
                  return;
                }
                if (event.key === "ArrowUp") {
                  event.preventDefault();
                  setCommandPaletteSelectedIndex((prev) => Math.max(0, prev - 1));
                  return;
                }
                if (event.key === "Enter") {
                  event.preventDefault();
                  closeCommandPalette();
                  selectedCommand?.run();
                }
              }}
              style={{
                width: "100%",
                padding: "10px 12px",
                borderRadius: 10,
                border: "1px solid #cbd5e1",
                fontSize: 14,
                outline: "none"
              }}
            />
            <ul style={{ margin: "10px 0 0", padding: 0, listStyle: "none" }}>
              {commandPaletteResults.map((cmd, idx) => (
                <li
                  key={cmd.id}
                  style={{
                    padding: "8px 10px",
                    borderRadius: 10,
                    cursor: "pointer",
                    background: idx === clampedCommandPaletteIndex ? "#e0e7ff" : "transparent"
                  }}
                  onMouseMove={() => setCommandPaletteSelectedIndex(idx)}
                  onMouseDown={(event) => {
                    event.preventDefault();
                    closeCommandPalette();
                    cmd.run();
                  }}
                >
                  <div>{cmd.title}</div>
                  {cmd.secondaryText ? (
                    <div style={{ marginTop: 2, fontSize: 12, color: "#64748b", fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace" }}>
                      {cmd.secondaryText}
                    </div>
                  ) : null}
                </li>
              ))}
            </ul>
          </div>
        </div>
      ) : null}
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
  const defaultCellFontFamily = useMemo(() => readRootCssVar("--font-mono", "ui-monospace, monospace"), []);
  const defaultHeaderFontFamily = useMemo(() => readRootCssVar("--font-sans", "system-ui"), []);

  return (
    <div style={{ padding: 24, fontFamily: "system-ui, sans-serif" }}>
      <h1 style={{ margin: 0 }}>Formula (Web Preview)</h1>
      <p style={{ marginTop: 8, color: "var(--formula-grid-cell-text, #475569)", opacity: 0.75 }}>
        Engine: <strong data-testid="engine-status">ready (mock)</strong>
      </p>

      <div data-testid="grid" style={{ marginTop: 16, height: 560 }}>
        <CanvasGrid
          provider={provider}
          rowCount={rowCount}
          colCount={colCount}
          frozenRows={frozenRows}
          frozenCols={frozenCols}
          defaultCellFontFamily={defaultCellFontFamily}
          defaultHeaderFontFamily={defaultHeaderFontFamily}
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
