import React, { useCallback, useEffect, useMemo, useState } from "react";

import { colToName } from "@formula/spreadsheet-frontend/a1";
import { getSheetNameValidationErrorMessage } from "@formula/workbook-backend";

import { t, tWithVars } from "../../i18n/index.js";
import { parseA1, parseRangeA1 } from "../../document/coords.js";

import type { PivotTableConfig } from "./types";
import { PivotBuilderPanel } from "./PivotBuilderPanel.js";
import { toRustPivotConfig } from "./pivotConfigMapping.js";

import type { PivotTableSummary } from "../../tauri/pivotBackend.js";
import { TauriPivotBackend } from "../../tauri/pivotBackend.js";
import { applyPivotCellUpdates } from "../../pivots/applyUpdates.js";
import * as nativeDialogs from "../../tauri/nativeDialogs.js";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver";
import { formatSheetNameForA1 } from "../../sheet/formatSheetNameForA1.js";

type RangeRect = { startRow: number; startCol: number; endRow: number; endCol: number };

type SelectionSnapshot = { sheetId: string; range: RangeRect };

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

type Props = {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  getDocumentController: () => any;
  getActiveSheetId?: () => string;
  getSelection?: () => SelectionSnapshot | null;
  invoke?: TauriInvoke;
  drainBackendSync?: () => Promise<void>;
  sheetNameResolver?: SheetNameResolver | null;
};

function cellToA1(row: number, col: number): string {
  return `${colToName(col)}${row + 1}`;
}

function rangeToA1(range: RangeRect): string {
  const start = cellToA1(range.startRow, range.startCol);
  const end = cellToA1(range.endRow, range.endCol);
  return start === end ? start : `${start}:${end}`;
}

function normalizeHeaderValue(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number") {
    if (Number.isFinite(value) && Math.abs(value - Math.round(value)) < Number.EPSILON) {
      return String(Math.round(value));
    }
    return String(value);
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "object") {
    const maybeText = (value as any).text;
    if (typeof maybeText === "string") return maybeText;
    try {
      return JSON.stringify(value);
    } catch {
      return String(value);
    }
  }
  return String(value);
}

function dedupeStrings(values: string[]): { ok: boolean; values: string[] } {
  const seen = new Set<string>();
  for (const v of values) {
    if (seen.has(v)) return { ok: false, values };
    seen.add(v);
  }
  return { ok: true, values };
}

function keyPartForCell(state: { value: unknown; formula: string | null }, uniqueSalt: string): string {
  if (state.formula) {
    // We don't have computed values here; treat formula cells as "could be unique"
    // to avoid under-estimating the pivot output size when checking destination overlap.
    return `formula:${state.formula}:${uniqueSalt}`;
  }
  const v = state.value ?? null;
  if (v == null) return "blank";
  if (typeof v === "string") return `s:${v}`;
  if (typeof v === "number") return `n:${v}`;
  if (typeof v === "boolean") return `b:${v ? 1 : 0}`;
  try {
    return `j:${JSON.stringify(v)}`;
  } catch {
    return `o:${String(v)}`;
  }
}

function estimatePivotOutputRect(params: {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  document: any;
  sheetId: string;
  source: RangeRect;
  availableFields: string[];
  config: PivotTableConfig;
  destination: { sheetId: string; startRow: number; startCol: number };
}): { startRow: number; startCol: number; endRow: number; endCol: number; cellCount: number } {
  const dataRowCount = Math.max(0, params.source.endRow - params.source.startRow); // exclude header row
  const valueFieldCount = params.config.valueFields.length;
  const rowFieldCount = params.config.rowFields.length;
  const colFieldCount = params.config.columnFields.length;

  const colByField = new Map<string, number>();
  for (let i = 0; i < params.availableFields.length; i += 1) {
    colByField.set(params.availableFields[i]!, params.source.startCol + i);
  }

  const rowKeyCount = (() => {
    if (dataRowCount === 0) return 0;
    if (rowFieldCount === 0) return 1;
    const keys = new Set<string>();
    for (let rOff = 0; rOff < dataRowCount; rOff += 1) {
      const row = params.source.startRow + 1 + rOff;
      const parts: string[] = [];
      for (const f of params.config.rowFields) {
        const col = colByField.get(f.sourceField);
        if (col == null) continue;
        const state = params.document.getCell(params.sheetId, { row, col }) as { value: unknown; formula: string | null };
        parts.push(keyPartForCell(state, `${row},${col}`));
      }
      keys.add(parts.join("\u0000"));
    }
    return Math.max(1, keys.size);
  })();

  const colKeyCount = (() => {
    if (dataRowCount === 0) return 0;
    if (colFieldCount === 0) return 1;
    const keys = new Set<string>();
    for (let rOff = 0; rOff < dataRowCount; rOff += 1) {
      const row = params.source.startRow + 1 + rOff;
      const parts: string[] = [];
      for (const f of params.config.columnFields) {
        const col = colByField.get(f.sourceField);
        if (col == null) continue;
        const state = params.document.getCell(params.sheetId, { row, col }) as { value: unknown; formula: string | null };
        parts.push(keyPartForCell(state, `${row},${col}`));
      }
      keys.add(parts.join("\u0000"));
    }
    return Math.max(1, keys.size);
  })();

  const rowLabelWidth = params.config.layout === "compact" ? 1 : rowFieldCount;

  const subtotalRows =
    params.config.subtotals !== "none" && rowFieldCount > 1
      ? // Upper bound: could emit at most one subtotal row per leaf row in the worst case.
        rowKeyCount
      : 0;

  const outputRows = 1 + rowKeyCount + subtotalRows + (params.config.grandTotals.rows ? 1 : 0);
  const outputCols =
    rowLabelWidth + colKeyCount * valueFieldCount + (params.config.grandTotals.columns ? valueFieldCount : 0);

  const startRow = params.destination.startRow;
  const startCol = params.destination.startCol;
  const endRow = startRow + Math.max(1, outputRows) - 1;
  const endCol = startCol + Math.max(1, outputCols) - 1;
  const cellCount = Math.max(1, outputRows) * Math.max(1, outputCols);

  return { startRow, startCol, endRow, endCol, cellCount };
}

export function PivotBuilderPanelContainer(props: Props) {
  const doc = props.getDocumentController();
  const sheetNameResolver = props.sheetNameResolver ?? null;

  const activeSheetId = props.getActiveSheetId?.() ?? doc?.getSheetIds?.()?.[0] ?? "Sheet1";

  const [sourceSheetId, setSourceSheetId] = useState<string>(activeSheetId);
  const [sourceRange, setSourceRange] = useState<RangeRect>({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
  const [sourceRangeText, setSourceRangeText] = useState<string>("A1");
  const [sourceError, setSourceError] = useState<string | null>(null);

  const [availableFields, setAvailableFields] = useState<string[]>([]);
  const [fieldsError, setFieldsError] = useState<string | null>(null);

  const [pivotName, setPivotName] = useState<string>("Pivot Table 1");
  const [destinationKind, setDestinationKind] = useState<"new" | "existing">("new");
  const [newSheetName, setNewSheetName] = useState<string>("Pivot Table");
  const [destSheetId, setDestSheetId] = useState<string>(activeSheetId);
  const [destCellA1, setDestCellA1] = useState<string>("A1");

  const [pivots, setPivots] = useState<PivotTableSummary[]>([]);
  const [busy, setBusy] = useState<{ kind: "create" } | { kind: "refresh"; pivotId: string } | null>(null);
  const [actionError, setActionError] = useState<string | null>(null);

  const sheetIds: string[] = (() => {
    const ids = doc?.getSheetIds?.() ?? [];
    return ids.length > 0 ? ids : ["Sheet1"];
  })();

  const sheetDisplayName = useCallback(
    (sheetId: string): string => {
      const id = String(sheetId ?? "").trim();
      if (!id) return "";
      return sheetNameResolver?.getSheetNameById(id) ?? id;
    },
    [sheetNameResolver],
  );

  const existingSheetNames = useMemo(
    () => sheetIds.map((id) => sheetDisplayName(id)).filter(Boolean),
    [sheetDisplayName, sheetIds],
  );

  const canEditCell: ((cell: { sheetId: string; row: number; col: number }) => boolean) | null =
    typeof (doc as any)?.canEditCell === "function" ? ((doc as any).canEditCell as any) : null;

  const ensureUpdatesEditable = useCallback(
    (updates: any[] | null | undefined): boolean => {
      if (!canEditCell) return true;
      if (!Array.isArray(updates) || updates.length === 0) return true;

      for (const update of updates) {
        const sheetId = String(update?.sheet_id ?? "").trim();
        const row = Number(update?.row);
        const col = Number(update?.col);
        if (!sheetId) continue;
        if (!Number.isInteger(row) || row < 0) continue;
        if (!Number.isInteger(col) || col < 0) continue;

        if (!canEditCell({ sheetId, row, col })) {
          return false;
        }
      }
      return true;
    },
    [canEditCell],
  );

  const resolveBackend = useCallback((): TauriPivotBackend | null => {
    try {
      return new TauriPivotBackend({ invoke: props.invoke });
    } catch {
      return null;
    }
  }, [props.invoke]);

  const useCurrentSelection = useCallback(() => {
    const sel = props.getSelection?.();
    if (!sel) return;
    setSourceSheetId(sel.sheetId);
    setSourceRange(sel.range);
    setSourceRangeText(rangeToA1(sel.range));
    setSourceError(null);
  }, [props]);

  // Prefill from the current selection when mounted.
  useEffect(() => {
    useCurrentSelection();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Allow external triggers (e.g. command palette) to re-sync the source range from the
  // latest selection without needing a full remount.
  useEffect(() => {
    if (typeof window === "undefined") return;
    const handler = () => useCurrentSelection();
    window.addEventListener("pivot-builder:use-selection", handler as any);
    return () => window.removeEventListener("pivot-builder:use-selection", handler as any);
  }, [useCurrentSelection]);

  const parseSourceRangeText = useCallback(
    (text: string) => {
      const raw = text.trim();
      if (!raw) {
        setSourceError(t("pivotBuilder.source.error.empty"));
        return;
      }
      try {
        const parsed = parseRangeA1(raw.includes(":") ? raw : `${raw}:${raw}`);
        const range = {
          startRow: parsed.start.row,
          startCol: parsed.start.col,
          endRow: parsed.end.row,
          endCol: parsed.end.col,
        };
        setSourceRange(range);
        setSourceError(null);
      } catch (err: any) {
        setSourceError(err?.message ?? String(err));
      }
    },
    [],
  );

  // Derive fields from table metadata (preferred) or header row.
  useEffect(() => {
    let cancelled = false;

    const update = async () => {
      setFieldsError(null);

      const rows = sourceRange.endRow - sourceRange.startRow + 1;
      const cols = sourceRange.endCol - sourceRange.startCol + 1;
      if (rows < 2 || cols < 1) {
        setAvailableFields([]);
        setFieldsError(t("pivotBuilder.source.error.tooSmall"));
        return;
      }

      const backend = resolveBackend();
      if (backend) {
        try {
          const tables = await backend.listTables();
          const table = tables.find((t) => {
            return (
              t.sheet_id === sourceSheetId &&
              sourceRange.startRow >= t.start_row &&
              sourceRange.startCol >= t.start_col &&
              sourceRange.endRow <= t.end_row &&
              sourceRange.endCol <= t.end_col
            );
          });
          if (table) {
            const fullRange: RangeRect = {
              startRow: table.start_row,
              startCol: table.start_col,
              endRow: table.end_row,
              endCol: table.end_col,
            };

            if (!cancelled) {
              const isSameRange =
                sourceRange.startRow === fullRange.startRow &&
                sourceRange.startCol === fullRange.startCol &&
                sourceRange.endRow === fullRange.endRow &&
                sourceRange.endCol === fullRange.endCol;
              if (!isSameRange) {
                setSourceRange(fullRange);
                setSourceRangeText(rangeToA1(fullRange));
              }
              setAvailableFields(table.columns);
            }
            return;
          }
        } catch {
          // Ignore table lookup failures (e.g. non-tauri contexts).
        }
      }

      const headers: string[] = [];
      for (let c = sourceRange.startCol; c <= sourceRange.endCol; c += 1) {
        const state = doc.getCell(sourceSheetId, { row: sourceRange.startRow, col: c }) as {
          value: unknown;
          formula: string | null;
        };
        const header = normalizeHeaderValue(state?.value ?? null).trim();
        headers.push(header);
      }

      const { ok } = dedupeStrings(headers);
      if (!ok) {
        setAvailableFields([]);
        setFieldsError(t("pivotBuilder.source.error.duplicateHeaders"));
        return;
      }
      if (headers.some((h) => h.trim() === "")) {
        setAvailableFields([]);
        setFieldsError(t("pivotBuilder.source.error.blankHeaders"));
        return;
      }

      if (!cancelled) {
        setAvailableFields(headers);
      }
    };

    void update();

    return () => {
      cancelled = true;
    };
  }, [doc, resolveBackend, sourceRange, sourceSheetId]);

  const loadPivotList = useCallback(async () => {
    const backend = resolveBackend();
    if (!backend) return;
    try {
      const list = await backend.listPivotTables();
      setPivots(Array.isArray(list) ? list : []);
    } catch {
      setPivots([]);
    }
  }, [resolveBackend]);

  useEffect(() => {
    void loadPivotList();
  }, [loadPivotList]);

  const newSheetNameError = useMemo(() => {
    if (destinationKind !== "new") return null;
    return getSheetNameValidationErrorMessage(newSheetName, { existingNames: existingSheetNames });
  }, [destinationKind, existingSheetNames, newSheetName]);

  const canCreate = !busy && !sourceError && !fieldsError && availableFields.length > 0 && !newSheetNameError;

  const destinationSummary = useMemo(() => {
    if (destinationKind === "new") {
      return `${formatSheetNameForA1(newSheetName)}!${destCellA1}`;
    }
    const sheetName = sheetDisplayName(destSheetId);
    return `${formatSheetNameForA1(sheetName || destSheetId)}!${destCellA1}`;
  }, [destCellA1, destSheetId, destinationKind, newSheetName, sheetDisplayName]);

  const guardDestination = useCallback(
    async (cfg: PivotTableConfig, dest: { sheetId: string; startRow: number; startCol: number }): Promise<boolean> => {
      if (availableFields.length === 0) return false;
      const rect = estimatePivotOutputRect({
        document: doc,
        sheetId: sourceSheetId,
        source: sourceRange,
        availableFields,
        config: cfg,
        destination: dest,
      });

      if (canEditCell) {
        // Always validate at least the anchor cell; scanning everything can be expensive.
        if (!canEditCell({ sheetId: dest.sheetId, row: dest.startRow, col: dest.startCol })) {
          await nativeDialogs.alert(t("pivotBuilder.destination.error.protected"));
          return false;
        }

        // For modest pivots, validate the full output rect so we don't partially apply updates
        // that get filtered by `DocumentController.canEditCell`.
        if (rect.cellCount <= 10_000) {
          for (let r = rect.startRow; r <= rect.endRow; r += 1) {
            for (let c = rect.startCol; c <= rect.endCol; c += 1) {
              if (!canEditCell({ sheetId: dest.sheetId, row: r, col: c })) {
                await nativeDialogs.alert(t("pivotBuilder.destination.error.protected"));
                return false;
              }
            }
          }
        }
      }

      // For large pivots, avoid scanning the full output region; just require an explicit confirmation.
      if (rect.cellCount > 10_000) {
        return nativeDialogs.confirm(tWithVars("pivotBuilder.destination.confirm.large", { destination: destinationSummary }));
      }

      let nonEmpty = 0;
      for (let r = rect.startRow; r <= rect.endRow; r += 1) {
        for (let c = rect.startCol; c <= rect.endCol; c += 1) {
          const state = doc.getCell(dest.sheetId, { row: r, col: c }) as { value: unknown; formula: string | null };
          if (state?.formula != null || state?.value != null) {
            nonEmpty += 1;
            if (nonEmpty >= 1) break;
          }
        }
        if (nonEmpty >= 1) break;
      }

      if (nonEmpty > 0) {
        return nativeDialogs.confirm(
          tWithVars("pivotBuilder.destination.confirm.overwrite", { destination: destinationSummary }),
        );
      }

      return true;
    },
    [availableFields, canEditCell, destinationSummary, doc, sourceRange, sourceSheetId],
  );

  const createPivot = useCallback(
    async (cfg: PivotTableConfig) => {
      setActionError(null);

      const backend = resolveBackend();
      if (!backend) {
        setActionError(t("pivotBuilder.backendUnavailable"));
        return;
      }

      const rows = sourceRange.endRow - sourceRange.startRow + 1;
      const cols = sourceRange.endCol - sourceRange.startCol + 1;
      if (rows < 2 || cols < 1) {
        setActionError(t("pivotBuilder.source.error.tooSmall"));
        return;
      }
      if (availableFields.length === 0) {
        setActionError(fieldsError ?? t("pivotBuilder.source.error.noHeaders"));
        return;
      }

      if (destinationKind === "new") {
        const sheetError = getSheetNameValidationErrorMessage(newSheetName, { existingNames: existingSheetNames });
        if (sheetError) {
          setActionError(sheetError);
          return;
        }
      }

      let destinationSheetIdResolved = destSheetId;
      let destinationStart = { row: 0, col: 0 };
      try {
        destinationStart = parseA1(destCellA1);
      } catch (err: any) {
        setActionError(err?.message ?? String(err));
        return;
      }

      if (destinationKind === "new") {
        try {
          const ids = doc?.getSheetIds?.() ?? [];
          const orderedIds = ids.length > 0 ? ids : activeSheetId ? [activeSheetId] : [];
          const activeIndex = activeSheetId ? orderedIds.indexOf(activeSheetId) : -1;
          const insertIndex = activeIndex >= 0 ? activeIndex + 1 : orderedIds.length;

          const info = await backend.addSheet(newSheetName, { index: insertIndex });
          destinationSheetIdResolved = info.id || newSheetName;

          // Ensure the sheet exists in the local DocumentController so downstream calls (like
          // destination validation and pivot updates) don't materialize it at the wrong position.
          try {
            doc?.addSheet?.({ sheetId: destinationSheetIdResolved, name: info.name, insertAfterId: activeSheetId }, { source: "pivot" });
          } catch {
            // ignore
          }
        } catch (err: any) {
          setActionError(err?.message ?? String(err));
          return;
        }
      }

      const dest = { sheetId: destinationSheetIdResolved, startRow: destinationStart.row, startCol: destinationStart.col };
      if (!(await guardDestination(cfg, dest))) return;

      setBusy({ kind: "create" });

      try {
        // Ensure any queued sheet edits are flushed before the backend computes the pivot.
        await new Promise<void>((resolve) => queueMicrotask(resolve));
        await props.drainBackendSync?.();

        const response = await backend.createPivotTable({
          name: pivotName.trim() || "Pivot Table",
          source_sheet_id: sourceSheetId,
          source_range: {
            start_row: sourceRange.startRow,
            start_col: sourceRange.startCol,
            end_row: sourceRange.endRow,
            end_col: sourceRange.endCol,
          },
          destination: { sheet_id: destinationSheetIdResolved, row: destinationStart.row, col: destinationStart.col },
          config: toRustPivotConfig(cfg) as unknown as Record<string, unknown>,
        });

        if (!ensureUpdatesEditable(response.updates as any)) {
          setActionError(t("pivotBuilder.destination.error.protected"));
          return;
        }

        doc.beginBatch({ label: "Create pivot table" });
        let committed = false;
        try {
          applyPivotCellUpdates(doc, response.updates);
          committed = true;
        } finally {
          if (committed) doc.endBatch();
          else doc.cancelBatch();
        }

        await loadPivotList();
      } catch (err: any) {
        setActionError(err?.message ?? String(err));
      } finally {
        setBusy(null);
      }
    },
    [
      availableFields.length,
      activeSheetId,
      destCellA1,
      destSheetId,
      destinationKind,
      doc,
      existingSheetNames,
      fieldsError,
      guardDestination,
      loadPivotList,
      newSheetName,
      pivotName,
      props.drainBackendSync,
      ensureUpdatesEditable,
      resolveBackend,
      sourceRange,
      sourceSheetId,
    ],
  );

  const refreshPivot = useCallback(
    async (pivotId: string) => {
      setActionError(null);

      const backend = resolveBackend();
      if (!backend) {
        setActionError(t("pivotBuilder.backendUnavailable"));
        return;
      }

      setBusy({ kind: "refresh", pivotId });
      try {
        await new Promise<void>((resolve) => queueMicrotask(resolve));
        await props.drainBackendSync?.();

        const updates = await backend.refreshPivotTable(pivotId);

        if (!ensureUpdatesEditable(updates as any)) {
          setActionError(t("pivotBuilder.destination.error.protected"));
          return;
        }

        doc.beginBatch({ label: "Refresh pivot table" });
        let committed = false;
        try {
          applyPivotCellUpdates(doc, updates);
          committed = true;
        } finally {
          if (committed) doc.endBatch();
          else doc.cancelBatch();
        }
        await loadPivotList();
      } catch (err: any) {
        setActionError(err?.message ?? String(err));
      } finally {
        setBusy(null);
      }
    },
    [doc, ensureUpdatesEditable, loadPivotList, props.drainBackendSync, resolveBackend],
  );

  return (
    <div style={{ padding: 12, display: "grid", gap: 12 }}>
      <div style={{ display: "grid", gap: 10, borderBottom: "1px solid var(--border)", paddingBottom: 12 }}>
        <div style={{ display: "grid", gap: 6 }}>
          <div style={{ fontSize: 12, color: "var(--text-secondary)" }}>{t("pivotBuilder.source.title")}</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
            <label style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
              <span style={{ fontSize: 12, color: "var(--text-secondary)" }}>{t("pivotBuilder.source.sheetLabel")}</span>
              <select
                data-testid="pivot-source-sheet"
                value={sourceSheetId}
                onChange={(e) => setSourceSheetId(e.target.value)}
              >
                {sheetIds.map((id) => (
                  <option key={id} value={id}>
                    {sheetDisplayName(id) || id}
                  </option>
                ))}
              </select>
            </label>

            <label style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
              <span style={{ fontSize: 12, color: "var(--text-secondary)" }}>{t("pivotBuilder.source.rangeLabel")}</span>
              <input
                data-testid="pivot-source-range"
                value={sourceRangeText}
                onChange={(e) => setSourceRangeText(e.target.value)}
                onBlur={() => parseSourceRangeText(sourceRangeText)}
                style={{ width: 140 }}
              />
            </label>

            <button type="button" data-testid="pivot-use-selection" onClick={useCurrentSelection}>
              {t("pivotBuilder.source.useSelection")}
            </button>
          </div>
          {sourceError ? <div style={{ color: "var(--error)" }}>{sourceError}</div> : null}
          {fieldsError ? <div style={{ color: "var(--error)" }}>{fieldsError}</div> : null}
        </div>

        <div style={{ display: "grid", gap: 6 }}>
          <div style={{ fontSize: 12, color: "var(--text-secondary)" }}>{t("pivotBuilder.destination.title")}</div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
            <label style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
              <input
                data-testid="pivot-destination-new"
                type="radio"
                name="pivot-destination"
                checked={destinationKind === "new"}
                onChange={() => setDestinationKind("new")}
              />
              {t("pivotBuilder.destination.newSheet")}
            </label>
            <label style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
              <input
                data-testid="pivot-destination-existing"
                type="radio"
                name="pivot-destination"
                checked={destinationKind === "existing"}
                onChange={() => setDestinationKind("existing")}
              />
              {t("pivotBuilder.destination.existing")}
            </label>
          </div>

          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
            {destinationKind === "new" ? (
              <label style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                <span style={{ fontSize: 12, color: "var(--text-secondary)" }}>{t("pivotBuilder.destination.sheetName")}</span>
                <input
                  data-testid="pivot-destination-new-sheet-name"
                  value={newSheetName}
                  onChange={(e) => setNewSheetName(e.target.value)}
                  style={{ width: 200 }}
                />
              </label>
            ) : (
              <label style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                <span style={{ fontSize: 12, color: "var(--text-secondary)" }}>{t("pivotBuilder.destination.sheetLabel")}</span>
                <select
                  data-testid="pivot-destination-sheet"
                  value={destSheetId}
                  onChange={(e) => setDestSheetId(e.target.value)}
                >
                  {sheetIds.map((id) => (
                    <option key={id} value={id}>
                      {sheetDisplayName(id) || id}
                    </option>
                  ))}
                </select>
              </label>
            )}

            {destinationKind === "new" && newSheetNameError ? (
              <div style={{ color: "var(--error)" }}>{newSheetNameError}</div>
            ) : null}

            <label style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
              <span style={{ fontSize: 12, color: "var(--text-secondary)" }}>{t("pivotBuilder.destination.startCell")}</span>
              <input
                data-testid="pivot-destination-cell"
                value={destCellA1}
                onChange={(e) => setDestCellA1(e.target.value)}
                style={{ width: 80 }}
              />
            </label>
          </div>
        </div>

        <div style={{ display: "grid", gap: 6 }}>
          <div style={{ fontSize: 12, color: "var(--text-secondary)" }}>{t("pivotBuilder.name.title")}</div>
          <input
            data-testid="pivot-name"
            value={pivotName}
            onChange={(e) => setPivotName(e.target.value)}
            style={{ maxWidth: 360 }}
          />
        </div>

        {actionError ? <div style={{ color: "var(--error)" }}>{actionError}</div> : null}

        {busy ? (
          <div data-testid="pivot-progress" style={{ color: "var(--text-secondary)", fontSize: 12 }}>
            {busy.kind === "create"
              ? t("pivotBuilder.progress.creating")
              : t("pivotBuilder.progress.refreshing")}
          </div>
        ) : null}
      </div>

      <PivotBuilderPanel
        availableFields={availableFields}
        onCreate={(cfg) => void createPivot(cfg)}
        createDisabled={!canCreate}
        createLabel={busy?.kind === "create" ? t("pivotBuilder.progress.creating") : undefined}
      />

      <div style={{ borderTop: "1px solid var(--border)", paddingTop: 12, display: "grid", gap: 8 }}>
        <div style={{ fontSize: 12, fontWeight: 600, color: "var(--text-secondary)" }}>{t("pivotBuilder.pivots.title")}</div>
        {pivots.length === 0 ? (
          <div style={{ color: "var(--text-secondary)", fontSize: 12 }}>{t("pivotBuilder.pivots.empty")}</div>
        ) : (
          <div style={{ display: "grid", gap: 6 }}>
            {pivots.map((p) => (
              <div
                key={p.id}
                data-testid={`pivot-item-${p.id}`}
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  padding: 8,
                  border: "1px solid var(--border)",
                  borderRadius: 8,
                }}
              >
                <div style={{ display: "grid" }}>
                  <div style={{ fontWeight: 600 }}>{p.name}</div>
                  <div style={{ fontSize: 12, color: "var(--text-secondary)" }}>
                    {formatSheetNameForA1(sheetDisplayName(p.source_sheet_id) || p.source_sheet_id)}!{rangeToA1({
                      startRow: p.source_range.start_row,
                      startCol: p.source_range.start_col,
                      endRow: p.source_range.end_row,
                      endCol: p.source_range.end_col,
                    })}{" "}
                    → {formatSheetNameForA1(sheetDisplayName(p.destination.sheet_id) || p.destination.sheet_id)}!{cellToA1(p.destination.row, p.destination.col)}
                  </div>
                </div>
                <button
                  type="button"
                  data-testid={`pivot-refresh-${p.id}`}
                  onClick={() => void refreshPivot(p.id)}
                  disabled={busy != null}
                >
                  {t("pivotBuilder.pivots.refresh")}
                </button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
