import React, { useEffect, useMemo, useState } from "react";

import type { SheetNameResolver } from "../../sheet/sheetNameResolver";
import { formatSheetNameForA1 } from "../../sheet/formatSheetNameForA1.js";
import { formatA1 } from "../../document/coords.js";
import { t, tWithVars } from "../../i18n/index.js";

import { diffYjsWorkbookVersionAgainstCurrent } from "../../versioning/index.js";
import { FormulaDiffView } from "../../versioning/ui/FormulaDiffView.js";

type VersionManagerLike = {
  doc: { encodeState(): Uint8Array };
  getVersion(versionId: string): Promise<{ snapshot: Uint8Array } | null>;
};

type SheetViewMeta = { frozenRows: number; frozenCols: number; backgroundImageId?: string | null };
type AddedSheet = { id: string; name: string | null; afterIndex: number; visibility?: string; tabColor?: string | null; view?: SheetViewMeta };
type RemovedSheet = { id: string; name: string | null; beforeIndex: number; visibility?: string; tabColor?: string | null; view?: SheetViewMeta };
type RenamedSheet = { id: string; beforeName: string | null; afterName: string | null };
type MovedSheet = { id: string; beforeIndex: number; afterIndex: number };
type SheetMetaChange = { id: string; field: string; before: unknown; after: unknown };
type SheetOptionChangeLabel =
  | { kind: "added" }
  | { kind: "removed" }
  | { kind: "renamed" }
  | { kind: "reordered" }
  | { kind: "meta"; count: number };

type WorkbookDiff = {
  sheets: {
    added: AddedSheet[];
    removed: RemovedSheet[];
    renamed: RenamedSheet[];
    moved: MovedSheet[];
    metaChanged?: SheetMetaChange[];
  };
  cellsBySheet: Array<{
    sheetId: string;
    sheetName: string | null;
    diff: {
      added: CellChange[];
      removed: CellChange[];
      modified: CellChange[];
      moved: MoveChange[];
      formatOnly: CellChange[];
    };
  }>;
  comments: { added: unknown[]; removed: unknown[]; modified: unknown[] };
  metadata: { added: unknown[]; removed: unknown[]; modified: unknown[] };
  namedRanges: { added: unknown[]; removed: unknown[]; modified: unknown[] };
};

type CellRef = { row: number; col: number };

type CellChange = {
  cell: CellRef;
  oldValue?: unknown;
  newValue?: unknown;
  oldFormula?: string | null;
  newFormula?: string | null;
  oldEncrypted?: boolean;
  newEncrypted?: boolean;
  oldKeyId?: string | null;
  newKeyId?: string | null;
};

type MoveChange = {
  oldLocation: CellRef;
  newLocation: CellRef;
  value: unknown;
  formula?: string | null;
  encrypted?: boolean;
  keyId?: string | null;
};

function hasAnyCellChanges(diff: WorkbookDiff["cellsBySheet"][number]["diff"]) {
  return (
    diff.added.length > 0 ||
    diff.removed.length > 0 ||
    diff.modified.length > 0 ||
    diff.moved.length > 0 ||
    diff.formatOnly.length > 0
  );
}

function summarizeJson(value: unknown): string {
  if (value === null || value === undefined) return t("versionHistory.compare.value.empty");
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  try {
    const json = JSON.stringify(value);
    return json.length > 200 ? `${json.slice(0, 200)}…` : json;
  } catch {
    return String(value);
  }
}

function summarizeCellContent(opts: { value?: unknown; formula?: string | null; encrypted?: boolean; keyId?: string | null }): string {
  if (opts.encrypted)
    return opts.keyId
      ? tWithVars("versionHistory.compare.value.encryptedWithKeyId", { keyId: opts.keyId })
      : t("versionHistory.compare.value.encrypted");
  const formula = opts.formula ?? null;
  if (formula) return formula;
  if (opts.value === null || opts.value === undefined) return t("versionHistory.compare.value.empty");
  return summarizeJson(opts.value);
}

function formatSheetMetaField(field: string): string {
  if (field === "visibility") return t("versionHistory.compare.sheetMetaField.visibility");
  if (field === "tabColor") return t("versionHistory.compare.sheetMetaField.tabColor");
  if (field === "view.frozenRows") return t("versionHistory.compare.sheetMetaField.frozenRows");
  if (field === "view.frozenCols") return t("versionHistory.compare.sheetMetaField.frozenCols");
  if (field === "view.backgroundImageId") return t("versionHistory.compare.sheetMetaField.backgroundImageId");
  return field;
}
function sheetDisplayName(sheetId: string, fallbackName: string | null, sheetNameResolver: SheetNameResolver | null): string {
  return sheetNameResolver?.getSheetNameById(sheetId) ?? fallbackName ?? sheetId;
}

function formatSheetQualifiedA1(sheetName: string, cell: CellRef): string {
  return `${formatSheetNameForA1(sheetName)}!${formatA1(cell)}`;
}

function nextPaint(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof requestAnimationFrame === "function") {
      requestAnimationFrame(() => resolve());
      return;
    }
    setTimeout(resolve, 0);
  });
}

export function VersionHistoryCompareSection({
  versionId,
  versionManager,
  sheetNameResolver = null,
}: {
  versionId: string | null;
  versionManager: VersionManagerLike | null;
  sheetNameResolver?: SheetNameResolver | null;
}) {
  const [diff, setDiff] = useState<WorkbookDiff | null>(null);
  const [diffError, setDiffError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [selectedSheetId, setSelectedSheetId] = useState<string | null>(null);

  useEffect(() => {
    if (!versionId || !versionManager) {
      setDiff(null);
      setDiffError(null);
      setLoading(false);
      return;
    }

    let cancelled = false;
    setLoading(true);
    setDiffError(null);
    setDiff(null);

    void (async () => {
      try {
        // Let React paint a "Computing diff…" indicator before doing the heavy work.
        await nextPaint();
        const nextDiff = (await diffYjsWorkbookVersionAgainstCurrent({ versionManager, versionId })) as WorkbookDiff;
        if (cancelled) return;
        setDiff(nextDiff);
      } catch (e) {
        if (cancelled) return;
        setDiffError((e as Error).message);
      } finally {
        if (!cancelled) setLoading(false);
      }
    })().catch(() => {});

    return () => {
      cancelled = true;
    };
  }, [versionId, versionManager]);

  useEffect(() => {
    if (!diff) {
      setSelectedSheetId(null);
      return;
    }
    const entries = diff.cellsBySheet ?? [];
    if (entries.length === 0) {
      setSelectedSheetId(null);
      return;
    }

    const sheetChangedIds = new Set([
      ...(diff.sheets?.added ?? []).map((s) => s.id),
      ...(diff.sheets?.removed ?? []).map((s) => s.id),
      ...(diff.sheets?.renamed ?? []).map((s) => s.id),
      ...(diff.sheets?.moved ?? []).map((s) => s.id),
      ...(diff.sheets?.metaChanged ?? []).map((c) => c.id),
    ]);
    setSelectedSheetId((prev) => {
      if (prev && entries.some((e) => e.sheetId === prev)) return prev;
      const preferred = entries.find((e) => hasAnyCellChanges(e.diff) || sheetChangedIds.has(e.sheetId)) ?? entries[0];
      return preferred?.sheetId ?? null;
    });
  }, [diff]);

  const summary = useMemo(() => {
    if (!diff) return null;
    const cellsTotals = { added: 0, removed: 0, modified: 0, moved: 0, formatOnly: 0 };
    for (const entry of diff.cellsBySheet ?? []) {
      cellsTotals.added += entry.diff.added.length;
      cellsTotals.removed += entry.diff.removed.length;
      cellsTotals.modified += entry.diff.modified.length;
      cellsTotals.moved += entry.diff.moved.length;
      cellsTotals.formatOnly += entry.diff.formatOnly.length;
    }
    return {
      sheets: {
        added: diff.sheets?.added?.length ?? 0,
        removed: diff.sheets?.removed?.length ?? 0,
        renamed: diff.sheets?.renamed?.length ?? 0,
        moved: diff.sheets?.moved?.length ?? 0,
        metaChanged: diff.sheets?.metaChanged?.length ?? 0,
      },
      cells: cellsTotals,
      metadata: {
        added: diff.metadata?.added?.length ?? 0,
        removed: diff.metadata?.removed?.length ?? 0,
        modified: diff.metadata?.modified?.length ?? 0,
      },
      namedRanges: {
        added: diff.namedRanges?.added?.length ?? 0,
        removed: diff.namedRanges?.removed?.length ?? 0,
        modified: diff.namedRanges?.modified?.length ?? 0,
      },
      comments: {
        added: diff.comments?.added?.length ?? 0,
        removed: diff.comments?.removed?.length ?? 0,
        modified: diff.comments?.modified?.length ?? 0,
      },
    };
  }, [diff]);

  const sheetOptions = useMemo(() => {
    if (!diff) return [];
    const metaCounts = new Map<string, number>();
    for (const change of diff.sheets?.metaChanged ?? []) {
      metaCounts.set(change.id, (metaCounts.get(change.id) ?? 0) + 1);
    }

    const changeLabels = new Map<string, SheetOptionChangeLabel[]>();
    const addLabel = (id: string, label: SheetOptionChangeLabel) => {
      const labels = changeLabels.get(id) ?? [];
      labels.push(label);
      changeLabels.set(id, labels);
    };
    for (const sheet of diff.sheets?.added ?? []) addLabel(sheet.id, { kind: "added" });
    for (const sheet of diff.sheets?.removed ?? []) addLabel(sheet.id, { kind: "removed" });
    for (const sheet of diff.sheets?.renamed ?? []) addLabel(sheet.id, { kind: "renamed" });
    for (const sheet of diff.sheets?.moved ?? []) addLabel(sheet.id, { kind: "reordered" });
    for (const [id, count] of metaCounts) {
      addLabel(id, { kind: "meta", count });
    }

    return (diff.cellsBySheet ?? []).map((entry) => {
      const name = sheetDisplayName(entry.sheetId, entry.sheetName, sheetNameResolver);
      const cellChanges = hasAnyCellChanges(entry.diff);
      const labels = changeLabels.get(entry.sheetId) ?? [];
      // Always surface sheet-level ops (added/removed/renamed/reordered) even if the sheet
      // also has cell-level changes. For meta changes, only surface the count when the
      // sheet has no cell changes (otherwise the sheet is already highlighted by cell diffs).
      const displayLabels = cellChanges ? labels.filter((l) => l.kind !== "meta") : labels;
      const formattedLabels = displayLabels
        .map((label) => {
          if (label.kind === "added") return t("versionHistory.compare.sheetOption.change.added");
          if (label.kind === "removed") return t("versionHistory.compare.sheetOption.change.removed");
          if (label.kind === "renamed") return t("versionHistory.compare.sheetOption.change.renamed");
          if (label.kind === "reordered") return t("versionHistory.compare.sheetOption.change.reordered");
          return tWithVars("versionHistory.compare.sheetOption.change.metaWithCount", { count: label.count });
        })
        .join(", ");
      const suffix =
        formattedLabels.length > 0
          ? tWithVars("versionHistory.compare.sheetOption.changesSuffix", { changes: formattedLabels })
          : cellChanges
            ? ""
            : t("versionHistory.compare.sheetOption.noChangesSuffix");
      return { sheetId: entry.sheetId, displayName: `${name}${suffix}`, rawName: name };
    });
  }, [diff, sheetNameResolver]);

  const selectedSheet = useMemo(() => {
    if (!diff || !selectedSheetId) return null;
    return diff.cellsBySheet.find((s) => s.sheetId === selectedSheetId) ?? null;
  }, [diff, selectedSheetId]);

  return (
    <div className="collab-version-history__compare">
      <div className="collab-version-history__compare-header">
        <h4 className="collab-version-history__compare-title">{t("versionHistory.compare.title")}</h4>
      </div>

      {!versionId ? (
        <div className="collab-version-history__compare-empty">{t("versionHistory.compare.empty.noVersionSelected")}</div>
      ) : loading ? (
        <div className="collab-version-history__compare-loading">{t("versionHistory.compare.loading")}</div>
      ) : diffError ? (
        <div className="collab-version-history__compare-error">{diffError}</div>
      ) : !diff || !summary ? (
        <div className="collab-version-history__compare-empty">{t("versionHistory.compare.empty.noDiff")}</div>
      ) : (
        <>
          <div className="collab-version-history__compare-summary">
            <div className="collab-version-history__compare-summary-group">
              <div className="collab-version-history__compare-summary-label">{t("versionHistory.compare.summary.sheets")}</div>
              <div className="collab-version-history__compare-summary-badges">
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--added">
                  {tWithVars("versionHistory.compare.badge.added", { count: summary.sheets.added })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--removed">
                  {tWithVars("versionHistory.compare.badge.removed", { count: summary.sheets.removed })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  {tWithVars("versionHistory.compare.badge.renamed", { count: summary.sheets.renamed })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  {tWithVars("versionHistory.compare.badge.meta", { count: summary.sheets.metaChanged })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--moved">
                  {tWithVars("versionHistory.compare.badge.moved", { count: summary.sheets.moved })}
                </span>
              </div>
            </div>

            <div className="collab-version-history__compare-summary-group">
              <div className="collab-version-history__compare-summary-label">{t("versionHistory.compare.summary.cells")}</div>
              <div className="collab-version-history__compare-summary-badges">
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--added">
                  {tWithVars("versionHistory.compare.badge.added", { count: summary.cells.added })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--removed">
                  {tWithVars("versionHistory.compare.badge.removed", { count: summary.cells.removed })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  {tWithVars("versionHistory.compare.badge.modified", { count: summary.cells.modified })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--moved">
                  {tWithVars("versionHistory.compare.badge.moved", { count: summary.cells.moved })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--format">
                  {tWithVars("versionHistory.compare.badge.formatOnly", { count: summary.cells.formatOnly })}
                </span>
              </div>
            </div>

            <div className="collab-version-history__compare-summary-group">
              <div className="collab-version-history__compare-summary-label">
                {t("versionHistory.compare.summary.workbookMetadata")}
              </div>
              <div className="collab-version-history__compare-summary-badges">
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--added">
                  {tWithVars("versionHistory.compare.badge.added", { count: summary.metadata.added })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--removed">
                  {tWithVars("versionHistory.compare.badge.removed", { count: summary.metadata.removed })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  {tWithVars("versionHistory.compare.badge.modified", { count: summary.metadata.modified })}
                </span>
              </div>
            </div>

            <div className="collab-version-history__compare-summary-group">
              <div className="collab-version-history__compare-summary-label">{t("versionHistory.compare.summary.namedRanges")}</div>
              <div className="collab-version-history__compare-summary-badges">
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--added">
                  {tWithVars("versionHistory.compare.badge.added", { count: summary.namedRanges.added })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--removed">
                  {tWithVars("versionHistory.compare.badge.removed", { count: summary.namedRanges.removed })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  {tWithVars("versionHistory.compare.badge.modified", { count: summary.namedRanges.modified })}
                </span>
              </div>
            </div>

            <div className="collab-version-history__compare-summary-group">
              <div className="collab-version-history__compare-summary-label">{t("versionHistory.compare.summary.comments")}</div>
              <div className="collab-version-history__compare-summary-badges">
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--added">
                  {tWithVars("versionHistory.compare.badge.added", { count: summary.comments.added })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--removed">
                  {tWithVars("versionHistory.compare.badge.removed", { count: summary.comments.removed })}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  {tWithVars("versionHistory.compare.badge.modified", { count: summary.comments.modified })}
                </span>
              </div>
            </div>
          </div>

          <div className="collab-version-history__compare-controls">
            <label className="collab-version-history__compare-control">
              {t("versionHistory.compare.control.sheet")}{" "}
              <select
                className="collab-version-history__compare-select"
                value={selectedSheetId ?? ""}
                onChange={(e) => setSelectedSheetId(e.target.value || null)}
              >
                {sheetOptions.map((opt) => (
                  <option key={opt.sheetId} value={opt.sheetId}>
                    {opt.displayName}
                  </option>
                ))}
              </select>
            </label>
          </div>

          {!selectedSheet ? (
            <div className="collab-version-history__compare-empty">{t("versionHistory.compare.empty.noSheetSelected")}</div>
          ) : (
            <div className="collab-version-history__sheet-diff">
              {(() => {
                const sheetName = sheetDisplayName(selectedSheet.sheetId, selectedSheet.sheetName, sheetNameResolver);
                const sd = selectedSheet.diff;
                const groups: Array<{
                  key: string;
                  label: string;
                  className: string;
                  items: any[];
                }> = [
                  {
                    key: "added",
                    label: t("versionHistory.compare.changeType.added"),
                    className: "collab-version-history__diff-row--added",
                    items: sd.added,
                  },
                  {
                    key: "removed",
                    label: t("versionHistory.compare.changeType.removed"),
                    className: "collab-version-history__diff-row--removed",
                    items: sd.removed,
                  },
                  {
                    key: "modified",
                    label: t("versionHistory.compare.changeType.modified"),
                    className: "collab-version-history__diff-row--modified",
                    items: sd.modified,
                  },
                  {
                    key: "moved",
                    label: t("versionHistory.compare.changeType.moved"),
                    className: "collab-version-history__diff-row--moved",
                    items: sd.moved,
                  },
                  {
                    key: "formatOnly",
                    label: t("versionHistory.compare.changeType.formatOnly"),
                    className: "collab-version-history__diff-row--format",
                    items: sd.formatOnly,
                  },
                ];

                const any = groups.some((g) => g.items.length > 0);
                return (
                  <>
                    {(() => {
                      const added = diff.sheets?.added?.find((s) => s.id === selectedSheet.sheetId) ?? null;
                      const removed = diff.sheets?.removed?.find((s) => s.id === selectedSheet.sheetId) ?? null;
                      const renamed = diff.sheets?.renamed?.find((s) => s.id === selectedSheet.sheetId) ?? null;
                      const moved = diff.sheets?.moved?.find((s) => s.id === selectedSheet.sheetId) ?? null;
                      if (!added && !removed && !renamed && !moved) return null;
                      return (
                        <div className="collab-version-history__diff-group">
                          <div className="collab-version-history__diff-group-title">{t("versionHistory.compare.control.sheet")}</div>
                          <div className="collab-version-history__diff-group-list">
                            {added ? (
                              <div className="collab-version-history__diff-row collab-version-history__diff-row--added">
                                <div className="collab-version-history__diff-row-header">
                                  <span className="collab-version-history__diff-cell-ref">{t("versionHistory.compare.changeType.added")}</span>
                                </div>
                                <div className="collab-version-history__diff-row-body">
                                  <div className="collab-version-history__diff-row-values">
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">{t("versionHistory.compare.sheetDetails.name")}</span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(added.name)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">{t("versionHistory.compare.sheetDetails.index")}</span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(added.afterIndex)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">
                                        {t("versionHistory.compare.sheetMetaField.visibility")}:
                                      </span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(added.visibility)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">
                                        {t("versionHistory.compare.sheetMetaField.tabColor")}:
                                      </span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(added.tabColor ?? null)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">
                                        {t("versionHistory.compare.sheetMetaField.backgroundImageId")}:
                                      </span>{" "}
                                      <span className="collab-version-history__diff-value">
                                        {summarizeJson(added.view?.backgroundImageId ?? null)}
                                      </span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">
                                        {t("versionHistory.compare.sheetMetaField.frozenRows")}:
                                      </span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(added.view?.frozenRows ?? 0)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">
                                        {t("versionHistory.compare.sheetMetaField.frozenCols")}:
                                      </span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(added.view?.frozenCols ?? 0)}</span>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            ) : null}

                            {removed ? (
                              <div className="collab-version-history__diff-row collab-version-history__diff-row--removed">
                                <div className="collab-version-history__diff-row-header">
                                  <span className="collab-version-history__diff-cell-ref">{t("versionHistory.compare.changeType.removed")}</span>
                                </div>
                                <div className="collab-version-history__diff-row-body">
                                  <div className="collab-version-history__diff-row-values">
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">{t("versionHistory.compare.sheetDetails.name")}</span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(removed.name)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">{t("versionHistory.compare.sheetDetails.index")}</span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(removed.beforeIndex)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">
                                        {t("versionHistory.compare.sheetMetaField.visibility")}:
                                      </span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(removed.visibility)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">
                                        {t("versionHistory.compare.sheetMetaField.tabColor")}:
                                      </span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(removed.tabColor ?? null)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">
                                        {t("versionHistory.compare.sheetMetaField.backgroundImageId")}:
                                      </span>{" "}
                                      <span className="collab-version-history__diff-value">
                                        {summarizeJson(removed.view?.backgroundImageId ?? null)}
                                      </span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">
                                        {t("versionHistory.compare.sheetMetaField.frozenRows")}:
                                      </span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(removed.view?.frozenRows ?? 0)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">
                                        {t("versionHistory.compare.sheetMetaField.frozenCols")}:
                                      </span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(removed.view?.frozenCols ?? 0)}</span>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            ) : null}

                            {renamed ? (
                              <div className="collab-version-history__diff-row collab-version-history__diff-row--modified">
                                <div className="collab-version-history__diff-row-header">
                                  <span className="collab-version-history__diff-cell-ref">{t("versionHistory.compare.sheetChange.renamed")}</span>
                                </div>
                                <div className="collab-version-history__diff-row-body">
                                  <div className="collab-version-history__diff-row-values">
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">{t("versionHistory.compare.label.before")}</span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(renamed.beforeName)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">{t("versionHistory.compare.label.after")}</span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(renamed.afterName)}</span>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            ) : null}

                            {moved ? (
                              <div className="collab-version-history__diff-row collab-version-history__diff-row--moved">
                                <div className="collab-version-history__diff-row-header">
                                  <span className="collab-version-history__diff-cell-ref">{t("versionHistory.compare.sheetChange.reordered")}</span>
                                </div>
                                <div className="collab-version-history__diff-row-body">
                                  <div className="collab-version-history__diff-row-values">
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">
                                        {t("versionHistory.compare.sheetDetails.beforeIndex")}
                                      </span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(moved.beforeIndex)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">{t("versionHistory.compare.sheetDetails.afterIndex")}</span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(moved.afterIndex)}</span>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            ) : null}
                          </div>
                        </div>
                      );
                    })()}

                    {(() => {
                      const meta = (diff.sheets?.metaChanged ?? []).filter((c) => c.id === selectedSheet.sheetId);
                      if (meta.length === 0) return null;
                      return (
                        <div className="collab-version-history__diff-group">
                          <div className="collab-version-history__diff-group-title">
                            {tWithVars("versionHistory.compare.diffGroup.titleWithCount", {
                              label: t("versionHistory.compare.diffGroup.sheetMetadata"),
                              count: meta.length,
                            })}
                          </div>
                          <div className="collab-version-history__diff-group-list">
                            {meta.map((change, idx) => (
                              <div
                                key={`meta-${idx}`}
                                className="collab-version-history__diff-row collab-version-history__diff-row--modified"
                              >
                                <div className="collab-version-history__diff-row-header">
                                  <span className="collab-version-history__diff-cell-ref">{formatSheetMetaField(change.field)}</span>
                                </div>
                                <div className="collab-version-history__diff-row-body">
                                  <div className="collab-version-history__diff-row-values">
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">{t("versionHistory.compare.label.before")}</span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(change.before)}</span>
                                    </div>
                                    <div className="collab-version-history__diff-row-value">
                                      <span className="collab-version-history__diff-row-label">{t("versionHistory.compare.label.after")}</span>{" "}
                                      <span className="collab-version-history__diff-value">{summarizeJson(change.after)}</span>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      );
                    })()}

                    {!any ? (
                      <div className="collab-version-history__sheet-diff-empty">{t("versionHistory.compare.empty.noCellChangesOnSheet")}</div>
                    ) : null}

                    {groups.map((g) => {
                      if (g.items.length === 0) return null;
                      return (
                        <div key={g.key} className="collab-version-history__diff-group">
                          <div className="collab-version-history__diff-group-title">
                            {tWithVars("versionHistory.compare.diffGroup.titleWithCount", {
                              label: g.label,
                              count: g.items.length,
                            })}
                          </div>
                          <div className="collab-version-history__diff-group-list">
                            {g.key === "moved"
                              ? (g.items as MoveChange[]).map((change, idx) => {
                                  const from = formatSheetQualifiedA1(sheetName, change.oldLocation);
                                  const to = formatSheetQualifiedA1(sheetName, change.newLocation);
                                  return (
                                    <div key={`${g.key}-${idx}`} className={`collab-version-history__diff-row ${g.className}`}>
                                      <div className="collab-version-history__diff-row-header">
                                        <span className="collab-version-history__diff-cell-ref">
                                          {from} → {to}
                                        </span>
                                      </div>
                                      <div className="collab-version-history__diff-row-body">
                                        <span className="collab-version-history__diff-value">
                                          {summarizeCellContent({
                                            value: change.value,
                                            formula: change.formula ?? null,
                                            encrypted: change.encrypted,
                                            keyId: change.keyId ?? null,
                                          })}
                                        </span>
                                      </div>
                                    </div>
                                  );
                                })
                              : (g.items as CellChange[]).map((change, idx) => {
                                  const ref = formatSheetQualifiedA1(sheetName, change.cell);
                                  const isModified = g.key === "modified" || g.key === "formatOnly";
                                  const oldSummary = summarizeCellContent({
                                    value: change.oldValue,
                                    formula: change.oldFormula ?? null,
                                    encrypted: change.oldEncrypted,
                                    keyId: change.oldKeyId ?? null,
                                  });
                                  const newSummary = summarizeCellContent({
                                    value: change.newValue,
                                    formula: change.newFormula ?? null,
                                    encrypted: change.newEncrypted,
                                    keyId: change.newKeyId ?? null,
                                  });

                                  const formulaChanged = (change.oldFormula ?? null) !== (change.newFormula ?? null);

                                  return (
                                    <div key={`${g.key}-${idx}`} className={`collab-version-history__diff-row ${g.className}`}>
                                      <div className="collab-version-history__diff-row-header">
                                        <span className="collab-version-history__diff-cell-ref">{ref}</span>
                                      </div>
                                      <div className="collab-version-history__diff-row-body">
                                        {isModified ? (
                                          <>
                                            <div className="collab-version-history__diff-row-values">
                                              <div className="collab-version-history__diff-row-value">
                                                <span className="collab-version-history__diff-row-label">
                                                  {t("versionHistory.compare.label.before")}
                                                </span>{" "}
                                                <span className="collab-version-history__diff-value">{oldSummary}</span>
                                              </div>
                                              <div className="collab-version-history__diff-row-value">
                                                <span className="collab-version-history__diff-row-label">
                                                  {t("versionHistory.compare.label.after")}
                                                </span>{" "}
                                                <span className="collab-version-history__diff-value">{newSummary}</span>
                                              </div>
                                            </div>
                                            {formulaChanged ? (
                                              <div
                                                className="collab-version-history__formula-diff"
                                                aria-label={t("versionHistory.compare.aria.formulaDiff")}
                                              >
                                                <FormulaDiffView before={change.oldFormula ?? null} after={change.newFormula ?? null} />
                                              </div>
                                            ) : null}
                                          </>
                                        ) : g.key === "added" ? (
                                          <div className="collab-version-history__diff-row-value">
                                            <span className="collab-version-history__diff-row-label">
                                              {t("versionHistory.compare.label.after")}
                                            </span>{" "}
                                            <span className="collab-version-history__diff-value">{newSummary}</span>
                                          </div>
                                        ) : (
                                          <div className="collab-version-history__diff-row-value">
                                            <span className="collab-version-history__diff-row-label">
                                              {t("versionHistory.compare.label.before")}
                                            </span>{" "}
                                            <span className="collab-version-history__diff-value">{oldSummary}</span>
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  );
                                })}
                          </div>
                        </div>
                      );
                    })}
                  </>
                );
              })()}
            </div>
          )}
        </>
      )}
    </div>
  );
}
