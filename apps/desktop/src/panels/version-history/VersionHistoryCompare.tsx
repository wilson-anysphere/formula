import React, { useEffect, useMemo, useState } from "react";

import type { SheetNameResolver } from "../../sheet/sheetNameResolver";
import { formatSheetNameForA1 } from "../../sheet/formatSheetNameForA1.js";
import { formatA1 } from "../../document/coords.js";

import { diffYjsWorkbookVersionAgainstCurrent } from "../../versioning/index.js";
import { FormulaDiffView } from "../../versioning/ui/FormulaDiffView.js";

type VersionManagerLike = {
  doc: { encodeState(): Uint8Array };
  getVersion(versionId: string): Promise<{ snapshot: Uint8Array } | null>;
};

type SheetMetaChange = { id: string; field: string; before: unknown; after: unknown };

type WorkbookDiff = {
  sheets: { added: unknown[]; removed: unknown[]; renamed: unknown[]; moved: unknown[]; metaChanged?: SheetMetaChange[] };
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
  if (value === null || value === undefined) return "∅";
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
  if (opts.encrypted) return opts.keyId ? `Encrypted (${opts.keyId})` : "Encrypted";
  const formula = opts.formula ?? null;
  if (formula) return formula;
  if (opts.value === null || opts.value === undefined) return "∅";
  return summarizeJson(opts.value);
}

function formatSheetMetaField(field: string): string {
  if (field === "visibility") return "Visibility";
  if (field === "tabColor") return "Tab color";
  if (field === "view.frozenRows") return "Frozen rows";
  if (field === "view.frozenCols") return "Frozen columns";
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
    })();

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

    const metaChangedIds = new Set((diff.sheets?.metaChanged ?? []).map((c) => c.id));
    setSelectedSheetId((prev) => {
      if (prev && entries.some((e) => e.sheetId === prev)) return prev;
      const preferred = entries.find((e) => hasAnyCellChanges(e.diff) || metaChangedIds.has(e.sheetId)) ?? entries[0];
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
    return (diff.cellsBySheet ?? []).map((entry) => {
      const name = sheetDisplayName(entry.sheetId, entry.sheetName, sheetNameResolver);
      const metaCount = metaCounts.get(entry.sheetId) ?? 0;
      const cellChanges = hasAnyCellChanges(entry.diff);
      const suffix = cellChanges || metaCount > 0 ? (metaCount > 0 && !cellChanges ? ` (meta: ${metaCount})` : "") : " (no changes)";
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
        <h4 className="collab-version-history__compare-title">Compare</h4>
      </div>

      {!versionId ? (
        <div className="collab-version-history__compare-empty">Select a version above to compare it against the current workbook.</div>
      ) : loading ? (
        <div className="collab-version-history__compare-loading">Computing diff…</div>
      ) : diffError ? (
        <div className="collab-version-history__compare-error">{diffError}</div>
      ) : !diff || !summary ? (
        <div className="collab-version-history__compare-empty">No diff available.</div>
      ) : (
        <>
          <div className="collab-version-history__compare-summary">
            <div className="collab-version-history__compare-summary-group">
              <div className="collab-version-history__compare-summary-label">Sheets</div>
              <div className="collab-version-history__compare-summary-badges">
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--added">
                  Added: {summary.sheets.added}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--removed">
                  Removed: {summary.sheets.removed}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  Renamed: {summary.sheets.renamed}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  Meta: {summary.sheets.metaChanged}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--moved">
                  Moved: {summary.sheets.moved}
                </span>
              </div>
            </div>

            <div className="collab-version-history__compare-summary-group">
              <div className="collab-version-history__compare-summary-label">Cells</div>
              <div className="collab-version-history__compare-summary-badges">
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--added">
                  Added: {summary.cells.added}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--removed">
                  Removed: {summary.cells.removed}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  Modified: {summary.cells.modified}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--moved">
                  Moved: {summary.cells.moved}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--format">
                  Format-only: {summary.cells.formatOnly}
                </span>
              </div>
            </div>

            <div className="collab-version-history__compare-summary-group">
              <div className="collab-version-history__compare-summary-label">Workbook metadata</div>
              <div className="collab-version-history__compare-summary-badges">
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--added">
                  Added: {summary.metadata.added}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--removed">
                  Removed: {summary.metadata.removed}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  Modified: {summary.metadata.modified}
                </span>
              </div>
            </div>

            <div className="collab-version-history__compare-summary-group">
              <div className="collab-version-history__compare-summary-label">Named ranges</div>
              <div className="collab-version-history__compare-summary-badges">
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--added">
                  Added: {summary.namedRanges.added}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--removed">
                  Removed: {summary.namedRanges.removed}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  Modified: {summary.namedRanges.modified}
                </span>
              </div>
            </div>

            <div className="collab-version-history__compare-summary-group">
              <div className="collab-version-history__compare-summary-label">Comments</div>
              <div className="collab-version-history__compare-summary-badges">
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--added">
                  Added: {summary.comments.added}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--removed">
                  Removed: {summary.comments.removed}
                </span>
                <span className="collab-version-history__compare-badge collab-version-history__compare-badge--modified">
                  Modified: {summary.comments.modified}
                </span>
              </div>
            </div>
          </div>

          <div className="collab-version-history__compare-controls">
            <label className="collab-version-history__compare-control">
              Sheet{" "}
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
            <div className="collab-version-history__compare-empty">No sheet selected.</div>
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
                  { key: "added", label: "Added", className: "collab-version-history__diff-row--added", items: sd.added },
                  { key: "removed", label: "Removed", className: "collab-version-history__diff-row--removed", items: sd.removed },
                  { key: "modified", label: "Modified", className: "collab-version-history__diff-row--modified", items: sd.modified },
                  { key: "moved", label: "Moved", className: "collab-version-history__diff-row--moved", items: sd.moved },
                  { key: "formatOnly", label: "Format-only", className: "collab-version-history__diff-row--format", items: sd.formatOnly },
                ];

                const any = groups.some((g) => g.items.length > 0);

                 return (
                   <>
                     {(() => {
                       const meta = (diff.sheets?.metaChanged ?? []).filter((c) => c.id === selectedSheet.sheetId);
                       if (meta.length === 0) return null;
                       return (
                         <div className="collab-version-history__diff-group">
                           <div className="collab-version-history__diff-group-title">Sheet metadata ({meta.length})</div>
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
                                       <span className="collab-version-history__diff-row-label">Before:</span>{" "}
                                       <span className="collab-version-history__diff-value">{summarizeJson(change.before)}</span>
                                     </div>
                                     <div className="collab-version-history__diff-row-value">
                                       <span className="collab-version-history__diff-row-label">After:</span>{" "}
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

                     {!any ? <div className="collab-version-history__sheet-diff-empty">No cell changes on this sheet.</div> : null}

                     {groups.map((g) => {
                       if (g.items.length === 0) return null;
                       return (
                         <div key={g.key} className="collab-version-history__diff-group">
                          <div className="collab-version-history__diff-group-title">
                            {g.label} ({g.items.length})
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
                                                <span className="collab-version-history__diff-row-label">Before:</span>{" "}
                                                <span className="collab-version-history__diff-value">{oldSummary}</span>
                                              </div>
                                              <div className="collab-version-history__diff-row-value">
                                                <span className="collab-version-history__diff-row-label">After:</span>{" "}
                                                <span className="collab-version-history__diff-value">{newSummary}</span>
                                              </div>
                                             </div>
                                            {formulaChanged ? (
                                              <div className="collab-version-history__formula-diff" aria-label="Formula diff">
                                                <FormulaDiffView before={change.oldFormula ?? null} after={change.newFormula ?? null} />
                                              </div>
                                            ) : null}
                                          </>
                                        ) : g.key === "added" ? (
                                          <div className="collab-version-history__diff-row-value">
                                            <span className="collab-version-history__diff-row-label">After:</span>{" "}
                                            <span className="collab-version-history__diff-value">{newSummary}</span>
                                          </div>
                                        ) : (
                                          <div className="collab-version-history__diff-row-value">
                                            <span className="collab-version-history__diff-row-label">Before:</span>{" "}
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
