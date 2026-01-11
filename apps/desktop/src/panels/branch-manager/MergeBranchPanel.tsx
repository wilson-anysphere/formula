import React, { useEffect, useMemo, useState } from "react";

import { t, tWithVars } from "../../i18n/index.js";

export type Cell = { value?: unknown; formula?: string; format?: Record<string, unknown> };

export type MergeConflict =
  | {
      type: "cell";
      sheetId: string;
      cell: string;
      reason: string;
      base: Cell | null;
      ours: Cell | null;
      theirs: Cell | null;
    }
  | {
      type: "move";
      sheetId: string;
      cell: string;
      reason: "move-destination";
      base: Cell | null;
      ours: { to: string } | null;
      theirs: { to: string } | null;
    }
  | {
      type: "sheet";
      reason: "rename" | "order" | "presence";
      sheetId?: string;
      base: unknown;
      ours: unknown;
      theirs: unknown;
    }
  | {
      type: "metadata";
      key: string;
      base: unknown;
      ours: unknown;
      theirs: unknown;
    }
  | {
      type: "namedRange";
      key: string;
      base: unknown;
      ours: unknown;
      theirs: unknown;
    }
  | {
      type: "comment";
      id: string;
      base: unknown;
      ours: unknown;
      theirs: unknown;
    };

export type MergePreview = {
  merged: unknown;
  conflicts: MergeConflict[];
};

export type Actor = { userId: string; role: "owner" | "admin" | "editor" | "commenter" | "viewer" };

export type ConflictResolution = {
  conflictIndex: number;
  choice: "ours" | "theirs" | "manual";
  manualCell?: Cell | null;
  manualMoveTo?: string;
  manualSheetName?: string | null;
  manualSheetOrder?: string[];
  manualSheetState?: unknown;
  manualMetadataValue?: unknown;
  manualNamedRangeValue?: unknown;
  manualCommentValue?: unknown;
};

export type BranchService = {
  previewMerge(actor: Actor, input: { sourceBranch: string }): Promise<MergePreview>;
  merge(actor: Actor, input: { sourceBranch: string; resolutions: ConflictResolution[] }): Promise<unknown>;
};

function cellSummary(cell: Cell | null) {
  if (!cell) return "∅";
  if (cell.formula) return cell.formula;
  if (cell.value !== undefined) return JSON.stringify(cell.value);
  return "∅";
}

function jsonSummary(value: unknown) {
  if (value === null || value === undefined) return "∅";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);

  /**
   * @param {unknown} inner
   * @param {number} depth
   * @returns {unknown}
   */
  function preview(inner, depth) {
    if (inner === null || inner === undefined) return null;
    if (typeof inner === "string" || typeof inner === "number" || typeof inner === "boolean") return inner;
    if (depth >= 2) return "[Object]";

    if (Array.isArray(inner)) {
      const sliced = inner.slice(0, 20).map((v) => preview(v, depth + 1));
      return inner.length > 20 ? [...sliced, "…"] : sliced;
    }

    if (typeof inner !== "object") return String(inner);

    const obj = inner as Record<string, unknown>;

    // Special-case sheet presence conflicts: the cell map can be huge; avoid
    // traversing it for UI summaries.
    if (depth === 0 && typeof obj.meta === "object" && obj.meta !== null && "cells" in obj) {
      return { meta: preview(obj.meta, depth + 1), cells: "[cells]" };
    }

    /** @type {Record<string, unknown>} */
    const out: Record<string, unknown> = {};

    let count = 0;
    let hasMore = false;
    for (const key in obj) {
      if (!Object.prototype.hasOwnProperty.call(obj, key)) continue;
      if (count >= 20) {
        hasMore = true;
        break;
      }
      out[key] = preview(obj[key], depth + 1);
      count += 1;
    }
    if (hasMore) out["…"] = "…";
    return out;
  }

  try {
    const json = JSON.stringify(preview(value, 0));
    return json.length > 200 ? `${json.slice(0, 200)}…` : json;
  } catch {
    return String(value);
  }
}

function conflictHeader(c: MergeConflict) {
  if (c.type === "cell" || c.type === "move") {
    return `${c.sheetId}!${c.cell} (${c.reason})`;
  }
  if (c.type === "sheet") {
    if (c.reason === "rename") return `sheet rename: ${c.sheetId ?? "?"}`;
    if (c.reason === "order") return "sheet order";
    if (c.reason === "presence") return `sheet presence: ${c.sheetId ?? "?"}`;
    return "sheet";
  }
  if (c.type === "namedRange") return `named range: ${c.key}`;
  if (c.type === "comment") return `comment: ${c.id}`;
  if (c.type === "metadata") return `metadata: ${c.key}`;
  // Exhaustive fallback.
  return "conflict";
}

export function MergeBranchPanel({
  actor,
  branchService,
  sourceBranch,
  onClose
}: {
  actor: Actor;
  branchService: BranchService;
  sourceBranch: string;
  onClose: () => void;
}) {
  const [preview, setPreview] = useState<MergePreview | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [resolutions, setResolutions] = useState<Map<number, ConflictResolution>>(new Map());

  useEffect(() => {
    void (async () => {
      try {
        setError(null);
        setPreview(await branchService.previewMerge(actor, { sourceBranch }));
      } catch (e) {
        setError((e as Error).message);
      }
    })();
  }, [actor, branchService, sourceBranch]);

  const canManage = useMemo(() => actor.role === "owner" || actor.role === "admin", [actor.role]);

  if (!canManage) {
    return (
      <div style={{ padding: 12 }}>
        <div>{t("branchMerge.permissionWarning")}</div>
      </div>
    );
  }

  return (
    <div style={{ padding: 12, fontFamily: "system-ui, sans-serif" }}>
      <h3>{tWithVars("branchMerge.titleWithSource", { sourceBranch })}</h3>
      {error && <div style={{ color: "var(--error)" }}>{error}</div>}
      {!preview ? (
        <div>{t("branchMerge.loading")}</div>
      ) : (
        <>
          <div style={{ marginBottom: 8 }}>
            {tWithVars("branchMerge.conflictsCount", { count: preview.conflicts.length })}
          </div>

          {preview.conflicts.map((c, idx) => (
            <div
              key={`${c.type}-${idx}`}
              style={{ border: "1px solid var(--border)", padding: 8, marginBottom: 8 }}
            >
              <div style={{ fontWeight: 600 }}>{conflictHeader(c)}</div>
              {c.type === "cell" ? (
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
                  <div>
                    <div style={{ color: "var(--text-secondary)" }}>{t("branchMerge.conflict.base")}</div>
                    <div>{cellSummary(c.base)}</div>
                  </div>
                  <div>
                    <div style={{ color: "var(--text-secondary)" }}>{t("branchMerge.conflict.ours")}</div>
                    <div>{cellSummary(c.ours)}</div>
                  </div>
                  <div>
                    <div style={{ color: "var(--text-secondary)" }}>{t("branchMerge.conflict.theirs")}</div>
                    <div>{cellSummary(c.theirs)}</div>
                  </div>
                </div>
              ) : c.type === "move" ? (
                <div style={{ display: "flex", gap: 8 }}>
                  <div>{tWithVars("branchMerge.conflict.move.oursTo", { to: c.ours?.to ?? "?" })}</div>
                  <div>{tWithVars("branchMerge.conflict.move.theirsTo", { to: c.theirs?.to ?? "?" })}</div>
                </div>
              ) : (
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
                  <div>
                    <div style={{ color: "var(--text-secondary)" }}>{t("branchMerge.conflict.base")}</div>
                    <div>{jsonSummary(c.base)}</div>
                  </div>
                  <div>
                    <div style={{ color: "var(--text-secondary)" }}>{t("branchMerge.conflict.ours")}</div>
                    <div>{jsonSummary(c.ours)}</div>
                  </div>
                  <div>
                    <div style={{ color: "var(--text-secondary)" }}>{t("branchMerge.conflict.theirs")}</div>
                    <div>{jsonSummary(c.theirs)}</div>
                  </div>
                </div>
              )}

              <div style={{ marginTop: 8, display: "flex", gap: 8 }}>
                <button
                  onClick={() => {
                    setResolutions(new Map(resolutions).set(idx, { conflictIndex: idx, choice: "ours" }));
                  }}
                >
                  {t("branchMerge.chooseOurs")}
                </button>
                <button
                  onClick={() => {
                    setResolutions(new Map(resolutions).set(idx, { conflictIndex: idx, choice: "theirs" }));
                  }}
                >
                  {t("branchMerge.chooseTheirs")}
                </button>
                <button
                  onClick={() => {
                    try {
                      const manual =
                        c.type === "move"
                          ? window.prompt(t("branchMerge.prompt.moveDestination"), c.ours?.to ?? "")
                          : c.type === "sheet" && c.reason === "rename"
                            ? window.prompt(t("branchMerge.prompt.manualJson"), String(c.ours ?? ""))
                            : c.type === "sheet" && c.reason === "order"
                              ? window.prompt(t("branchMerge.prompt.manualJson"), JSON.stringify(c.ours ?? []))
                              : c.type === "sheet" && c.reason === "presence"
                                ? // Presence conflicts can embed large cell maps; avoid
                                  // pre-populating the prompt with a giant JSON blob.
                                  window.prompt(t("branchMerge.prompt.manualJson"), "")
                                : window.prompt(t("branchMerge.prompt.manualJson"), JSON.stringify(c.ours ?? null));

                      if (manual === null) return;

                      const resolution: ConflictResolution = { conflictIndex: idx, choice: "manual" };

                      if (c.type === "move") {
                        resolution.manualMoveTo = manual;
                      } else if (c.type === "cell") {
                        resolution.manualCell = manual ? (JSON.parse(manual) as Cell) : null;
                      } else if (c.type === "sheet" && c.reason === "rename") {
                        resolution.manualSheetName = manual.length > 0 ? manual : null;
                      } else if (c.type === "sheet" && c.reason === "order") {
                        resolution.manualSheetOrder = manual ? (JSON.parse(manual) as string[]) : [];
                      } else if (c.type === "sheet" && c.reason === "presence") {
                        resolution.manualSheetState = manual ? JSON.parse(manual) : null;
                      } else if (c.type === "metadata") {
                        resolution.manualMetadataValue = manual ? JSON.parse(manual) : null;
                      } else if (c.type === "namedRange") {
                        resolution.manualNamedRangeValue = manual ? JSON.parse(manual) : null;
                      } else if (c.type === "comment") {
                        resolution.manualCommentValue = manual ? JSON.parse(manual) : null;
                      }

                      setResolutions(new Map(resolutions).set(idx, resolution));
                    } catch (e) {
                      setError((e as Error).message);
                    }
                  }}
                >
                  {t("branchMerge.manual")}
                </button>
              </div>
            </div>
          ))}

          <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
            <button onClick={onClose}>{t("branchMerge.cancel")}</button>
            <button
              disabled={preview.conflicts.length !== resolutions.size}
              onClick={async () => {
                try {
                  setError(null);
                  await branchService.merge(actor, {
                    sourceBranch,
                    resolutions: Array.from(resolutions.values())
                  });
                  onClose();
                } catch (e) {
                  setError((e as Error).message);
                }
              }}
            >
              {t("branchMerge.applyMerge")}
            </button>
          </div>
        </>
      )}
    </div>
  );
}
