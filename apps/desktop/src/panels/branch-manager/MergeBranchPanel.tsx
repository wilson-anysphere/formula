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
    };

export type MergePreview = {
  merged: unknown;
  conflicts: MergeConflict[];
};

export type Actor = { userId: string; role: "owner" | "admin" | "editor" | "commenter" | "viewer" };

export type ConflictResolution =
  | { conflictIndex: number; choice: "ours" | "theirs" | "manual"; manualCell?: Cell | null }
  | { conflictIndex: number; choice: "ours" | "theirs" | "manual"; manualMoveTo?: string };

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
              <div style={{ fontWeight: 600 }}>
                {c.sheetId}!{c.cell} ({c.reason})
              </div>
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
              ) : (
                <div style={{ display: "flex", gap: 8 }}>
                  <div>{tWithVars("branchMerge.conflict.move.oursTo", { to: c.ours?.to ?? "?" })}</div>
                  <div>{tWithVars("branchMerge.conflict.move.theirsTo", { to: c.theirs?.to ?? "?" })}</div>
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
                    const manual =
                      c.type === "move"
                        ? window.prompt(t("branchMerge.prompt.moveDestination"), c.ours?.to ?? "")
                        : window.prompt(t("branchMerge.prompt.manualJson"), "");
                    if (manual === null) return;
                    if (c.type === "move") {
                      setResolutions(
                        new Map(resolutions).set(idx, {
                          conflictIndex: idx,
                          choice: "manual",
                          manualMoveTo: manual
                        })
                      );
                    } else {
                      setResolutions(
                        new Map(resolutions).set(idx, {
                          conflictIndex: idx,
                          choice: "manual",
                          manualCell: manual ? (JSON.parse(manual) as Cell) : null
                        })
                      );
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
