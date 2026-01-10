import React, { useEffect, useMemo, useState } from "react";

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

export type Actor = { userId: string; role: "owner" | "admin" | "editor" | "viewer" };

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
        <div>Merge requires owner/admin permissions.</div>
      </div>
    );
  }

  return (
    <div style={{ padding: 12, fontFamily: "system-ui, sans-serif" }}>
      <h3>Merge branch: {sourceBranch}</h3>
      {error && <div style={{ color: "var(--error)" }}>{error}</div>}
      {!preview ? (
        <div>Loading…</div>
      ) : (
        <>
          <div style={{ marginBottom: 8 }}>
            Conflicts: {preview.conflicts.length}
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
                    <div style={{ color: "var(--text-secondary)" }}>Base</div>
                    <div>{cellSummary(c.base)}</div>
                  </div>
                  <div>
                    <div style={{ color: "var(--text-secondary)" }}>Ours</div>
                    <div>{cellSummary(c.ours)}</div>
                  </div>
                  <div>
                    <div style={{ color: "var(--text-secondary)" }}>Theirs</div>
                    <div>{cellSummary(c.theirs)}</div>
                  </div>
                </div>
              ) : (
                <div style={{ display: "flex", gap: 8 }}>
                  <div>Ours → {c.ours?.to ?? "?"}</div>
                  <div>Theirs → {c.theirs?.to ?? "?"}</div>
                </div>
              )}

              <div style={{ marginTop: 8, display: "flex", gap: 8 }}>
                <button
                  onClick={() => {
                    setResolutions(new Map(resolutions).set(idx, { conflictIndex: idx, choice: "ours" }));
                  }}
                >
                  Choose ours
                </button>
                <button
                  onClick={() => {
                    setResolutions(new Map(resolutions).set(idx, { conflictIndex: idx, choice: "theirs" }));
                  }}
                >
                  Choose theirs
                </button>
                <button
                  onClick={() => {
                    const manual =
                      c.type === "move"
                        ? window.prompt("Move destination", c.ours?.to ?? "")
                        : window.prompt("Manual JSON cell value", "");
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
                  Manual…
                </button>
              </div>
            </div>
          ))}

          <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
            <button onClick={onClose}>Cancel</button>
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
              Apply merge
            </button>
          </div>
        </>
      )}
    </div>
  );
}
