import React, { useEffect, useMemo, useState } from "react";

import { t, tWithVars } from "../../i18n/index.js";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver";
import { formatSheetNameForA1 } from "../../sheet/formatSheetNameForA1.js";
import { showInputBox } from "../../extensions/ui.js";

import { FormulaDiffView } from "../../versioning/ui/FormulaDiffView.js";

export type Cell = { value?: unknown; formula?: string; format?: Record<string, unknown>; enc?: unknown };

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

type ManualCellDraft = {
  deleteCell: boolean;
  /**
   * When a conflict contains encrypted variants, manual resolution can choose
   * which encrypted payload to keep without asking the user to edit raw JSON.
   */
  encSource: "custom" | "base" | "ours" | "theirs";
  valueText: string;
  formulaText: string;
  formatText: string;
  formatError: string | null;
};

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function encKeyId(enc: unknown): string | null {
  if (!isPlainObject(enc)) return null;
  const maybe = (enc as Record<string, unknown>)["keyId"] ?? (enc as Record<string, unknown>)["kid"];
  return typeof maybe === "string" && maybe.trim().length > 0 ? maybe : null;
}

function encryptedCellText(enc: unknown): string {
  const keyId = encKeyId(enc);
  return keyId ? `Encrypted cell (keyId: ${keyId})` : "Encrypted cell";
}

function truncate(text: string, maxLen: number): string {
  return text.length > maxLen ? `${text.slice(0, maxLen)}…` : text;
}

function valueSummary(value: unknown): string {
  if (value === null || value === undefined) return "∅";
  try {
    return truncate(JSON.stringify(value), 200);
  } catch {
    return truncate(String(value), 200);
  }
}

function isSheetPresenceState(value: unknown): value is { meta: unknown; cells: unknown } {
  if (!isPlainObject(value)) return false;
  return isPlainObject((value as Record<string, unknown>).meta) && isPlainObject((value as Record<string, unknown>).cells);
}

function formatSummary(format: unknown): string {
  if (format === null || format === undefined) return "∅";
  try {
    return truncate(JSON.stringify(format, null, 2), 2000);
  } catch {
    return truncate(String(format), 2000);
  }
}

function jsonSummary(value: unknown) {
  if (value === null || value === undefined) return "∅";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);

  // Sheet presence conflicts can embed huge cell maps; show a compact summary and
  // avoid traversing `cells` for UI previews.
  if (isSheetPresenceState(value)) {
    return jsonSummary({
      meta: (value as any).meta,
      cells: "[cells]",
    });
  }

  const preview = (inner: unknown, depth: number): unknown => {
    if (inner === null || inner === undefined) return null;
    if (typeof inner === "string" || typeof inner === "number" || typeof inner === "boolean") return inner;
    if (depth >= 2) return "[Object]";

    if (Array.isArray(inner)) {
      const sliced = inner.slice(0, 20).map((v) => preview(v, depth + 1));
      return inner.length > 20 ? [...sliced, "…"] : sliced;
    }

    if (typeof inner !== "object") return String(inner);

    const obj = inner as Record<string, unknown>;

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
  };

  try {
    const json = JSON.stringify(preview(value, 0));
    return json.length > 200 ? `${json.slice(0, 200)}…` : json;
  } catch {
    return String(value);
  }
}

function conflictHeader(c: MergeConflict, sheetNameResolver: SheetNameResolver | null) {
  const displayName = (sheetId: string | null | undefined): string => {
    const id = String(sheetId ?? "").trim();
    if (!id) return "?";
    return sheetNameResolver?.getSheetNameById(id) ?? id;
  };

  if (c.type === "cell" || c.type === "move") {
    const sheetName = displayName(c.sheetId);
    return `${formatSheetNameForA1(sheetName)}!${c.cell} (${c.reason})`;
  }
  if (c.type === "sheet") {
    if (c.reason === "rename") return `sheet rename: ${displayName(c.sheetId)}`;
    if (c.reason === "order") return "sheet order";
    if (c.reason === "presence") return `sheet presence: ${displayName(c.sheetId)}`;
    return "sheet";
  }
  if (c.type === "namedRange") return `named range: ${c.key}`;
  if (c.type === "comment") return `comment: ${c.id}`;
  if (c.type === "metadata") return `metadata: ${c.key}`;
  // Exhaustive fallback.
  return "conflict";
}

function cellHasValue(cell: Cell | null): boolean {
  if (!cell) return false;
  // Treat null as empty to match DocumentController semantics.
  return cell.value !== null && cell.value !== undefined;
}

function cellHasFormula(cell: Cell | null): boolean {
  return Boolean(cell?.formula && cell.formula.trim().length > 0);
}

function cellHasEnc(cell: Cell | null): boolean {
  return cell?.enc !== null && cell?.enc !== undefined;
}

function stringifyForKey(value: unknown): string {
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function cellValueKey(cell: Cell | null): string {
  if (!cell) return "∅";
  if (cellHasEnc(cell)) return `enc:${stringifyForKey(cell.enc)}`;
  if (cellHasValue(cell)) return `value:${stringifyForKey(cell.value)}`;
  return "∅";
}

function CellInlineView({ cell }: { cell: Cell | null }) {
  if (!cell) return <span className="branch-merge__empty">∅</span>;
  if (cellHasEnc(cell)) return <span className="branch-merge__encrypted">{encryptedCellText(cell.enc)}</span>;
  if (cellHasFormula(cell)) {
    return <FormulaDiffView before={cell.formula ?? null} after={cell.formula ?? null} />;
  }
  if (cellHasValue(cell)) return <span className="branch-merge__value">{valueSummary(cell.value)}</span>;
  return <span className="branch-merge__empty">∅</span>;
}

function CellConflictColumn({
  label,
  cell,
  baseCell,
  baseFormula,
  showEnc,
  showFormula,
  showValue,
  showFormat,
  formulaMode,
}: {
  label: string;
  cell: Cell | null;
  baseCell: Cell | null;
  baseFormula: string | null;
  showEnc: boolean;
  showFormula: boolean;
  showValue: boolean;
  showFormat: boolean;
  formulaMode: "base" | "ours" | "theirs";
}) {
  const currentFormula = cell?.formula ?? null;
  const formulaOld = baseFormula;
  const formulaNew = formulaMode === "base" ? baseFormula : currentFormula;
  const showValueDiff = formulaMode !== "base" && cellValueKey(baseCell) !== cellValueKey(cell);

  return (
    <div className="branch-merge__cell-column">
      <div className="branch-merge__conflict-label">{label}</div>

      {showEnc ? (
        <div className="branch-merge__cell-section">
          <div className="branch-merge__cell-section-title">Encrypted</div>
          <div className="branch-merge__cell-section-body">
            {cellHasEnc(cell) ? <span className="branch-merge__encrypted">{encryptedCellText(cell?.enc)}</span> : "∅"}
          </div>
        </div>
      ) : null}

      {showFormula ? (
        <div className="branch-merge__cell-section">
          <div className="branch-merge__cell-section-title">Formula</div>
          <div className="branch-merge__cell-section-body">
            {cellHasEnc(cell) ? (
              <span className="branch-merge__encrypted">{encryptedCellText(cell?.enc)}</span>
            ) : (
              <FormulaDiffView before={formulaOld} after={formulaNew} />
            )}
          </div>
        </div>
      ) : null}

      {showValue ? (
        <div className="branch-merge__cell-section">
          <div className="branch-merge__cell-section-title">Value</div>
          <div className="branch-merge__cell-section-body">
            {showValueDiff ? (
              <>
                <span className={cellHasEnc(baseCell) ? "branch-merge__encrypted" : undefined}>
                  {cellHasEnc(baseCell)
                    ? encryptedCellText(baseCell?.enc)
                    : cellHasValue(baseCell)
                      ? valueSummary(baseCell?.value)
                      : "∅"}
                </span>
                <span className="branch-merge__value-diff-arrow"> → </span>
                <span className={cellHasEnc(cell) ? "branch-merge__encrypted" : undefined}>
                  {cellHasEnc(cell)
                    ? encryptedCellText(cell?.enc)
                    : cellHasValue(cell)
                      ? valueSummary(cell?.value)
                      : "∅"}
                </span>
              </>
            ) : cellHasEnc(cell) ? (
              <span className="branch-merge__encrypted">{encryptedCellText(cell?.enc)}</span>
            ) : cellHasValue(cell) ? (
              valueSummary(cell?.value)
            ) : (
              "∅"
            )}
          </div>
        </div>
      ) : null}

      {showFormat ? (
        <div className="branch-merge__cell-section">
          <div className="branch-merge__cell-section-title">Format</div>
          <pre className="branch-merge__cell-json">{formatSummary(cell?.format)}</pre>
        </div>
      ) : null}
    </div>
  );
}

function valueToEditorText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function parseValueFromEditorText(text: string): unknown {
  const trimmed = text.trim();
  if (!trimmed) return undefined;

  // Prefer plain strings by default; only parse JSON when the shape strongly
  // implies non-string intent.
  if (trimmed === "true") return true;
  if (trimmed === "false") return false;

  const looksNumber = /^-?(?:\d+|\d*\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed);
  if (looksNumber) return Number(trimmed);

  const looksJson = trimmed.startsWith("{") || trimmed.startsWith("[") || trimmed.startsWith('"');
  if (looksJson) {
    try {
      return JSON.parse(trimmed);
    } catch {
      // Fall through to a plain string.
    }
  }

  // Treat `null` as a literal string; users can clear the cell explicitly using
  // the "Delete cell" toggle.
  return trimmed;
}

function normalizeFormulaInput(text: string | null | undefined): string | null {
  const trimmed = String(text ?? "").trim();
  if (!trimmed) return null;
  const withoutEquals = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const body = withoutEquals.trim();
  if (!body) return null;
  return `=${body}`;
}

function normalizeManualCell(cell: Cell | null): Cell | null {
  if (!cell || typeof cell !== "object") return null;

  const out: Cell = {};

  if (cell.enc !== null && cell.enc !== undefined) out.enc = cell.enc;

  const formula = normalizeFormulaInput(cell.formula);
  if (formula) out.formula = formula;

  if (cell.value !== null && cell.value !== undefined) out.value = cell.value;

  if (cell.format !== null && cell.format !== undefined) out.format = cell.format;

  if (
    out.enc === undefined &&
    out.formula === undefined &&
    out.value === undefined &&
    out.format === undefined
  ) {
    return null;
  }

  // Enforce mutual exclusion between enc/formula/value.
  if (out.enc !== undefined) {
    delete out.formula;
    delete out.value;
  } else if (out.formula !== undefined) {
    delete out.value;
  }

  return out;
}

function initialDraftForCellConflict(
  conflict: Extract<MergeConflict, { type: "cell" }>,
  seed: Cell | null,
  seedSource: "base" | "ours" | "theirs"
): ManualCellDraft {
  const hasEnc = cellHasEnc(conflict.base) || cellHasEnc(conflict.ours) || cellHasEnc(conflict.theirs);

  return {
    deleteCell: seed === null,
    encSource: hasEnc && cellHasEnc(seed) ? seedSource : "custom",
    valueText: seed && seed.formula ? "" : valueToEditorText(seed?.value),
    formulaText: seed && seed.formula ? normalizeFormulaInput(seed.formula) ?? seed.formula : "",
    formatText: seed?.format ? JSON.stringify(seed.format, null, 2) : "",
    formatError: null,
  };
}

function manualCellFromDraft(
  conflict: Extract<MergeConflict, { type: "cell" }>,
  draft: ManualCellDraft
): Cell | null {
  if (draft.deleteCell) return null;

  const hasEnc = cellHasEnc(conflict.base) || cellHasEnc(conflict.ours) || cellHasEnc(conflict.theirs);

  if (hasEnc && draft.encSource !== "custom") {
    const chosen =
      draft.encSource === "base"
        ? conflict.base
        : draft.encSource === "ours"
          ? conflict.ours
          : conflict.theirs;
    const cell = chosen ? { ...chosen } : null;
    if (!cell) return null;

    const formatText = draft.formatText.trim();
    if (!formatText) {
      delete cell.format;
    } else {
      try {
        const parsed = JSON.parse(formatText);
        if (parsed !== null && parsed !== undefined) cell.format = parsed as Record<string, unknown>;
        else delete cell.format;
      } catch {
        // Validation is handled separately; ignore parse errors here.
      }
    }

    return normalizeManualCell(cell);
  }

  /** @type {Cell} */
  const cell: Cell = {};

  const formula = normalizeFormulaInput(draft.formulaText);
  if (formula) {
    cell.formula = formula;
  } else {
    const nextValue = parseValueFromEditorText(draft.valueText);
    if (nextValue !== undefined) cell.value = nextValue;
  }

  const formatText = draft.formatText.trim();
  if (formatText.length > 0) {
    try {
      const parsed = JSON.parse(formatText);
      if (parsed !== null && parsed !== undefined) cell.format = parsed as Record<string, unknown>;
    } catch {
      // Caller is expected to surface format parse errors via draft.formatError.
    }
  }

  return normalizeManualCell(cell);
}

export function MergeBranchPanel({
  actor,
  branchService,
  sourceBranch,
  sheetNameResolver = null,
  mutationsDisabled = false,
  onClose
}: {
  actor: Actor;
  branchService: BranchService;
  sourceBranch: string;
  sheetNameResolver?: SheetNameResolver | null;
  mutationsDisabled?: boolean;
  onClose: () => void;
}) {
  const [preview, setPreview] = useState<MergePreview | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [resolutions, setResolutions] = useState<Map<number, ConflictResolution>>(new Map());
  const [manualCellDrafts, setManualCellDrafts] = useState<Map<number, ManualCellDraft>>(new Map());

  // Reset any in-progress resolutions when switching merge targets so conflict indices
  // don't get applied to a different preview.
  useEffect(() => {
    setResolutions(new Map());
    setManualCellDrafts(new Map());
  }, [sourceBranch]);

  useEffect(() => {
    if (mutationsDisabled) return;
    void (async () => {
      try {
        setError(null);
        setPreview(await branchService.previewMerge(actor, { sourceBranch }));
      } catch (e) {
        setError((e as Error).message);
      }
    })();
  }, [actor, branchService, sourceBranch, mutationsDisabled]);

  const canManage = useMemo(() => actor.role === "owner" || actor.role === "admin", [actor.role]);

  const hasManualErrors = useMemo(() => {
    if (!preview) return false;
    for (const [idx, resolution] of resolutions) {
      if (resolution.choice !== "manual") continue;
      const conflict = preview.conflicts[idx];
      if (!conflict || conflict.type !== "cell") continue;
      const draft = manualCellDrafts.get(idx);
      if (draft?.formatError) return true;
    }
    return false;
  }, [manualCellDrafts, preview, resolutions]);

  if (!canManage) {
    return (
      <div className="branch-merge branch-merge--permission">
        <div className="branch-merge__permission-warning">{t("branchMerge.permissionWarning")}</div>
      </div>
    );
  }

  return (
    <div className="branch-merge">
      <h3 className="branch-merge__title">{tWithVars("branchMerge.titleWithSource", { sourceBranch })}</h3>
      {error && <div className="branch-merge__error">{error}</div>}
      {!preview ? (
        <div className="branch-merge__loading">{t("branchMerge.loading")}</div>
      ) : (
        <>
          <div className="branch-merge__summary">
            {tWithVars("branchMerge.conflictsCount", { count: preview.conflicts.length })}
          </div>

          {preview.conflicts.map((c, idx) => {
            const resolution = resolutions.get(idx);
            const selectedChoice = resolution?.choice ?? null;

            const choose = (choice: ConflictResolution["choice"]) => {
              if (mutationsDisabled) return;
              setResolutions((prev) => new Map(prev).set(idx, { conflictIndex: idx, choice }));
            };

            const cellDraft = manualCellDrafts.get(idx);

            const updateManualDraft = (next: ManualCellDraft, conflict: Extract<MergeConflict, { type: "cell" }>) => {
              if (mutationsDisabled) return;

              const formatText = next.formatText.trim();
              let formatError: string | null = null;
              if (!next.deleteCell && formatText.length > 0) {
                try {
                  JSON.parse(formatText);
                } catch (e) {
                  formatError = (e as Error).message;
                }
              }

              const normalized: ManualCellDraft = { ...next, formatError };

              setManualCellDrafts((prev) => new Map(prev).set(idx, normalized));
              setResolutions((prev) =>
                new Map(prev).set(idx, {
                  conflictIndex: idx,
                  choice: "manual",
                  manualCell: manualCellFromDraft(conflict, normalized),
                })
              );
            };

            return (
              <div key={`${c.type}-${idx}`} className="branch-merge__conflict">
                <div className="branch-merge__conflict-title">{conflictHeader(c, sheetNameResolver)}</div>

                {c.type === "cell" ? (
                  (() => {
                    const showEnc = cellHasEnc(c.base) || cellHasEnc(c.ours) || cellHasEnc(c.theirs);
                    const showFormula = cellHasFormula(c.base) || cellHasFormula(c.ours) || cellHasFormula(c.theirs);
                    const showValue = cellHasValue(c.base) || cellHasValue(c.ours) || cellHasValue(c.theirs);
                    const showFormat =
                      c.reason === "format" ||
                      c.base?.format !== undefined ||
                      c.ours?.format !== undefined ||
                      c.theirs?.format !== undefined;
                    const baseFormula = c.base?.formula ?? null;

                    return (
                      <div className="branch-merge__conflict-grid branch-merge__conflict-grid--cell">
                        <CellConflictColumn
                          label={t("branchMerge.conflict.base")}
                          cell={c.base}
                          baseCell={c.base}
                          baseFormula={baseFormula}
                          showEnc={showEnc}
                          showFormula={showFormula}
                          showValue={showValue}
                          showFormat={showFormat}
                          formulaMode="base"
                        />
                        <CellConflictColumn
                          label={t("branchMerge.conflict.ours")}
                          cell={c.ours}
                          baseCell={c.base}
                          baseFormula={baseFormula}
                          showEnc={showEnc}
                          showFormula={showFormula}
                          showValue={showValue}
                          showFormat={showFormat}
                          formulaMode="ours"
                        />
                        <CellConflictColumn
                          label={t("branchMerge.conflict.theirs")}
                          cell={c.theirs}
                          baseCell={c.base}
                          baseFormula={baseFormula}
                          showEnc={showEnc}
                          showFormula={showFormula}
                          showValue={showValue}
                          showFormat={showFormat}
                          formulaMode="theirs"
                        />
                      </div>
                    );
                  })()
                ) : c.type === "move" ? (
                  <div className="branch-merge__conflict-move">
                    <div className="branch-merge__conflict-move-base">
                      <div className="branch-merge__conflict-label">{t("branchMerge.conflict.base")}</div>
                      <div className="branch-merge__conflict-value">
                        <CellInlineView cell={c.base} />
                      </div>
                    </div>
                    <div className="branch-merge__conflict-move-dest">
                      <div>{tWithVars("branchMerge.conflict.move.oursTo", { to: c.ours?.to ?? "?" })}</div>
                      <div>{tWithVars("branchMerge.conflict.move.theirsTo", { to: c.theirs?.to ?? "?" })}</div>
                    </div>
                  </div>
                ) : (
                  <div className="branch-merge__conflict-grid">
                    <div>
                      <div className="branch-merge__conflict-label">{t("branchMerge.conflict.base")}</div>
                      <div className="branch-merge__conflict-value">{jsonSummary(c.base)}</div>
                    </div>
                    <div>
                      <div className="branch-merge__conflict-label">{t("branchMerge.conflict.ours")}</div>
                      <div className="branch-merge__conflict-value">{jsonSummary(c.ours)}</div>
                    </div>
                    <div>
                      <div className="branch-merge__conflict-label">{t("branchMerge.conflict.theirs")}</div>
                      <div className="branch-merge__conflict-value">{jsonSummary(c.theirs)}</div>
                    </div>
                  </div>
                )}

                <div className="branch-merge__resolution-actions">
                  <button
                    disabled={mutationsDisabled}
                    data-selected={selectedChoice === "ours"}
                    aria-pressed={selectedChoice === "ours"}
                    onClick={() => choose("ours")}
                  >
                    {t("branchMerge.chooseOurs")}
                  </button>
                  <button
                    disabled={mutationsDisabled}
                    data-selected={selectedChoice === "theirs"}
                    aria-pressed={selectedChoice === "theirs"}
                    onClick={() => choose("theirs")}
                  >
                    {t("branchMerge.chooseTheirs")}
                  </button>
                  <button
                    disabled={mutationsDisabled}
                    data-selected={selectedChoice === "manual"}
                    aria-pressed={selectedChoice === "manual"}
                    onClick={async () => {
                      if (mutationsDisabled) return;
                      try {
                        if (c.type === "cell") {
                          // Seed structured editor from ours (merge default).
                          const seed = c.ours ? { ...c.ours } : null;
                          const draft = manualCellDrafts.get(idx) ?? initialDraftForCellConflict(c, seed, "ours");
                          updateManualDraft(draft, c);
                          return;
                        }

                        const manual =
                          c.type === "move"
                            ? await showInputBox({
                                prompt: t("branchMerge.prompt.moveDestination"),
                                value: c.ours?.to ?? "",
                              })
                            : c.type === "sheet" && c.reason === "rename"
                              ? await showInputBox({
                                  prompt: t("branchMerge.prompt.manualJson"),
                                  value: String(c.ours ?? ""),
                                })
                              : c.type === "sheet" && c.reason === "order"
                                ? await showInputBox({
                                    prompt: t("branchMerge.prompt.manualJson"),
                                    value: JSON.stringify(c.ours ?? [], null, 2),
                                    type: "textarea",
                                  })
                                : c.type === "sheet" && c.reason === "presence"
                                  ? // Presence conflicts can embed large cell maps; avoid
                                    // pre-populating the prompt with a giant JSON blob.
                                    await showInputBox({
                                      prompt: t("branchMerge.prompt.manualJson"),
                                      value: "",
                                      type: "textarea",
                                    })
                                  : await showInputBox({
                                      prompt: t("branchMerge.prompt.manualJson"),
                                      value: JSON.stringify(c.ours ?? null, null, 2),
                                      type: "textarea",
                                    });

                        if (manual === null) return;

                        const next: ConflictResolution = { conflictIndex: idx, choice: "manual" };

                        if (c.type === "move") {
                          next.manualMoveTo = manual;
                        } else if (c.type === "sheet" && c.reason === "rename") {
                          next.manualSheetName = manual.length > 0 ? manual : null;
                        } else if (c.type === "sheet" && c.reason === "order") {
                          next.manualSheetOrder = manual ? (JSON.parse(manual) as string[]) : [];
                        } else if (c.type === "sheet" && c.reason === "presence") {
                          next.manualSheetState = manual ? JSON.parse(manual) : null;
                        } else if (c.type === "metadata") {
                          next.manualMetadataValue = manual ? JSON.parse(manual) : null;
                        } else if (c.type === "namedRange") {
                          next.manualNamedRangeValue = manual ? JSON.parse(manual) : null;
                        } else if (c.type === "comment") {
                          next.manualCommentValue = manual ? JSON.parse(manual) : null;
                        }

                        setResolutions((prev) => new Map(prev).set(idx, next));
                      } catch (e) {
                        setError((e as Error).message);
                      }
                    }}
                  >
                    {t("branchMerge.manual")}
                  </button>
                </div>

                {c.type === "cell" && selectedChoice === "manual" ? (
                  (() => {
                    const draft = cellDraft ?? initialDraftForCellConflict(c, c.ours ? { ...c.ours } : null, "ours");
                    const hasEnc = cellHasEnc(c.base) || cellHasEnc(c.ours) || cellHasEnc(c.theirs);
                    const contentLocked = hasEnc && draft.encSource !== "custom";
                    const contentDisabled = mutationsDisabled || draft.deleteCell || contentLocked;
                    const formatDisabled = mutationsDisabled || draft.deleteCell;

                    return (
                      <div className="branch-merge__manual-cell-editor">
                        <label className="branch-merge__manual-cell-toggle">
                          <input
                            type="checkbox"
                            disabled={mutationsDisabled}
                            checked={draft.deleteCell}
                            onChange={(e) => updateManualDraft({ ...draft, deleteCell: e.target.checked }, c)}
                          />
                          Delete cell
                        </label>

                        {hasEnc ? (
                          <div className="branch-merge__manual-cell-row">
                            <div className="branch-merge__manual-cell-label">Encrypted</div>
                            <select
                              value={draft.encSource}
                              disabled={mutationsDisabled || draft.deleteCell}
                              onChange={(e) => {
                                const encSource = e.target.value as ManualCellDraft["encSource"];
                                let next: ManualCellDraft = { ...draft, deleteCell: false, encSource };

                                if (encSource !== "custom") {
                                  const chosen =
                                    encSource === "base" ? c.base : encSource === "ours" ? c.ours : c.theirs;
                                  next = {
                                    ...next,
                                    // Encrypted payloads cannot be edited, so clear the content fields.
                                    formulaText: "",
                                    valueText: "",
                                    formatText: chosen?.format ? JSON.stringify(chosen.format, null, 2) : "",
                                  };
                                }

                                updateManualDraft(next, c);
                              }}
                            >
                              <option value="custom">Custom (unencrypted)</option>
                              {cellHasEnc(c.base) ? <option value="base">Use base</option> : null}
                              {cellHasEnc(c.ours) ? <option value="ours">Use ours</option> : null}
                              {cellHasEnc(c.theirs) ? <option value="theirs">Use theirs</option> : null}
                            </select>
                          </div>
                        ) : null}

                        <div className="branch-merge__manual-cell-row">
                          <div className="branch-merge__manual-cell-label">Formula</div>
                          <input
                            value={draft.formulaText}
                            disabled={contentDisabled}
                            placeholder="=SUM(A1:A10)"
                            onChange={(e) => updateManualDraft({ ...draft, formulaText: e.target.value }, c)}
                          />
                        </div>

                        <div className="branch-merge__manual-cell-row">
                          <div className="branch-merge__manual-cell-label">Value</div>
                          <input
                            value={draft.valueText}
                            disabled={contentDisabled || draft.formulaText.trim().length > 0}
                            placeholder="123"
                            onChange={(e) => updateManualDraft({ ...draft, valueText: e.target.value }, c)}
                          />
                        </div>

                        <div className="branch-merge__manual-cell-row branch-merge__manual-cell-row--format">
                          <div className="branch-merge__manual-cell-label">Format (JSON)</div>
                          <textarea
                            value={draft.formatText}
                            disabled={formatDisabled}
                            onChange={(e) => updateManualDraft({ ...draft, formatText: e.target.value }, c)}
                          />
                        </div>

                        {draft.formatError ? (
                          <div className="branch-merge__manual-cell-error">{draft.formatError}</div>
                        ) : null}
                      </div>
                    );
                  })()
                ) : null}
              </div>
            );
          })}

          <div className="branch-merge__footer-actions">
            <button onClick={onClose}>{t("branchMerge.cancel")}</button>
            <button
              disabled={mutationsDisabled || preview.conflicts.length !== resolutions.size || hasManualErrors}
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
