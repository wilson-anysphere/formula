import { showToast } from "../extensions/ui.js";
import { cellToA1, rangeToA1 } from "../selection/a1";

type RejectionReason = "permission" | "encryption" | "unknown";
type RejectionKind = "cell" | "format" | "rangeRun" | "drawing" | "chart" | "undoRedo" | "unknown";

// Editing surfaces may call this helper in response to every key press (e.g. typing into a
// read-only sheet). To avoid spamming users with identical warnings, throttle repeated toasts.
const REJECTION_TOAST_THROTTLE_MS = 1_000;
let lastToastMessage: string | null = null;
let lastToastTime = 0;
let lastToastRoot: HTMLElement | null = null;

function isCellDelta(delta: any): delta is { sheetId?: string; row: number; col: number } {
  return delta != null && typeof delta === "object" && Number.isInteger(delta.row) && Number.isInteger(delta.col);
}

function isRangeRunDelta(
  delta: any,
): delta is { sheetId?: string; col: number; startRow: number; endRowExclusive: number } {
  return (
    delta != null &&
    typeof delta === "object" &&
    Number.isInteger(delta.col) &&
    Number.isInteger(delta.startRow) &&
    Number.isInteger(delta.endRowExclusive)
  );
}

function isFormatDelta(delta: any): boolean {
  return delta != null && typeof delta === "object" && typeof delta.layer === "string";
}

function inferRejectionReason(rejected: any[]): RejectionReason {
  for (const delta of rejected) {
    const reason = typeof delta?.rejectionReason === "string" ? delta.rejectionReason : null;
    if (reason === "encryption") return "encryption";
  }
  for (const delta of rejected) {
    const reason = typeof delta?.rejectionReason === "string" ? delta.rejectionReason : null;
    if (reason === "permission") return "permission";
  }
  return "unknown";
}

function inferRejectionKind(rejected: any[]): RejectionKind {
  for (const delta of rejected) {
    const kind = typeof delta?.rejectionKind === "string" ? delta.rejectionKind : null;
    if (kind === "cell" || kind === "format" || kind === "rangeRun" || kind === "drawing" || kind === "chart" || kind === "undoRedo")
      return kind;
  }

  // Backwards compatibility: callers may be using a binder that doesn't annotate deltas.
  for (const delta of rejected) {
    if (isCellDelta(delta)) return "cell";
  }
  for (const delta of rejected) {
    if (isRangeRunDelta(delta)) return "rangeRun";
  }
  for (const delta of rejected) {
    if (isFormatDelta(delta)) return "format";
  }

  return "unknown";
}

function describeRejectedTarget(kind: RejectionKind, rejected: any[]): string | null {
  if (kind === "cell") {
    const first = rejected.find((d) => isCellDelta(d)) ?? null;
    if (!first) return null;
    if (first.row < 0 || first.col < 0) return null;
    return cellToA1({ row: first.row, col: first.col });
  }

  if (kind === "rangeRun") {
    const first = rejected.find((d) => isRangeRunDelta(d)) ?? null;
    if (!first) return null;
    if (first.col < 0 || first.startRow < 0) return null;
    const endRow = first.endRowExclusive - 1;
    if (!Number.isInteger(endRow) || endRow < first.startRow) return null;
    return rangeToA1({ startRow: first.startRow, startCol: first.col, endRow, endCol: first.col });
  }

  if (kind === "drawing" || kind === "chart" || kind === "undoRedo") {
    return null;
  }

  return null;
}

/**
 * Best-effort UX for binder edit rejections.
 *
 * The collaboration binder will revert local UI state when an edit is rejected (permissions,
 * missing encryption key, etc). Without feedback, this can look like the UI "snapped back".
 */
export function showCollabEditRejectedToast(rejected: any[]): void {
  if (!Array.isArray(rejected) || rejected.length === 0) return;

  // Tests (and some desktop integration points) recreate `#toast-root` between runs. If we keep
  // throttling state across different roots, a toast that would normally appear can be suppressed
  // because the previous run emitted the same message moments earlier.
  //
  // In the real app the root is long-lived, so this preserves the intended spam protection while
  // keeping tests deterministic.
  if (lastToastRoot && !lastToastRoot.isConnected) {
    lastToastRoot = null;
    lastToastMessage = null;
    lastToastTime = 0;
  }
  const toastRoot = (() => {
    try {
      return document.getElementById("toast-root");
    } catch {
      return null;
    }
  })();
  if (toastRoot !== lastToastRoot) {
    lastToastRoot = toastRoot;
    lastToastMessage = null;
    lastToastTime = 0;
  }

  const reason = inferRejectionReason(rejected);
  const kind = inferRejectionKind(rejected);
  const target = describeRejectedTarget(kind, rejected);

  const message = (() => {
    if (reason === "encryption") {
      return target ? `Missing encryption key for protected cell (${target})` : "Missing encryption key for protected cell";
    }

    if (kind === "format" || kind === "rangeRun") {
      return target
        ? `Read-only: you don't have permission to change formatting (${target})`
        : "Read-only: you don't have permission to change formatting";
    }

    if (kind === "drawing") {
      return "Read-only: you don't have permission to edit drawings";
    }

    if (kind === "chart") {
      return "Read-only: you don't have permission to edit charts";
    }

    if (kind === "undoRedo") {
      return "Read-only: you don't have permission to undo/redo";
    }

    // Default to a simple "read-only" message for cell edits.
    return target ? `Read-only: you don't have permission to edit that cell (${target})` : "Read-only: you don't have permission to edit that cell";
  })();

  const now = Date.now();
  const canThrottle = now > 0 && lastToastTime > 0;
  if (canThrottle && message === lastToastMessage && now - lastToastTime < REJECTION_TOAST_THROTTLE_MS) {
    return;
  }

  try {
    showToast(message, "warning");
    lastToastMessage = message;
    lastToastTime = now;
  } catch {
    // `showToast` requires a #toast-root; some test-only contexts don't include it.
  }
}
