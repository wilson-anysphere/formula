/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { showCollabEditRejectedToast } from "../editRejectionToast";

describe("collab edit rejection toast", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    document.body.innerHTML = `<div id="toast-root"></div>`;
  });

  afterEach(() => {
    vi.clearAllTimers();
    vi.useRealTimers();
  });

  it("shows a read-only toast for permission-rejected cell edits", () => {
    showCollabEditRejectedToast([
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        rejectionKind: "cell",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("Read-only");
    expect(content).toContain("A1");
  });

  it("shows a missing encryption key toast for encryption-rejected cell edits", () => {
    showCollabEditRejectedToast([
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        rejectionKind: "cell",
        rejectionReason: "encryption",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("Missing encryption key");
    expect(content).toContain("A1");
  });

  it("shows a formatting toast for rejected format edits", () => {
    showCollabEditRejectedToast([
      {
        sheetId: "Sheet1",
        layer: "sheet",
        beforeStyleId: 1,
        afterStyleId: 2,
        rejectionKind: "format",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("formatting");
  });

  it("shows a drawing toast for rejected drawing edits", () => {
    showCollabEditRejectedToast([
      {
        rejectionKind: "drawing",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("edit drawings");
  });

  it("throttles repeated drawing toasts", () => {
    vi.setSystemTime(1_000);
    showCollabEditRejectedToast([{ rejectionKind: "drawing", rejectionReason: "permission" }]);
    showCollabEditRejectedToast([{ rejectionKind: "drawing", rejectionReason: "permission" }]);

    expect(document.querySelectorAll('[data-testid="toast"]')).toHaveLength(1);

    vi.setSystemTime(2_000);
    showCollabEditRejectedToast([{ rejectionKind: "drawing", rejectionReason: "permission" }]);
    expect(document.querySelectorAll('[data-testid="toast"]')).toHaveLength(2);
  });

  it("shows a chart toast for rejected chart edits", () => {
    showCollabEditRejectedToast([
      {
        rejectionKind: "chart",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("edit charts");
  });

  it("shows an undo/redo toast for rejected undo/redo actions", () => {
    showCollabEditRejectedToast([
      {
        rejectionKind: "undoRedo",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("undo/redo");
  });

  it("shows a row/col visibility toast for rejected hide/unhide actions", () => {
    showCollabEditRejectedToast([
      {
        rejectionKind: "rowColVisibility",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("hide");
  });
});
