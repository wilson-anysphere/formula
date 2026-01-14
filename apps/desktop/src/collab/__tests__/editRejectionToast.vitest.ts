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

  it("includes the encrypted payload key id in the toast when available", () => {
    showCollabEditRejectedToast([
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        rejectionKind: "cell",
        rejectionReason: "encryption",
        encryptionKeyId: "k-range-1",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("Missing encryption key");
    expect(content).toContain("k-range-1");
  });

  it("shows an actionable toast when the encrypted payload schema is unsupported", () => {
    showCollabEditRejectedToast([
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        rejectionKind: "cell",
        rejectionReason: "encryption",
        encryptionPayloadUnsupported: true,
        encryptionKeyId: "k-range-1",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("unsupported format");
    expect(content).toContain("Update Formula");
    expect(content).toContain("k-range-1");
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

  it("shows a formatting defaults toast for rejected formatting-default changes", () => {
    showCollabEditRejectedToast([
      {
        rejectionKind: "formatDefaults",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("formatting defaults");
  });

  it("shows an insert pictures toast for rejected picture insertion", () => {
    showCollabEditRejectedToast([
      {
        rejectionKind: "insertPictures",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("insert pictures");
  });

  it("shows a background image toast for rejected sheet background changes", () => {
    showCollabEditRejectedToast([
      {
        rejectionKind: "backgroundImage",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("sheet background");
  });

  it("shows a sort toast for rejected sort actions", () => {
    showCollabEditRejectedToast([
      {
        rejectionKind: "sort",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("sort");
  });

  it("shows a merge cells toast for rejected merge actions", () => {
    showCollabEditRejectedToast([
      {
        rejectionKind: "mergeCells",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("merge cells");
  });

  it("shows a fill cells toast for rejected fill actions", () => {
    showCollabEditRejectedToast([
      {
        rejectionKind: "fillCells",
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("fill cells");
  });

  it.each([
    ["cellContents", "edit cell contents"],
    ["editCells", "edit cells"],
    ["insertRows", "insert rows"],
    ["insertColumns", "insert columns"],
    ["deleteRows", "delete rows"],
    ["deleteColumns", "delete columns"],
    ["pageSetup", "page setup"],
    ["printAreaSet", "set a print area"],
    ["printAreaClear", "clear the print area"],
    ["printAreaEdit", "edit the print area"],
  ] as const)("shows a %s toast for rejected actions", (rejectionKind, expectedText) => {
    showCollabEditRejectedToast([
      {
        rejectionKind,
        rejectionReason: "permission",
      },
    ]);

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain(expectedText);
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
