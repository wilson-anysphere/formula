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
});

