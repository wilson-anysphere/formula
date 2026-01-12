import { describe, expect, it } from "vitest";

import { reconcileClipboardCopyContextForPaste } from "../clipboardPasteContext";

describe("clipboardPasteContext", () => {
  it("clears stale internal clipboard context when the clipboard no longer matches", () => {
    const initialContext = {
      range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
      payload: { text: "A\tB", html: "<table><tr><td>A</td><td>B</td></tr></table>" },
      cells: [[{ value: "A", formula: null, styleId: 0 }]],
    };

    const { isInternalPaste, nextContext } = reconcileClipboardCopyContextForPaste(initialContext, { text: "X\tY" });

    expect(isInternalPaste).toBe(false);
    expect(nextContext).toBeNull();
  });

  it("treats rtf-only clipboard reads as internal when rtf matches the last internal copy payload", () => {
    const initialContext = {
      range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
      payload: { text: "A", html: "<table><tr><td>A</td></tr></table>", rtf: "{\\rtf1\\ansi A}" },
      cells: [[{ value: "A", formula: null, styleId: 0 }]],
    };

    const { isInternalPaste, nextContext } = reconcileClipboardCopyContextForPaste(initialContext, { rtf: "{\\rtf1\\ansi A}" });

    expect(isInternalPaste).toBe(true);
    expect(nextContext).toEqual(initialContext);
  });

  it("treats rtf-only clipboard reads as internal when extracted text matches the last internal plain-text payload", () => {
    const initialContext = {
      range: { startRow: 0, endRow: 0, startCol: 0, endCol: 1 },
      payload: {
        text: "A\tB",
        html: "<table><tr><td>A</td><td>B</td></tr></table>",
        rtf: "{\\rtf1\\ansi A\\cell B\\row}",
      },
      cells: [[{ value: "A", formula: null, styleId: 0 }, { value: "B", formula: null, styleId: 0 }]],
    };

    // Different RTF payload (uses \tab/\par), but extracts to the same TSV text.
    const content = { rtf: "{\\rtf1\\ansi\\deff0\\uc1\\pard A\\tab B\\par}" };
    const { isInternalPaste, nextContext } = reconcileClipboardCopyContextForPaste(initialContext, content);

    expect(isInternalPaste).toBe(true);
    expect(nextContext).toEqual(initialContext);
  });

  it("treats image-only clipboard reads (pngBase64) as usable and clears stale context", () => {
    const initialContext = {
      range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
      payload: { text: "A" },
      cells: [[{ value: "A", formula: null, styleId: 0 }]],
    };

    const { isInternalPaste, nextContext } = reconcileClipboardCopyContextForPaste(initialContext, {
      pngBase64: "data:image/png;base64,AAAA",
    });

    expect(isInternalPaste).toBe(false);
    expect(nextContext).toBeNull();
  });

  it("does not clear internal clipboard context when clipboard read yields no usable content", () => {
    const initialContext = {
      range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
      payload: { text: "A\tB" },
      cells: [[{ value: "A", formula: null, styleId: 0 }]],
    };

    const { isInternalPaste, nextContext } = reconcileClipboardCopyContextForPaste(initialContext, {});

    expect(isInternalPaste).toBe(false);
    expect(nextContext).toEqual(initialContext);
  });

  it("treats trailing newlines in clipboard text as internal when the content matches", () => {
    const initialContext = {
      range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
      payload: { text: "A\tB" },
      cells: [[{ value: "A", formula: null, styleId: 0 }]],
    };

    const { isInternalPaste, nextContext } = reconcileClipboardCopyContextForPaste(initialContext, { text: "A\tB\n" });

    expect(isInternalPaste).toBe(true);
    expect(nextContext).toEqual(initialContext);
  });
});
