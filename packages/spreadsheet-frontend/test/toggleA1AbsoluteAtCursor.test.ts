import { describe, expect, it } from "vitest";

import { toggleA1AbsoluteAtCursor } from "../src/toggleA1AbsoluteAtCursor";

describe("toggleA1AbsoluteAtCursor", () => {
  it("cycles a single-cell ref through absolute modes at the caret", () => {
    const step1 = toggleA1AbsoluteAtCursor("=A1", 2, 2);
    expect(step1).not.toBeNull();
    expect(step1?.text).toBe("=$A$1");
    expect(step1?.cursorStart).toBe(4);
    expect(step1?.cursorEnd).toBe(4);

    const step2 = toggleA1AbsoluteAtCursor(step1!.text, step1!.cursorStart, step1!.cursorEnd);
    expect(step2?.text).toBe("=A$1");
    expect(step2?.cursorStart).toBe(3);
    expect(step2?.cursorEnd).toBe(3);

    const step3 = toggleA1AbsoluteAtCursor(step2!.text, step2!.cursorStart, step2!.cursorEnd);
    expect(step3?.text).toBe("=$A1");
    expect(step3?.cursorStart).toBe(3);
    expect(step3?.cursorEnd).toBe(3);

    const step4 = toggleA1AbsoluteAtCursor(step3!.text, step3!.cursorStart, step3!.cursorEnd);
    expect(step4?.text).toBe("=A1");
    expect(step4?.cursorStart).toBe(2);
    expect(step4?.cursorEnd).toBe(2);
  });

  it("cycles a range ref and applies the mode to both endpoints", () => {
    const input = "=SUM(A1:B2)";
    // Caret inside the range (before "B").
    const step1 = toggleA1AbsoluteAtCursor(input, 8, 8);
    expect(step1?.text).toBe("=SUM($A$1:$B$2)");
    // Cursor should still be inside the reference token (before "B").
    expect(step1?.cursorStart).toBe(11);
    expect(step1?.cursorEnd).toBe(11);
  });

  it("preserves sheet qualifiers (unquoted + quoted)", () => {
    const unquoted = toggleA1AbsoluteAtCursor("=Sheet1!A1", 9, 9);
    expect(unquoted?.text).toBe("=Sheet1!$A$1");

    const quoted = toggleA1AbsoluteAtCursor("=SUM('My Sheet'!A1:B2)", 17, 17);
    expect(quoted?.text).toBe("=SUM('My Sheet'!$A$1:$B$2)");
  });

  it("keeps full-token selections selecting the full toggled token", () => {
    const res = toggleA1AbsoluteAtCursor("=A1", 1, 3);
    expect(res?.text).toBe("=$A$1");
    expect(res?.cursorStart).toBe(1);
    expect(res?.cursorEnd).toBe(5);
  });
});

