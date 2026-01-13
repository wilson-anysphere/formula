import { describe, expect, it } from "vitest";

import { tryInsertCollabSheet } from "../collabSheetMutations";

type PlainSheetEntry = { id: string; name: string; visibility: string };

function makePlainSession(initial: PlainSheetEntry[], role: "editor" | "viewer" | "commenter" = "editor") {
  const data = initial.map((s) => ({ ...s }));
  const sheets = {
    get length() {
      return data.length;
    },
    get(idx: number) {
      return data[idx];
    },
    insert(idx: number, entries: any[]) {
      data.splice(idx, 0, ...entries);
    },
    toArray() {
      return data.slice();
    },
  };

  return {
    sheets,
    transactLocal: (fn: () => void) => fn(),
    isReadOnly: () => role !== "editor",
    getRole: () => role,
  };
}

describe("tryInsertCollabSheet (lightweight session.sheets)", () => {
  it.each(["viewer", "commenter"] as const)("blocks inserts when role is %s", (role) => {
    const session = makePlainSession(
      [
        { id: "a", name: "A", visibility: "visible" },
        { id: "b", name: "B", visibility: "visible" },
      ],
      role,
    );

    const before = session.sheets.toArray();
    const result = tryInsertCollabSheet({
      session: session as any,
      sheetId: "new",
      name: "New Sheet",
      visibility: "visible",
      insertAfterSheetId: "a",
    });

    expect(result.inserted).toBe(false);
    expect(session.sheets.toArray()).toEqual(before);
  });

  it("inserts plain entries when role is editor", () => {
    const session = makePlainSession([
      { id: "a", name: "A", visibility: "visible" },
      { id: "b", name: "B", visibility: "visible" },
    ]);

    const result = tryInsertCollabSheet({
      session: session as any,
      sheetId: "new",
      name: "New Sheet",
      visibility: "visible",
      insertAfterSheetId: "a",
    });

    expect(result).toEqual({ inserted: true, index: 1 });
    expect(session.sheets.toArray()).toEqual([
      { id: "a", name: "A", visibility: "visible" },
      { id: "new", name: "New Sheet", visibility: "visible" },
      { id: "b", name: "B", visibility: "visible" },
    ]);
  });
});

