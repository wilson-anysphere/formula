import * as Y from "yjs";
import { describe, expect, it } from "vitest";

import { createCollabSession } from "@formula/collab-session";

import {
  createMetadataManagerForSessionWithPermissions,
  createNamedRangeManagerForSessionWithPermissions,
  createSheetManagerForSessionWithPermissions,
} from "../src/index.ts";

describe("@formula/collab-workbook permission-aware managers", () => {
  it("allows editors to mutate workbook sheets/metadata/namedRanges", () => {
    const doc = new Y.Doc();
    const session = createCollabSession({ doc });
    session.setPermissions({ role: "editor", userId: "u-editor", rangeRestrictions: [] });

    const sheetMgr = createSheetManagerForSessionWithPermissions(session);
    sheetMgr.addSheet({ id: "s2", name: "Second" });
    expect(sheetMgr.list().map((s) => s.id)).toContain("s2");
    expect(session.sheets.toArray().some((s: any) => s?.get?.("id") === "s2")).toBe(true);

    const metadataMgr = createMetadataManagerForSessionWithPermissions(session);
    metadataMgr.set("foo", "bar");
    expect(metadataMgr.get("foo")).toBe("bar");
    expect(session.metadata.get("foo")).toBe("bar");

    const namedRangeMgr = createNamedRangeManagerForSessionWithPermissions(session);
    namedRangeMgr.set("MyRange", "A1:B2");
    expect(namedRangeMgr.get("MyRange")).toBe("A1:B2");
    expect(session.namedRanges.get("MyRange")).toBe("A1:B2");
  });

  it.each(["viewer", "commenter"] as const)("prevents %s from mutating workbook sheets", (role) => {
    const doc = new Y.Doc();
    const session = createCollabSession({ doc });
    session.setPermissions({ role, userId: "u-readonly", rangeRestrictions: [] });

    const sheetMgr = createSheetManagerForSessionWithPermissions(session);
    const beforeIds = sheetMgr.list().map((s) => s.id);

    expect(() => sheetMgr.addSheet({ id: "s2", name: "Second" })).toThrow(/read-?only/i);

    // No mutation should have occurred.
    expect(sheetMgr.list().map((s) => s.id)).toEqual(beforeIds);
    expect(session.sheets.toArray().some((s: any) => s?.get?.("id") === "s2")).toBe(false);
  });

  it.each(["viewer", "commenter"] as const)("prevents %s from mutating workbook metadata", (role) => {
    const doc = new Y.Doc();
    const session = createCollabSession({ doc });
    session.setPermissions({ role, userId: "u-readonly", rangeRestrictions: [] });

    const metadataMgr = createMetadataManagerForSessionWithPermissions(session);
    expect(() => metadataMgr.set("foo", "bar")).toThrow(/read-?only/i);
    expect(metadataMgr.get("foo")).toBeUndefined();
    expect(session.metadata.get("foo")).toBeUndefined();
  });

  it.each(["viewer", "commenter"] as const)("prevents %s from mutating workbook named ranges", (role) => {
    const doc = new Y.Doc();
    const session = createCollabSession({ doc });
    session.setPermissions({ role, userId: "u-readonly", rangeRestrictions: [] });

    const namedRangeMgr = createNamedRangeManagerForSessionWithPermissions(session);
    expect(() => namedRangeMgr.set("MyRange", "A1:B2")).toThrow(/read-?only/i);
    expect(namedRangeMgr.get("MyRange")).toBeUndefined();
    expect(session.namedRanges.get("MyRange")).toBeUndefined();
  });
});

