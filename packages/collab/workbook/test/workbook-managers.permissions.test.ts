import * as Y from "yjs";
import { describe, expect, it } from "vitest";

import {
  createCollabSession,
  createMetadataManagerForSessionWithPermissions as createMetadataManagerForSessionWithPermissionsFromSession,
  createNamedRangeManagerForSessionWithPermissions as createNamedRangeManagerForSessionWithPermissionsFromSession,
  createSheetManagerForSessionWithPermissions as createSheetManagerForSessionWithPermissionsFromSession,
} from "@formula/collab-session";

import {
  createMetadataManagerForSessionWithPermissions as createMetadataManagerForSessionWithPermissionsFromWorkbook,
  createNamedRangeManagerForSessionWithPermissions as createNamedRangeManagerForSessionWithPermissionsFromWorkbook,
  createSheetManagerForSessionWithPermissions as createSheetManagerForSessionWithPermissionsFromWorkbook,
} from "../src/index.ts";

const SOURCES = [
  {
    label: "@formula/collab-workbook",
    createSheetManagerForSessionWithPermissions: createSheetManagerForSessionWithPermissionsFromWorkbook,
    createMetadataManagerForSessionWithPermissions: createMetadataManagerForSessionWithPermissionsFromWorkbook,
    createNamedRangeManagerForSessionWithPermissions: createNamedRangeManagerForSessionWithPermissionsFromWorkbook,
  },
  {
    label: "@formula/collab-session",
    createSheetManagerForSessionWithPermissions: createSheetManagerForSessionWithPermissionsFromSession,
    createMetadataManagerForSessionWithPermissions: createMetadataManagerForSessionWithPermissionsFromSession,
    createNamedRangeManagerForSessionWithPermissions: createNamedRangeManagerForSessionWithPermissionsFromSession,
  },
] as const;

describe.each(SOURCES)("$label permission-aware workbook managers", (source) => {
  it("allows editors to mutate workbook sheets/metadata/namedRanges", () => {
    const doc = new Y.Doc();
    const session = createCollabSession({ doc });
    session.setPermissions({ role: "editor", userId: "u-editor", rangeRestrictions: [] });

    const sheetMgr = source.createSheetManagerForSessionWithPermissions(session);
    sheetMgr.addSheet({ id: "s2", name: "Second" });
    expect(sheetMgr.list().map((s) => s.id)).toContain("s2");
    expect(session.sheets.toArray().some((s: any) => s?.get?.("id") === "s2")).toBe(true);

    const metadataMgr = source.createMetadataManagerForSessionWithPermissions(session);
    metadataMgr.set("foo", "bar");
    expect(metadataMgr.get("foo")).toBe("bar");
    expect(session.metadata.get("foo")).toBe("bar");

    const namedRangeMgr = source.createNamedRangeManagerForSessionWithPermissions(session);
    namedRangeMgr.set("MyRange", "A1:B2");
    expect(namedRangeMgr.get("MyRange")).toBe("A1:B2");
    expect(session.namedRanges.get("MyRange")).toBe("A1:B2");
  });

  it.each(["viewer", "commenter"] as const)("prevents %s from mutating workbook sheets", (role) => {
    const doc = new Y.Doc();
    const session = createCollabSession({ doc });
    session.setPermissions({ role, userId: "u-readonly", rangeRestrictions: [] });

    const sheetMgr = source.createSheetManagerForSessionWithPermissions(session);
    const beforeIds = sheetMgr.list().map((s) => s.id);
    const beforeSheet1Name = sheetMgr.list().find((s) => s.id === "Sheet1")?.name ?? null;

    expect(() => sheetMgr.addSheet({ id: "s2", name: "Second" })).toThrow(/read-?only/i);
    // Direct Yjs mutations should also be blocked.
    expect(() => (sheetMgr.sheets as any).push([new Y.Map()])).toThrow(/read-?only/i);

    const sheet1 = sheetMgr.getById("Sheet1");
    expect(sheet1).toBeTruthy();
    expect(() => (sheet1 as any).set("name", "Hacked")).toThrow(/read-?only/i);

    // No mutation should have occurred.
    expect(sheetMgr.list().map((s) => s.id)).toEqual(beforeIds);
    expect(sheetMgr.list().find((s) => s.id === "Sheet1")?.name ?? null).toBe(beforeSheet1Name);
    expect(session.sheets.toArray().some((s: any) => s?.get?.("id") === "s2")).toBe(false);
  });

  it.each(["viewer", "commenter"] as const)("prevents %s from mutating workbook metadata", (role) => {
    const doc = new Y.Doc();
    const session = createCollabSession({ doc });
    session.setPermissions({ role, userId: "u-readonly", rangeRestrictions: [] });

    const metadataMgr = source.createMetadataManagerForSessionWithPermissions(session);
    expect(() => metadataMgr.set("foo", "bar")).toThrow(/read-?only/i);
    expect(() => (metadataMgr.metadata as any).set("foo", "bar")).toThrow(/read-?only/i);
    expect(metadataMgr.get("foo")).toBeUndefined();
    expect(session.metadata.get("foo")).toBeUndefined();
  });

  it.each(["viewer", "commenter"] as const)("prevents %s from mutating workbook named ranges", (role) => {
    const doc = new Y.Doc();
    const session = createCollabSession({ doc });
    session.setPermissions({ role, userId: "u-readonly", rangeRestrictions: [] });

    const namedRangeMgr = source.createNamedRangeManagerForSessionWithPermissions(session);
    expect(() => namedRangeMgr.set("MyRange", "A1:B2")).toThrow(/read-?only/i);
    expect(() => (namedRangeMgr.namedRanges as any).set("MyRange", "A1:B2")).toThrow(/read-?only/i);
    expect(namedRangeMgr.get("MyRange")).toBeUndefined();
    expect(session.namedRanges.get("MyRange")).toBeUndefined();
  });
});
