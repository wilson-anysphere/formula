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
    // Seed a nested Yjs type so we can ensure read-only callers can't mutate
    // workbook sheet metadata indirectly (via nested maps/arrays).
    const seededSheet1 = session.sheets.get(0) as any;
    expect(seededSheet1).toBeTruthy();
    doc.transact(() => {
      seededSheet1.set("nested", new Y.Map());
    });

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

    const nested = (sheet1 as any).get("nested");
    expect(nested).toBeTruthy();
    expect(() => (nested as any).set("x", 1)).toThrow(/read-?only/i);

    // No mutation should have occurred.
    expect(sheetMgr.list().map((s) => s.id)).toEqual(beforeIds);
    expect(sheetMgr.list().find((s) => s.id === "Sheet1")?.name ?? null).toBe(beforeSheet1Name);
    expect(session.sheets.toArray().some((s: any) => s?.get?.("id") === "s2")).toBe(false);
  });

  it.each(["viewer", "commenter"] as const)("prevents %s from mutating workbook metadata", (role) => {
    const doc = new Y.Doc();
    // Seed a nested Yjs type in metadata (common for workbook-level UI state).
    const nestedArr = new Y.Array<unknown>();
    nestedArr.push([new Y.Map()]);
    doc.transact(() => {
      doc.getMap("metadata").set("nested", nestedArr);
    });
    const session = createCollabSession({ doc });
    session.setPermissions({ role, userId: "u-readonly", rangeRestrictions: [] });

    const metadataMgr = source.createMetadataManagerForSessionWithPermissions(session);
    expect(() => metadataMgr.set("foo", "bar")).toThrow(/read-?only/i);
    expect(() => (metadataMgr.metadata as any).set("foo", "bar")).toThrow(/read-?only/i);
    const nested = (metadataMgr.metadata as any).get("nested");
    expect(nested).toBeTruthy();
    expect(() => (nested as any).push([new Y.Map()])).toThrow(/read-?only/i);
    // Ensure slice/map read helpers also return guarded nested values.
    const sliced = (nested as any).slice();
    expect(Array.isArray(sliced)).toBe(true);
    expect(() => sliced[0].set("x", 1)).toThrow(/read-?only/i);
    expect(() => (nested as any).map((v: any) => v.set("x", 1))).toThrow(/read-?only/i);
    expect(metadataMgr.get("foo")).toBeUndefined();
    expect(session.metadata.get("foo")).toBeUndefined();
    expect((doc.getMap("metadata").get("nested") as any)?.length ?? 0).toBe(1);
  });

  it.each(["viewer", "commenter"] as const)("prevents %s from mutating workbook named ranges", (role) => {
    const doc = new Y.Doc();
    // Seed a nested Yjs type as a named range value to ensure nested types are guarded.
    doc.transact(() => {
      doc.getMap("namedRanges").set("Nested", new Y.Map());
    });
    const session = createCollabSession({ doc });
    session.setPermissions({ role, userId: "u-readonly", rangeRestrictions: [] });

    const namedRangeMgr = source.createNamedRangeManagerForSessionWithPermissions(session);
    expect(() => namedRangeMgr.set("MyRange", "A1:B2")).toThrow(/read-?only/i);
    expect(() => (namedRangeMgr.namedRanges as any).set("MyRange", "A1:B2")).toThrow(/read-?only/i);
    const nested = (namedRangeMgr.namedRanges as any).get("Nested");
    expect(nested).toBeTruthy();
    expect(() => (nested as any).set("x", 1)).toThrow(/read-?only/i);
    expect(namedRangeMgr.get("MyRange")).toBeUndefined();
    expect(session.namedRanges.get("MyRange")).toBeUndefined();
  });
});
