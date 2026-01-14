import { describe, expect, it } from "vitest";

import { getWorkbookFileMetadataFromWorkbookInfo, splitWorkbookPath } from "../workbookFileMetadata";

describe("workbookFileMetadata", () => {
  it("returns null for empty paths", () => {
    expect(splitWorkbookPath("")).toBeNull();
    expect(splitWorkbookPath("   ")).toBeNull();
  });

  it("parses POSIX paths into directory + filename", () => {
    expect(splitWorkbookPath("/tmp/book.xlsx")).toEqual({ directory: "/tmp/", filename: "book.xlsx" });
  });

  it("parses Windows paths into directory + filename", () => {
    expect(splitWorkbookPath("C:\\Users\\me\\Book1.xlsx")).toEqual({
      directory: "C:\\Users\\me\\",
      filename: "Book1.xlsx",
    });
  });

  it("treats filename-only paths as directory=\"\" + filename", () => {
    expect(splitWorkbookPath("Book1.xlsx")).toEqual({ directory: "", filename: "Book1.xlsx" });
  });

  it("prefers WorkbookInfo.path over origin_path", () => {
    expect(
      getWorkbookFileMetadataFromWorkbookInfo({
        path: "/a/b/primary.xlsx",
        origin_path: "/x/y/origin.xlsx",
      }),
    ).toEqual({ directory: "/a/b/", filename: "primary.xlsx" });
  });

  it("falls back to origin_path when path is missing", () => {
    expect(getWorkbookFileMetadataFromWorkbookInfo({ path: null, origin_path: "origin.xlsx" })).toEqual({
      directory: "",
      filename: "origin.xlsx",
    });
  });

  it("returns null metadata when workbook info has no usable path", () => {
    expect(getWorkbookFileMetadataFromWorkbookInfo(null)).toEqual({ directory: null, filename: null });
    expect(getWorkbookFileMetadataFromWorkbookInfo({ path: null, origin_path: null })).toEqual({ directory: null, filename: null });
    expect(getWorkbookFileMetadataFromWorkbookInfo({ path: "   ", origin_path: "" })).toEqual({ directory: null, filename: null });
  });
});
