import { describe, expect, it } from "vitest";

import { computeDlpCacheKey } from "../dlpCacheKey.js";

describe("computeDlpCacheKey", () => {
  it("is stable across classification record ordering", () => {
    const base = {
      documentId: "doc",
      policy: { rules: { a: 1 } },
      includeRestrictedContent: false,
    };

    const r1 = { selector: { scope: "cell", sheetId: "Sheet1", row: 0, col: 0 }, classification: { level: "Public" } };
    const r2 = { selector: { scope: "cell", sheetId: "Sheet1", row: 1, col: 0 }, classification: { level: "Restricted" } };

    const key1 = computeDlpCacheKey({ ...base, classificationRecords: [r1, r2] });
    const key2 = computeDlpCacheKey({ ...base, classificationRecords: [r2, r1] });

    expect(key1).toEqual(key2);
  });

  it("returns a precomputed cacheKey when present", () => {
    const dlp = { cacheKey: "dlp:precomputed" };
    expect(computeDlpCacheKey(dlp)).toBe("dlp:precomputed");
  });

  it("does not reuse cacheKey when only a classification store is provided (store may change)", () => {
    const base = {
      document_id: "doc",
      policy: { rules: { a: 1 } },
      include_restricted_content: false,
      // If this value were trusted unconditionally, it could cause stale cache reuse when the
      // underlying store contents change.
      cacheKey: "dlp:stale",
    };

    let records: any[] = [
      { selector: { scope: "cell", sheetId: "Sheet1", row: 0, col: 0 }, classification: { level: "Public" } },
    ];
    const classification_store = {
      list: () => records,
    };

    const dlp = { ...base, classification_store };
    const key1 = computeDlpCacheKey(dlp);
    expect(key1).not.toBe("dlp:stale");

    records = [
      { selector: { scope: "cell", sheetId: "Sheet1", row: 0, col: 0 }, classification: { level: "Restricted" } },
    ];
    const key2 = computeDlpCacheKey(dlp);
    expect(key1).not.toEqual(key2);
  });

  it("changes when includeRestrictedContent changes", () => {
    const base = {
      documentId: "doc",
      policy: { rules: { a: 1 } },
      classificationRecords: [],
    };

    const keyFalse = computeDlpCacheKey({ ...base, includeRestrictedContent: false });
    const keyTrue = computeDlpCacheKey({ ...base, includeRestrictedContent: true });
    expect(keyFalse).not.toEqual(keyTrue);
  });

  it("changes when classification changes", () => {
    const base = {
      documentId: "doc",
      policy: { rules: { a: 1 } },
      includeRestrictedContent: false,
    };

    const selector = { scope: "cell", sheetId: "Sheet1", row: 0, col: 0 };

    const keyPublic = computeDlpCacheKey({
      ...base,
      classificationRecords: [{ selector, classification: { level: "Public" } }],
    });
    const keyRestricted = computeDlpCacheKey({
      ...base,
      classificationRecords: [{ selector, classification: { level: "Restricted" } }],
    });

    expect(keyPublic).not.toEqual(keyRestricted);
  });

  it("changes when policy changes", () => {
    const base = {
      documentId: "doc",
      classificationRecords: [],
      includeRestrictedContent: false,
    };

    const key1 = computeDlpCacheKey({ ...base, policy: { rules: { maxAllowed: "Confidential" } } });
    const key2 = computeDlpCacheKey({ ...base, policy: { rules: { maxAllowed: "Restricted" } } });

    expect(key1).not.toEqual(key2);
  });
});
