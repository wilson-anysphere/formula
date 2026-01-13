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

