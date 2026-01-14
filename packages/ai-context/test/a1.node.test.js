import assert from "node:assert/strict";
import test from "node:test";

import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS, parseA1Range, rangeToA1 } from "../src/a1.js";

test("A1 utilities: roundtrips unquoted identifier-like sheet names", () => {
  const input = "Sheet1!A1:B2";
  const parsed = parseA1Range(input);
  assert.deepStrictEqual(parsed, {
    sheetName: "Sheet1",
    startRow: 0,
    startCol: 0,
    endRow: 1,
    endCol: 1,
  });
  assert.equal(rangeToA1(parsed), input);
});

test("A1 utilities: roundtrips Excel-style quoted sheet names", () => {
  const input = "'My Sheet'!A1:B2";
  const parsed = parseA1Range(input);
  assert.equal(parsed.sheetName, "My Sheet");
  assert.equal(rangeToA1(parsed), input);
});

test("A1 utilities: roundtrips Excel-style quoted sheet names with embedded quotes", () => {
  const input = "'Bob''s Sheet'!A1";
  const parsed = parseA1Range(input);
  assert.equal(parsed.sheetName, "Bob's Sheet");
  assert.equal(rangeToA1(parsed), input);
});

test("A1 utilities: accepts legacy unquoted sheet names with spaces", () => {
  const parsed = parseA1Range("My Sheet!A1");
  assert.equal(parsed.sheetName, "My Sheet");
  assert.equal(rangeToA1(parsed), "'My Sheet'!A1");
});

test("A1 utilities: rangeToA1 accepts already-quoted sheetName input (back-compat)", () => {
  assert.equal(
    rangeToA1({
      sheetName: "'My Sheet'",
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
    }),
    "'My Sheet'!A1",
  );

  assert.equal(
    rangeToA1({
      sheetName: "'Bob''s Sheet'",
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
    }),
    "'Bob''s Sheet'!A1",
  );
});

test("A1 utilities: quotes reserved/ambiguous sheet names (TRUE, A1, R1C1, leading digits)", () => {
  assert.equal(
    rangeToA1({
      sheetName: "TRUE",
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
    }),
    "'TRUE'!A1",
  );

  assert.equal(
    rangeToA1({
      sheetName: "A1",
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
    }),
    "'A1'!A1",
  );

  assert.equal(
    rangeToA1({
      sheetName: "R1C1",
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
    }),
    "'R1C1'!A1",
  );

  assert.equal(
    rangeToA1({
      sheetName: "1Sheet",
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
    }),
    "'1Sheet'!A1",
  );
});

test("A1 utilities: parses absolute refs with $ (cell + range) and canonicalizes output", () => {
  assert.deepStrictEqual(parseA1Range("$A$1"), {
    sheetName: undefined,
    startRow: 0,
    startCol: 0,
    endRow: 0,
    endCol: 0,
  });
  assert.equal(rangeToA1(parseA1Range("$A$1")), "A1");

  assert.equal(rangeToA1(parseA1Range("Sheet1!$A$1:$B$2")), "Sheet1!A1:B2");
  assert.equal(rangeToA1(parseA1Range("'My Sheet'!$a$1:$b$2")), "'My Sheet'!A1:B2");
});

test("A1 utilities: accepts lower-case column letters", () => {
  assert.equal(rangeToA1(parseA1Range("a1:b2")), "A1:B2");
  assert.equal(rangeToA1(parseA1Range("sheet1!a1")), "sheet1!A1");
});

test("A1 utilities: parses whole-column ranges (A:C) using Excel max row limits", () => {
  const parsed = parseA1Range("A:C");
  assert.deepStrictEqual(parsed, {
    sheetName: undefined,
    startRow: 0,
    startCol: 0,
    endRow: EXCEL_MAX_ROWS - 1,
    endCol: 2,
  });
  assert.equal(rangeToA1(parsed), "A:C");
  assert.equal(rangeToA1(parseA1Range("Sheet1!c:a")), "Sheet1!A:C");
});

test("A1 utilities: parses whole-row ranges (1:10) using Excel max column limits", () => {
  const parsed = parseA1Range("1:10");
  assert.deepStrictEqual(parsed, {
    sheetName: undefined,
    startRow: 0,
    startCol: 0,
    endRow: 9,
    endCol: EXCEL_MAX_COLS - 1,
  });
  assert.equal(rangeToA1(parsed), "1:10");
  assert.equal(rangeToA1(parseA1Range("'My Sheet'!10:1")), "'My Sheet'!1:10");
});

test("A1 utilities: rejects invalid A1 references", () => {
  assert.throws(() => parseA1Range(""));
  assert.throws(() => parseA1Range("!A1"));

  // Invalid cell refs.
  assert.throws(() => parseA1Range("A0"));
  assert.throws(() => parseA1Range("XFE1")); // beyond XFD
  assert.throws(() => parseA1Range("A1048577")); // beyond max row

  // Invalid row/col ranges.
  assert.throws(() => parseA1Range("A"));
  assert.throws(() => parseA1Range("0:1"));
  assert.throws(() => parseA1Range("A:1"));
  assert.throws(() => parseA1Range("A1:B"));
  assert.throws(() => parseA1Range("A:"));
  assert.throws(() => parseA1Range(":A"));
});

