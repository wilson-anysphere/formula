import assert from "node:assert/strict";
import test from "node:test";

import { normalizePostgresPlaceholders } from "../../src/folding/placeholders.js";

test("placeholders: converts ? placeholders to $n for Postgres", () => {
  const sql = "SELECT * FROM t WHERE a = ? AND b >= ?";
  assert.equal(normalizePostgresPlaceholders(sql, 2), "SELECT * FROM t WHERE a = $1 AND b >= $2");
});

test("placeholders: does not rewrite Postgres jsonb ? operator", () => {
  const sql = "SELECT * FROM t WHERE data ? 'key' AND a = ?";
  assert.equal(normalizePostgresPlaceholders(sql, 1), "SELECT * FROM t WHERE data ? 'key' AND a = $1");
});

test("placeholders: ignores question marks inside string literals", () => {
  const sql = "SELECT '?' AS q WHERE a = ?";
  assert.equal(normalizePostgresPlaceholders(sql, 1), "SELECT '?' AS q WHERE a = $1");
});

test("placeholders: converts placeholders following LIKE/LIMIT/CASE keywords", () => {
  const sql = "SELECT * FROM t WHERE name LIKE ? LIMIT ? OFFSET ? AND (CASE WHEN ok THEN ? ELSE ? END) IS NOT NULL";
  assert.equal(
    normalizePostgresPlaceholders(sql, 5),
    "SELECT * FROM t WHERE name LIKE $1 LIMIT $2 OFFSET $3 AND (CASE WHEN ok THEN $4 ELSE $5 END) IS NOT NULL",
  );
});

test("placeholders: rewrites placeholders inside CAST()", () => {
  const sql = 'SELECT CAST(? AS DOUBLE PRECISION) AS n, CAST(? AS TEXT) AS s';
  assert.equal(normalizePostgresPlaceholders(sql, 2), 'SELECT CAST($1 AS DOUBLE PRECISION) AS n, CAST($2 AS TEXT) AS s');
});

test("placeholders: ignores question marks inside comments", () => {
  const sql = "SELECT * FROM t -- ? comment\nWHERE a = ? /* ? block */";
  assert.equal(normalizePostgresPlaceholders(sql, 1), "SELECT * FROM t -- ? comment\nWHERE a = $1 /* ? block */");
});
