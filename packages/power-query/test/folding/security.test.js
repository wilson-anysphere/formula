import assert from "node:assert/strict";
import test from "node:test";

import { getSqlDialect } from "../../src/folding/dialect.js";
import { QueryFoldingEngine } from "../../src/folding/sql.js";

/**
 * Deterministic RNG for property-ish tests.
 * @param {number} seed
 */
function makeRng(seed) {
  let state = seed >>> 0;
  return () => {
    // LCG parameters from Numerical Recipes.
    state = (1664525 * state + 1013904223) >>> 0;
    return state / 0x100000000;
  };
}

/**
 * @param {() => number} rng
 * @param {number} maxLen
 */
function randomIdentifier(rng, maxLen) {
  const alphabet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 _-"`[];/\\\n\t';
  const len = Math.floor(rng() * maxLen);
  let out = "";
  for (let i = 0; i < len; i++) {
    out += alphabet[Math.floor(rng() * alphabet.length)];
  }
  return out;
}

/**
 * @param {string} inner
 * @param {string} quoteChar
 */
function assertNoUnescaped(inner, quoteChar) {
  for (let i = 0; i < inner.length; i++) {
    if (inner[i] !== quoteChar) continue;
    assert.equal(inner[i + 1], quoteChar, `found unescaped ${quoteChar} at index ${i}`);
    i += 1;
  }
}

test("security: dialect quoteIdentifier escapes embedded quotes/backticks (property test)", () => {
  const rng = makeRng(0xdecafbad);
  const dialects = ["postgres", "sqlite", "mysql", "sqlserver"];
  for (const name of dialects) {
    const dialect = getSqlDialect(/** @type {any} */ (name));
    const quoteChar = name === "mysql" ? "`" : name === "sqlserver" ? "]" : '"';
    const openChar = name === "sqlserver" ? "[" : quoteChar;
    for (let i = 0; i < 500; i++) {
      const ident = randomIdentifier(rng, 40);
      const quoted = dialect.quoteIdentifier(ident);
      assert.ok(quoted.startsWith(openChar) && quoted.endsWith(quoteChar), `bad quotes for ${name}: ${quoted}`);

      const inner = quoted.slice(1, -1);
      assertNoUnescaped(inner, quoteChar);

      const roundTrip = inner.replaceAll(quoteChar + quoteChar, quoteChar);
      assert.equal(roundTrip, ident);
    }
  }
});

test("security: predicate values are parameterized (no SQL injection via values)", () => {
  const folding = new QueryFoldingEngine();
  const payload = "x'); DROP TABLE sales; --";
  const query = {
    id: "q_injection",
    name: "Injection",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [{ id: "s1", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: payload } } }],
  };

  const plan = folding.compile(query);
  assert.equal(plan.type, "sql");
  assert.ok(plan.sql.includes("?"));
  assert.ok(!plan.sql.includes(payload));
  assert.deepEqual(plan.params, [payload]);
});

test("security: addColumn string literals are parameterized (no SQL injection via formula)", () => {
  const folding = new QueryFoldingEngine();
  const payload = "DROP TABLE users";
  const query = {
    id: "q_formula_injection",
    name: "Formula injection",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Region"] } },
      { id: "s2", name: "Add", operation: { type: "addColumn", name: "Injected", formula: `"${payload}"` } },
    ],
  };

  const plan = folding.compile(query);
  assert.equal(plan.type, "sql");
  assert.ok(plan.sql.includes("?"));
  assert.ok(!plan.sql.includes(payload));
  assert.deepEqual(plan.params, [payload]);
});

test("security: addColumn literals are always parameterized + preserve placeholder ordering (property test)", () => {
  const rng = makeRng(0x5151757);
  const folding = new QueryFoldingEngine();

  for (let i = 0; i < 250; i++) {
    const payload = `__payload__${i}:${randomIdentifier(rng, 60)}`;
    const query = {
      id: `q_formula_prop_${i}`,
      name: "Formula injection (property)",
      source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
      steps: [
        { id: "s1", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } } },
        { id: "s2", name: "Add", operation: { type: "addColumn", name: "Injected", formula: JSON.stringify(payload) } },
      ],
    };

    const plan = folding.compile(query);
    assert.equal(plan.type, "sql");
    assert.ok(plan.sql.includes("?"));
    assert.ok(!plan.sql.includes(payload));
    assert.deepEqual(plan.params, [payload, "East"]);
  }
});
