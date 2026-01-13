import assert from "node:assert/strict";
import test from "node:test";

import { parsePartialFormula } from "../src/formulaPartialParser.js";

function createRng(seed) {
  /** @type {number} */
  let state = seed >>> 0;

  /** @returns {number} uint32 */
  function nextU32() {
    // xorshift32
    state ^= state << 13;
    state >>>= 0;
    state ^= state >>> 17;
    state >>>= 0;
    state ^= state << 5;
    state >>>= 0;
    return state;
  }

  /**
   * @param {number} maxExclusive
   * @returns {number}
   */
  function int(maxExclusive) {
    if (!Number.isFinite(maxExclusive) || maxExclusive <= 0) return 0;
    return nextU32() % maxExclusive;
  }

  /**
   * @template T
   * @param {T[]} arr
   * @returns {T}
   */
  function pick(arr) {
    return arr[int(arr.length)];
  }

  return { nextU32, int, pick };
}

const FUNCTION_NAMES = [
  "SUM",
  "IF",
  "VLOOKUP",
  "XLOOKUP",
  "_xlfn.XLOOKUP",
  "_xlfn.TAKE",
  "INDEX",
  "MATCH",
  // Include something that should *not* be treated as a function name token.
  "A1",
];

const CHAR_POOL = [
  ..."ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789",
  " ",
  "\t",
  "\n",
  "=",
  "+",
  "-",
  "*",
  "/",
  "^",
  "&",
  "%",
  "<",
  ">",
  ",",
  ";",
  ":",
  "(",
  ")",
  "[",
  "]",
  "{",
  "}",
  ".",
  "_",
  "$",
  "!",
  "?",
  "@",
  "#",
  "~",
  "`",
  "\\",
  '"',
  "'",
];

/**
 * @param {{ rng: ReturnType<typeof createRng>, maxLen: number }} params
 */
function randomInput({ rng, maxLen }) {
  // Ensure we exercise formula-specific code paths frequently enough to catch
  // any index-math bugs.
  const buildMode = rng.int(6);
  let s = "";

  // 0..2: structured-ish formula strings
  if (buildMode <= 2) {
    const fn = rng.pick(FUNCTION_NAMES);
    s = `=${fn}(`;
  } else if (buildMode === 3) {
    // A quoted sheet name pattern.
    s = `='My Sheet''Name'!`;
    if (rng.int(2) === 0) s += rng.pick(FUNCTION_NAMES);
  } else if (buildMode === 4) {
    // Non-formula input (no leading "=").
    s = "";
  } else {
    // Random, but often formula-like.
    s = rng.int(2) === 0 ? "=" : "";
  }

  const remaining = Math.max(0, maxLen - s.length);
  const extraLen = rng.int(remaining + 1);
  for (let i = 0; i < extraLen; i++) {
    s += rng.pick(CHAR_POOL);
  }
  return s;
}

test("parsePartialFormula never throws on arbitrary input (deterministic fuzz)", () => {
  // Keep runtime bounded; this runs on CI and locally.
  const N = 1000;
  const MAX_LEN = 120;
  const rng = createRng(0xdeadbeef);

  const functionRegistry = {
    // Minimal stub: mark only the first arg as a "range arg".
    isRangeArg(_fnName, argIndex) {
      return argIndex === 0;
    },
  };

  for (let i = 0; i < N; i++) {
    const input = randomInput({ rng, maxLen: MAX_LEN });
    const len = input.length;
    const positions = [
      0,
      Math.floor(len / 2),
      len,
      rng.int(len + 1),
    ];

    for (const cursorPosition of positions) {
      let result;
      try {
        result = parsePartialFormula(input, cursorPosition, functionRegistry);
      } catch (err) {
        assert.fail(
          `parsePartialFormula threw for input=${JSON.stringify(input)} cursor=${cursorPosition}\n` +
            `error: ${err?.stack ?? String(err)}`,
        );
      }

      assert.ok(result && typeof result === "object", "Expected an object result");
      assert.equal(typeof result.isFormula, "boolean", "Expected result.isFormula to be boolean");
      assert.equal(typeof result.inFunctionCall, "boolean", "Expected result.inFunctionCall to be boolean");
    }
  }
});

