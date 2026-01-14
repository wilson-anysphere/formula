import assert from "node:assert/strict";
import test from "node:test";

import { detectCellMoves } from "./moves.js";

test("detectCellMoves: does not match encrypted markers to plaintext cells", () => {
  const base = {
    A1: { enc: null, format: { bold: true } },
  };

  // Delete the encrypted marker at A1 and add a plaintext (format-only) cell elsewhere.
  const next = {
    B2: { format: { bold: true } },
  };

  const moves = detectCellMoves(base, next);
  assert.equal(moves.size, 0);
});

test("detectCellMoves: still detects moves for encrypted markers", () => {
  const base = {
    A1: { enc: null, format: { bold: true } },
  };

  const next = {
    B2: { enc: null, format: { bold: true } },
  };

  const moves = detectCellMoves(base, next);
  assert.equal(moves.size, 1);
  assert.equal(moves.get("A1"), "B2");
});

