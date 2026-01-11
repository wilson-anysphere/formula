import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";

import { FORMULA_API_DTS } from "@formula/scripting";

test("FORMULA_API_DTS stays in sync with packages/scripting/formula.d.ts", async () => {
  const fileUrl = new URL("../packages/scripting/formula.d.ts", import.meta.url);
  const source = await readFile(fileUrl, "utf8");

  assert.equal(FORMULA_API_DTS.trimEnd(), source.trimEnd());
});

