import assert from "node:assert/strict";
import test from "node:test";

import { postProcessGeneratedCode } from "../src/postprocess.js";

test("postProcessGeneratedCode ensures `def main` scripts call main() under __main__", async () => {
  const raw = [
    "import formula",
    "",
    "def main():",
    "    sheet = formula.active_sheet",
    "    sheet[\"A1\"] = 1",
    "",
  ].join("\n");

  const processed = await postProcessGeneratedCode({ code: raw, target: "python" });
  assert.match(processed, /if __name__ == \"__main__\":/);
  assert.match(processed, /\n\s*main\(\)\s*$/m);
});

