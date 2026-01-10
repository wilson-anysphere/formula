import assert from "node:assert/strict";
import test from "node:test";

import { VbaMigrator } from "../src/converter.js";

class MockChatClient {
  async chat() {
    return {
      message: {
        role: "assistant",
        content: [
          "sheet = formula.active_sheet",
          'sheet.Range("A1").Value = 1',
        ].join("\n"),
      },
    };
  }
}

test("VbaMigrator can use an LLMClient.chat()-style interface", async () => {
  const migrator = new VbaMigrator({ llm: new MockChatClient() });
  const module = { name: "Module1", code: "Sub Main()\nEnd Sub\n" };
  const result = await migrator.convertModule(module, { target: "python" });
  assert.match(result.code, /^import formula/m);
  assert.match(result.code, /sheet\["A1"\]\s*=\s*1/);
});

