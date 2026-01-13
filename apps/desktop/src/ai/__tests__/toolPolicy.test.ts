import { describe, expect, it } from "vitest";

import { InMemoryWorkbook, SpreadsheetLLMToolExecutor } from "../../../../../packages/ai-tools/src/index.js";

import { getDesktopToolPolicy } from "../toolPolicy.js";

describe("getDesktopToolPolicy", () => {
  it("agent mode denies fetch_external_data even when host enables it", () => {
    const spreadsheet = new InMemoryWorkbook();

    const hostEnabled = new SpreadsheetLLMToolExecutor(spreadsheet, {
      allow_external_data: true,
      allowed_external_hosts: ["example.com"]
    });
    expect(hostEnabled.tools.map((tool) => tool.name)).toContain("fetch_external_data");

    const agentPolicy = new SpreadsheetLLMToolExecutor(spreadsheet, {
      allow_external_data: true,
      allowed_external_hosts: ["example.com"],
      toolPolicy: getDesktopToolPolicy({ mode: "agent" })
    });
    expect(agentPolicy.tools.map((tool) => tool.name)).not.toContain("fetch_external_data");
  });
});

