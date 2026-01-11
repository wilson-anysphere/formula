import { describe, expect, it } from "vitest";

import { decideAllowedTools } from "../src/llm/toolPolicy.js";

describe("tool policy", () => {
  it("chat analysis questions only expose read+compute tools", () => {
    const policy = decideAllowedTools({
      mode: "chat",
      user_text: "What's the average of column B?",
      has_attachments: false,
      allow_external_data: false
    });

    expect(policy.allowed_tools).toEqual(["read_range", "filter_range", "detect_anomalies", "compute_statistics"]);
  });

  it("chat mutation requests expose write tools (but not network by default)", () => {
    const policy = decideAllowedTools({
      mode: "chat",
      user_text: "Set A1 to 99",
      has_attachments: false,
      allow_external_data: false
    });

    expect(policy.allowed_tools).toEqual(["read_range", "write_cell", "set_range", "filter_range"]);
    expect(policy.allowed_tools).not.toContain("fetch_external_data");
  });

  it("fetch_external_data is only exposed when explicitly requested and allowed", () => {
    const requestedButDisabled = decideAllowedTools({
      mode: "chat",
      user_text: "Fetch data from https://example.com/data.json and put it in Sheet1!A1",
      has_attachments: false,
      allow_external_data: false
    });
    expect(requestedButDisabled.allowed_tools).not.toContain("fetch_external_data");

    const requestedAndEnabled = decideAllowedTools({
      mode: "chat",
      user_text: "Fetch data from https://example.com/data.json and put it in Sheet1!A1",
      has_attachments: false,
      allow_external_data: true
    });
    expect(requestedAndEnabled.allowed_tools).toContain("fetch_external_data");
  });

  it("inline_edit never exposes network tools and only exposes formatting when requested", () => {
    const format = decideAllowedTools({
      mode: "inline_edit",
      user_text: "Make the selected cells bold",
      has_attachments: true,
      allow_external_data: true
    });
    expect(format.allowed_tools).toEqual(["read_range", "write_cell", "set_range", "apply_formatting"]);
    expect(format.allowed_tools).not.toContain("fetch_external_data");

    const noFormat = decideAllowedTools({
      mode: "inline_edit",
      user_text: "Uppercase all values",
      has_attachments: true,
      allow_external_data: true
    });
    expect(noFormat.allowed_tools).toEqual(["read_range", "write_cell", "set_range"]);
    expect(noFormat.allowed_tools).not.toContain("apply_formatting");
  });

  it("cell_function never exposes mutation or network tools", () => {
    const policy = decideAllowedTools({
      mode: "cell_function",
      user_text: "Fetch data from https://example.com and put it in A1",
      has_attachments: false,
      allow_external_data: true
    });

    expect(policy.allowed_tools).toEqual(["read_range", "filter_range", "detect_anomalies", "compute_statistics"]);
    expect(policy.allowed_tools).not.toContain("fetch_external_data");
    expect(policy.allowed_tools).not.toContain("write_cell");
    expect(policy.allowed_tools).not.toContain("set_range");
  });
});

