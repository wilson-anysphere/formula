import type { ToolPolicy, ToolName } from "../../../../packages/ai-tools/src/tool-schema.js";

export type DesktopAiMode = "chat" | "inline_edit" | "agent";

const CHAT_READ_ONLY_POLICY: ToolPolicy = {
  allowCategories: ["read", "compute"],
  mutationsAllowed: false,
  externalNetworkAllowed: false
};

const CHAT_MUTATION_POLICY: ToolPolicy = {
  allowCategories: ["read", "compute", "mutate", "format"],
  externalNetworkAllowed: false
};

const INLINE_EDIT_ALLOWED_TOOLS: ToolName[] = [
  "read_range",
  "write_cell",
  "set_range",
  "apply_formula_column",
  "sort_range",
  "filter_range",
  "apply_formatting",
  "detect_anomalies",
  "compute_statistics"
];

const INLINE_EDIT_POLICY: ToolPolicy = {
  allowTools: INLINE_EDIT_ALLOWED_TOOLS,
  externalNetworkAllowed: false
};

const AGENT_POLICY: ToolPolicy = {};

const CHAT_MUTATION_INTENT_RE = new RegExp(
  String.raw`\b(` +
    [
      "replace",
      "fill",
      "write",
      "format",
      "insert",
      "update",
      "set",
      "edit",
      "change",
      "delete",
      "remove",
      "clear",
      "sort",
      "create",
      "add",
      "apply",
      "chart",
      "plot",
      "graph",
      "pivot"
    ].join("|") +
    String.raw`)\b`,
  "i"
);

export function getDesktopToolPolicy(params: { mode: DesktopAiMode; prompt?: string }): ToolPolicy {
  switch (params.mode) {
    case "agent":
      return AGENT_POLICY;
    case "inline_edit":
      return INLINE_EDIT_POLICY;
    case "chat": {
      const prompt = String(params.prompt ?? "").trim();
      if (prompt && CHAT_MUTATION_INTENT_RE.test(prompt)) return CHAT_MUTATION_POLICY;
      return CHAT_READ_ONLY_POLICY;
    }
    default: {
      const exhaustive: never = params.mode;
      return exhaustive;
    }
  }
}
