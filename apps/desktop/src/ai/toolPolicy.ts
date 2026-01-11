import type { ToolPolicy, ToolName } from "../../../../packages/ai-tools/src/tool-schema.js";

export type DesktopAiMode = "chat" | "inline_edit" | "agent";

const CHAT_READ_ONLY_POLICY: ToolPolicy = {
  allowCategories: ["read", "analysis"],
  externalNetworkAllowed: false
};

const CHAT_WRITE_INTENT_RE = new RegExp(
  String.raw`\b(` +
    [
      "replace",
      "fill",
      "write",
      "insert",
      "update",
      "set",
      "edit",
      "change",
      "delete",
      "remove",
      "clear",
      "sort",
      "apply"
    ].join("|") +
    String.raw`)\b`,
  "i"
);

const CHAT_FORMAT_INTENT_RE = /\b(format|bold|italic|underline|font|color|colour|highlight|background|number\s*format|currency|percent|percentage|align|alignment)\b/i;
const CHAT_CHART_INTENT_RE = /\b(chart|plot|graph)\b/i;
const CHAT_PIVOT_INTENT_RE = /\bpivot\b/i;

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

export function getDesktopToolPolicy(params: { mode: DesktopAiMode; prompt?: string }): ToolPolicy {
  switch (params.mode) {
    case "agent":
      return AGENT_POLICY;
    case "inline_edit":
      return INLINE_EDIT_POLICY;
    case "chat": {
      const prompt = String(params.prompt ?? "").trim();
      if (!prompt) return CHAT_READ_ONLY_POLICY;

      const wantsWrite = CHAT_WRITE_INTENT_RE.test(prompt);
      const wantsFormat = CHAT_FORMAT_INTENT_RE.test(prompt);
      const wantsChart = CHAT_CHART_INTENT_RE.test(prompt);
      const wantsPivot = CHAT_PIVOT_INTENT_RE.test(prompt);

      if (!wantsWrite && !wantsFormat && !wantsChart && !wantsPivot) return CHAT_READ_ONLY_POLICY;

      const allowCategories: NonNullable<ToolPolicy["allowCategories"]> = ["read", "analysis"];
      if (wantsWrite) allowCategories.push("mutate");
      if (wantsFormat) allowCategories.push("format");
      if (wantsChart) allowCategories.push("chart");
      if (wantsPivot) allowCategories.push("pivot");
      return { allowCategories, externalNetworkAllowed: false };
    }
    default: {
      const exhaustive: never = params.mode;
      return exhaustive;
    }
  }
}
