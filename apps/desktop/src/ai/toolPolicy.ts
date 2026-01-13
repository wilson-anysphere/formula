import { decideAllowedTools } from "../../../../packages/ai-tools/src/llm/toolPolicy.js";
import { ToolNameSchema, type ToolName, type ToolPolicy } from "../../../../packages/ai-tools/src/tool-schema.js";

export type DesktopAiMode = "chat" | "inline_edit" | "agent";

const CHAT_READ_ONLY_TOOLS: ToolName[] = ["read_range", "filter_range", "detect_anomalies", "compute_statistics"];

function uniqueInToolOrder(names: Iterable<ToolName>): ToolName[] {
  const set = new Set(names);
  return (ToolNameSchema.options as ToolName[]).filter((name) => set.has(name));
}

// Agent mode should be least-privilege by default. Even if a host integration later
// enables `allow_external_data`, we still explicitly deny external network tools for
// desktop agents unless/until we add a dedicated, user-visible approval flow.
const AGENT_POLICY: ToolPolicy = { externalNetworkAllowed: false };

export function getDesktopToolPolicy(params: {
  mode: DesktopAiMode;
  prompt?: string;
  hasAttachments?: boolean;
}): ToolPolicy {
  switch (params.mode) {
    case "agent":
      return AGENT_POLICY;
    case "inline_edit": {
      const prompt = String(params.prompt ?? "").trim();
      const policy = decideAllowedTools({
        mode: "inline_edit",
        user_text: prompt,
        // Inline edit is always attached to a concrete selection.
        has_attachments: true,
        allow_external_data: false
      });
      return { allowTools: policy.allowed_tools, externalNetworkAllowed: false };
    }
    case "chat": {
      const prompt = String(params.prompt ?? "").trim();
      const policy = decideAllowedTools({
        mode: "chat",
        user_text: prompt,
        has_attachments: Boolean(params.hasAttachments),
        // Desktop chat blocks external fetch by default (host-controlled, agent-only for now).
        allow_external_data: false
      });
      // Always allow safe read+analysis tools for grounding and post-response verification,
      // even when the user's prompt is an edit instruction.
      const allowTools = uniqueInToolOrder([...CHAT_READ_ONLY_TOOLS, ...policy.allowed_tools]);
      return { allowTools, externalNetworkAllowed: false };
    }
    default: {
      const exhaustive: never = params.mode;
      return exhaustive;
    }
  }
}
