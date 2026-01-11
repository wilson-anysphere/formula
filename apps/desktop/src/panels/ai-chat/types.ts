export type Attachment =
  | { type: "range"; reference: string; data?: unknown }
  | { type: "formula"; reference: string; data?: { formula: string } }
  | { type: "table"; reference: string; data?: unknown }
  | { type: "chart"; reference: string; data?: unknown };

export type ChatRole = "user" | "assistant" | "tool";

export interface ChatVerification {
  needs_tools: boolean;
  used_tools: boolean;
  verified: boolean;
  confidence: number;
  warnings: string[];
}

export interface ChatMessage {
  id: string;
  role: ChatRole;
  content: string;
  attachments?: Attachment[];
  verification?: ChatVerification;
  pending?: boolean;
  requiresApproval?: boolean;
}
