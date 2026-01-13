import { describe, expect, it } from "vitest";

import { planTokenBudget } from "./budgetPlanner.js";

describe("budgetPlanner", () => {
  it("computes totals, overhead estimates, remaining tokens, and section allocations", () => {
    const estimator = {
      estimateTextTokens: (text: string) => (text ? 3 : 0),
      estimateMessageTokens: (_message: any) => 5,
      estimateMessagesTokens: (messages: any[]) => (Array.isArray(messages) ? messages.length * 5 : 0),
    };

    const plan = planTokenBudget({
      maxContextTokens: 100,
      reserveForOutputTokens: 20,
      systemPrompt: "system prompt",
      tools: [{ name: "read_range", description: "Read some cells" }],
      messages: [
        { role: "user", content: "Hello" },
        { role: "assistant", content: "Hi" },
      ],
      estimator: estimator as any,
      sectionTargets: { schema: 50, retrieved: 50, samples: 50 },
    });

    expect(plan.total).toBe(100);
    expect(plan.reserved).toBe(20);
    expect(plan.available).toBe(80);

    // Custom estimator returns deterministic constants for easy arithmetic.
    expect(plan.systemPromptTokens).toBe(3);
    expect(plan.toolDefinitionTokens).toBe(3);
    expect(plan.messageTokens).toBe(10);
    expect(plan.fixedPromptTokens).toBe(16);
    expect(plan.remainingForContextTokens).toBe(64);

    expect(plan.sections).toBeTruthy();
    expect(plan.sections!.allocatedTokens).toBe(64);
    expect(plan.sections!.unallocatedTokens).toBe(0);

    // Total target is 150, scaled down to 64. All targets are equal so the extra
    // rounding token goes to the alphabetically-first key (retrieved).
    expect(plan.sections!.allocationByKey).toEqual({
      retrieved: 22,
      samples: 21,
      schema: 21,
    });
  });

  it("handles missing/empty inputs", () => {
    const plan = planTokenBudget({
      maxContextTokens: 123,
      reserveForOutputTokens: 10,
    });

    expect(plan.total).toBe(123);
    expect(plan.reserved).toBe(10);
    expect(plan.available).toBe(113);

    expect(plan.systemPromptTokens).toBe(0);
    expect(plan.toolDefinitionTokens).toBe(0);
    expect(plan.messageTokens).toBe(0);
    expect(plan.fixedPromptTokens).toBe(0);
    expect(plan.remainingForContextTokens).toBe(113);
    expect(plan.sections).toBeUndefined();
  });

  it("respects a custom TokenEstimator", () => {
    const charEstimator = {
      estimateTextTokens: (text: string) => String(text ?? "").length,
      estimateMessageTokens: (message: any) => String(message?.content ?? "").length,
      estimateMessagesTokens: (messages: any[]) =>
        Array.isArray(messages) ? messages.reduce((sum, m) => sum + String(m?.content ?? "").length, 0) : 0,
    };

    const systemPrompt = "x".repeat(10);
    const plan = planTokenBudget({
      maxContextTokens: 100,
      reserveForOutputTokens: 0,
      systemPrompt,
      estimator: charEstimator as any,
    });

    expect(plan.systemPromptTokens).toBe(10);
    expect(plan.remainingForContextTokens).toBe(90);
  });
});
