import assert from "node:assert/strict";
import test from "node:test";

import { planTokenBudget } from "../src/budgetPlanner.js";

test("budgetPlanner: computes totals, overhead estimates, remaining tokens, and section allocations", () => {
  const estimator = {
    estimateTextTokens: (text) => (text ? 3 : 0),
    estimateMessageTokens: (_message) => 5,
    estimateMessagesTokens: (messages) => (Array.isArray(messages) ? messages.length * 5 : 0),
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
    estimator,
    sectionTargets: { schema: 50, retrieved: 50, samples: 50 },
  });

  assert.equal(plan.total, 100);
  assert.equal(plan.reserved, 20);
  assert.equal(plan.available, 80);

  // Custom estimator returns deterministic constants for easy arithmetic.
  assert.equal(plan.systemPromptTokens, 3);
  assert.equal(plan.toolDefinitionTokens, 3);
  assert.equal(plan.messageTokens, 10);
  assert.equal(plan.fixedPromptTokens, 16);
  assert.equal(plan.remainingForContextTokens, 64);

  assert.ok(plan.sections);
  assert.equal(plan.sections.allocatedTokens, 64);
  assert.equal(plan.sections.unallocatedTokens, 0);

  // Total target is 150, scaled down to 64. All targets are equal so the extra
  // rounding token goes to the alphabetically-first key (retrieved).
  assert.deepStrictEqual(plan.sections.allocationByKey, {
    retrieved: 22,
    samples: 21,
    schema: 21,
  });
});

test("budgetPlanner: handles missing/empty inputs", () => {
  const plan = planTokenBudget({
    maxContextTokens: 123,
    reserveForOutputTokens: 10,
  });

  assert.equal(plan.total, 123);
  assert.equal(plan.reserved, 10);
  assert.equal(plan.available, 113);

  assert.equal(plan.systemPromptTokens, 0);
  assert.equal(plan.toolDefinitionTokens, 0);
  assert.equal(plan.messageTokens, 0);
  assert.equal(plan.fixedPromptTokens, 0);
  assert.equal(plan.remainingForContextTokens, 113);
  assert.equal(plan.sections, undefined);
});

test("budgetPlanner: respects a custom TokenEstimator", () => {
  const charEstimator = {
    estimateTextTokens: (text) => String(text ?? "").length,
    estimateMessageTokens: (message) => String(message?.content ?? "").length,
    estimateMessagesTokens: (messages) =>
      Array.isArray(messages) ? messages.reduce((sum, m) => sum + String(m?.content ?? "").length, 0) : 0,
  };

  const systemPrompt = "x".repeat(10);
  const plan = planTokenBudget({
    maxContextTokens: 100,
    reserveForOutputTokens: 0,
    systemPrompt,
    estimator: charEstimator,
  });

  assert.equal(plan.systemPromptTokens, 10);
  assert.equal(plan.remainingForContextTokens, 90);
});

