import test from "node:test";
import assert from "node:assert/strict";

import { FindReplaceController } from "./findReplaceController.js";

test("FindReplaceController blocks replace operations when canReplace denies", async () => {
  const calls = { beginBatch: 0, endBatch: 0 };
  const toasts = [];

  const controller = new FindReplaceController({
    workbook: {},
    beginBatch: () => {
      calls.beginBatch += 1;
    },
    endBatch: () => {
      calls.endBatch += 1;
    },
    canReplace: () => ({ allowed: false, reason: "Read-only: blocked" }),
    showToast: (message, type) => {
      toasts.push({ message, type });
    },
  });

  controller.query = "a";
  controller.replacement = "b";

  const result = await controller.replaceAll();
  assert.equal(result, null);
  assert.equal(calls.beginBatch, 0);
  assert.equal(calls.endBatch, 0);
  assert.deepEqual(toasts, [{ message: "Read-only: blocked", type: "warning" }]);
});

test("FindReplaceController blocks replaceNext operations when canReplace denies", async () => {
  const toasts = [];

  const controller = new FindReplaceController({
    workbook: {},
    canReplace: () => false,
    showToast: (message, type) => {
      toasts.push({ message, type });
    },
  });

  controller.query = "a";
  controller.replacement = "b";

  const result = await controller.replaceNext();
  assert.equal(result, null);
  assert.equal(toasts.length, 1);
  assert.equal(toasts[0]?.type, "warning");
  assert.equal(toasts[0]?.message, "Replacing cells is not allowed.");
});

