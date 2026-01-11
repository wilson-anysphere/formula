import { describe, expect, it } from "vitest";

import { createDesktopDlpContext } from "./desktopDlp.js";

import { createMemoryStorage } from "../../../../packages/security/dlp/src/classificationStore.js";
import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";

describe("createDesktopDlpContext", () => {
  it("uses the active org id stored in localStorage when orgId is not provided", () => {
    const storage = createMemoryStorage();
    storage.setItem("dlp:activeOrgId", "acme");

    const ctx = createDesktopDlpContext({ documentId: "doc-0", storage });
    expect(ctx.orgId).toBe("acme");
  });

  it("does not throw when stored policies are invalid", () => {
    const storage = createMemoryStorage();
    // LocalPolicyStore will parse this successfully, but mergePolicies/validatePolicy should reject it.
    storage.setItem("dlp:orgPolicy:default", JSON.stringify("not-a-policy"));

    const ctx = createDesktopDlpContext({ documentId: "doc-1", storage });
    expect(ctx.documentId).toBe("doc-1");
    expect(ctx.policy?.rules?.[DLP_ACTION.AI_CLOUD_PROCESSING]).toBeDefined();
  });

  it("tolerates storage implementations that throw", () => {
    const throwingStorage = {
      getItem() {
        throw new Error("boom");
      },
      setItem() {
        throw new Error("boom");
      },
      removeItem() {
        throw new Error("boom");
      },
    };

    const ctx = createDesktopDlpContext({ documentId: "doc-2", storage: throwingStorage as any });
    expect(ctx.documentId).toBe("doc-2");
    expect(ctx.policy?.rules?.[DLP_ACTION.AI_CLOUD_PROCESSING]).toBeDefined();
  });
});
