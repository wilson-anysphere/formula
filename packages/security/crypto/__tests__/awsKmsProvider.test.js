import test from "node:test";
import assert from "node:assert/strict";
import { AwsKmsProvider } from "../kms/providers.js";

async function hasAwsSdk() {
  try {
    await import("@aws-sdk/client-kms");
    return true;
  } catch {
    return false;
  }
}

test("AwsKmsProvider requires region", () => {
  assert.throws(() => new AwsKmsProvider({}), /requires region/i);
});

test("AwsKmsProvider.wrapKey requires keyId", async () => {
  const kms = new AwsKmsProvider({ region: "us-east-1", keyId: null });
  await assert.rejects(() => kms.wrapKey({ plaintextKey: Buffer.alloc(32) }), /requires keyId/i);
});

test("AwsKmsProvider throws helpful error when @aws-sdk/client-kms is missing", async (t) => {
  if (await hasAwsSdk()) {
    t.skip("Environment already provides @aws-sdk/client-kms; skipping missing-sdk error assertion.");
    return;
  }

  const kms = new AwsKmsProvider({ region: "us-east-1", keyId: "alias/example" });
  await assert.rejects(
    () => kms.wrapKey({ plaintextKey: Buffer.alloc(32), encryptionContext: { orgId: "org-1" } }),
    /Install @aws-sdk\/client-kms/i
  );
});
