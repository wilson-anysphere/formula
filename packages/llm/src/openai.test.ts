import { describe, expect, it } from "vitest";

import { OpenAIClient } from "./openai.js";

function withProcessUndefined<T>(fn: () => T): T {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, "process");
  try {
    Object.defineProperty(globalThis, "process", { value: undefined, configurable: true });
    return fn();
  } finally {
    if (descriptor) {
      Object.defineProperty(globalThis, "process", descriptor);
    } else {
      // @ts-expect-error - `process` is not part of the DOM typings.
      delete globalThis.process;
    }
  }
}

describe("OpenAIClient", () => {
  it("uses OPENAI_API_KEY from process.env in Node when present", () => {
    const originalKey = process.env.OPENAI_API_KEY;
    try {
      process.env.OPENAI_API_KEY = "env-test";
      const client = new OpenAIClient();
      expect((client as any).apiKey).toBe("env-test");
    } finally {
      if (originalKey === undefined) {
        delete process.env.OPENAI_API_KEY;
      } else {
        process.env.OPENAI_API_KEY = originalKey;
      }
    }
  });

  it("can be constructed when `process` is undefined if apiKey is provided", () => {
    const error = withProcessUndefined(() => {
      try {
        new OpenAIClient({ apiKey: "test" });
        return undefined;
      } catch (err) {
        return err;
      }
    });

    expect(error).toBeUndefined();
  });

  it("throws a friendly missing-key error when `process` is undefined", () => {
    const error = withProcessUndefined(() => {
      try {
        new OpenAIClient();
        return undefined;
      } catch (err) {
        return err;
      }
    });

    expect(error).toBeInstanceOf(Error);
    expect(error).not.toBeInstanceOf(ReferenceError);
    expect((error as Error).message).toMatch(/apiKey/);
    expect((error as Error).message).toMatch(/OPENAI_API_KEY/);
  });
});
