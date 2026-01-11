import fs from "node:fs";

import { describe, expect, it } from "vitest";

import { selectorKey as clientSelectorKey } from "./src/selectors.js";
import { validatePolicy as clientValidatePolicy } from "./src/policy.js";
import { evaluatePolicy as clientEvaluatePolicy } from "./src/policyEngine.js";
import clientCore from "./src/core.js";

import * as serverDlp from "../../../services/api/src/dlp/dlp";

type Vectors = {
  selectorKey: Array<{ name: string; selector: unknown; expected: string }>;
  policyValidation: Array<{
    name: string;
    policy: unknown;
    expectValid: boolean;
    errorContains?: string;
  }>;
  policyEvaluation: Array<{
    name: string;
    policy: unknown;
    action: string;
    classification: unknown;
    options?: { includeRestrictedContent?: boolean };
    expected: {
      decision: string;
      maxAllowed: string | null;
      classification: { level: string; labels: string[] };
    };
  }>;
  redaction: Array<{ name: string; value: unknown; expected: string }>;
};

function loadVectors(): Vectors {
  const url = new URL("../../../tests/dlp/dlp_shared_vectors.json", import.meta.url);
  return JSON.parse(fs.readFileSync(url, "utf8")) as Vectors;
}

function runValidation(fn: (value: unknown) => void, value: unknown): { ok: true } | { ok: false; message: string } {
  try {
    fn(value);
    return { ok: true };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return { ok: false, message };
  }
}

describe("DLP shared core vectors", () => {
  const vectors = loadVectors();

  it("selectorKey vectors match across client + server", () => {
    for (const v of vectors.selectorKey) {
      const clientKey = clientSelectorKey(v.selector);
      const serverKey = serverDlp.selectorKey(v.selector);
      expect(clientKey, v.name).toBe(serverKey);
      expect(clientKey, v.name).toBe(v.expected);
    }
  });

  it("policy validation vectors match across client + server", () => {
    for (const v of vectors.policyValidation) {
      const client = runValidation(clientValidatePolicy, v.policy);
      const server = runValidation(serverDlp.validateDlpPolicy, v.policy);

      expect(client.ok, `${v.name}: client.ok`).toBe(v.expectValid);
      expect(server.ok, `${v.name}: server.ok`).toBe(v.expectValid);

      if (!v.expectValid && v.errorContains) {
        if (!client.ok) expect(client.message, `${v.name}: client.message`).toContain(v.errorContains);
        if (!server.ok) expect(server.message, `${v.name}: server.message`).toContain(v.errorContains);
      }
    }
  });

  it("policy evaluation vectors match across client + server", () => {
    for (const v of vectors.policyEvaluation) {
      const client = clientEvaluatePolicy({
        action: v.action,
        classification: v.classification,
        policy: v.policy,
        options: v.options,
      });

      const server = serverDlp.evaluatePolicy({
        action: v.action,
        classification: v.classification,
        policy: v.policy,
        options: v.options,
      });

      expect(client, `${v.name}: client vs server`).toEqual(server);
      expect(
        { decision: client.decision, maxAllowed: client.maxAllowed, classification: client.classification },
        `${v.name}: matches expected`,
      ).toEqual(v.expected);
    }
  });

  it("redaction vectors match across client + server", () => {
    for (const v of vectors.redaction) {
      expect(clientCore.redact(v.value, null), `${v.name}: client`).toBe(serverDlp.redact(v.value, null));
      expect(serverDlp.redact(v.value, null), `${v.name}: server`).toBe(v.expected);
    }
  });
});

