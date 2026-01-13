import { describe, expect, it } from "vitest";

import { classifyText, redactText } from "./dlp.js";

describe("dlp heuristic", () => {
  it("only flags Luhn-valid credit card numbers (keeps formatting variants)", () => {
    const validSpaced = "My card is 4111 1111 1111 1111.";
    const validDashed = "Or use 5500-0000-0000-0004 for testing.";
    const invalid = "Not a real card: 4111 1111 1111 1112.";
    const obviouslyInvalid = "Also not a card: 0000 0000 0000 0000.";
    // Luhn-valid but unrealistic (common in ids/timestamps). Should not be flagged.
    const luhnValidButNotCard = "id=1000000000009";

    expect(classifyText(validSpaced).findings).toContain("credit_card");
    expect(classifyText(validDashed).findings).toContain("credit_card");
    expect(classifyText(invalid).findings).not.toContain("credit_card");
    expect(classifyText(obviouslyInvalid).findings).not.toContain("credit_card");
    expect(classifyText(luhnValidButNotCard).findings).not.toContain("credit_card");
  });

  it("detects phone numbers (international + US)", () => {
    const text = "Call (415) 555-2671 or +44 20 7946 0958.";
    expect(classifyText(text).findings).toContain("phone_number");
    expect(redactText(text)).toBe("Call [REDACTED_PHONE] or [REDACTED_PHONE].");
  });

  it("detects and redacts US phone numbers with extensions", () => {
    const text = "Dial (415) 555-2671 ext 1234 for support.";
    expect(classifyText(text).findings).toContain("phone_number");
    expect(redactText(text)).toBe("Dial [REDACTED_PHONE] for support.");
  });

  it("does not treat arithmetic expressions like phone numbers (reduced false positives)", () => {
    const formulaLike = "=A1+12345678901";
    expect(classifyText(formulaLike).findings).not.toContain("phone_number");
    expect(redactText(formulaLike)).toBe(formulaLike);
  });

  it("does not treat large numeric constants in formulas like phone numbers", () => {
    const formulaLike = "=SUM(A1)+12345678901";
    expect(classifyText(formulaLike).findings).not.toContain("phone_number");
    expect(redactText(formulaLike)).toBe(formulaLike);
  });

  it("does not treat Excel-style '=+<number>' formulas like phone numbers", () => {
    const formulaLike = "=+12345678901";
    expect(classifyText(formulaLike).findings).not.toContain("phone_number");
    expect(redactText(formulaLike)).toBe(formulaLike);
  });

  it("detects common API keys/tokens with conservative patterns", () => {
    const awsAccessKeyId = "AKIAIOSFODNN7EXAMPLE";
    const ghToken = "ghp_123456789012345678901234567890123456";
    const ghFineGrained = `github_pat_${"A".repeat(82)}`;
    const googleApiKey = `AIza${"A".repeat(35)}`;
    const stripeSecretKey = `sk_live_${"a".repeat(24)}`;
    const slackToken = "xoxb-123456789012-123456789012-abcdefghijklmnopqrstuvwxyzABCDEF";
    const text = `Keys: ${awsAccessKeyId} ${ghToken} ${ghFineGrained} ${googleApiKey} ${stripeSecretKey} ${slackToken}`;

    expect(classifyText(text).findings).toContain("api_key");
  });

  it("detects and redacts private key blocks", () => {
    const key = `-----BEGIN PRIVATE KEY-----\nabc123\n-----END PRIVATE KEY-----`;
    expect(classifyText(key).findings).toContain("private_key");
    expect(redactText(key)).toBe("[REDACTED_PRIVATE_KEY]");
  });

  it("detects and validates IBANs (mod 97)", () => {
    const valid = "GB82 WEST 1234 5698 7654 32";
    const invalid = "GB82 WEST 1234 5698 7654 33";
    // Valid example that ends with a letter.
    const validEndsWithLetter = "MT84 MALT 0110 0001 2345 MTLC AST0 01S";

    expect(classifyText(`iban=${valid}`).findings).toContain("iban");
    expect(classifyText(`iban=${invalid}`).findings).not.toContain("iban");
    expect(classifyText(`iban=${validEndsWithLetter}`).findings).toContain("iban");
    expect(redactText(`iban=${validEndsWithLetter}`)).toBe("iban=[REDACTED_IBAN]");
  });

  it("redacts new detector types without leaking substrings", () => {
    const cc = "4111 1111 1111 1111";
    const phone = "(415) 555-2671";
    const iban = "GB82 WEST 1234 5698 7654 32";
    const key = "AKIAIOSFODNN7EXAMPLE";

    const input = `cc=${cc} phone=${phone} iban=${iban} key=${key}`;
    const out = redactText(input);

    expect(out).toBe(
      "cc=[REDACTED_CREDIT_CARD] phone=[REDACTED_PHONE] iban=[REDACTED_IBAN] key=[REDACTED_API_KEY]",
    );
    expect(out).not.toContain(cc);
    expect(out).not.toContain(phone);
    expect(out).not.toContain(iban);
    expect(out).not.toContain(key);
  });

  it("does not redact invalid credit card numbers (reduced false positives)", () => {
    const invalid = "4111 1111 1111 1112";
    expect(redactText(`cc=${invalid}`)).toBe(`cc=${invalid}`);
  });

  it("does not redact Luhn-valid but unrealistic card-like ids (reduced false positives)", () => {
    const id = "1000000000009";
    expect(classifyText(`id=${id}`).findings).not.toContain("credit_card");
    expect(redactText(`id=${id}`)).toBe(`id=${id}`);
  });
});
