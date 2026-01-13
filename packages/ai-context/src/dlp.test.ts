import { describe, expect, it } from "vitest";

import { classifyText, redactText } from "./dlp.js";

describe("dlp heuristic", () => {
  it("only flags Luhn-valid credit card numbers (keeps formatting variants)", () => {
    const validSpaced = "My card is 4111 1111 1111 1111.";
    const validDashed = "Or use 5500-0000-0000-0004 for testing.";
    const invalid = "Not a real card: 4111 1111 1111 1112.";
    const obviouslyInvalid = "Also not a card: 0000 0000 0000 0000.";

    expect(classifyText(validSpaced).findings).toContain("credit_card");
    expect(classifyText(validDashed).findings).toContain("credit_card");
    expect(classifyText(invalid).findings).not.toContain("credit_card");
    expect(classifyText(obviouslyInvalid).findings).not.toContain("credit_card");
  });

  it("detects phone numbers (international + US)", () => {
    const text = "Call (415) 555-2671 or +44 20 7946 0958.";
    expect(classifyText(text).findings).toContain("phone_number");
  });

  it("detects and redacts US phone numbers with extensions", () => {
    const text = "Dial (415) 555-2671 ext 1234 for support.";
    expect(classifyText(text).findings).toContain("phone_number");
    expect(redactText(text)).toBe("Dial [REDACTED_PHONE] for support.");
  });

  it("detects common API keys/tokens with conservative patterns", () => {
    const awsAccessKeyId = "AKIAIOSFODNN7EXAMPLE";
    const ghToken = "ghp_123456789012345678901234567890123456";
    const slackToken = "xoxb-123456789012-123456789012-abcdefghijklmnopqrstuvwxyzABCDEF";
    const text = `Keys: ${awsAccessKeyId} ${ghToken} ${slackToken}`;

    expect(classifyText(text).findings).toContain("api_key");
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
});
