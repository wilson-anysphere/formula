import assert from "node:assert/strict";
import test from "node:test";

import { classifyText, redactText } from "../src/dlp.js";

test("dlp heuristic: only flags Luhn-valid credit card numbers (keeps formatting variants)", () => {
  const validSpaced = "My card is 4111 1111 1111 1111.";
  const validDashed = "Or use 5500-0000-0000-0004 for testing.";
  const validAmex = "Amex: 3782 822463 10005";
  const validDiscover = "Discover: 6011 1111 1111 1117";
  const validJcb = "JCB: 3530 1113 3330 0000";
  const validDiners = "Diners: 3056 930902 5904";
  const invalid = "Not a real card: 4111 1111 1111 1112.";
  const obviouslyInvalid = "Also not a card: 0000 0000 0000 0000.";
  // Luhn-valid but unrealistic (common in ids/timestamps). Should not be flagged.
  const luhnValidButNotCard = "id=1000000000009";

  assert.ok(classifyText(validSpaced).findings.includes("credit_card"));
  assert.ok(classifyText(validDashed).findings.includes("credit_card"));
  assert.ok(classifyText(validAmex).findings.includes("credit_card"));
  assert.ok(classifyText(validDiscover).findings.includes("credit_card"));
  assert.ok(classifyText(validJcb).findings.includes("credit_card"));
  assert.ok(classifyText(validDiners).findings.includes("credit_card"));
  assert.ok(!classifyText(invalid).findings.includes("credit_card"));
  assert.ok(!classifyText(obviouslyInvalid).findings.includes("credit_card"));
  assert.ok(!classifyText(luhnValidButNotCard).findings.includes("credit_card"));
});

test("dlp heuristic: detects phone numbers (international + US)", () => {
  const text = "Call (415) 555-2671 or +44 20 7946 0958.";
  assert.ok(classifyText(text).findings.includes("phone_number"));
  assert.equal(redactText(text), "Call [REDACTED_PHONE] or [REDACTED_PHONE].");
});

test("dlp heuristic: detects and redacts US phone numbers with extensions", () => {
  const text = "Dial (415) 555-2671 ext 1234 for support.";
  assert.ok(classifyText(text).findings.includes("phone_number"));
  assert.equal(redactText(text), "Dial [REDACTED_PHONE] for support.");
});

test("dlp heuristic: does not treat arithmetic expressions like phone numbers (reduced false positives)", () => {
  const formulaLike = "=A1+12345678901";
  assert.ok(!classifyText(formulaLike).findings.includes("phone_number"));
  assert.equal(redactText(formulaLike), formulaLike);
});

test("dlp heuristic: does not treat large numeric constants in formulas like phone numbers", () => {
  const formulaLike = "=SUM(A1)+12345678901";
  assert.ok(!classifyText(formulaLike).findings.includes("phone_number"));
  assert.equal(redactText(formulaLike), formulaLike);
});

test("dlp heuristic: does not treat Excel-style '=+<number>' formulas like phone numbers", () => {
  const formulaLike = "=+12345678901";
  assert.ok(!classifyText(formulaLike).findings.includes("phone_number"));
  assert.equal(redactText(formulaLike), formulaLike);
});

test("dlp heuristic: detects common API keys/tokens with conservative patterns", () => {
  const awsAccessKeyId = "AKIAIOSFODNN7EXAMPLE";
  const ghToken = "ghp_123456789012345678901234567890123456";
  const ghFineGrained = `github_pat_${"A".repeat(82)}`;
  const googleApiKey = `AIza${"A".repeat(35)}`;
  const stripeSecretKey = `sk_live_${"a".repeat(24)}`;
  const slackToken = "xoxb-123456789012-123456789012-abcdefghijklmnopqrstuvwxyzABCDEF";
  const text = `Keys: ${awsAccessKeyId} ${ghToken} ${ghFineGrained} ${googleApiKey} ${stripeSecretKey} ${slackToken}`;

  assert.ok(classifyText(text).findings.includes("api_key"));
});

test("dlp heuristic: detects and redacts private key blocks", () => {
  const key = `-----BEGIN PRIVATE KEY-----\nabc123\n-----END PRIVATE KEY-----`;
  assert.ok(classifyText(key).findings.includes("private_key"));
  assert.equal(redactText(key), "[REDACTED_PRIVATE_KEY]");
});

test("dlp heuristic: detects and validates IBANs (mod 97)", () => {
  const valid = "GB82 WEST 1234 5698 7654 32";
  const invalid = "GB82 WEST 1234 5698 7654 33";
  // Valid example that ends with a letter.
  const validEndsWithLetter = "MT84 MALT 0110 0001 2345 MTLC AST0 01S";

  assert.ok(classifyText(`iban=${valid}`).findings.includes("iban"));
  assert.ok(!classifyText(`iban=${invalid}`).findings.includes("iban"));
  assert.ok(classifyText(`iban=${validEndsWithLetter}`).findings.includes("iban"));
  assert.equal(redactText(`iban=${validEndsWithLetter}`), "iban=[REDACTED_IBAN]");
});

test("dlp heuristic: redacts new detector types without leaking substrings", () => {
  const cc = "4111 1111 1111 1111";
  const phone = "(415) 555-2671";
  const iban = "GB82 WEST 1234 5698 7654 32";
  const key = "AKIAIOSFODNN7EXAMPLE";

  const input = `cc=${cc} phone=${phone} iban=${iban} key=${key}`;
  const out = redactText(input);

  assert.equal(
    out,
    "cc=[REDACTED_CREDIT_CARD] phone=[REDACTED_PHONE] iban=[REDACTED_IBAN] key=[REDACTED_API_KEY]",
  );
  assert.ok(!out.includes(cc));
  assert.ok(!out.includes(phone));
  assert.ok(!out.includes(iban));
  assert.ok(!out.includes(key));
});

test("dlp heuristic: does not redact invalid credit card numbers (reduced false positives)", () => {
  const invalid = "4111 1111 1111 1112";
  assert.equal(redactText(`cc=${invalid}`), `cc=${invalid}`);
});

test("dlp heuristic: does not redact Luhn-valid but unrealistic card-like ids (reduced false positives)", () => {
  const id = "1000000000009";
  assert.ok(!classifyText(`id=${id}`).findings.includes("credit_card"));
  assert.equal(redactText(`id=${id}`), `id=${id}`);
});

