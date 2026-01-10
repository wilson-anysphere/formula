/**
 * Very small DLP classifier/redactor intended to complement a richer policy
 * engine (Task 93). This keeps RAG retrieval safe by default.
 */

const EMAIL_RE = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;
const SSN_RE = /\b\d{3}-\d{2}-\d{4}\b/g;
const CREDIT_CARD_RE = /\b(?:\d[ -]*?){13,16}\b/g;

function hasMatch(re, text) {
  re.lastIndex = 0;
  return re.test(text);
}

/**
 * @param {string} text
 */
export function classifyText(text) {
  const findings = [];
  if (hasMatch(EMAIL_RE, text)) findings.push("email");
  if (hasMatch(SSN_RE, text)) findings.push("ssn");
  if (hasMatch(CREDIT_CARD_RE, text)) findings.push("credit_card");
  const level = findings.length > 0 ? "sensitive" : "public";
  return { level, findings };
}

/**
 * @param {string} text
 */
export function redactText(text) {
  return String(text)
    .replace(EMAIL_RE, "[REDACTED_EMAIL]")
    .replace(SSN_RE, "[REDACTED_SSN]")
    .replace(CREDIT_CARD_RE, "[REDACTED_CREDIT_CARD]");
}
