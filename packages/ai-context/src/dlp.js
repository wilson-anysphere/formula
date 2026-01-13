/**
 * Very small DLP classifier/redactor intended to complement a richer policy
 * engine (Task 93). This keeps RAG retrieval safe by default.
 */

const EMAIL_RE = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;
const SSN_RE = /\b\d{3}-\d{2}-\d{4}\b/g;
// Multi-line private key blocks (PEM/OpenSSH/PGP). Low-false-positive detector.
const PRIVATE_KEY_BLOCK_RE =
  /-----BEGIN (?:[A-Z0-9]+ )*PRIVATE KEY-----[\s\S]*?-----END (?:[A-Z0-9]+ )*PRIVATE KEY-----/g;
const PGP_PRIVATE_KEY_BLOCK_RE =
  /-----BEGIN PGP PRIVATE KEY BLOCK-----[\s\S]*?-----END PGP PRIVATE KEY BLOCK-----/g;
// Candidate detector only. Use Luhn validation before classifying/redacting.
//
// Keep the historical 13-16 digit range to avoid matching long numeric ids, while still
// supporting common card formats and separators (spaces/dashes).
const CREDIT_CARD_CANDIDATE_RE = /\b(?:\d[ -]*?){13,16}\b/g;

// Phone numbers are noisy; keep the patterns conservative and validate digit count.
// - International: require '+' prefix (E.164-like)
// - US: require separators / parentheses, optionally prefixed with +1
// NOTE: Phone patterns are particularly prone to false positives in spreadsheets (e.g.
// formulas that contain `+123...`). Keep this conservative:
// - Require a plausible delimiter (whitespace/opening punctuation/assignment) before '+'.
// - Avoid matching arithmetic operators like `)` or `*` that often appear in formulas.
const PHONE_INTL_CANDIDATE_RE = /(^|[\s"'({\[<:=,])(\+\d[\d\s().-]{7,}\d)/g;
const PHONE_US_CANDIDATE_RE = /(?:\+1[\s.-]?)?(?:\(\d{3}\)|\d{3})[\s.-]\d{3}[\s.-]\d{4}(?:\s*(?:ext|x|#)\s*\d{1,5})?/g;

// Conservative token detectors. Keep patterns specific to reduce false positives.
const API_KEY_RE =
  /\b(?:AKIA|ASIA)[0-9A-Z]{16}\b|\bgh[pousr]_[A-Za-z0-9]{36}\b|\bgithub_pat_[A-Za-z0-9_]{82}\b|\bAIza[0-9A-Za-z_-]{35}\b|\b(?:sk|rk)_(?:test|live)_[0-9a-zA-Z]{24}\b|\bxox[baprs]-\d{10,13}-\d{10,13}-[A-Za-z0-9]{24,64}\b/g;

// Candidate detector only. Validate IBAN checksum (mod 97) before classifying/redacting.
//
// This regex is intentionally broad (IBAN lengths are country-specific). We rely on
// checksum validation + prefix extraction to avoid false positives and to prevent the
// greedy match from swallowing adjacent tokens like `key=...`.
const IBAN_CANDIDATE_RE = /\b[A-Z]{2}\d{2}(?:[ ]?[A-Z0-9]){11,30}\b/gi;

function hasMatch(re, text) {
  re.lastIndex = 0;
  return re.test(text);
}

/**
 * @param {string} value
 */
function digitsOnly(value) {
  return String(value).replace(/\D/g, "");
}

/**
 * @param {string} digits
 */
function isAllSameDigit(digits) {
  if (!digits) return false;
  for (let i = 1; i < digits.length; i++) {
    if (digits[i] !== digits[0]) return false;
  }
  return true;
}

/**
 * Luhn checksum validation.
 * @param {string} digits
 */
function luhnIsValid(digits) {
  let sum = 0;
  let shouldDouble = false;
  for (let i = digits.length - 1; i >= 0; i--) {
    const code = digits.charCodeAt(i) - 48;
    if (code < 0 || code > 9) return false;
    let n = code;
    if (shouldDouble) {
      n *= 2;
      if (n > 9) n -= 9;
    }
    sum += n;
    shouldDouble = !shouldDouble;
  }
  return sum % 10 === 0;
}

/**
 * @param {string} candidate
 */
function isValidCreditCard(candidate) {
  const digits = digitsOnly(candidate);
  if (digits.length < 13 || digits.length > 16) return false;
  // All-same digits are almost certainly noise.
  if (isAllSameDigit(digits)) return false;
  // Further reduce false positives by only accepting well-known card IIN ranges.
  if (!matchesKnownCardBrand(digits)) return false;
  return luhnIsValid(digits);
}

/**
 * Heuristic allowlist for major card network IIN ranges.
 *
 * This intentionally does not attempt to cover every possible issuer/network; the goal
 * is to reduce false positives when scanning arbitrary spreadsheet text.
 *
 * @param {string} digits
 */
function matchesKnownCardBrand(digits) {
  const len = digits.length;
  const d2 = len >= 2 ? Number.parseInt(digits.slice(0, 2), 10) : NaN;
  const d3 = len >= 3 ? Number.parseInt(digits.slice(0, 3), 10) : NaN;
  const d4 = len >= 4 ? Number.parseInt(digits.slice(0, 4), 10) : NaN;

  // Visa: 13 or 16 digits, starts with 4.
  if (digits[0] === "4") return len === 13 || len === 16;

  // MasterCard: 16 digits, 51-55 or 2221-2720.
  if (len === 16) {
    if (d2 >= 51 && d2 <= 55) return true;
    if (d4 >= 2221 && d4 <= 2720) return true;
  }

  // American Express: 15 digits, 34 or 37.
  if (len === 15 && (d2 === 34 || d2 === 37)) return true;

  // Discover: 16 digits, 6011, 65, 644-649.
  if (len === 16) {
    if (digits.startsWith("6011")) return true;
    if (d2 === 65) return true;
    if (d3 >= 644 && d3 <= 649) return true;
  }

  // JCB: 16 digits, 3528-3589.
  if (len === 16 && d4 >= 3528 && d4 <= 3589) return true;

  // Diners Club: 14 digits, 300-305, 36, 38, 39.
  if (len === 14) {
    if (d3 >= 300 && d3 <= 305) return true;
    if (d2 === 36 || d2 === 38 || d2 === 39) return true;
  }

  return false;
}

/**
 * @param {string} text
 */
function hasValidCreditCard(text) {
  CREDIT_CARD_CANDIDATE_RE.lastIndex = 0;
  for (let match; (match = CREDIT_CARD_CANDIDATE_RE.exec(text)); ) {
    if (isValidCreditCard(match[0])) return true;
  }
  return false;
}

/**
 * @param {string} candidate
 */
function isValidPhone(candidate) {
  const raw = String(candidate);
  if (raw.includes("+")) {
    const digits = digitsOnly(raw);
    // International (+ prefix required by regex): allow 10-15 digits as a conservative E.164-ish bound.
    if (digits.length < 10 || digits.length > 15) return false;
    return digits[0] !== "0";
  }

  // Strip common US extension suffixes (e.g. "ext 123", "x123").
  const main = raw.replace(/\s*(?:ext|x|#)\s*\d{1,5}$/i, "");
  const digits = digitsOnly(main);
  // US-ish. Accept 10 digits, or 11 with leading 1.
  if (digits.length === 10) return true;
  if (digits.length === 11 && digits[0] === "1") return true;
  return false;
}

/**
 * @param {string} text
 */
function hasValidPhoneNumber(text) {
  PHONE_INTL_CANDIDATE_RE.lastIndex = 0;
  for (let match; (match = PHONE_INTL_CANDIDATE_RE.exec(text)); ) {
    const candidate = match[2] ?? match[0];
    // Avoid formula noise like `=+12345678901` (common Excel/Sheets pattern).
    if (match.index === 0 && match[1] === "=") continue;
    if (isValidPhone(candidate)) return true;
  }
  PHONE_US_CANDIDATE_RE.lastIndex = 0;
  for (let match; (match = PHONE_US_CANDIDATE_RE.exec(text)); ) {
    if (isValidPhone(match[0])) return true;
  }
  return false;
}

/**
 * @param {string} candidate
 */
function normalizeIban(candidate) {
  return String(candidate).replace(/\s+/g, "").toUpperCase();
}

/**
 * IBAN checksum validation (ISO 13616 mod 97).
 * @param {string} candidate
 */
function isValidIban(candidate) {
  const iban = normalizeIban(candidate);
  if (iban.length < 15 || iban.length > 34) return false;
  if (!/^[A-Z]{2}\d{2}[A-Z0-9]+$/.test(iban)) return false;

  const rearranged = iban.slice(4) + iban.slice(0, 4);
  let remainder = 0;
  for (let i = 0; i < rearranged.length; i++) {
    const ch = rearranged[i];
    const code = ch.charCodeAt(0);
    if (code >= 48 && code <= 57) {
      remainder = (remainder * 10 + (code - 48)) % 97;
      continue;
    }
    // A=10, B=11, ..., Z=35
    if (code >= 65 && code <= 90) {
      const value = code - 55;
      remainder = (remainder * 100 + value) % 97;
      continue;
    }
    return false;
  }

  return remainder === 1;
}

/**
 * IBAN candidate regexes can be greedy when the input contains adjacent identifiers
 * (e.g. `iban=GB82 ... 32 key=...`). Attempt to extract the longest checksum-valid
 * IBAN prefix without leaking it in redaction output.
 *
 * @param {string} candidate
 * @returns {{ iban: string, suffix: string } | null}
 */
function extractValidIban(candidate) {
  const raw = String(candidate);
  const normalized = normalizeIban(raw);
  const maxLen = Math.min(34, normalized.length);
  for (let len = maxLen; len >= 15; len--) {
    // The candidate regex ensures the string begins with country+checksum digits.
    const prefixNormalized = normalized.slice(0, len);
    if (!isValidIban(prefixNormalized)) continue;

    // Map the normalized prefix length back to a slice boundary in the raw string
    // (which may include spaces).
    let seen = 0;
    let cut = raw.length;
    for (let i = 0; i < raw.length; i++) {
      const ch = raw[i];
      const code = ch.charCodeAt(0);
      const isAlphaNum =
        (code >= 48 && code <= 57) || // 0-9
        (code >= 65 && code <= 90) || // A-Z
        (code >= 97 && code <= 122); // a-z
      if (!isAlphaNum) continue;
      seen++;
      if (seen === len) {
        cut = i + 1;
        break;
      }
    }

    return { iban: raw.slice(0, cut), suffix: raw.slice(cut) };
  }

  return null;
}

/**
 * @param {string} text
 */
function hasValidIban(text) {
  IBAN_CANDIDATE_RE.lastIndex = 0;
  for (let match; (match = IBAN_CANDIDATE_RE.exec(text)); ) {
    if (extractValidIban(match[0])) return true;
  }
  return false;
}

/**
 * @param {string} text
 */
export function classifyText(text) {
  const findings = [];
  if (hasMatch(PRIVATE_KEY_BLOCK_RE, text) || hasMatch(PGP_PRIVATE_KEY_BLOCK_RE, text)) findings.push("private_key");
  if (hasMatch(EMAIL_RE, text)) findings.push("email");
  if (hasMatch(SSN_RE, text)) findings.push("ssn");
  if (hasValidCreditCard(text)) findings.push("credit_card");
  if (hasValidPhoneNumber(text)) findings.push("phone_number");
  if (hasMatch(API_KEY_RE, text)) findings.push("api_key");
  if (hasValidIban(text)) findings.push("iban");
  const level = findings.length > 0 ? "sensitive" : "public";
  return { level, findings };
}

/**
 * @param {string} text
 */
export function redactText(text) {
  return String(text)
    .replace(PRIVATE_KEY_BLOCK_RE, "[REDACTED_PRIVATE_KEY]")
    .replace(PGP_PRIVATE_KEY_BLOCK_RE, "[REDACTED_PRIVATE_KEY]")
    .replace(API_KEY_RE, "[REDACTED_API_KEY]")
    .replace(EMAIL_RE, "[REDACTED_EMAIL]")
    .replace(SSN_RE, "[REDACTED_SSN]")
    .replace(IBAN_CANDIDATE_RE, (match) => {
      const extracted = extractValidIban(match);
      if (!extracted) return match;
      return `[REDACTED_IBAN]${extracted.suffix}`;
    })
    .replace(PHONE_INTL_CANDIDATE_RE, (match, prefix, candidate, offset) => {
      // Avoid formula noise like `=+12345678901` (common Excel/Sheets pattern).
      if (offset === 0 && prefix === "=") return match;
      if (!candidate || !isValidPhone(candidate)) return match;
      return `${prefix}[REDACTED_PHONE]`;
    })
    .replace(PHONE_US_CANDIDATE_RE, (match) => (isValidPhone(match) ? "[REDACTED_PHONE]" : match))
    .replace(CREDIT_CARD_CANDIDATE_RE, (match) => (isValidCreditCard(match) ? "[REDACTED_CREDIT_CARD]" : match));
}
