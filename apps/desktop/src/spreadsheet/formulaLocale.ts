export type FormulaLocaleId = "en-US" | "de-DE" | "fr-FR" | "es-ES";

/**
 * Best-effort locale-id normalization compatible with the formula engine's locale registry.
 *
 * The desktop app receives locale IDs from a mix of sources:
 * - BCP-47 language tags ("de-DE", "fr", "es-MX")
 * - POSIX locale IDs ("de_DE.UTF-8", "de_DE@euro")
 *
 * JS APIs like `Intl.NumberFormat` only accept BCP-47 tags, so we normalize inputs into
 * a simplified BCP-47 shape (`language` or `language-REGION`) and then, when needed,
 * map them onto the small set of formula locales the engine ships (`en-US`, `de-DE`, ...).
 *
 * Keep this logic aligned with `crates/formula-engine/tests/locale_id_normalization.rs`.
 */

/**
 * Normalize an arbitrary locale ID into a simplified BCP-47 tag usable with `Intl.*`.
 *
 * - Trims whitespace
 * - Converts `_` to `-`
 * - Strips POSIX encoding/modifier suffixes (".UTF-8", "@euro")
 * - Ignores BCP-47 extensions/variants, preserving only `language` and an optional `REGION`
 */
export function normalizeLocaleId(localeId: string | null | undefined): string | null {
  const raw = String(localeId ?? "").trim();
  if (!raw) return null;

  // Strip common POSIX suffixes (encoding/modifier), e.g. `de_DE.UTF-8`, `de_DE@euro`.
  let base = raw;
  for (const sep of [".", "@"]) {
    const idx = base.indexOf(sep);
    if (idx >= 0) base = base.slice(0, idx);
  }

  base = base.replaceAll("_", "-");

  const parts = base.split("-").filter(Boolean);
  if (parts.length === 0) return null;

  const language = parts[0]!.toLowerCase();

  // Find the first region-like subtag before any BCP-47 extension singleton (e.g. `u`, `x`).
  let region: string | null = null;
  for (let i = 1; i < parts.length; i += 1) {
    const part = parts[i]!;
    if (part.length === 1) break; // extension singleton
    if (/^[A-Za-z]{2}$/.test(part)) {
      region = part.toUpperCase();
      break;
    }
    if (/^\d{3}$/.test(part)) {
      region = part;
      break;
    }
  }

  return region ? `${language}-${region}` : language;
}

/**
 * Normalize an arbitrary locale ID to one of the formula locales the engine supports.
 *
 * Returns `null` when the locale is unknown/unsupported.
 */
export function normalizeFormulaLocaleId(localeId: string | null | undefined): FormulaLocaleId | null {
  const normalized = normalizeLocaleId(localeId);
  if (!normalized) return null;

  const lang = normalized.split("-")[0] ?? "";
  switch (lang.toLowerCase()) {
    case "de":
      return "de-DE";
    case "fr":
      return "fr-FR";
    case "es":
      return "es-ES";
    case "en":
      // The engine treats `en-GB` as an alias for the canonical formula locale (English names + `,` separators).
      return "en-US";
    default:
      return null;
  }
}

