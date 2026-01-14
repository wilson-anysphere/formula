// Keep in sync with `crates/formula-engine/src/locale/registry.rs` (`ALL_LOCALES`).
export type FormulaLocaleId =
  | "en-US"
  | "ja-JP"
  | "zh-CN"
  | "ko-KR"
  | "zh-TW"
  | "de-DE"
  | "fr-FR"
  | "es-ES";

/**
 * Best-effort locale-id normalization compatible with the formula engine's locale registry.
 *
 * The desktop app receives locale IDs from a mix of sources:
 * - BCP-47 language tags ("de-DE", "fr", "es-MX")
 * - POSIX locale IDs ("de_DE.UTF-8", "de_DE@euro")
 *
 * JS APIs like `Intl.NumberFormat` only accept BCP-47 tags, so we normalize inputs into
 * a simplified BCP-47 shape (`language` or `language-REGION`) and then, when needed,
 * map them onto the set of formula locales the engine ships (`en-US`, `de-DE`, `ja-JP`, ...).
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
  // POSIX "C locale" aliases show up in some environments; treat them as English so callers
  // can safely pass the result to `Intl.*` APIs.
  if (language === "c" || language === "posix") return "en-US";

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

type LocaleKeyParts = {
  lang: string;
  script: string | null;
  region: string | null;
};

function normalizeLocaleKey(id: string | null | undefined): string | null {
  const trimmed = String(id ?? "").trim();
  if (!trimmed) return null;

  // Strip POSIX encoding/modifier suffixes (".UTF-8", "@euro").
  let key = trimmed;
  for (const sep of [".", "@"]) {
    const idx = key.indexOf(sep);
    if (idx >= 0) key = key.slice(0, idx);
  }

  key = key.replaceAll("_", "-").trim();
  return key ? key : null;
}

function parseLocaleKey(key: string): LocaleKeyParts | null {
  // Parse BCP-47 tags and variants such as `de-CH-1996` or `zh-Hant-u-nu-latn`. We only care
  // about the language + optional script/region, ignoring variants/extensions.
  const parts = key.split("-").filter(Boolean);
  if (parts.length === 0) return null;

  const lang = parts[0]!.toLowerCase();

  let script: string | null = null;
  let region: string | null = null;

  let next = parts[1] ?? null;
  // Optional script subtag (4 alpha characters) comes before the region.
  if (next && /^[A-Za-z]{4}$/.test(next)) {
    script = next.toLowerCase();
    next = parts[2] ?? null;
  }

  if (next) {
    if (/^[A-Za-z]{2}$/.test(next)) region = next.toLowerCase();
    else if (/^\d{3}$/.test(next)) region = next;
  }

  return { lang, script, region };
}

/**
 * Normalize an arbitrary locale ID to one of the formula locales the engine supports.
 *
 * Returns `null` when the locale is unknown/unsupported.
 */
export function normalizeFormulaLocaleId(localeId: string | null | undefined): FormulaLocaleId | null {
  // Fast path: many callers (formula engine tooling, formula bar) already provide canonical
  // engine locale IDs. Avoid the heavier normalization/parsing work in that common case.
  switch (localeId) {
    case "en-US":
    case "ja-JP":
    case "zh-CN":
    case "ko-KR":
    case "zh-TW":
    case "de-DE":
    case "fr-FR":
    case "es-ES":
      return localeId;
    default:
      break;
  }

  const key = normalizeLocaleKey(localeId);
  if (!key) return null;
  const parts = parseLocaleKey(key);
  if (!parts) return null;

  // Map language/region variants onto the small set of engine-supported locales.
  // For example, `fr-CA` still resolves to `fr-FR`, and `de-AT` resolves to `de-DE`.
  switch (parts.lang) {
    // Many POSIX environments report locale as `C` / `POSIX` for the default "C locale".
    // Treat these as English (United States) so callers don't need to special-case.
    case "en":
    case "c":
    case "posix":
      // The engine treats `en-GB` (and other English locales) as an alias for the canonical
      // formula locale (English names + `,` separators).
      return "en-US";
    case "ja":
      return "ja-JP";
    case "zh": {
      // Prefer explicit region codes when present.
      //
      // Otherwise, use the BCP-47 script subtag:
      // - `zh-Hant` is Traditional Chinese, commonly associated with `zh-TW`.
      // - `zh-Hans` is Simplified Chinese, commonly associated with `zh-CN`.
      const region = parts.region;
      if (region) {
        return region === "tw" || region === "hk" || region === "mo" ? "zh-TW" : "zh-CN";
      }
      return parts.script === "hant" ? "zh-TW" : "zh-CN";
    }
    case "ko":
      return "ko-KR";
    case "de":
      return "de-DE";
    case "fr":
      return "fr-FR";
    case "es":
      return "es-ES";
    default:
      return null;
  }
}
