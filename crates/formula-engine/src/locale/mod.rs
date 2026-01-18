mod registry;
mod translate;
mod value_locale;

use std::borrow::Cow;
use std::sync::OnceLock;

pub use registry::{get_locale, iter_locales, FormulaLocale, DE_DE, EN_US, ES_ES, FR_FR};
pub use translate::{
    canonicalize_formula, canonicalize_formula_with_style, localize_formula,
    localize_formula_with_style,
};
pub use value_locale::{DateOrder, ValueLocaleConfig};

/// Resolve the workbook's legacy Windows text codepage for a locale id.
///
/// Excel's legacy DBCS (`*B`) text functions (`LENB`, `LEFTB`, etc.) and `ASC`/`DBCS` conversions
/// depend on a workbook-level "text codepage". The engine models this as a Windows code page
/// number (e.g. 932 for Shift_JIS).
///
/// This helper maps the engine's canonical locale ids onto the Windows codepages that Excel uses
/// for common DBCS locales:
/// - `ja-JP` → `932` (Shift_JIS)
/// - `zh-CN` → `936` (GBK)
/// - `ko-KR` → `949` (Korean)
/// - `zh-TW` → `950` (Big5)
///
/// For all other (including unknown) locale ids, this returns the default single-byte codepage
/// `1252` (Windows-1252).
///
/// The input may be a raw OS/browser locale tag (e.g. `"ja"`, `"zh-HK"`, `"de_DE.UTF-8"`). It is
/// normalized using [`get_locale`] so aliases resolve to the engine's canonical locale ids.
pub fn text_codepage_for_locale_id(locale_id: &str) -> u16 {
    match get_locale(locale_id).map(|locale| locale.id) {
        Some("ja-JP") => 932,
        Some("zh-CN") => 936,
        Some("ko-KR") => 949,
        Some("zh-TW") => 950,
        _ => 1252,
    }
}

/// Return the list of locale identifiers supported by the formula engine.
///
/// The order is deterministic so downstreams can safely embed it in user-facing error messages and
/// tests.
pub fn supported_locale_ids() -> &'static [&'static str] {
    static IDS: OnceLock<Box<[&'static str]>> = OnceLock::new();
    IDS.get_or_init(|| {
        let mut ids: Vec<&'static str> = Vec::new();
        let (_, upper) = iter_locales().size_hint();
        if let Some(capacity) = upper {
            if ids.try_reserve_exact(capacity).is_err() {
                debug_assert!(
                    false,
                    "allocation failed (supported_locale_ids, capacity={capacity})"
                );
                return Vec::new().into_boxed_slice();
            }
        }
        for locale in iter_locales() {
            ids.push(locale.id);
        }
        ids.sort_unstable();
        ids.into_boxed_slice()
    })
}

#[derive(Debug, Clone, Copy)]
struct LocaleKeyParts<'a> {
    lang: &'a str,
    script: Option<&'a str>,
    region: Option<&'a str>,
}

fn normalize_locale_key(id: &str) -> Option<Cow<'_, str>> {
    let trimmed = id.trim();
    if trimmed.is_empty() {
        return None;
    }

    let suffix_start = trimmed
        .find(|c| matches!(c, '.' | '@'))
        .unwrap_or(trimmed.len());
    let trimmed = &trimmed[..suffix_start];
    if trimmed.is_empty() {
        return None;
    }

    let needs_transform = trimmed
        .as_bytes()
        .iter()
        .any(|b| *b == b'_' || b.is_ascii_uppercase());
    if !needs_transform {
        return Some(Cow::Borrowed(trimmed));
    }

    // Normalize common locale tag spellings:
    // - treat `-` and `_` as equivalent
    // - match case-insensitively
    //
    // Note: this intentionally supports a small set of locales the engine ships with, plus
    // best-effort normalization for common OS / browser locale tags.
    let mut key = String::new();
    if key.try_reserve_exact(trimmed.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (normalize_locale_key, len={})",
            trimmed.len()
        );
        return None;
    }
    for ch in trimmed.chars() {
        let ch = match ch {
            '_' => '-',
            other => other,
        };
        key.push(ch.to_ascii_lowercase());
    }
    Some(Cow::Owned(key))
}

fn parse_locale_key(key: &str) -> Option<LocaleKeyParts<'_>> {
    // Parse BCP-47 tags and variants such as `de-CH-1996` or `fr-Latn-FR-u-nu-latn`. We only care
    // about the language + optional script/region, ignoring variants/extensions.
    let mut parts = key.split('-').filter(|p| !p.is_empty());
    let lang = parts.next()?;
    let mut next = parts.next();
    // Optional script subtag (4 alpha characters) comes before the region.
    let script = next.filter(|p| p.len() == 4 && p.chars().all(|c| c.is_ascii_alphabetic()));
    if script.is_some() {
        next = parts.next();
    }

    let region = next.filter(|p| {
        (p.len() == 2 && p.chars().all(|c| c.is_ascii_alphabetic()))
            || (p.len() == 3 && p.chars().all(|c| c.is_ascii_digit()))
    });

    Some(LocaleKeyParts {
        lang,
        script,
        region,
    })
}

fn normalize_locale_id(id: &str) -> Option<&'static str> {
    let trimmed = id.trim();
    if trimmed.is_empty() {
        return None;
    }

    // Fast path for canonical ids (common in engine/wire formats): avoid allocation + lowercasing.
    match trimmed {
        "en-US" => return Some("en-US"),
        "de-DE" => return Some("de-DE"),
        "fr-FR" => return Some("fr-FR"),
        "es-ES" => return Some("es-ES"),
        "ja-JP" => return Some("ja-JP"),
        "zh-CN" => return Some("zh-CN"),
        "zh-TW" => return Some("zh-TW"),
        "ko-KR" => return Some("ko-KR"),
        // Common aliases for the formula parsing locale (still English function names + `,`).
        "en-GB" | "en-UK" | "en-AU" | "en-NZ" | "en-IE" | "en-ZA" => return Some("en-US"),
        // POSIX "C locale" aliases.
        "C" | "POSIX" => return Some("en-US"),
        _ => {}
    }

    let key = normalize_locale_key(id)?;
    let parts = parse_locale_key(key.as_ref())?;

    // Map language/region variants onto the small set of engine-supported locales.
    // For example, `fr-CA` still resolves to `fr-FR`, and `de-AT` resolves to `de-DE`.
    match parts.lang {
        // Many POSIX environments report locale as `C` / `POSIX` for the default "C locale".
        // Treat these as English (United States) so callers don't need to special-case.
        "en" | "c" | "posix" => Some("en-US"),
        "ja" => Some("ja-JP"),
        "zh" => {
            // Prefer explicit region codes when present.
            //
            // Otherwise, use the BCP-47 script subtag:
            // - `zh-Hant` is Traditional Chinese, commonly associated with `zh-TW`.
            // - `zh-Hans` is Simplified Chinese, commonly associated with `zh-CN`.
            match parts.region {
                Some("tw") | Some("hk") | Some("mo") => Some("zh-TW"),
                Some(_) => Some("zh-CN"),
                None => match parts.script {
                    Some("hant") => Some("zh-TW"),
                    Some("hans") => Some("zh-CN"),
                    _ => Some("zh-CN"),
                },
            }
        }
        "ko" => Some("ko-KR"),
        "de" => Some("de-DE"),
        "fr" => Some("fr-FR"),
        "es" => Some("es-ES"),
        _ => None,
    }
}
