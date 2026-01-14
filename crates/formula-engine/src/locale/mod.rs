mod registry;
mod translate;
mod value_locale;

pub use registry::{get_locale, iter_locales, FormulaLocale, DE_DE, EN_US, ES_ES, FR_FR};
pub use translate::{
    canonicalize_formula, canonicalize_formula_with_style, localize_formula,
    localize_formula_with_style,
};
pub use value_locale::{DateOrder, ValueLocaleConfig};

#[derive(Debug, Clone, Copy)]
struct LocaleKeyParts<'a> {
    lang: &'a str,
    script: Option<&'a str>,
    region: Option<&'a str>,
}

fn normalize_locale_key(id: &str) -> Option<String> {
    let trimmed = id.trim();
    if trimmed.is_empty() {
        return None;
    }

    // Normalize common locale tag spellings:
    // - treat `-` and `_` as equivalent
    // - match case-insensitively
    //
    // Note: this intentionally supports a small set of locales the engine ships with, plus
    // best-effort normalization for common OS / browser locale tags.
    let mut key = String::with_capacity(trimmed.len());
    for ch in trimmed.chars() {
        let ch = match ch {
            '_' => '-',
            other => other,
        };
        key.push(ch.to_ascii_lowercase());
    }

    // Handle common POSIX locale tags like `en_US.UTF-8` or `de_DE@euro` by dropping the encoding /
    // modifier suffix. (Browser/BCP-47 tags don't use these, but it's a cheap compatibility win.)
    if let Some(idx) = key.find('.') {
        key.truncate(idx);
    }
    if let Some(idx) = key.find('@') {
        key.truncate(idx);
    }

    Some(key)
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
    let key = normalize_locale_key(id)?;
    let parts = parse_locale_key(&key)?;

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
