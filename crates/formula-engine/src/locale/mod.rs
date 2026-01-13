mod registry;
mod translate;
mod value_locale;

pub use registry::{get_locale, FormulaLocale, DE_DE, EN_US, ES_ES, FR_FR};
pub use translate::{
    canonicalize_formula, canonicalize_formula_with_style, localize_formula,
    localize_formula_with_style,
};
pub use value_locale::{DateOrder, ValueLocaleConfig};

fn normalize_locale_id(id: &str) -> Option<&'static str> {
    let trimmed = id.trim();
    if trimmed.is_empty() {
        return None;
    }

    // Normalize common locale tag spellings:
    // - treat `-` and `_` as equivalent
    // - match case-insensitively
    //
    // This intentionally supports a small set of locales the engine ships with.
    let mut key = String::with_capacity(trimmed.len());
    for ch in trimmed.chars() {
        let ch = match ch {
            '_' => '-',
            other => other,
        };
        key.push(ch.to_ascii_lowercase());
    }

    match key.as_str() {
        "en-us" | "en" => Some("en-US"),
        "de-de" | "de" => Some("de-DE"),
        "fr-fr" | "fr" => Some("fr-FR"),
        "es-es" | "es" => Some("es-ES"),
        _ => None,
    }
}
