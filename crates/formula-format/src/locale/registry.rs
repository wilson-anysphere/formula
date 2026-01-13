/// Minimal locale information used by [`format_number`] for "plain number" rendering.
///
/// This is intentionally lightweight and independent from the richer [`crate::Locale`] used by the
/// full Excel format-code engine. `NumberLocale` only carries the decimal/thousands separators
/// needed by the UI/formula bar when no explicit number format code is available.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NumberLocale {
    pub id: &'static str,
    pub decimal_separator: char,
    pub thousands_separator: Option<char>,
}

pub static EN_US: NumberLocale = NumberLocale {
    id: "en-US",
    decimal_separator: crate::Locale::en_us().decimal_sep,
    thousands_separator: Some(crate::Locale::en_us().thousands_sep),
};

/// British English uses the same separators as `en-US`.
pub static EN_GB: NumberLocale = NumberLocale {
    id: "en-GB",
    decimal_separator: crate::Locale::en_us().decimal_sep,
    thousands_separator: Some(crate::Locale::en_us().thousands_sep),
};

pub static DE_DE: NumberLocale = NumberLocale {
    id: "de-DE",
    decimal_separator: crate::Locale::de_de().decimal_sep,
    thousands_separator: Some(crate::Locale::de_de().thousands_sep),
};

/// French (France).
///
/// We use U+00A0 NO-BREAK SPACE as the thousands separator. Some environments prefer U+202F
/// NARROW NO-BREAK SPACE; if we ever need to distinguish, we can add a separate entry, but U+00A0
/// is widely supported and matches `crate::Locale::fr_fr()`.
pub static FR_FR: NumberLocale = NumberLocale {
    id: "fr-FR",
    decimal_separator: crate::Locale::fr_fr().decimal_sep,
    thousands_separator: Some(crate::Locale::fr_fr().thousands_sep),
};

pub static ES_ES: NumberLocale = NumberLocale {
    id: "es-ES",
    decimal_separator: crate::Locale::es_es().decimal_sep,
    thousands_separator: Some(crate::Locale::es_es().thousands_sep),
};

/// Spanish (Mexico) commonly uses `.` for decimals and `,` for thousands grouping.
pub static ES_MX: NumberLocale = NumberLocale {
    id: "es-MX",
    decimal_separator: crate::Locale::en_us().decimal_sep,
    thousands_separator: Some(crate::Locale::en_us().thousands_sep),
};

pub static IT_IT: NumberLocale = NumberLocale {
    id: "it-IT",
    decimal_separator: crate::Locale::it_it().decimal_sep,
    thousands_separator: Some(crate::Locale::it_it().thousands_sep),
};

/// Swiss German / Swiss-style number separators (`'` grouping, `.` decimal).
pub static DE_CH: NumberLocale = NumberLocale {
    id: "de-CH",
    decimal_separator: '.',
    thousands_separator: Some('\''),
};

/// Swiss French / Swiss-style number separators (`'` grouping, `.` decimal).
pub static FR_CH: NumberLocale = NumberLocale {
    id: "fr-CH",
    decimal_separator: '.',
    thousands_separator: Some('\''),
};

/// Swiss Italian / Swiss-style number separators (`'` grouping, `.` decimal).
pub static IT_CH: NumberLocale = NumberLocale {
    id: "it-CH",
    decimal_separator: '.',
    thousands_separator: Some('\''),
};

fn normalize_locale_id(id: &str) -> Option<&'static str> {
    let trimmed = id.trim();
    if trimmed.is_empty() {
        return None;
    }

    // Normalize common locale tag spellings:
    // - treat `-` and `_` as equivalent
    // - match case-insensitively
    //
    // This intentionally supports a small set of locales the formatter ships with.
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

    // Drop BCP-47 extensions (`en-US-u-nu-latn`, `fr-FR-x-private`, ...). For our purposes we only
    // care about the language/region portion.
    if let Some(idx) = key.find("-u-") {
        key.truncate(idx);
    }
    if let Some(idx) = key.find("-x-") {
        key.truncate(idx);
    }

    match key.as_str() {
        "en-us" | "en" => Some("en-US"),
        "en-gb" | "en-uk" => Some("en-GB"),
        "de-de" | "de" => Some("de-DE"),
        "de-ch" => Some("de-CH"),
        "fr-fr" | "fr" => Some("fr-FR"),
        "fr-ch" => Some("fr-CH"),
        "es-es" | "es" => Some("es-ES"),
        "es-mx" => Some("es-MX"),
        "it-it" | "it" => Some("it-IT"),
        "it-ch" => Some("it-CH"),
        _ => {
            // Fall back to the language part for region-specific variants we don't explicitly list
            // (e.g. `fr-CA`, `de-AT`, `en-AU`).
            let lang = key.split('-').next().unwrap_or("");
            match lang {
                "en" => Some("en-US"),
                "de" => Some("de-DE"),
                "fr" => Some("fr-FR"),
                "es" => Some("es-ES"),
                "it" => Some("it-IT"),
                _ => None,
            }
        }
    }
}

pub fn get_locale(id: &str) -> Option<&'static NumberLocale> {
    match normalize_locale_id(id)? {
        "en-US" => Some(&EN_US),
        "en-GB" => Some(&EN_GB),
        "de-DE" => Some(&DE_DE),
        "de-CH" => Some(&DE_CH),
        "fr-FR" => Some(&FR_FR),
        "fr-CH" => Some(&FR_CH),
        "es-ES" => Some(&ES_ES),
        "es-MX" => Some(&ES_MX),
        "it-IT" => Some(&IT_IT),
        "it-CH" => Some(&IT_CH),
        _ => None,
    }
}

/// Convert a full format-code [`crate::Locale`] into a `NumberLocale`.
///
/// When the separators match one of the built-in locales, the returned `id` is a canonical BCP-47
/// tag (e.g. `"fr-FR"`). Otherwise the `id` is `"und"` and separators are copied from `locale`.
pub fn number_locale_from_locale(locale: crate::Locale) -> NumberLocale {
    if locale == crate::Locale::en_us() {
        EN_US
    } else if locale == crate::Locale::de_de() {
        DE_DE
    } else if locale == crate::Locale::fr_fr() {
        FR_FR
    } else if locale == crate::Locale::es_es() {
        ES_ES
    } else if locale == crate::Locale::it_it() {
        IT_IT
    } else {
        NumberLocale {
            id: "und",
            decimal_separator: locale.decimal_sep,
            thousands_separator: Some(locale.thousands_sep),
        }
    }
}

/// Format a plain number using locale-specific separators.
///
/// This is intentionally not a full Excel number-format implementation yet; it
/// covers the most visible internationalization differences (thousands/decimal
/// separators) needed by the UI and formula bar.
pub fn format_number(value: f64, locale: &NumberLocale) -> String {
    // Avoid displaying negative zero (can show up after floating point operations).
    if value == 0.0 {
        return "0".to_string();
    }

    let s = value.to_string();

    // Preserve scientific notation as-is except for applying locale decimal
    // separator to the mantissa.
    if let Some((mantissa, exp)) = split_exponent(&s) {
        let mantissa = format_number_mantissa(mantissa, locale);
        return format!("{mantissa}{exp}");
    }

    format_number_mantissa(&s, locale)
}

fn split_exponent(s: &str) -> Option<(&str, &str)> {
    if let Some(idx) = s.find('e') {
        Some((&s[..idx], &s[idx..]))
    } else if let Some(idx) = s.find('E') {
        Some((&s[..idx], &s[idx..]))
    } else {
        None
    }
}

fn format_number_mantissa(mantissa: &str, locale: &NumberLocale) -> String {
    // Handle sign separately so grouping code only sees digits.
    let (sign, unsigned) = match mantissa.strip_prefix('-') {
        Some(rest) => ("-", rest),
        None => ("", mantissa),
    };

    // If the mantissa isn't a plain digit string (NaN/inf), leave unchanged.
    let (int_part, frac_part) = unsigned.split_once('.').unwrap_or((unsigned, ""));
    if !int_part.chars().all(|c| c.is_ascii_digit()) || !frac_part.chars().all(|c| c.is_ascii_digit())
    {
        return mantissa.to_string();
    }

    let grouped_int = match locale.thousands_separator {
        Some(sep) => group_thousands(int_part, sep),
        None => int_part.to_string(),
    };

    if frac_part.is_empty() {
        format!("{sign}{grouped_int}")
    } else {
        format!(
            "{sign}{grouped_int}{}{}",
            locale.decimal_separator, frac_part
        )
    }
}

fn group_thousands(int_part: &str, sep: char) -> String {
    let bytes = int_part.as_bytes();
    let len = bytes.len();
    if len <= 3 {
        return int_part.to_string();
    }

    let mut out = String::with_capacity(len + len / 3);
    let mut first_group = len % 3;
    if first_group == 0 {
        first_group = 3;
    }

    out.push_str(&int_part[..first_group]);
    let mut idx = first_group;
    while idx < len {
        out.push(sep);
        out.push_str(&int_part[idx..idx + 3]);
        idx += 3;
    }

    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalizes_locale_ids_like_formula_engine() {
        assert_eq!(normalize_locale_id("en-us"), Some("en-US"));
        assert_eq!(normalize_locale_id("en_US"), Some("en-US"));
        assert_eq!(normalize_locale_id("en_US.UTF-8"), Some("en-US"));
        assert_eq!(normalize_locale_id("en_US@posix"), Some("en-US"));
        assert_eq!(normalize_locale_id("en"), Some("en-US"));
        assert_eq!(normalize_locale_id("en-GB"), Some("en-GB"));
        assert_eq!(normalize_locale_id("en_uk"), Some("en-GB"));
        assert_eq!(normalize_locale_id("en-AU"), Some("en-US"));
        assert_eq!(normalize_locale_id("de"), Some("de-DE"));
        assert_eq!(normalize_locale_id("de-AT"), Some("de-DE"));
        assert_eq!(normalize_locale_id("de_ch"), Some("de-CH"));
        assert_eq!(normalize_locale_id("fr_fr"), Some("fr-FR"));
        assert_eq!(normalize_locale_id("fr-CA"), Some("fr-FR"));
        assert_eq!(normalize_locale_id("fr_CH"), Some("fr-CH"));
        assert_eq!(normalize_locale_id("es-ES"), Some("es-ES"));
        assert_eq!(normalize_locale_id("es-MX"), Some("es-MX"));
        assert_eq!(normalize_locale_id("es-AR"), Some("es-ES"));
        assert_eq!(normalize_locale_id("it_it"), Some("it-IT"));
        assert_eq!(normalize_locale_id("it-CH"), Some("it-CH"));
        assert_eq!(normalize_locale_id("fr-FR-u-nu-latn"), Some("fr-FR"));
        assert_eq!(normalize_locale_id(""), None);
    }

    #[test]
    fn formats_scientific_mantissa_with_locale_decimal_separator() {
        // `format_number` relies on `f64::to_string()`, which may or may not produce exponent
        // notation. Test the exponent handling logic directly so it stays covered.
        let (mantissa, exp) = split_exponent("-1.23E-6").unwrap();
        let mantissa = format_number_mantissa(mantissa, &DE_DE);
        assert_eq!(format!("{mantissa}{exp}"), "-1,23E-6");
    }

    #[test]
    fn built_in_number_locales_match_format_locale_separators() {
        let en = crate::Locale::en_us();
        assert_eq!(EN_US.decimal_separator, en.decimal_sep);
        assert_eq!(EN_US.thousands_separator, Some(en.thousands_sep));
        assert_eq!(EN_GB.decimal_separator, en.decimal_sep);
        assert_eq!(EN_GB.thousands_separator, Some(en.thousands_sep));

        let de = crate::Locale::de_de();
        assert_eq!(DE_DE.decimal_separator, de.decimal_sep);
        assert_eq!(DE_DE.thousands_separator, Some(de.thousands_sep));

        let fr = crate::Locale::fr_fr();
        assert_eq!(FR_FR.decimal_separator, fr.decimal_sep);
        assert_eq!(FR_FR.thousands_separator, Some(fr.thousands_sep));

        let es = crate::Locale::es_es();
        assert_eq!(ES_ES.decimal_separator, es.decimal_sep);
        assert_eq!(ES_ES.thousands_separator, Some(es.thousands_sep));

        let it = crate::Locale::it_it();
        assert_eq!(IT_IT.decimal_separator, it.decimal_sep);
        assert_eq!(IT_IT.thousands_separator, Some(it.thousands_sep));
    }
}
