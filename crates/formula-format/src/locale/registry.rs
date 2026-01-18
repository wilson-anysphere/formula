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

/// Portuguese (Portugal).
pub static PT_PT: NumberLocale = NumberLocale {
    id: "pt-PT",
    decimal_separator: crate::Locale::es_es().decimal_sep,
    thousands_separator: Some(crate::Locale::es_es().thousands_sep),
};

/// Portuguese (Brazil).
pub static PT_BR: NumberLocale = NumberLocale {
    id: "pt-BR",
    decimal_separator: crate::Locale::es_es().decimal_sep,
    thousands_separator: Some(crate::Locale::es_es().thousands_sep),
};

/// Dutch (Netherlands).
pub static NL_NL: NumberLocale = NumberLocale {
    id: "nl-NL",
    decimal_separator: crate::Locale::es_es().decimal_sep,
    thousands_separator: Some(crate::Locale::es_es().thousands_sep),
};

/// Dutch (Belgium).
pub static NL_BE: NumberLocale = NumberLocale {
    id: "nl-BE",
    decimal_separator: crate::Locale::es_es().decimal_sep,
    thousands_separator: Some(crate::Locale::es_es().thousands_sep),
};

/// Japanese (Japan) commonly uses `.` for decimals and `,` for thousands grouping.
pub static JA_JP: NumberLocale = NumberLocale {
    id: "ja-JP",
    decimal_separator: crate::Locale::en_us().decimal_sep,
    thousands_separator: Some(crate::Locale::en_us().thousands_sep),
};

/// Chinese (China) commonly uses `.` for decimals and `,` for thousands grouping.
pub static ZH_CN: NumberLocale = NumberLocale {
    id: "zh-CN",
    decimal_separator: crate::Locale::en_us().decimal_sep,
    thousands_separator: Some(crate::Locale::en_us().thousands_sep),
};

/// Chinese (Taiwan) commonly uses `.` for decimals and `,` for thousands grouping.
pub static ZH_TW: NumberLocale = NumberLocale {
    id: "zh-TW",
    decimal_separator: crate::Locale::en_us().decimal_sep,
    thousands_separator: Some(crate::Locale::en_us().thousands_sep),
};

/// Chinese (Hong Kong) commonly uses `.` for decimals and `,` for thousands grouping.
pub static ZH_HK: NumberLocale = NumberLocale {
    id: "zh-HK",
    decimal_separator: crate::Locale::en_us().decimal_sep,
    thousands_separator: Some(crate::Locale::en_us().thousands_sep),
};

/// Chinese (Singapore) commonly uses `.` for decimals and `,` for thousands grouping.
pub static ZH_SG: NumberLocale = NumberLocale {
    id: "zh-SG",
    decimal_separator: crate::Locale::en_us().decimal_sep,
    thousands_separator: Some(crate::Locale::en_us().thousands_sep),
};

/// Chinese (Macau) commonly uses `.` for decimals and `,` for thousands grouping.
pub static ZH_MO: NumberLocale = NumberLocale {
    id: "zh-MO",
    decimal_separator: crate::Locale::en_us().decimal_sep,
    thousands_separator: Some(crate::Locale::en_us().thousands_sep),
};

/// Korean (Korea) commonly uses `.` for decimals and `,` for thousands grouping.
pub static KO_KR: NumberLocale = NumberLocale {
    id: "ko-KR",
    decimal_separator: crate::Locale::en_us().decimal_sep,
    thousands_separator: Some(crate::Locale::en_us().thousands_sep),
};

/// Russian (Russia) commonly uses `,` for decimals and NBSP for thousands grouping.
pub static RU_RU: NumberLocale = NumberLocale {
    id: "ru-RU",
    decimal_separator: crate::Locale::fr_fr().decimal_sep,
    thousands_separator: Some(crate::Locale::fr_fr().thousands_sep),
};

/// Polish (Poland) commonly uses `,` for decimals and NBSP for thousands grouping.
pub static PL_PL: NumberLocale = NumberLocale {
    id: "pl-PL",
    decimal_separator: crate::Locale::fr_fr().decimal_sep,
    thousands_separator: Some(crate::Locale::fr_fr().thousands_sep),
};

/// Swedish (Sweden) commonly uses `,` for decimals and NBSP for thousands grouping.
pub static SV_SE: NumberLocale = NumberLocale {
    id: "sv-SE",
    decimal_separator: crate::Locale::fr_fr().decimal_sep,
    thousands_separator: Some(crate::Locale::fr_fr().thousands_sep),
};

/// Norwegian Bokmål (Norway) commonly uses `,` for decimals and NBSP for thousands grouping.
pub static NB_NO: NumberLocale = NumberLocale {
    id: "nb-NO",
    decimal_separator: crate::Locale::fr_fr().decimal_sep,
    thousands_separator: Some(crate::Locale::fr_fr().thousands_sep),
};

/// Danish (Denmark) commonly uses `,` for decimals and `.` for thousands grouping.
pub static DA_DK: NumberLocale = NumberLocale {
    id: "da-DK",
    decimal_separator: crate::Locale::es_es().decimal_sep,
    thousands_separator: Some(crate::Locale::es_es().thousands_sep),
};

/// Turkish (Türkiye) commonly uses `,` for decimals and `.` for thousands grouping.
pub static TR_TR: NumberLocale = NumberLocale {
    id: "tr-TR",
    decimal_separator: crate::Locale::es_es().decimal_sep,
    thousands_separator: Some(crate::Locale::es_es().thousands_sep),
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
    // This intentionally supports a small set of locales the formatter ships with (plus a
    // best-effort language/region fallback for common variants).
    let mut key = String::new();
    let _ = key.try_reserve(trimmed.len());
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

    // Parse BCP-47 tags and variants such as `de-CH-1996` or `fr-Latn-FR-u-nu-latn`. We only care
    // about the language + optional region, ignoring script/variants/extensions.
    let mut parts = key.split('-').filter(|p| !p.is_empty());
    let lang = parts.next()?;
    let mut next = parts.next();
    // Optional script subtag (4 alpha characters) comes before the region.
    if next.is_some_and(|p| p.len() == 4 && p.chars().all(|c| c.is_ascii_alphabetic())) {
        next = parts.next();
    }

    let region = next.filter(|p| {
        (p.len() == 2 && p.chars().all(|c| c.is_ascii_alphabetic()))
            || (p.len() == 3 && p.chars().all(|c| c.is_ascii_digit()))
    });

    match lang {
        "en" => match region {
            Some("gb") | Some("uk") => Some("en-GB"),
            _ => Some("en-US"),
        },
        "de" => match region {
            Some("ch") => Some("de-CH"),
            _ => Some("de-DE"),
        },
        "fr" => match region {
            Some("ch") => Some("fr-CH"),
            _ => Some("fr-FR"),
        },
        "es" => match region {
            Some("mx") => Some("es-MX"),
            _ => Some("es-ES"),
        },
        "it" => match region {
            Some("ch") => Some("it-CH"),
            _ => Some("it-IT"),
        },
        "pt" => match region {
            Some("br") => Some("pt-BR"),
            _ => Some("pt-PT"),
        },
        "nl" => match region {
            Some("be") => Some("nl-BE"),
            _ => Some("nl-NL"),
        },
        "ja" => Some("ja-JP"),
        "zh" => match region {
            Some("tw") => Some("zh-TW"),
            Some("hk") => Some("zh-HK"),
            Some("sg") => Some("zh-SG"),
            Some("mo") => Some("zh-MO"),
            _ => Some("zh-CN"),
        },
        "ko" => Some("ko-KR"),
        "ru" => Some("ru-RU"),
        "pl" => Some("pl-PL"),
        "sv" => Some("sv-SE"),
        "nb" | "no" | "nn" => Some("nb-NO"),
        "da" => Some("da-DK"),
        "tr" => Some("tr-TR"),
        _ => None,
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
        "pt-PT" => Some(&PT_PT),
        "pt-BR" => Some(&PT_BR),
        "nl-NL" => Some(&NL_NL),
        "nl-BE" => Some(&NL_BE),
        "ja-JP" => Some(&JA_JP),
        "zh-CN" => Some(&ZH_CN),
        "zh-TW" => Some(&ZH_TW),
        "zh-HK" => Some(&ZH_HK),
        "zh-SG" => Some(&ZH_SG),
        "zh-MO" => Some(&ZH_MO),
        "ko-KR" => Some(&KO_KR),
        "ru-RU" => Some(&RU_RU),
        "pl-PL" => Some(&PL_PL),
        "sv-SE" => Some(&SV_SE),
        "nb-NO" => Some(&NB_NO),
        "da-DK" => Some(&DA_DK),
        "tr-TR" => Some(&TR_TR),
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

    let mut out = String::new();
    let _ = out.try_reserve(len + len / 3);
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
        assert_eq!(normalize_locale_id("en-GB-oed"), Some("en-GB"));
        assert_eq!(normalize_locale_id("de"), Some("de-DE"));
        assert_eq!(normalize_locale_id("de-AT"), Some("de-DE"));
        assert_eq!(normalize_locale_id("de_ch"), Some("de-CH"));
        assert_eq!(normalize_locale_id("de-CH-1996"), Some("de-CH"));
        assert_eq!(normalize_locale_id("fr_fr"), Some("fr-FR"));
        assert_eq!(normalize_locale_id("fr-CA"), Some("fr-FR"));
        assert_eq!(normalize_locale_id("fr_CH"), Some("fr-CH"));
        assert_eq!(normalize_locale_id("fr-CH-1996"), Some("fr-CH"));
        assert_eq!(normalize_locale_id("es-ES"), Some("es-ES"));
        assert_eq!(normalize_locale_id("es-MX"), Some("es-MX"));
        assert_eq!(normalize_locale_id("es-AR"), Some("es-ES"));
        assert_eq!(normalize_locale_id("pt-PT"), Some("pt-PT"));
        assert_eq!(normalize_locale_id("pt_BR"), Some("pt-BR"));
        assert_eq!(normalize_locale_id("pt-AO"), Some("pt-PT"));
        assert_eq!(normalize_locale_id("nl"), Some("nl-NL"));
        assert_eq!(normalize_locale_id("nl_BE"), Some("nl-BE"));
        assert_eq!(normalize_locale_id("it_it"), Some("it-IT"));
        assert_eq!(normalize_locale_id("it-CH"), Some("it-CH"));
        assert_eq!(normalize_locale_id("ja-JP"), Some("ja-JP"));
        assert_eq!(normalize_locale_id("zh-CN"), Some("zh-CN"));
        assert_eq!(normalize_locale_id("zh-Hant-TW"), Some("zh-TW"));
        assert_eq!(normalize_locale_id("zh-HK"), Some("zh-HK"));
        assert_eq!(normalize_locale_id("zh-SG"), Some("zh-SG"));
        assert_eq!(normalize_locale_id("zh-MO"), Some("zh-MO"));
        assert_eq!(normalize_locale_id("ko-KR"), Some("ko-KR"));
        assert_eq!(normalize_locale_id("ru-RU"), Some("ru-RU"));
        assert_eq!(normalize_locale_id("pl-PL"), Some("pl-PL"));
        assert_eq!(normalize_locale_id("sv-SE"), Some("sv-SE"));
        assert_eq!(normalize_locale_id("nb-NO"), Some("nb-NO"));
        assert_eq!(normalize_locale_id("no-NO"), Some("nb-NO"));
        assert_eq!(normalize_locale_id("da-DK"), Some("da-DK"));
        assert_eq!(normalize_locale_id("tr-TR"), Some("tr-TR"));
        assert_eq!(normalize_locale_id("fr-FR-u-nu-latn"), Some("fr-FR"));
        assert_eq!(normalize_locale_id("fr-Latn-FR-u-nu-latn"), Some("fr-FR"));
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
        assert_eq!(JA_JP.decimal_separator, en.decimal_sep);
        assert_eq!(JA_JP.thousands_separator, Some(en.thousands_sep));
        assert_eq!(ZH_CN.decimal_separator, en.decimal_sep);
        assert_eq!(ZH_CN.thousands_separator, Some(en.thousands_sep));
        assert_eq!(ZH_TW.decimal_separator, en.decimal_sep);
        assert_eq!(ZH_TW.thousands_separator, Some(en.thousands_sep));
        assert_eq!(ZH_HK.decimal_separator, en.decimal_sep);
        assert_eq!(ZH_HK.thousands_separator, Some(en.thousands_sep));
        assert_eq!(ZH_SG.decimal_separator, en.decimal_sep);
        assert_eq!(ZH_SG.thousands_separator, Some(en.thousands_sep));
        assert_eq!(ZH_MO.decimal_separator, en.decimal_sep);
        assert_eq!(ZH_MO.thousands_separator, Some(en.thousands_sep));
        assert_eq!(KO_KR.decimal_separator, en.decimal_sep);
        assert_eq!(KO_KR.thousands_separator, Some(en.thousands_sep));

        let de = crate::Locale::de_de();
        assert_eq!(DE_DE.decimal_separator, de.decimal_sep);
        assert_eq!(DE_DE.thousands_separator, Some(de.thousands_sep));

        let fr = crate::Locale::fr_fr();
        assert_eq!(FR_FR.decimal_separator, fr.decimal_sep);
        assert_eq!(FR_FR.thousands_separator, Some(fr.thousands_sep));
        assert_eq!(RU_RU.decimal_separator, fr.decimal_sep);
        assert_eq!(RU_RU.thousands_separator, Some(fr.thousands_sep));
        assert_eq!(PL_PL.decimal_separator, fr.decimal_sep);
        assert_eq!(PL_PL.thousands_separator, Some(fr.thousands_sep));
        assert_eq!(SV_SE.decimal_separator, fr.decimal_sep);
        assert_eq!(SV_SE.thousands_separator, Some(fr.thousands_sep));
        assert_eq!(NB_NO.decimal_separator, fr.decimal_sep);
        assert_eq!(NB_NO.thousands_separator, Some(fr.thousands_sep));

        let es = crate::Locale::es_es();
        assert_eq!(ES_ES.decimal_separator, es.decimal_sep);
        assert_eq!(ES_ES.thousands_separator, Some(es.thousands_sep));
        assert_eq!(PT_PT.decimal_separator, es.decimal_sep);
        assert_eq!(PT_PT.thousands_separator, Some(es.thousands_sep));
        assert_eq!(PT_BR.decimal_separator, es.decimal_sep);
        assert_eq!(PT_BR.thousands_separator, Some(es.thousands_sep));
        assert_eq!(NL_NL.decimal_separator, es.decimal_sep);
        assert_eq!(NL_NL.thousands_separator, Some(es.thousands_sep));
        assert_eq!(NL_BE.decimal_separator, es.decimal_sep);
        assert_eq!(NL_BE.thousands_separator, Some(es.thousands_sep));
        assert_eq!(DA_DK.decimal_separator, es.decimal_sep);
        assert_eq!(DA_DK.thousands_separator, Some(es.thousands_sep));
        assert_eq!(TR_TR.decimal_separator, es.decimal_sep);
        assert_eq!(TR_TR.thousands_separator, Some(es.thousands_sep));

        let it = crate::Locale::it_it();
        assert_eq!(IT_IT.decimal_separator, it.decimal_sep);
        assert_eq!(IT_IT.thousands_separator, Some(it.thousands_sep));
    }
}
