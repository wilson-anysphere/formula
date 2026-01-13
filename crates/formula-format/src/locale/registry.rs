/// Minimal locale information used by [`format_number`] for "plain number" rendering.
///
/// This is intentionally lightweight and independent from the richer [`crate::Locale`] used by the
/// full Excel format-code engine. `NumberLocale` only carries the decimal/thousands separators
/// needed by the UI/formula bar when no explicit number format code is available.
#[derive(Debug)]
pub struct NumberLocale {
    pub id: &'static str,
    pub decimal_separator: char,
    pub thousands_separator: Option<char>,
}

pub static EN_US: NumberLocale = NumberLocale {
    id: "en-US",
    decimal_separator: '.',
    thousands_separator: Some(','),
};

/// British English uses the same separators as `en-US`.
pub static EN_GB: NumberLocale = NumberLocale {
    id: "en-GB",
    decimal_separator: '.',
    thousands_separator: Some(','),
};

pub static DE_DE: NumberLocale = NumberLocale {
    id: "de-DE",
    decimal_separator: ',',
    thousands_separator: Some('.'),
};

/// French (France).
///
/// We use U+00A0 NO-BREAK SPACE as the thousands separator. Some environments prefer U+202F
/// NARROW NO-BREAK SPACE; if we ever need to distinguish, we can add a separate entry, but U+00A0
/// is widely supported and matches `crate::Locale::fr_fr()`.
pub static FR_FR: NumberLocale = NumberLocale {
    id: "fr-FR",
    decimal_separator: ',',
    thousands_separator: Some('\u{00A0}'),
};

pub static ES_ES: NumberLocale = NumberLocale {
    id: "es-ES",
    decimal_separator: ',',
    thousands_separator: Some('.'),
};

pub static IT_IT: NumberLocale = NumberLocale {
    id: "it-IT",
    decimal_separator: ',',
    thousands_separator: Some('.'),
};

pub fn get_locale(id: &str) -> Option<&'static NumberLocale> {
    match id {
        "en-US" => Some(&EN_US),
        "en-GB" => Some(&EN_GB),
        "de-DE" => Some(&DE_DE),
        "fr-FR" => Some(&FR_FR),
        "es-ES" => Some(&ES_ES),
        "it-IT" => Some(&IT_IT),
        _ => None,
    }
}

/// Format a plain number using locale-specific separators.
///
/// This is intentionally not a full Excel number-format implementation yet; it
/// covers the most visible internationalization differences (thousands/decimal
/// separators) needed by the UI and formula bar.
pub fn format_number(value: f64, locale: &NumberLocale) -> String {
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
