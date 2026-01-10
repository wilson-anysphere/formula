/// Locale configuration for parsing and rendering formulas.
///
/// The formula engine keeps function identifiers in a canonical (Excel
/// "English") form internally. Locales define how to translate between the
/// canonical and localized function names, plus punctuation differences.
#[derive(Debug)]
pub struct FormulaLocale {
    pub id: &'static str,
    /// Decimal separator used in numeric literals (e.g. `1,23` in `de-DE`).
    pub decimal_separator: char,
    /// Argument separator used in function calls.
    pub argument_separator: char,
    /// Thousands separator for numeric literals. Excel's formula language does
    /// not consistently allow grouping separators in all locales; we only use
    /// this opportunistically when it does not conflict with the argument
    /// separator.
    pub thousands_separator: Option<char>,
    /// `true` when this locale is right-to-left in the UI (formula language is
    /// still left-to-right).
    pub is_rtl: bool,
    /// Mapping table between canonical (English) names and localized names.
    ///
    /// The first element of each pair is the canonical function name.
    pub function_name_map: &'static [(&'static str, &'static str)],
}

impl FormulaLocale {
    /// Translate an input function name into canonical form.
    pub fn canonical_function_name(&self, name: &str) -> String {
        // Excel function names are case-insensitive; store canonical names as
        // uppercase ASCII.
        let upper = name.to_ascii_uppercase();
        for (canonical, localized) in self.function_name_map {
            if localized.eq_ignore_ascii_case(&upper) {
                return (*canonical).to_string();
            }
        }
        upper
    }

    /// Translate a canonical function name into its localized display form.
    pub fn localized_function_name(&self, canonical: &str) -> String {
        let upper = canonical.to_ascii_uppercase();
        for (canon, localized) in self.function_name_map {
            if canon.eq_ignore_ascii_case(&upper) {
                return (*localized).to_string();
            }
        }
        upper
    }

    /// Returns a thousands separator that is safe to consume inside numeric
    /// literals for this locale.
    ///
    /// In `en-US` the thousands separator is `,` which would be ambiguous with
    /// the argument separator. We therefore disable thousands separators in the
    /// parser when they collide with the argument separator.
    pub fn numeric_thousands_separator(&self) -> Option<char> {
        match self.thousands_separator {
            Some(sep) if sep != self.argument_separator => Some(sep),
            _ => None,
        }
    }
}

const DE_DE_FUNCTIONS: &[(&str, &str)] = &[
    ("SUM", "SUMME"),
    ("AVERAGE", "MITTELWERT"),
    ("MIN", "MIN"),
    ("MAX", "MAX"),
];

/// English (United States) uses `.` for decimals and `,` for arguments.
pub static EN_US: FormulaLocale = FormulaLocale {
    id: "en-US",
    decimal_separator: '.',
    argument_separator: ',',
    thousands_separator: Some(','),
    is_rtl: false,
    function_name_map: &[],
};

/// German (Germany) matches Excel's common localization:
/// - `,` decimal separator
/// - `;` argument separator
/// - Localized function names (e.g. `SUMME`)
pub static DE_DE: FormulaLocale = FormulaLocale {
    id: "de-DE",
    decimal_separator: ',',
    argument_separator: ';',
    thousands_separator: Some('.'),
    is_rtl: false,
    function_name_map: DE_DE_FUNCTIONS,
};

pub fn get_locale(id: &str) -> Option<&'static FormulaLocale> {
    match id {
        "en-US" => Some(&EN_US),
        "de-DE" => Some(&DE_DE),
        _ => None,
    }
}

