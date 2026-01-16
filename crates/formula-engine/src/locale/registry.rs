use crate::value::{casefold, with_casefolded_key};
use crate::value::ErrorKind;
use crate::LocaleConfig;

use std::collections::HashMap;
use std::sync::OnceLock;

#[derive(Debug)]
struct FunctionTranslationMaps {
    canon_to_loc: HashMap<String, &'static str>,
    loc_to_canon: HashMap<String, &'static str>,
}

/// Translation table for Excel function identifiers.
///
/// Data is stored outside the Rust source in simple TSV files under `src/locale/data/`.
/// See `src/locale/data/README.md` for the TSV format and the generator scripts that keep
/// these tables complete and normalized.
/// This keeps the code small and provides a straightforward path to scale to hundreds
/// of translated functions by generating the TSV from upstream sources (e.g. Office
/// function translation lists) without hand-editing Rust tables.
#[derive(Debug)]
struct FunctionTranslations {
    data_tsv: &'static str,
    maps: OnceLock<FunctionTranslationMaps>,
}

impl FunctionTranslations {
    const fn new(data_tsv: &'static str) -> Self {
        Self {
            data_tsv,
            maps: OnceLock::new(),
        }
    }

    fn maps(&self) -> &FunctionTranslationMaps {
        self.maps.get_or_init(|| {
            let mut canon_to_loc = HashMap::new();
            let mut loc_to_canon = HashMap::new();
            // Track the exact line that introduced each key so we can produce actionable
            // diagnostics if the TSV contains duplicate entries.
            let mut canon_line: HashMap<String, (usize, &'static str)> = HashMap::new();
            let mut loc_line: HashMap<String, (usize, &'static str)> = HashMap::new();

            for (idx, raw_line) in self.data_tsv.lines().enumerate() {
                let line_no = idx + 1;
                let line = raw_line.trim();
                if line.is_empty() || line.starts_with('#') {
                    continue;
                }

                let mut parts = raw_line.split('\t');
                let canon = parts.next().unwrap_or("");
                let loc = parts.next().unwrap_or_else(|| {
                    panic!("invalid function translation line (expected TSV) at line {line_no}: {raw_line:?}")
                });
                if parts.next().is_some() {
                    panic!(
                        "invalid function translation line (too many columns) at line {line_no}: {raw_line:?}"
                    );
                }
                let canon = canon.trim();
                let loc = loc.trim();
                if canon.is_empty() || loc.is_empty() {
                    panic!(
                        "invalid function translation line (empty entry) at line {line_no}: {raw_line:?}"
                    );
                }

                let canon_key = casefold(canon);
                let loc_key = casefold(loc);

                if let Some((prev_no, prev_line)) = canon_line.get(&canon_key) {
                    panic!(
                        "duplicate canonical function translation key {canon_key:?}\n  first: line {prev_no}: {prev_line:?}\n  second: line {line_no}: {line:?}"
                    );
                }
                if let Some((prev_no, prev_line)) = loc_line.get(&loc_key) {
                    panic!(
                        "duplicate localized function translation key {loc_key:?}\n  first: line {prev_no}: {prev_line:?}\n  second: line {line_no}: {line:?}"
                    );
                }

                canon_line.insert(canon_key.clone(), (line_no, line));
                loc_line.insert(loc_key.clone(), (line_no, line));

                canon_to_loc.insert(canon_key, loc);
                loc_to_canon.insert(loc_key, canon);
            }

            FunctionTranslationMaps {
                canon_to_loc,
                loc_to_canon,
            }
        })
    }

    fn localized_to_canonical(&self, localized_key: &str) -> Option<&'static str> {
        self.maps().loc_to_canon.get(localized_key).copied()
    }

    fn canonical_to_localized(&self, canonical_key: &str) -> Option<&'static str> {
        self.maps().canon_to_loc.get(canonical_key).copied()
    }
}

static EMPTY_FUNCTIONS: FunctionTranslations = FunctionTranslations::new("");
// Locale TSVs live in `src/locale/data/`. See `src/locale/data/README.md` for
// contributor docs (format, completeness requirements, and generators).
static DE_DE_FUNCTIONS: FunctionTranslations =
    FunctionTranslations::new(include_str!("data/de-DE.tsv"));
static FR_FR_FUNCTIONS: FunctionTranslations =
    FunctionTranslations::new(include_str!("data/fr-FR.tsv"));
static ES_ES_FUNCTIONS: FunctionTranslations =
    FunctionTranslations::new(include_str!("data/es-ES.tsv"));

#[derive(Debug)]
struct ErrorTranslationMaps {
    /// Case-folded canonical error literal -> preferred localized spelling.
    canon_to_loc: HashMap<String, &'static str>,
    /// Case-folded localized error literal -> canonical error literal.
    loc_to_canon: HashMap<String, &'static str>,
}

/// Translation table for Excel error literals (e.g. `#VALUE!`).
///
/// Data is stored outside the Rust source in TSV files under `src/locale/data/`.
/// See `src/locale/data/README.md` for the TSV format and the generator scripts.
///
/// Unlike function translations, error translation TSVs can contain *multiple* localized spellings
/// for the same canonical error literal to support Excel-compatible alias spellings (e.g. Spanish
/// inverted punctuation variants).
#[derive(Debug)]
struct ErrorTranslations {
    data_tsv: &'static str,
    maps: OnceLock<ErrorTranslationMaps>,
}

impl ErrorTranslations {
    const fn new(data_tsv: &'static str) -> Self {
        Self {
            data_tsv,
            maps: OnceLock::new(),
        }
    }

    fn maps(&self) -> &ErrorTranslationMaps {
        self.maps.get_or_init(|| {
            let mut canon_to_loc = HashMap::new();
            let mut loc_to_canon = HashMap::new();
            // Track the exact line that introduced each localized key so we can produce actionable
            // diagnostics if the TSV contains duplicates (which would make parsing ambiguous).
            let mut loc_line: HashMap<String, (usize, &'static str)> = HashMap::new();

            for (idx, raw_line) in self.data_tsv.lines().enumerate() {
                let line_no = idx + 1;
                let trimmed = raw_line.trim();
                if trimmed.is_empty() {
                    continue;
                }
                // Error literals themselves start with `#`, so treat comments as `#` followed by
                // whitespace (or a bare `#`) rather than treating all `#` lines as comments.
                let is_comment = trimmed == "#"
                    || (trimmed.starts_with('#')
                        && trimmed
                            .chars()
                            .nth(1)
                            .is_some_and(|c| c.is_whitespace()));
                if is_comment {
                    continue;
                }

                let mut parts = raw_line.split('\t');
                let canon = parts.next().unwrap_or("");
                let loc = parts.next().unwrap_or_else(|| {
                    panic!(
                        "invalid error translation line (expected TSV) at line {line_no}: {raw_line:?}"
                    )
                });
                if parts.next().is_some() {
                    panic!(
                        "invalid error translation line (too many columns) at line {line_no}: {raw_line:?}"
                    );
                }

                let canon = canon.trim();
                let loc = loc.trim();
                if canon.is_empty() || loc.is_empty() {
                    panic!(
                        "invalid error translation line (empty entry) at line {line_no}: {raw_line:?}"
                    );
                }
                if !canon.starts_with('#') || !loc.starts_with('#') {
                    panic!(
                        "invalid error translation line (expected error literals to start with '#') at line {line_no}: {raw_line:?}"
                    );
                }

                let canon_key = casefold(canon);
                let loc_key = casefold(loc);

                if let Some((prev_no, prev_line)) = loc_line.get(&loc_key) {
                    panic!(
                        "duplicate localized error translation key {loc_key:?}\n  first: line {prev_no}: {prev_line:?}\n  second: line {line_no}: {raw_line:?}"
                    );
                }
                loc_line.insert(loc_key.clone(), (line_no, raw_line));

                // For canonical->localized, the first spelling wins and becomes the preferred
                // localized display form.
                canon_to_loc.entry(canon_key).or_insert(loc);
                // For localized->canonical, accept all localized spellings.
                loc_to_canon.insert(loc_key, canon);
            }

            ErrorTranslationMaps {
                canon_to_loc,
                loc_to_canon,
            }
        })
    }

    fn localized_to_canonical(&self, localized: &str) -> Option<&'static str> {
        with_casefolded_key(localized, |key| self.maps().loc_to_canon.get(key).copied())
    }

    fn canonical_to_localized(&self, canonical: &str) -> Option<&'static str> {
        with_casefolded_key(canonical, |key| self.maps().canon_to_loc.get(key).copied())
    }
}

static EMPTY_ERRORS: ErrorTranslations = ErrorTranslations::new("");
// Locale error TSVs live in `src/locale/data/`. See `src/locale/data/README.md` for
// contributor docs (format, completeness requirements, and generators).
static DE_DE_ERRORS: ErrorTranslations =
    ErrorTranslations::new(include_str!("data/de-DE.errors.tsv"));
static FR_FR_ERRORS: ErrorTranslations =
    ErrorTranslations::new(include_str!("data/fr-FR.errors.tsv"));
static ES_ES_ERRORS: ErrorTranslations =
    ErrorTranslations::new(include_str!("data/es-ES.errors.tsv"));

/// Locale configuration for parsing and rendering formulas.
///
/// The formula engine keeps function identifiers in a canonical (Excel "English") form internally.
/// Locales define how to translate between the canonical and localized function names, plus
/// punctuation differences.
#[derive(Debug)]
pub struct FormulaLocale {
    pub id: &'static str,
    /// Separators and punctuation used by the lexer/parser for this locale.
    pub config: LocaleConfig,
    /// `true` when this locale is right-to-left in the UI (formula language is still left-to-right).
    pub is_rtl: bool,
    /// Localized boolean literals (Excel keywords).
    pub boolean_true: &'static str,
    pub boolean_false: &'static str,
    errors: &'static ErrorTranslations,
    functions: &'static FunctionTranslations,
}

impl FormulaLocale {
    /// Translate an input function name into canonical form.
    pub fn canonical_function_name(&self, name: &str) -> String {
        let (has_prefix, base) = split_xlfn_prefix(name);
        let mut out = String::new();
        if has_prefix {
            out.push_str("_xlfn.");
        }
        with_casefolded_key(base, |folded| {
            let mapped = self
                .functions
                .localized_to_canonical(folded)
                .unwrap_or(folded);
            out.push_str(mapped);
        });
        out
    }

    /// Translate a canonical function name into its localized display form.
    pub fn localized_function_name(&self, canonical: &str) -> String {
        let (has_prefix, base) = split_xlfn_prefix(canonical);
        let mut out = String::new();
        if has_prefix {
            out.push_str("_xlfn.");
        }
        with_casefolded_key(base, |folded| {
            let mapped = self
                .functions
                .canonical_to_localized(folded)
                .unwrap_or(folded);
            out.push_str(mapped);
        });
        out
    }

    pub fn canonical_boolean_literal(&self, ident: &str) -> Option<bool> {
        // Excel treats keywords case-insensitively across Unicode. Keep boolean keyword matching
        // consistent with function translation keys by using the same Unicode-aware case folding.
        with_casefolded_key(ident, |folded| {
            if folded == self.boolean_true {
                Some(true)
            } else if folded == self.boolean_false {
                Some(false)
            } else {
                None
            }
        })
    }

    pub fn localized_boolean_literal(&self, value: bool) -> &'static str {
        if value {
            self.boolean_true
        } else {
            self.boolean_false
        }
    }

    pub fn canonical_error_literal(&self, localized: &str) -> Option<&'static str> {
        if let Some(canonical) = self.errors.localized_to_canonical(localized) {
            return Some(normalize_canonical_error_literal(canonical));
        }

        // Excel also accepts canonical error literals even in localized formulas (and `#N/A!` as an
        // alias for `#N/A`).
        ErrorKind::from_code(localized).map(|kind| kind.as_code())
    }

    pub fn localized_error_literal(&self, canonical: &str) -> Option<&'static str> {
        // Normalize canonical spellings (`#N/A!` -> `#N/A`) so they localize consistently.
        let canonical = normalize_canonical_error_literal(canonical);
        // Prefer the locale-specific translation when present; otherwise fall back to the canonical
        // spelling.
        self.errors
            .canonical_to_localized(canonical)
            .or_else(|| ErrorKind::from_code(canonical).map(|kind| kind.as_code()))
    }
}

fn normalize_canonical_error_literal(code: &str) -> &str {
    if code.eq_ignore_ascii_case("#N/A!") {
        "#N/A"
    } else {
        code
    }
}

fn split_xlfn_prefix(name: &str) -> (bool, &str) {
    const PREFIX: &str = "_xlfn.";
    let Some(prefix) = name.get(..PREFIX.len()) else {
        return (false, name);
    };
    if prefix.eq_ignore_ascii_case(PREFIX) {
        (true, &name[PREFIX.len()..])
    } else {
        (false, name)
    }
}

/// English (United States) uses `.` for decimals and `,` for arguments.
pub static EN_US: FormulaLocale = FormulaLocale {
    id: "en-US",
    config: LocaleConfig::en_us(),
    is_rtl: false,
    boolean_true: "TRUE",
    boolean_false: "FALSE",
    errors: &EMPTY_ERRORS,
    functions: &EMPTY_FUNCTIONS,
};

/// Japanese (Japan).
///
/// This is a minimal locale registration to allow selecting the workbook locale id (e.g. from
/// `formula-wasm set_locale_id`) so codepage-driven DBCS semantics can be enabled by higher layers.
///
/// Punctuation/function names are currently treated as en-US.
pub static JA_JP: FormulaLocale = FormulaLocale {
    id: "ja-JP",
    config: LocaleConfig::en_us(),
    is_rtl: false,
    boolean_true: "TRUE",
    boolean_false: "FALSE",
    errors: &EMPTY_ERRORS,
    functions: &EMPTY_FUNCTIONS,
};

/// Chinese (Simplified, China).
///
/// Minimal locale registration; punctuation/function names are currently treated as en-US.
pub static ZH_CN: FormulaLocale = FormulaLocale {
    id: "zh-CN",
    config: LocaleConfig::en_us(),
    is_rtl: false,
    boolean_true: "TRUE",
    boolean_false: "FALSE",
    errors: &EMPTY_ERRORS,
    functions: &EMPTY_FUNCTIONS,
};

/// Korean (Korea).
///
/// Minimal locale registration; punctuation/function names are currently treated as en-US.
pub static KO_KR: FormulaLocale = FormulaLocale {
    id: "ko-KR",
    config: LocaleConfig::en_us(),
    is_rtl: false,
    boolean_true: "TRUE",
    boolean_false: "FALSE",
    errors: &EMPTY_ERRORS,
    functions: &EMPTY_FUNCTIONS,
};

/// Chinese (Traditional, Taiwan).
///
/// Minimal locale registration; punctuation/function names are currently treated as en-US.
pub static ZH_TW: FormulaLocale = FormulaLocale {
    id: "zh-TW",
    config: LocaleConfig::en_us(),
    is_rtl: false,
    boolean_true: "TRUE",
    boolean_false: "FALSE",
    errors: &EMPTY_ERRORS,
    functions: &EMPTY_FUNCTIONS,
};

/// German (Germany) matches Excel's common localization:
/// - `,` decimal separator
/// - `;` argument separator
/// - `\` array column separator
/// - Localized function names (e.g. `SUMME`)
pub static DE_DE: FormulaLocale = FormulaLocale {
    id: "de-DE",
    config: LocaleConfig::de_de(),
    is_rtl: false,
    boolean_true: "WAHR",
    boolean_false: "FALSCH",
    errors: &DE_DE_ERRORS,
    functions: &DE_DE_FUNCTIONS,
};

/// French (France).
pub static FR_FR: FormulaLocale = FormulaLocale {
    id: "fr-FR",
    config: LocaleConfig::fr_fr(),
    is_rtl: false,
    boolean_true: "VRAI",
    boolean_false: "FAUX",
    errors: &FR_FR_ERRORS,
    functions: &FR_FR_FUNCTIONS,
};

/// Spanish (Spain).
pub static ES_ES: FormulaLocale = FormulaLocale {
    id: "es-ES",
    config: LocaleConfig::es_es(),
    is_rtl: false,
    boolean_true: "VERDADERO",
    boolean_false: "FALSO",
    errors: &ES_ES_ERRORS,
    functions: &ES_ES_FUNCTIONS,
};

static ALL_LOCALES: [&FormulaLocale; 8] = [
    &EN_US, &JA_JP, &ZH_CN, &KO_KR, &ZH_TW, &DE_DE, &FR_FR, &ES_ES,
];

/// Enumerate every [`FormulaLocale`] that the engine ships with.
///
/// This iterator is the source of truth for the set of *canonical* locale ids returned by
/// [`get_locale`] (after locale id normalization).
///
/// Notes:
/// - The list includes "minimal" DBCS locales (`ja-JP`, `zh-CN`, `zh-TW`, `ko-KR`). These currently
///   reuse `en-US` punctuation/function names, but are still distinct locale registrations so higher
///   layers can enable workbook semantics that depend on the locale id (e.g. codepage-driven DBCS
///   behavior).
/// - Callers should not rely on the iteration order. Sort by [`FormulaLocale::id`] when producing
///   user-visible output.
pub fn iter_locales() -> impl Iterator<Item = &'static FormulaLocale> {
    ALL_LOCALES.iter().copied()
}

pub fn get_locale(id: &str) -> Option<&'static FormulaLocale> {
    let normalized = super::normalize_locale_id(id)?;
    ALL_LOCALES
        .iter()
        .copied()
        .find(|locale| locale.id == normalized)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::any::Any;
    use std::panic::AssertUnwindSafe;

    fn panic_message(err: &(dyn Any + Send)) -> String {
        if let Some(msg) = err.downcast_ref::<&str>() {
            (*msg).to_string()
        } else if let Some(msg) = err.downcast_ref::<String>() {
            msg.clone()
        } else {
            "<non-string panic>".to_string()
        }
    }

    #[test]
    fn duplicate_canonical_key_panics_with_diagnostics() {
        let translations = FunctionTranslations::new(
            "\
SUM\tSOMME
SUM\tSUMA
",
        );
        let err = std::panic::catch_unwind(AssertUnwindSafe(|| {
            translations.maps();
        }))
        .expect_err("expected duplicate canonical key to panic");

        let msg = panic_message(&*err);
        assert!(msg.contains("duplicate canonical function translation key"));
        assert!(msg.contains("\"SUM\""));
        assert!(msg.contains("line 1"));
        assert!(msg.contains("line 2"));
        assert!(msg.contains("SUM\\tSOMME"));
        assert!(msg.contains("SUM\\tSUMA"));
    }

    #[test]
    fn duplicate_localized_key_panics_with_diagnostics() {
        let translations = FunctionTranslations::new(
            "\
SUM\tSOMME
AVERAGE\tSOMME
",
        );
        let err = std::panic::catch_unwind(AssertUnwindSafe(|| {
            translations.maps();
        }))
        .expect_err("expected duplicate localized key to panic");

        let msg = panic_message(&*err);
        assert!(msg.contains("duplicate localized function translation key"));
        assert!(msg.contains("\"SOMME\""));
        assert!(msg.contains("line 1"));
        assert!(msg.contains("line 2"));
        assert!(msg.contains("SUM\\tSOMME"));
        assert!(msg.contains("AVERAGE\\tSOMME"));
    }

    #[test]
    fn function_translation_rejects_extra_tsv_columns() {
        let translations = FunctionTranslations::new("SUM\tSOMME\tEXTRA\n");
        let err = std::panic::catch_unwind(AssertUnwindSafe(|| {
            translations.maps();
        }))
        .expect_err("expected extra TSV columns to panic");
        let msg = panic_message(&*err);
        assert!(msg.contains("too many columns"));
        assert!(msg.contains("line 1"));
        assert!(msg.contains("SUM\\tSOMME\\tEXTRA"));
    }

    #[test]
    fn function_translation_rejects_trailing_empty_column() {
        let translations = FunctionTranslations::new("SUM\tSOMME\t\n");
        let err = std::panic::catch_unwind(AssertUnwindSafe(|| {
            translations.maps();
        }))
        .expect_err("expected trailing empty TSV column to panic");
        let msg = panic_message(&*err);
        assert!(msg.contains("too many columns"));
        assert!(msg.contains("line 1"));
        assert!(msg.contains("SUM\\tSOMME\\t"));
    }

    #[test]
    fn duplicate_canonical_error_key_is_allowed_and_prefers_first_spelling() {
        let translations = ErrorTranslations::new(
            "\
# Canonical\tLocalized
#VALUE!\t#WERT!
#VALUE!\t#VALEUR!
",
        );
        // First localized spelling should be preferred for canonical->localized.
        assert_eq!(
            translations.canonical_to_localized("#VALUE!"),
            Some("#WERT!")
        );
        // All localized spellings should be accepted for localized->canonical.
        assert_eq!(
            translations.localized_to_canonical("#WERT!"),
            Some("#VALUE!")
        );
        assert_eq!(
            translations.localized_to_canonical("#VALEUR!"),
            Some("#VALUE!")
        );
    }

    #[test]
    fn duplicate_localized_error_key_panics_with_diagnostics() {
        let translations = ErrorTranslations::new(
            "\
#VALUE!\t#WERT!
#REF!\t#WERT!
",
        );
        let err = std::panic::catch_unwind(AssertUnwindSafe(|| {
            translations.maps();
        }))
        .expect_err("expected duplicate localized key to panic");

        let msg = panic_message(&*err);
        assert!(msg.contains("duplicate localized error translation key"));
        assert!(msg.contains("\"#WERT!\""));
        assert!(msg.contains("line 1"));
        assert!(msg.contains("line 2"));
        assert!(msg.contains("#VALUE!\\t#WERT!"));
        assert!(msg.contains("#REF!\\t#WERT!"));
    }

    #[test]
    fn error_translation_rejects_extra_tsv_columns() {
        let translations = ErrorTranslations::new("#VALUE!\t#WERT!\tEXTRA\n");
        let err = std::panic::catch_unwind(AssertUnwindSafe(|| {
            translations.maps();
        }))
        .expect_err("expected extra TSV columns to panic");
        let msg = panic_message(&*err);
        assert!(msg.contains("too many columns"));
        assert!(msg.contains("line 1"));
        assert!(msg.contains("#VALUE!\\t#WERT!\\tEXTRA"));
    }

    #[test]
    fn error_translation_rejects_trailing_empty_column() {
        let translations = ErrorTranslations::new("#VALUE!\t#WERT!\t\n");
        let err = std::panic::catch_unwind(AssertUnwindSafe(|| {
            translations.maps();
        }))
        .expect_err("expected trailing empty TSV columns to panic");
        let msg = panic_message(&*err);
        assert!(msg.contains("too many columns"));
        assert!(msg.contains("line 1"));
        assert!(msg.contains("#VALUE!\\t#WERT!\\t"));
    }

    #[test]
    fn error_translation_rejects_non_error_literals() {
        let translations = ErrorTranslations::new("VALUE\t#WERT!\n");
        let err = std::panic::catch_unwind(AssertUnwindSafe(|| {
            translations.maps();
        }))
        .expect_err("expected non-error literal columns to panic");
        let msg = panic_message(&*err);
        assert!(msg.contains("expected error literals to start with '#'"));
        assert!(msg.contains("line 1"));
        assert!(msg.contains("VALUE\\t#WERT!"));
    }

    #[test]
    fn canonical_boolean_literal_uses_unicode_case_folding() {
        // Function translation keys use Unicode-aware uppercasing for case-insensitive matching.
        // Boolean keyword translation should behave the same so locales with non-ASCII spellings
        // still accept mixed-case input.
        let test_locale = FormulaLocale {
            id: "test",
            config: LocaleConfig::en_us(),
            is_rtl: false,
            boolean_true: "ÄPFEL",
            boolean_false: "ÖL",
            errors: &EMPTY_ERRORS,
            functions: &EMPTY_FUNCTIONS,
        };

        assert_eq!(test_locale.canonical_boolean_literal("äpfel"), Some(true));
        assert_eq!(test_locale.canonical_boolean_literal("Äpfel"), Some(true));
        assert_eq!(test_locale.canonical_boolean_literal("öl"), Some(false));
        assert_eq!(test_locale.canonical_boolean_literal("apfel"), None);
    }
}
