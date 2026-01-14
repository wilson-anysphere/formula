use crate::LocaleConfig;
use crate::value::casefold;

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

                let (canon, loc) = line.split_once('\t').unwrap_or_else(|| {
                    panic!(
                        "invalid function translation line (expected TSV) at line {line_no}: {line:?}"
                    )
                });
                let canon = canon.trim();
                let loc = loc.trim();
                if canon.is_empty() || loc.is_empty() {
                    panic!(
                        "invalid function translation line (empty entry) at line {line_no}: {line:?}"
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
    /// Mapping table between canonical (English) error literals and localized error literals.
    ///
    /// Each entry is `(canonical, localized)`, e.g. `("#VALUE!", "#WERT!")`.
    pub error_literal_map: &'static [(&'static str, &'static str)],
    functions: &'static FunctionTranslations,
}

impl FormulaLocale {
    /// Translate an input function name into canonical form.
    pub fn canonical_function_name(&self, name: &str) -> String {
        let (has_prefix, base) = split_xlfn_prefix(name);
        let folded = casefold_ident(base);

        let mapped = self
            .functions
            .localized_to_canonical(&folded)
            .unwrap_or(folded.as_str());

        let mut out = String::new();
        if has_prefix {
            out.push_str("_xlfn.");
        }
        out.push_str(mapped);
        out
    }

    /// Translate a canonical function name into its localized display form.
    pub fn localized_function_name(&self, canonical: &str) -> String {
        let (has_prefix, base) = split_xlfn_prefix(canonical);
        let folded = casefold_ident(base);

        let mapped = self
            .functions
            .canonical_to_localized(&folded)
            .unwrap_or(folded.as_str());

        let mut out = String::new();
        if has_prefix {
            out.push_str("_xlfn.");
        }
        out.push_str(mapped);
        out
    }

    pub fn canonical_boolean_literal(&self, ident: &str) -> Option<bool> {
        if ident.eq_ignore_ascii_case(self.boolean_true) {
            Some(true)
        } else if ident.eq_ignore_ascii_case(self.boolean_false) {
            Some(false)
        } else {
            None
        }
    }

    pub fn localized_boolean_literal(&self, value: bool) -> &'static str {
        if value {
            self.boolean_true
        } else {
            self.boolean_false
        }
    }

    pub fn canonical_error_literal(&self, localized: &str) -> Option<&'static str> {
        if localized.is_ascii() {
            for (canonical, loc) in self.error_literal_map {
                if loc.eq_ignore_ascii_case(localized) {
                    return Some(*canonical);
                }
            }
            return None;
        }

        // Use Unicode-aware case folding for localized spellings with non-ASCII characters.
        // This matches the behavior used for other case-insensitive comparisons in the engine
        // (e.g. criteria matching).
        let localized_fold = crate::value::casefold(localized);
        for (canonical, loc) in self.error_literal_map {
            if crate::value::casefold(loc) == localized_fold {
                return Some(*canonical);
            }
        }
        None
    }

    pub fn localized_error_literal(&self, canonical: &str) -> Option<&'static str> {
        if canonical.is_ascii() {
            for (canon, localized) in self.error_literal_map {
                if canon.eq_ignore_ascii_case(canonical) {
                    return Some(*localized);
                }
            }
            return None;
        }

        // Canonical spellings are ASCII today, but keep this Unicode-aware for future additions.
        let canonical_fold = crate::value::casefold(canonical);
        for (canon, localized) in self.error_literal_map {
            if crate::value::casefold(canon) == canonical_fold {
                return Some(*localized);
            }
        }
        None
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

fn casefold_ident(ident: &str) -> String {
    // Locale translation needs case-insensitive matching that behaves like Excel.
    // Use Unicode-aware uppercasing (`ß` -> `SS`, `ä` -> `Ä`, ...) for non-ASCII.
    if ident.is_ascii() {
        ident.to_ascii_uppercase()
    } else {
        ident.chars().flat_map(|ch| ch.to_uppercase()).collect()
    }
}

/// English (United States) uses `.` for decimals and `,` for arguments.
pub static EN_US: FormulaLocale = FormulaLocale {
    id: "en-US",
    config: LocaleConfig::en_us(),
    is_rtl: false,
    boolean_true: "TRUE",
    boolean_false: "FALSE",
    error_literal_map: &[],
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
    // For TSV format + generation workflow, see `src/locale/data/README.md`.
    error_literal_map: &[
        ("#VALUE!", "#WERT!"),
        ("#REF!", "#BEZUG!"),
        ("#SPILL!", "#ÜBERLAUF!"),
        ("#GETTING_DATA", "#DATEN_ABRUFEN"),
    ],
    functions: &DE_DE_FUNCTIONS,
};

/// French (France).
pub static FR_FR: FormulaLocale = FormulaLocale {
    id: "fr-FR",
    config: LocaleConfig::fr_fr(),
    is_rtl: false,
    boolean_true: "VRAI",
    boolean_false: "FAUX",
    // For TSV format + generation workflow, see `src/locale/data/README.md`.
    error_literal_map: &[
        ("#VALUE!", "#VALEUR!"),
        ("#NAME?", "#NOM?"),
        ("#GETTING_DATA", "#OBTENTION_DONNEES"),
    ],
    functions: &FR_FR_FUNCTIONS,
};

/// Spanish (Spain).
pub static ES_ES: FormulaLocale = FormulaLocale {
    id: "es-ES",
    config: LocaleConfig::es_es(),
    is_rtl: false,
    boolean_true: "VERDADERO",
    boolean_false: "FALSO",
    // For TSV format + generation workflow, see `src/locale/data/README.md`.
    error_literal_map: &[
        ("#VALUE!", "#¡VALOR!"),
        ("#NAME?", "#¿NOMBRE?"),
        ("#GETTING_DATA", "#OBTENIENDO_DATOS"),
    ],
    functions: &ES_ES_FUNCTIONS,
};

pub fn get_locale(id: &str) -> Option<&'static FormulaLocale> {
    match super::normalize_locale_id(id)? {
        "en-US" => Some(&EN_US),
        "de-DE" => Some(&DE_DE),
        "fr-FR" => Some(&FR_FR),
        "es-ES" => Some(&ES_ES),
        _ => None,
    }
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
}
