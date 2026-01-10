use crate::LocaleConfig;

use std::collections::HashMap;
use std::sync::OnceLock;

#[derive(Debug)]
struct FunctionTranslationMaps {
    canon_to_loc: HashMap<&'static str, &'static str>,
    loc_to_canon: HashMap<&'static str, &'static str>,
}

/// Translation table for Excel function identifiers.
///
/// Data is stored outside the Rust source in simple TSV files under `locale/data/`.
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

            for line in self.data_tsv.lines() {
                let line = line.trim();
                if line.is_empty() || line.starts_with('#') {
                    continue;
                }

                let (canon, loc) = line.split_once('\t').unwrap_or_else(|| {
                    panic!("invalid function translation line (expected TSV): {line:?}")
                });
                let canon = canon.trim();
                let loc = loc.trim();
                if canon.is_empty() || loc.is_empty() {
                    panic!("invalid function translation line (empty entry): {line:?}");
                }

                canon_to_loc.insert(canon, loc);
                loc_to_canon.insert(loc, canon);
            }

            FunctionTranslationMaps {
                canon_to_loc,
                loc_to_canon,
            }
        })
    }

    fn localized_to_canonical(&self, localized_upper: &str) -> Option<&'static str> {
        self.maps().loc_to_canon.get(localized_upper).copied()
    }

    fn canonical_to_localized(&self, canonical_upper: &str) -> Option<&'static str> {
        self.maps().canon_to_loc.get(canonical_upper).copied()
    }
}

static EMPTY_FUNCTIONS: FunctionTranslations = FunctionTranslations::new("");
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
    functions: &'static FunctionTranslations,
}

impl FormulaLocale {
    /// Translate an input function name into canonical form.
    pub fn canonical_function_name(&self, name: &str) -> String {
        let (has_prefix, base) = split_xlfn_prefix(name);
        let upper = base.to_ascii_uppercase();

        let mapped = self
            .functions
            .localized_to_canonical(&upper)
            .unwrap_or(upper.as_str());

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
        let upper = base.to_ascii_uppercase();

        let mapped = self
            .functions
            .canonical_to_localized(&upper)
            .unwrap_or(upper.as_str());

        let mut out = String::new();
        if has_prefix {
            out.push_str("_xlfn.");
        }
        out.push_str(mapped);
        out
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
    functions: &DE_DE_FUNCTIONS,
};

/// French (France).
pub static FR_FR: FormulaLocale = FormulaLocale {
    id: "fr-FR",
    config: LocaleConfig {
        decimal_separator: ',',
        arg_separator: ';',
        array_col_separator: '\\',
        array_row_separator: ';',
        // Often a space (or non-breaking space) in the UI, but ambiguous in formulas.
        thousands_separator: None,
    },
    is_rtl: false,
    functions: &FR_FR_FUNCTIONS,
};

/// Spanish (Spain).
pub static ES_ES: FormulaLocale = FormulaLocale {
    id: "es-ES",
    config: LocaleConfig {
        decimal_separator: ',',
        arg_separator: ';',
        array_col_separator: '\\',
        array_row_separator: ';',
        thousands_separator: Some('.'),
    },
    is_rtl: false,
    functions: &ES_ES_FUNCTIONS,
};

pub fn get_locale(id: &str) -> Option<&'static FormulaLocale> {
    match id {
        "en-US" => Some(&EN_US),
        "de-DE" => Some(&DE_DE),
        "fr-FR" => Some(&FR_FR),
        "es-ES" => Some(&ES_ES),
        _ => None,
    }
}
