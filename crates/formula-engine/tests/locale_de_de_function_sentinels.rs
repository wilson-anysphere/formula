use formula_engine::locale;

/// Regression test: Excel's de-DE locale uses strongly localized spellings for core worksheet
/// functions. If our translation sources are incomplete or drift from Excel, these functions can
/// silently fall back to English, breaking localized editing + round-tripping.
#[test]
fn de_de_core_function_spellings_match_excel() {
    let mappings = [
        // Source of truth: `src/locale/data/sources/de-DE.json` (extracted from Excel and then
        // normalized to omit identity mappings).
        ("SUM", "SUMME"),
        ("IF", "WENN"),
        ("TRUE", "WAHR"),
        ("FALSE", "FALSCH"),
        ("COUNTIF", "ZÄHLENWENN"),
        ("VLOOKUP", "SVERWEIS"),
        ("HLOOKUP", "WVERWEIS"),
        ("IFERROR", "WENNFEHLER"),
        ("XLOOKUP", "XVERWEIS"),
        ("TEXTJOIN", "TEXTVERKETTEN"),
    ];

    for (canonical, localized) in mappings {
        // These should not be identity-mapped in German Excel; an identity mapping is a strong
        // signal that our locale table is incomplete/regressed.
        assert_ne!(
            canonical, localized,
            "test setup error: expected a non-identity mapping"
        );

        assert_eq!(locale::DE_DE.localized_function_name(canonical), localized);
        assert_eq!(locale::DE_DE.canonical_function_name(localized), canonical);

        // Ensure the formula translation pipeline (not just the raw name tables) round-trips.
        let canonical_formula = format!("={}()", canonical);
        let localized_formula = format!("={}()", localized);
        assert_eq!(
            locale::localize_formula(&canonical_formula, &locale::DE_DE).unwrap(),
            localized_formula
        );
        assert_eq!(
            locale::canonicalize_formula(&localized_formula, &locale::DE_DE).unwrap(),
            canonical_formula
        );
    }
}
