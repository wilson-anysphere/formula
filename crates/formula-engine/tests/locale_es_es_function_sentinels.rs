use formula_engine::locale;

/// Regression test: Excel's es-ES locale uses strongly localized spellings for core financial
/// functions. If our translation tables are incomplete, these functions silently fall back to
/// identity mappings (English), which breaks Excel compatibility.
#[test]
fn locale_parsing_es_es_financial_function_spellings_match_excel() {
    let mappings = [
        // Source of truth: `src/locale/data/sources/es-ES.json` (extracted from Excel).
        ("NPV", "VNA"),
        ("IRR", "TIR"),
        ("PV", "VA"),
        ("FV", "VF"),
        ("PMT", "PAGO"),
        ("RATE", "TASA"),
        // High-signal non-periodic variants.
        ("XIRR", "TIR.NO.PER"),
        ("XNPV", "VNA.NO.PER"),
    ];

    for (canonical, localized) in mappings {
        // These functions should not be identity-mapped in Spanish Excel; an identity mapping is a
        // strong signal that our locale table is incomplete/regressed.
        assert_ne!(
            canonical, localized,
            "test setup error: expected a non-identity mapping"
        );

        assert_eq!(locale::ES_ES.localized_function_name(canonical), localized);
        assert_eq!(locale::ES_ES.canonical_function_name(localized), canonical);

        // Ensure the formula translation pipeline (not just the raw name tables) round-trips.
        let canonical_formula = format!("={}()", canonical);
        let localized_formula = format!("={}()", localized);
        assert_eq!(
            locale::localize_formula(&canonical_formula, &locale::ES_ES).unwrap(),
            localized_formula
        );
        assert_eq!(
            locale::canonicalize_formula(&localized_formula, &locale::ES_ES).unwrap(),
            canonical_formula
        );
    }

    // Also validate that we round-trip locale punctuation (argument + decimal separators) while
    // translating these function names.
    assert_eq!(
        locale::localize_formula("=NPV(0.1,1,2)", &locale::ES_ES).unwrap(),
        "=VNA(0,1;1;2)"
    );
    assert_eq!(
        locale::canonicalize_formula("=VNA(0,1;1;2)", &locale::ES_ES).unwrap(),
        "=NPV(0.1,1,2)"
    );

    // Dotted localized names should also translate correctly with arguments.
    assert_eq!(
        locale::localize_formula("=XNPV(0.1,1,2)", &locale::ES_ES).unwrap(),
        "=VNA.NO.PER(0,1;1;2)"
    );
    assert_eq!(
        locale::canonicalize_formula("=VNA.NO.PER(0,1;1;2)", &locale::ES_ES).unwrap(),
        "=XNPV(0.1,1,2)"
    );
}
