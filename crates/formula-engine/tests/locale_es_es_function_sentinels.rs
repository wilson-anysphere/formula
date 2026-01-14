use formula_engine::locale;

/// Regression test: Excel's es-ES locale uses strongly localized spellings for core worksheet
/// functions (including boolean TRUE()/FALSE()).
///
/// If our translation sources are incomplete or drift from Excel, these can silently fall back to
/// English, breaking localized editing + round-tripping.
#[test]
fn locale_parsing_es_es_core_function_spellings_match_excel() {
    let mappings = [
        // Source of truth: `src/locale/data/sources/es-ES.json` (extracted from Excel).
        ("SUM", "SUMA"),
        ("IF", "SI"),
        ("TRUE", "VERDADERO"),
        ("FALSE", "FALSO"),
        ("COUNTIF", "CONTAR.SI"),
        ("VLOOKUP", "BUSCARV"),
        ("HLOOKUP", "BUSCARH"),
        ("IFERROR", "SI.ERROR"),
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
}

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

/// Regression test: the es-ES locale strongly localizes many common statistical and forecasting
/// function spellings. These are high-signal sentinels because missing entries silently degrade to
/// identity mappings (English).
#[test]
fn locale_parsing_es_es_statistical_and_forecasting_spellings_match_excel() {
    let mappings = [
        // Forecasting
        ("FORECAST", "PRONOSTICO"),
        ("FORECAST.LINEAR", "PRONOSTICO.LINEAL"),
        ("FORECAST.ETS", "PRONOSTICO.ETS"),
        ("FORECAST.ETS.CONFINT", "PRONOSTICO.ETS.INT.CONFIANZA"),
        ("FORECAST.ETS.SEASONALITY", "PRONOSTICO.ETS.ESTACIONALIDAD"),
        ("FORECAST.ETS.STAT", "PRONOSTICO.ETS.ESTADISTICA"),
        // Statistical distributions + tests
        ("CHISQ.DIST", "DISTR.CHI.CUAD.N"),
        ("CHISQ.DIST.RT", "DISTR.CHI.CUAD.DER.N"),
        ("CHISQ.INV", "INV.CHI.CUAD.N"),
        ("CHISQ.INV.RT", "INV.CHI.CUAD.DER.N"),
        ("CHISQ.TEST", "PRUEBA.CHI.CUAD.N"),
        ("F.DIST", "DISTR.F.N"),
        ("F.DIST.RT", "DISTR.F.DER.N"),
        ("F.TEST", "PRUEBA.F.N"),
        ("F.INV", "INV.F.N"),
        ("F.INV.RT", "INV.F.DER.N"),
        ("FTEST", "PRUEBA.F"),
        ("FINV", "INV.F"),
        ("GAMMA.DIST", "DISTR.GAMMA.N"),
        ("GAMMA.INV", "INV.GAMMA.N"),
        ("LOGNORM.DIST", "DISTR.LOGNORM.N"),
        ("LOGNORM.INV", "INV.LOGNORM.N"),
        ("LOGINV", "INV.LOGNORM"),
        ("NEGBINOM.DIST", "DISTR.BINOM.NEG.N"),
        ("POISSON", "DISTR.POISSON"),
        ("POISSON.DIST", "DISTR.POISSON.N"),
        ("T.DIST.2T", "DISTR.T.2C.N"),
        ("T.DIST.RT", "DISTR.T.DER.N"),
        ("T.INV", "INV.T.N"),
        ("Z.TEST", "PRUEBA.Z.N"),
        // Bond/coupon functions (commonly missing when Excel treats them as _xludf)
        ("COUPDAYBS", "DIAS.CUPON.INI"),
        ("COUPDAYS", "DIAS.CUPON"),
        ("COUPDAYSNC", "DIAS.CUPON.SIG"),
        ("COUPNCD", "FECHA.CUPON.SIG"),
        ("COUPNUM", "NUM.CUPONES"),
        ("COUPPCD", "FECHA.CUPON.ANT"),
        // Other common regression candidates
        ("INTERCEPT", "INTERSECCION.EJE"),
        ("IPMT", "PAGOINT"),
        ("MIRR", "TIRM"),
        ("MINVERSE", "MINVERSA"),
        // Legacy distribution spellings
        ("WEIBULL", "DISTR.WEIBULL"),
    ];

    for (canonical, localized) in mappings {
        assert_ne!(
            canonical, localized,
            "test setup error: expected a non-identity mapping"
        );

        assert_eq!(locale::ES_ES.localized_function_name(canonical), localized);
        assert_eq!(locale::ES_ES.canonical_function_name(localized), canonical);

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
}

/// Guard against accidentally committing an incomplete `es-ES` mapping (missing entries silently
/// fall back to canonical/English names, which looks "complete" but breaks Excel compatibility).
#[test]
fn locale_es_es_identity_mapping_rate_is_not_suspiciously_high() {
    // Keep the bar conservative to avoid brittle exact counts; we just want to catch obvious
    // regressions back to mostly-English tables.
    const MAX_IDENTITY_RATE: f64 = 0.35;
    const ES_ES_TSV: &str = include_str!("../src/locale/data/es-ES.tsv");

    let mut total = 0usize;
    let mut identity = 0usize;
    for line in ES_ES_TSV.lines() {
        if line.is_empty() || line.starts_with('#') {
            continue;
        }
        let (canonical, localized) = line.split_once('\t').unwrap_or_else(|| {
            panic!("invalid TSV entry (expected Canonical<TAB>Localized): {line}")
        });
        total += 1;
        if canonical == localized {
            identity += 1;
        }
    }

    let identity_rate = identity as f64 / total as f64;
    assert!(
        identity_rate <= MAX_IDENTITY_RATE,
        "es-ES identity mapping rate too high ({identity}/{total} = {identity_rate:.3}); this suggests the locale table is incomplete and falling back to English"
    );
}

/// Secondary robustness check: detect cases where `es-ES` is identity-mapped (English fallback)
/// even though *both* `de-DE` and `fr-FR` have non-identity spellings for the same canonical
/// function.
///
/// We expect this set to be small: while Spanish sometimes keeps English names (especially for
/// niche functions), widespread identity mappings here are a strong signal that `es-ES` was
/// extracted from an older Excel build that treated newer functions as unknown and returned
/// canonical names.
#[test]
fn locale_es_es_suspicious_identity_mappings_are_rare() {
    const DE_DE_TSV: &str = include_str!("../src/locale/data/de-DE.tsv");
    const ES_ES_TSV: &str = include_str!("../src/locale/data/es-ES.tsv");
    const FR_FR_TSV: &str = include_str!("../src/locale/data/fr-FR.tsv");

    fn parse(tsv: &str) -> std::collections::BTreeMap<String, String> {
        let mut out = std::collections::BTreeMap::new();
        for line in tsv.lines() {
            if line.is_empty() || line.starts_with('#') {
                continue;
            }
            let (canon, loc) = line.split_once('\t').unwrap_or_else(|| {
                panic!("invalid TSV entry (expected Canonical<TAB>Localized): {line}")
            });
            out.insert(canon.to_string(), loc.to_string());
        }
        out
    }

    let de = parse(DE_DE_TSV);
    let es = parse(ES_ES_TSV);
    let fr = parse(FR_FR_TSV);

    let mut base = 0usize;
    let mut suspicious: Vec<String> = Vec::new();
    for (canon, es_loc) in es {
        let Some(de_loc) = de.get(&canon) else {
            continue;
        };
        let Some(fr_loc) = fr.get(&canon) else {
            continue;
        };
        if de_loc != &canon && fr_loc != &canon {
            base += 1;
            if es_loc == canon {
                suspicious.push(canon);
            }
        }
    }

    assert!(
        base > 0,
        "test invariant: expected some de-DE/fr-FR translated functions"
    );

    let suspicious_count = suspicious.len();
    let ratio = suspicious_count as f64 / base as f64;

    // Keep this guard intentionally loose: Spanish can legitimately keep English names for a few
    // functions even when other locales translate them. We just want to catch big regressions.
    const MAX_SUSPICIOUS_RATIO: f64 = 0.10;

    if ratio > MAX_SUSPICIOUS_RATIO {
        suspicious.sort();
        let sample = suspicious
            .iter()
            .take(25)
            .cloned()
            .collect::<Vec<_>>()
            .join(", ");
        panic!(
            "es-ES has too many suspicious identity mappings vs de-DE+fr-FR ({suspicious_count}/{base} = {ratio:.3}). Examples: {sample}"
        );
    }
}
