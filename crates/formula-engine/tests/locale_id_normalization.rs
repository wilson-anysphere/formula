use formula_engine::locale::{self, ValueLocaleConfig};

#[test]
fn get_locale_normalizes_locale_ids() {
    for (input, expected) in [
        // Exact IDs still work.
        ("en-US", &locale::EN_US),
        ("de-DE", &locale::DE_DE),
        ("fr-FR", &locale::FR_FR),
        ("es-ES", &locale::ES_ES),
        // Trim whitespace.
        ("  de-DE  ", &locale::DE_DE),
        // Treat `-` and `_` as equivalent.
        ("de_DE", &locale::DE_DE),
        // Match case-insensitively.
        ("DE-de", &locale::DE_DE),
        // Language-only fallbacks.
        ("de", &locale::DE_DE),
        ("fr", &locale::FR_FR),
        ("es", &locale::ES_ES),
        ("en", &locale::EN_US),
    ] {
        let locale = locale::get_locale(input).unwrap_or_else(|| {
            panic!("expected locale for {input:?}, got None");
        });
        assert!(
            std::ptr::eq(locale, expected),
            "expected {input:?} to resolve to {:?}, got {:?}",
            expected.id,
            locale.id
        );
    }

    // Unknown locales should stay unknown.
    for input in ["de-AT", "en-GB", "zz-ZZ", "", "   "] {
        assert!(
            locale::get_locale(input).is_none(),
            "expected None for {input:?}"
        );
    }
}

#[test]
fn value_locale_config_for_locale_id_normalizes_locale_ids() {
    for (input, expected) in [
        // Exact IDs still work.
        ("en-US", ValueLocaleConfig::en_us()),
        ("de-DE", ValueLocaleConfig::de_de()),
        ("fr-FR", ValueLocaleConfig::fr_fr()),
        ("es-ES", ValueLocaleConfig::es_es()),
        // Trim whitespace.
        ("  fr-FR  ", ValueLocaleConfig::fr_fr()),
        // Treat `-` and `_` as equivalent.
        ("es_ES", ValueLocaleConfig::es_es()),
        // Match case-insensitively.
        ("EN_us", ValueLocaleConfig::en_us()),
        // Language-only fallbacks.
        ("de", ValueLocaleConfig::de_de()),
        ("fr", ValueLocaleConfig::fr_fr()),
        ("es", ValueLocaleConfig::es_es()),
        ("en", ValueLocaleConfig::en_us()),
    ] {
        assert_eq!(
            ValueLocaleConfig::for_locale_id(input),
            Some(expected),
            "unexpected ValueLocaleConfig for {input:?}"
        );
    }

    // Unknown locales should stay unknown.
    for input in ["de-AT", "en-GB", "zz-ZZ", "", "   "] {
        assert_eq!(
            ValueLocaleConfig::for_locale_id(input),
            None,
            "expected None for {input:?}"
        );
    }
}
