use formula_engine::locale::{self, ValueLocaleConfig};

#[test]
fn get_locale_normalizes_locale_ids() {
    for (input, expected) in [
        // Exact IDs still work.
        ("en-US", &locale::EN_US),
        ("de-DE", &locale::DE_DE),
        ("fr-FR", &locale::FR_FR),
        ("es-ES", &locale::ES_ES),
        // POSIX locale IDs with encoding / modifier suffix.
        ("de_DE.UTF-8", &locale::DE_DE),
        ("de_DE@euro", &locale::DE_DE),
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
        // Language/region fallbacks.
        ("fr-CA", &locale::FR_FR),
        ("de-AT", &locale::DE_DE),
        ("es-MX", &locale::ES_ES),
        // `en-GB` is accepted as an alias for the formula locale (English function names +
        // `,` separators).
        ("en-GB", &locale::EN_US),
        // Ignore BCP-47 variants/extensions.
        ("fr-FR-u-nu-latn", &locale::FR_FR),
        ("de-CH-1996", &locale::DE_DE),
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
    for input in ["pt-BR", "it-IT", "zz-ZZ", "", "   "] {
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
        ("en-GB", ValueLocaleConfig::en_gb()),
        ("de-DE", ValueLocaleConfig::de_de()),
        ("fr-FR", ValueLocaleConfig::fr_fr()),
        ("es-ES", ValueLocaleConfig::es_es()),
        // POSIX locale IDs with encoding / modifier suffix.
        ("de_DE.UTF-8", ValueLocaleConfig::de_de()),
        ("de_DE@euro", ValueLocaleConfig::de_de()),
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
        // Language/region fallbacks.
        ("fr-CA", ValueLocaleConfig::fr_fr()),
        ("de-AT", ValueLocaleConfig::de_de()),
        ("es-MX", ValueLocaleConfig::es_es()),
        // Ignore BCP-47 variants/extensions.
        ("fr-FR-u-nu-latn", ValueLocaleConfig::fr_fr()),
        ("de-CH-1996", ValueLocaleConfig::de_de()),
    ] {
        assert_eq!(
            ValueLocaleConfig::for_locale_id(input),
            Some(expected),
            "unexpected ValueLocaleConfig for {input:?}"
        );
    }

    // Unknown locales should stay unknown.
    for input in ["pt-BR", "it-IT", "zz-ZZ", "", "   "] {
        assert_eq!(
            ValueLocaleConfig::for_locale_id(input),
            None,
            "expected None for {input:?}"
        );
    }
}
