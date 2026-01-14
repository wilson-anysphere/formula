use formula_engine::locale::{self, ValueLocaleConfig};

#[test]
fn locale_id_normalization_get_locale_normalizes_locale_ids() {
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
        ("en-UK", &locale::EN_US),
        // Other English locales still use the `en-US` formula parsing locale, even though value
        // parsing may use a different date order.
        ("en-AU", &locale::EN_US),
        ("en-NZ", &locale::EN_US),
        ("en-IE", &locale::EN_US),
        ("en-ZA", &locale::EN_US),
        // POSIX "C locale" aliases.
        ("C", &locale::EN_US),
        ("POSIX", &locale::EN_US),
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

    // Chinese script subtags should influence the default region when none is provided.
    assert_eq!(
        locale::get_locale("zh-Hant").expect("expected locale").id,
        "zh-TW"
    );
    assert_eq!(
        locale::get_locale("zh-Hans").expect("expected locale").id,
        "zh-CN"
    );
    // Script subtags should still work when the tag includes extensions.
    assert_eq!(
        locale::get_locale("zh-Hant-u-nu-latn")
            .expect("expected locale")
            .id,
        "zh-TW"
    );
}

#[test]
fn locale_id_normalization_value_locale_config_for_locale_id_normalizes_locale_ids() {
    for (input, expected) in [
        // Exact IDs still work.
        ("en-US", ValueLocaleConfig::en_us()),
        ("en-GB", ValueLocaleConfig::en_gb()),
        ("en-UK", ValueLocaleConfig::en_gb()),
        ("en-AU", ValueLocaleConfig::en_gb()),
        ("en-NZ", ValueLocaleConfig::en_gb()),
        ("en-IE", ValueLocaleConfig::en_gb()),
        ("en-ZA", ValueLocaleConfig::en_gb()),
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
        // POSIX "C locale" aliases.
        ("C", ValueLocaleConfig::en_us()),
        ("POSIX", ValueLocaleConfig::en_us()),
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

    // Chinese script subtags should influence the default region when none is provided.
    assert_eq!(
        ValueLocaleConfig::for_locale_id("zh-Hant"),
        Some(ValueLocaleConfig::zh_tw())
    );
    assert_eq!(
        ValueLocaleConfig::for_locale_id("zh-Hans"),
        Some(ValueLocaleConfig::zh_cn())
    );
    // Script subtags should still work when the tag includes extensions.
    assert_eq!(
        ValueLocaleConfig::for_locale_id("zh-Hant-u-nu-latn"),
        Some(ValueLocaleConfig::zh_tw())
    );
}
