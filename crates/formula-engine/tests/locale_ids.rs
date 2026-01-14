use formula_engine::locale::{self, ValueLocaleConfig};

#[test]
fn dbcs_locale_ids_are_registered() {
    for id in ["ja-JP", "zh-CN", "ko-KR", "zh-TW"] {
        assert!(
            ValueLocaleConfig::for_locale_id(id).is_some(),
            "ValueLocaleConfig missing locale id: {id}"
        );
        assert!(
            locale::get_locale(id).is_some(),
            "Formula locale registry missing locale id: {id}"
        );
    }
}

#[test]
fn iter_locales_includes_all_expected_locale_ids() {
    let supported: Vec<&'static str> = locale::iter_locales().map(|locale| locale.id).collect();

    for id in [
        "en-US", "de-DE", "fr-FR", "es-ES", "ja-JP", "zh-CN", "zh-TW", "ko-KR",
    ] {
        assert!(
            supported.contains(&id),
            "iter_locales() missing locale id: {id}"
        );
    }
}
