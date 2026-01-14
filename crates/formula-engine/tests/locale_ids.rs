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

