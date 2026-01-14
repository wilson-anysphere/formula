use formula_engine::LocaleConfig;

mod locale_parse_number {
    use super::LocaleConfig;

    #[test]
    fn de_de_parses_thousands_and_decimal() {
        let locale = LocaleConfig::de_de();

        assert_eq!(locale.parse_number("1.234,56"), Some(1234.56));
        assert_eq!(locale.parse_number("1.234.567"), Some(1234567.0));

        // `.` could be a thousands separator or a decimal point in this locale. Reject ambiguous
        // inputs with multiple possible interpretations.
        assert_eq!(locale.parse_number("1.234.56"), None);

        // Always accept canonical `.` decimal separator regardless of locale.
        assert_eq!(locale.parse_number("1.5"), Some(1.5));
    }

    #[test]
    fn de_de_parses_exponent_forms() {
        let locale = LocaleConfig::de_de();

        assert_eq!(locale.parse_number("1,23E3"), Some(1230.0));
        assert_eq!(locale.parse_number("1.23E3"), Some(1230.0));
    }

    #[test]
    fn fr_fr_parses_nbsp_and_narrow_nbsp_thousands_separators() {
        let locale = LocaleConfig::fr_fr();

        assert_eq!(locale.parse_number("1\u{00A0}234,56"), Some(1234.56));
        assert_eq!(locale.parse_number("1\u{202F}234,56"), Some(1234.56));

        // ASCII space is meaningful in the formula language (range intersection), so we do not
        // treat it as a valid grouping separator.
        assert_eq!(locale.parse_number("1 234,56"), None);
    }

    #[test]
    fn es_es_parses_thousands_and_decimal() {
        let locale = LocaleConfig::es_es();

        assert_eq!(locale.parse_number("1.234,56"), Some(1234.56));
    }

    #[test]
    fn en_us_parses_canonical_numbers() {
        let locale = LocaleConfig::en_us();

        assert_eq!(locale.parse_number("1.5"), Some(1.5));
        assert_eq!(locale.parse_number("1.23E3"), Some(1230.0));
    }
}
