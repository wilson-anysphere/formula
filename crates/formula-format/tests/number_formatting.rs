use formula_format::locale;

#[test]
fn formats_en_us_thousands_and_decimals() {
    assert_eq!(locale::format_number(1234.56, &locale::EN_US), "1,234.56");
    assert_eq!(locale::format_number(-12.0, &locale::EN_US), "-12");
}

#[test]
fn formats_de_de_thousands_and_decimals() {
    assert_eq!(locale::format_number(1234.56, &locale::DE_DE), "1.234,56");
    assert_eq!(locale::format_number(0.5, &locale::DE_DE), "0,5");
}

