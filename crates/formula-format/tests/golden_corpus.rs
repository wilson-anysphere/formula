use formula_format::{format_value, DateSystem, FormatOptions, Locale, Value};

#[test]
fn golden_corpus_numbers_and_datetime() {
    let en_1900 = FormatOptions {
        locale: Locale::en_us(),
        date_system: DateSystem::Excel1900,
    };

    let cases: &[(Value<'_>, &str, &FormatOptions, &str)] = &[
        // Basic numeric rules.
        (Value::Number(0.0), "0", &en_1900, "0"),
        (Value::Number(-2.0), "0", &en_1900, "-2"),
        (Value::Number(0.5), "#.00", &en_1900, ".50"),
        (Value::Number(0.0), "#.00", &en_1900, ".00"),
        // Percent scaling.
        (Value::Number(0.256), "0%", &en_1900, "26%"),
        // Scaling commas.
        (Value::Number(1500.0), "0,", &en_1900, "2"),
        (Value::Number(1234567.0), "#,##0,", &en_1900, "1,235"),
        // Escaped literals: `\\%` renders a percent sign without scaling.
        (Value::Number(5.0), "0\\%", &en_1900, "5%"),
        // `?` placeholders render spaces for insignificant digits.
        (Value::Number(1.2), "0.??", &en_1900, "1.2 "),
        // Conditional sections (conditions evaluated in-order, first unconditional section is else).
        (Value::Number(150.0), "[>=100]0.0;[Red]0.0", &en_1900, "150.0"),
        (Value::Number(50.0), "[>=100]0.0;[Red]0.0", &en_1900, "50.0"),
        (Value::Number(-50.0), "[>=100]0.0;[Red]0.0", &en_1900, "50.0"),
        // Date/time.
        (Value::Number(1.0), "yyyy-mm-dd", &en_1900, "1900-01-01"),
        (Value::Number(1.0), "d-mmm-yy", &en_1900, "1-Jan-00"),
        (Value::Number(1.0), "ddd", &en_1900, "Mon"),
        (Value::Number(1.0), "dddd", &en_1900, "Monday"),
        (Value::Number(1.5), "h:mm", &en_1900, "12:00"),
        (Value::Number(1.75), "h:mm:ss", &en_1900, "18:00:00"),
        (Value::Number(1.5), "[h]:mm", &en_1900, "36:00"),
        (Value::Number(1.0 / 24.0), "[hh]", &en_1900, "01"),
        (Value::Number(1.0 / 1440.0), "[mm]", &en_1900, "01"),
        (Value::Number(1.0 / 86_400.0), "[ss]", &en_1900, "01"),
        // Fractional seconds with more than 3 digits.
        (
            Value::Number(1.2345 / 86_400.0),
            "mm:ss.0000",
            &en_1900,
            "00:01.2345",
        ),
        // Negative serial dates render ##### like Excel.
        (Value::Number(-1.0), "m/d/yyyy", &en_1900, "#####"),
    ];

    for (value, format_code, options, expected) in cases {
        let rendered = format_value(*value, Some(*format_code), options);
        assert_eq!(
            rendered.text,
            *expected,
            "value={value:?} format={format_code:?}"
        );
    }
}
