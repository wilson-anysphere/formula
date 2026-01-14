use formula_model::drawings::ImageId;
use formula_model::{
    validate_value, CellRef, CellValue, DataValidation, DataValidationContext,
    DataValidationErrorAlert, DataValidationErrorKind, DataValidationErrorStyle,
    DataValidationKind, DataValidationOperator, EntityValue, ImageValue, Range, RecordValue,
    Worksheet,
};

struct TestCtx;

impl DataValidationContext for TestCtx {
    fn resolve_list_source(&self, formula: &str) -> Option<Vec<String>> {
        match formula {
            "MyRange" => Some(vec!["x".to_string(), "y".to_string()]),
            _ => None,
        }
    }

    fn eval_custom_formula(&self, formula: &str, candidate: &CellValue) -> Option<bool> {
        match formula {
            "POSITIVE" => match candidate {
                CellValue::Number(n) => Some(*n > 0.0),
                CellValue::String(s) => s.trim().parse::<f64>().ok().map(|n| n > 0.0),
                _ => Some(false),
            },
            _ => None,
        }
    }
}

fn dv(
    kind: DataValidationKind,
    operator: Option<DataValidationOperator>,
    formula1: &str,
    formula2: Option<&str>,
) -> DataValidation {
    DataValidation {
        kind,
        operator,
        formula1: formula1.to_string(),
        formula2: formula2.map(|s| s.to_string()),
        allow_blank: false,
        show_input_message: false,
        show_error_message: false,
        show_drop_down: false,
        input_message: None,
        error_alert: None,
    }
}

#[test]
fn decimal_operators_match_expected_semantics() {
    let ctx = TestCtx;

    assert!(
        validate_value(
            &dv(
                DataValidationKind::Decimal,
                Some(DataValidationOperator::Between),
                "5",
                Some("10")
            ),
            &CellValue::Number(7.0),
            &ctx
        )
        .ok
    );
    assert!(
        !validate_value(
            &dv(
                DataValidationKind::Decimal,
                Some(DataValidationOperator::Between),
                "5",
                Some("10")
            ),
            &CellValue::Number(4.0),
            &ctx
        )
        .ok
    );

    assert!(
        validate_value(
            &dv(
                DataValidationKind::Decimal,
                Some(DataValidationOperator::NotBetween),
                "5",
                Some("10")
            ),
            &CellValue::Number(4.0),
            &ctx
        )
        .ok
    );
    assert!(
        !validate_value(
            &dv(
                DataValidationKind::Decimal,
                Some(DataValidationOperator::NotBetween),
                "5",
                Some("10")
            ),
            &CellValue::Number(7.0),
            &ctx
        )
        .ok
    );

    assert!(
        validate_value(
            &dv(
                DataValidationKind::Decimal,
                Some(DataValidationOperator::Equal),
                "5",
                None
            ),
            &CellValue::String("5".to_string()),
            &ctx
        )
        .ok
    );

    assert!(
        !validate_value(
            &dv(
                DataValidationKind::Decimal,
                Some(DataValidationOperator::NotEqual),
                "5",
                None
            ),
            &CellValue::Number(5.0),
            &ctx
        )
        .ok
    );

    assert!(
        validate_value(
            &dv(
                DataValidationKind::Decimal,
                Some(DataValidationOperator::GreaterThan),
                "5",
                None
            ),
            &CellValue::Number(6.0),
            &ctx
        )
        .ok
    );

    assert!(
        validate_value(
            &dv(
                DataValidationKind::Decimal,
                Some(DataValidationOperator::GreaterThanOrEqual),
                "5",
                None
            ),
            &CellValue::Number(5.0),
            &ctx
        )
        .ok
    );

    assert!(
        validate_value(
            &dv(
                DataValidationKind::Decimal,
                Some(DataValidationOperator::LessThan),
                "5",
                None
            ),
            &CellValue::Number(4.0),
            &ctx
        )
        .ok
    );

    assert!(
        validate_value(
            &dv(
                DataValidationKind::Decimal,
                Some(DataValidationOperator::LessThanOrEqual),
                "5",
                None
            ),
            &CellValue::Number(5.0),
            &ctx
        )
        .ok
    );
}

#[test]
fn whole_number_validation_coerces_strings_and_checks_integerness() {
    let ctx = TestCtx;

    let rule = dv(
        DataValidationKind::Whole,
        Some(DataValidationOperator::Between),
        "1",
        Some("10"),
    );
    assert!(validate_value(&rule, &CellValue::String("3".to_string()), &ctx).ok);

    let bad = validate_value(&rule, &CellValue::Number(3.2), &ctx);
    assert_eq!(bad.ok, false);
    assert_eq!(
        bad.error_kind,
        Some(DataValidationErrorKind::NotWholeNumber)
    );
}

#[test]
fn allow_blank_short_circuits_validation() {
    let ctx = TestCtx;

    let mut rule = dv(
        DataValidationKind::Whole,
        Some(DataValidationOperator::Between),
        "1",
        Some("10"),
    );
    rule.allow_blank = true;
    assert!(validate_value(&rule, &CellValue::Empty, &ctx).ok);

    rule.allow_blank = false;
    let result = validate_value(&rule, &CellValue::Empty, &ctx);
    assert_eq!(result.ok, false);
    assert_eq!(
        result.error_kind,
        Some(DataValidationErrorKind::BlankNotAllowed)
    );
}

#[test]
fn allow_blank_treats_empty_rich_display_string_as_blank() {
    let ctx = TestCtx;

    let mut rule = dv(
        DataValidationKind::Whole,
        Some(DataValidationOperator::Between),
        "1",
        Some("10"),
    );
    rule.allow_blank = true;
    assert!(validate_value(&rule, &CellValue::Entity(EntityValue::new("")), &ctx).ok);
}

#[test]
fn allow_blank_does_not_treat_record_display_field_entity_as_blank() {
    let ctx = TestCtx;

    let mut rule = dv(
        DataValidationKind::Whole,
        Some(DataValidationOperator::Between),
        "1",
        Some("10"),
    );
    rule.allow_blank = true;

    let record = CellValue::Record(
        RecordValue::default()
            .with_display_field("company")
            .with_field("company", CellValue::Entity(EntityValue::new("Apple"))),
    );
    let result = validate_value(&rule, &record, &ctx);
    assert_eq!(result.ok, false);
    assert_eq!(
        result.error_kind,
        Some(DataValidationErrorKind::TypeMismatch)
    );
}

#[test]
fn list_validation_supports_constants_and_callback_sources() {
    let ctx = TestCtx;

    let constant = dv(DataValidationKind::List, None, "\"a,b,c\"", None);
    assert!(validate_value(&constant, &CellValue::String("B".to_string()), &ctx).ok);
    let bad = validate_value(&constant, &CellValue::String("d".to_string()), &ctx);
    assert_eq!(bad.ok, false);
    assert_eq!(bad.error_kind, Some(DataValidationErrorKind::NotInList));

    let from_range = dv(DataValidationKind::List, None, "MyRange", None);
    assert!(validate_value(&from_range, &CellValue::String("y".to_string()), &ctx).ok);
    assert!(!validate_value(&from_range, &CellValue::String("z".to_string()), &ctx).ok);
}

#[test]
fn list_validation_is_case_insensitive_for_unicode_text() {
    let ctx = TestCtx;

    // Unicode-aware case-insensitive matching should behave like Excel (ß -> SS).
    let constant = dv(DataValidationKind::List, None, "\"Maß\"", None);
    assert!(validate_value(&constant, &CellValue::String("MASS".to_string()), &ctx).ok);
}

#[test]
fn list_and_text_validations_use_rich_value_display_strings() {
    let ctx = TestCtx;

    let list = dv(DataValidationKind::List, None, "\"a,b,c\"", None);
    assert!(validate_value(&list, &CellValue::Entity(EntityValue::new("B")), &ctx).ok);

    let record = CellValue::Record(
        RecordValue::default()
            .with_display_field("name")
            .with_field("name", "d"),
    );
    let result = validate_value(&list, &record, &ctx);
    assert_eq!(result.ok, false);
    assert_eq!(result.error_kind, Some(DataValidationErrorKind::NotInList));

    let image = CellValue::Image(ImageValue {
        image_id: ImageId::new("image1.png"),
        alt_text: Some("B".to_string()),
        width: None,
        height: None,
    });
    assert!(validate_value(&list, &image, &ctx).ok);

    let len_rule = dv(
        DataValidationKind::TextLength,
        Some(DataValidationOperator::Equal),
        "2",
        None,
    );
    assert!(validate_value(&len_rule, &CellValue::Entity(EntityValue::new("🙂")), &ctx).ok);
    assert!(
        validate_value(
            &len_rule,
            &CellValue::Image(ImageValue {
                image_id: ImageId::new("image1.png"),
                alt_text: Some("🙂".to_string()),
                width: None,
                height: None,
            }),
            &ctx
        )
        .ok
    );
}

fn excel_serial(date: chrono::NaiveDate) -> f64 {
    let base = chrono::NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
    let days = (date - base).num_days() as f64;
    let leap_bug_cutover = chrono::NaiveDate::from_ymd_opt(1900, 3, 1).unwrap();
    if date >= leap_bug_cutover {
        days + 1.0
    } else {
        days
    }
}

#[test]
fn date_and_time_validations_accept_strings_and_serial_numbers() {
    let ctx = TestCtx;

    let date_rule = dv(
        DataValidationKind::Date,
        Some(DataValidationOperator::Between),
        "\"2020-01-01\"",
        Some("\"2020-01-31\""),
    );
    assert!(
        validate_value(
            &date_rule,
            &CellValue::String("2020-01-15".to_string()),
            &ctx
        )
        .ok
    );

    let jan_15 = excel_serial(chrono::NaiveDate::from_ymd_opt(2020, 1, 15).unwrap());
    assert!(validate_value(&date_rule, &CellValue::Number(jan_15), &ctx).ok);

    assert!(
        !validate_value(
            &date_rule,
            &CellValue::String("2020-02-01".to_string()),
            &ctx
        )
        .ok
    );

    let time_rule = dv(
        DataValidationKind::Time,
        Some(DataValidationOperator::Between),
        "\"09:00\"",
        Some("\"17:00\""),
    );
    assert!(validate_value(&time_rule, &CellValue::String("12:00".to_string()), &ctx).ok);
    assert!(validate_value(&time_rule, &CellValue::Number(0.5), &ctx).ok);
    assert!(validate_value(&time_rule, &CellValue::Number(1.5), &ctx).ok);
    assert!(!validate_value(&time_rule, &CellValue::String("08:59".to_string()), &ctx).ok);
}

#[test]
fn numeric_validation_parses_rich_value_display_strings() {
    let ctx = TestCtx;

    let rule = dv(
        DataValidationKind::Decimal,
        Some(DataValidationOperator::Equal),
        "123",
        None,
    );
    assert!(validate_value(&rule, &CellValue::Entity(EntityValue::new("123")), &ctx).ok);
    assert!(
        validate_value(
            &rule,
            &CellValue::Image(ImageValue {
                image_id: ImageId::new("image1.png"),
                alt_text: Some("123".to_string()),
                width: None,
                height: None,
            }),
            &ctx
        )
        .ok
    );
}

#[test]
fn text_length_uses_utf16_code_units_like_excel() {
    let ctx = TestCtx;

    let len_rule = dv(
        DataValidationKind::TextLength,
        Some(DataValidationOperator::Equal),
        "2",
        None,
    );

    assert!(validate_value(&len_rule, &CellValue::String("🙂".to_string()), &ctx).ok);
    assert!(!validate_value(&len_rule, &CellValue::String("a".to_string()), &ctx).ok);
}

#[test]
fn error_alert_style_and_messages_are_resolved() {
    let ctx = TestCtx;

    let mut rule = dv(DataValidationKind::List, None, "\"a,b\"", None);
    rule.show_error_message = true;
    rule.error_alert = Some(DataValidationErrorAlert {
        style: DataValidationErrorStyle::Warning,
        title: Some("Oops".to_string()),
        body: Some("Choose from the list".to_string()),
    });

    let result = validate_value(&rule, &CellValue::String("c".to_string()), &ctx);
    assert_eq!(result.ok, false);
    assert_eq!(result.error_kind, Some(DataValidationErrorKind::NotInList));
    assert_eq!(result.error_style, Some(DataValidationErrorStyle::Warning));
    assert_eq!(result.error_title.as_deref(), Some("Oops"));
    assert_eq!(
        result.error_message.as_deref(),
        Some("Choose from the list")
    );
}

#[test]
fn custom_validation_uses_callback_hook() {
    let ctx = TestCtx;

    let rule = dv(DataValidationKind::Custom, None, "POSITIVE", None);
    assert!(validate_value(&rule, &CellValue::Number(1.0), &ctx).ok);

    let bad = validate_value(&rule, &CellValue::Number(-1.0), &ctx);
    assert_eq!(bad.ok, false);
    assert_eq!(
        bad.error_kind,
        Some(DataValidationErrorKind::CustomFormulaFalse)
    );

    let unresolved = dv(DataValidationKind::Custom, None, "UNKNOWN", None);
    let bad = validate_value(&unresolved, &CellValue::Number(1.0), &ctx);
    assert_eq!(bad.ok, false);
    assert_eq!(
        bad.error_kind,
        Some(DataValidationErrorKind::CustomFormulaUnresolved)
    );
}

#[test]
fn worksheet_stores_validations_and_respects_merged_cell_anchors() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet
        .merge_range(Range::from_a1("A1:B1").unwrap())
        .expect("merge");

    let rule = dv(DataValidationKind::List, None, "\"a,b\"", None);
    let id = sheet.add_data_validation(vec![Range::from_a1("B1").unwrap()], rule);

    let a1 = CellRef::from_a1("A1").unwrap();
    let b1 = CellRef::from_a1("B1").unwrap();
    assert_eq!(sheet.data_validations_for_cell(a1).len(), 1);
    assert_eq!(sheet.data_validations_for_cell(b1).len(), 1);

    assert!(sheet.remove_data_validation(id));
    assert!(sheet.data_validations_for_cell(a1).is_empty());
}
