use formula_engine::eval::parse_a1;
use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn filter_filters_rows_from_range() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", true).unwrap();
    engine.set_cell_value("Sheet1", "B2", false).unwrap();
    engine.set_cell_value("Sheet1", "B3", true).unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=FILTER(A1:A3,B1:B3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("D2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
}

#[test]
fn filter_filters_columns_from_range() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 4.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 6.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", "=FILTER(A1:C2,{1,0,1})")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "E1").expect("spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("F2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(6.0));
}

#[test]
fn filter_if_empty_branch_and_calc_error() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", false).unwrap();
    engine.set_cell_value("Sheet1", "B2", false).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=FILTER(A1:A2,B1:B2,\"none\")")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Text("none".to_string())
    );

    engine
        .set_cell_formula("Sheet1", "E1", "=FILTER(A1:A2,B1:B2)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Error(ErrorKind::Calc)
    );
}

#[test]
fn sort_sorts_numbers_and_text() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=SORT(A1:A3)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));

    engine.set_cell_value("Sheet1", "B1", "b").unwrap();
    engine.set_cell_value("Sheet1", "B2", "A").unwrap();
    engine.set_cell_value("Sheet1", "B3", "c").unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=SORT(B1:B3)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Text("A".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "D2"),
        Value::Text("b".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Text("c".to_string())
    );
}

#[test]
fn sort_sorts_rich_values_by_display_string() {
    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("b")))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", Value::Record(RecordValue::new("A")))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A3", Value::Entity(EntityValue::new("c")))
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=SORT(A1:A3)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Record(RecordValue::new("A"))
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C2"),
        Value::Entity(EntityValue::new("b"))
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C3"),
        Value::Entity(EntityValue::new("c"))
    );
}

#[test]
fn sort_uses_record_display_field_for_keying() {
    let mut engine = Engine::new();

    let mut record_a = RecordValue::new("zzz").field("Name", "A");
    record_a.display_field = Some("name".to_string());
    let record_a = Value::Record(record_a);

    let mut record_b = RecordValue::new("aaa").field("Name", "B");
    record_b.display_field = Some("NAME".to_string());
    let record_b = Value::Record(record_b);

    engine
        .set_cell_value("Sheet1", "A1", record_a.clone())
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", record_b.clone())
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=SORT(A1:A2)")
        .unwrap();
    engine.recalculate_single_threaded();

    // Sort should use the `Name` field ("A"/"B"), not the fallback `display` string ("zzz"/"aaa").
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), record_a);
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), record_b);
}

#[test]
fn sort_is_case_insensitive_for_unicode_text_like_excel() {
    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("Straße")))
        .unwrap();
    engine.set_cell_value("Sheet1", "A2", "STRASSE").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=SORT(A1:A2)")
        .unwrap();
    engine.recalculate_single_threaded();

    // Excel-style case folding treats ß like SS. Entities/records should behave text-like, so the
    // two values compare equal and SORT preserves their original order.
    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Entity(EntityValue::new("Straße"))
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C2"),
        Value::Text("STRASSE".to_string())
    );
}

#[test]
fn sort_orders_field_error_by_error_code() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SORT({#FIELD!;#VALUE!;#DIV/0!})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Error(ErrorKind::Field)
    );
}

#[test]
fn sort_record_display_field_error_is_an_error_key() {
    let mut engine = Engine::new();

    let mut record_field_err =
        RecordValue::new("aaa").field("Name", Value::Error(ErrorKind::Field));
    record_field_err.display_field = Some("Name".to_string());
    let record_field_err = Value::Record(record_field_err);

    let mut record_value_err =
        RecordValue::new("zzz").field("Name", Value::Error(ErrorKind::Value));
    record_value_err.display_field = Some("Name".to_string());
    let record_value_err = Value::Record(record_value_err);

    engine
        .set_cell_value("Sheet1", "A1", record_field_err.clone())
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", record_value_err.clone())
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=SORT(A1:A2)")
        .unwrap();
    engine.recalculate_single_threaded();

    // Sorting should use the display_field-derived errors (VALUE before FIELD), not the fallback
    // record.display ("aaa"/"zzz").
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), record_value_err);
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), record_field_err);
}

#[test]
fn sort_by_col_sorts_columns() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 20.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", "=SORT(A1:C2,1,1,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(30.0));
}

#[test]
fn sort_supports_multi_key_sort_index_and_order_vectors() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "a").unwrap();
    engine.set_cell_value("Sheet1", "A2", "b").unwrap();
    engine.set_cell_value("Sheet1", "A3", "c").unwrap();
    engine.set_cell_value("Sheet1", "A4", "d").unwrap();

    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B4", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=SORT(A1:B4,{2,1},{1,-1})")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Text("c".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "D2"),
        Value::Text("b".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(1.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Text("a".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(2.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "D4"),
        Value::Text("d".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "E4"), Value::Number(3.0));
}

#[test]
fn unique_by_row_and_column_and_exactly_once() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A5", 2.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=UNIQUE(A1:A5)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));

    engine
        .set_cell_formula("Sheet1", "D1", "=UNIQUE(A1:A5,,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(3.0));

    engine.set_cell_value("Sheet1", "E1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "F1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "G1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "E2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "F2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "G2", 4.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "I1", "=UNIQUE(E1:G2,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "I1").expect("spill range");
    assert_eq!(start, parse_a1("I1").unwrap());
    assert_eq!(end, parse_a1("J2").unwrap());

    // First two columns are duplicates; UNIQUE by_col keeps the first occurrence.
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J2"), Value::Number(4.0));
}

#[test]
fn unique_uses_record_display_field_for_deduping() {
    let mut engine = Engine::new();

    let mut record_a = RecordValue::new("display_a").field("Name", "Same");
    record_a.display_field = Some("name".to_string());
    let record_a = Value::Record(record_a);

    let mut record_b = RecordValue::new("display_b").field("Name", "Same");
    record_b.display_field = Some("NAME".to_string());
    let record_b = Value::Record(record_b);

    engine
        .set_cell_value("Sheet1", "A1", record_a.clone())
        .unwrap();
    engine.set_cell_value("Sheet1", "A2", record_b).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=UNIQUE(A1:A2)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C1").unwrap());

    // UNIQUE should use the `Name` field value ("Same") to dedupe, keeping the first occurrence.
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), record_a);
}

#[test]
fn unique_record_display_field_numeric_is_locale_aware() {
    let mut engine = Engine::new();
    // Use a locale with a different decimal separator to ensure we use the workbook locale when
    // stringifying numeric display_field values.
    assert!(engine.set_value_locale_id("de-DE"));

    let mut record_a = RecordValue::new("display_a").field("Name", 1.5);
    record_a.display_field = Some("Name".to_string());
    let record_a = Value::Record(record_a);

    // In de-DE, 1.5 is formatted as "1,5" under General.
    let mut record_b = RecordValue::new("display_b").field("Name", "1,5");
    record_b.display_field = Some("Name".to_string());

    engine
        .set_cell_value("Sheet1", "A1", record_a.clone())
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", Value::Record(record_b))
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=UNIQUE(A1:A2)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C1").unwrap());

    // If numeric stringification were not locale-aware, we'd get "1.5" and the two records would
    // not dedupe. With locale-aware General formatting, they dedupe and keep the first record.
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), record_a);
}

#[test]
fn spill_blocked_spill_and_footprint_change_recalculate_dependents() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", true).unwrap();
    engine.set_cell_value("Sheet1", "B2", true).unwrap();
    engine.set_cell_value("Sheet1", "B3", true).unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=FILTER(A1:A3,B1:B3)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(3.0));

    // Block the spill.
    engine.set_cell_value("Sheet1", "D2", 99.0).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Spill)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(99.0));

    // Clear the blocker and shrink/expand the spill while ensuring dependents recalculate.
    engine.set_cell_value("Sheet1", "D2", Value::Blank).unwrap();
    engine.set_cell_value("Sheet1", "B2", false).unwrap();
    engine.set_cell_value("Sheet1", "B3", false).unwrap();
    engine.set_cell_formula("Sheet1", "E1", "=D3").unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Blank);

    // Expanding the footprint should update E1 via the new spill cell dependency.
    engine.set_cell_value("Sheet1", "B2", true).unwrap();
    engine.set_cell_value("Sheet1", "B3", true).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(3.0));
}

#[test]
fn sort_accepts_arrays_produced_by_expressions() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", true).unwrap();
    engine.set_cell_value("Sheet1", "B2", false).unwrap();
    engine.set_cell_value("Sheet1", "B3", true).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=SORT(FILTER(A1:A3,B1:B3))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
}

#[test]
fn sort_rejects_invalid_order_and_index() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=SORT(A1:A2,0)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Error(ErrorKind::Value)
    );

    engine
        .set_cell_formula("Sheet1", "D1", "=SORT(A1:A2,1,0)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn sort_treats_blank_optional_args_as_defaults() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();

    // Leave sort_order blank to reach by_col argument; Excel treats this as omitted.
    engine
        .set_cell_formula("Sheet1", "C1", "=SORT(A1:A3,1,,FALSE)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));
}

#[test]
fn sort_does_not_treat_blank_cell_references_as_omitted_args() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", Value::Blank).unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", "=SORT(A1:A3,C1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Error(ErrorKind::Value)
    );

    engine
        .set_cell_formula("Sheet1", "F1", "=SORT(A1:A3,1,C1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "F1"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn sortby_sorts_rows_and_columns() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 20.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=SORTBY(A1:A3,B1:B3)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(1.0));

    // Column sort: sort the columns of A1:C1 based on a 1x3 key vector.
    engine.set_cell_value("Sheet1", "F1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "G1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "H1", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "F2", "=SORTBY(F1:H1,{3,1,2})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H2"), Value::Number(1.0));
}

#[test]
fn sortby_accepts_single_cell_sort_order_references() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 200.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 300.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 2.0).unwrap();

    engine.set_cell_value("Sheet1", "C1", -1.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", "=SORTBY(A1:A3,B1:B3,C1)")
        .unwrap();
    engine.recalculate_single_threaded();

    // Sort descending by B (3,2,1) => A values (200,300,100).
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(200.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(300.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(100.0));
}

#[test]
fn sortby_disambiguates_optional_sort_order_args() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 200.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 300.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 2.0).unwrap();

    engine.set_cell_value("Sheet1", "C1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 0.0).unwrap();

    // Excel treats the 3rd argument (C1:C3) as the second by_array (not sort_order1) because it is a range.
    engine
        .set_cell_formula("Sheet1", "E1", "=SORTBY(A1:A3,B1:B3,C1:C3)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(200.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(100.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(300.0));
}

#[test]
fn expand_expands_arrays_with_padding_and_defaults() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={1,2;3,4}")
        .unwrap();
    engine.recalculate_single_threaded();

    engine
        .set_cell_formula("Sheet1", "D1", "=EXPAND(A1:B2,3,4,0)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F3"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G3"), Value::Number(0.0));

    // Omit cols via blank placeholder to reach pad_with, using the default col count (2).
    engine
        .set_cell_formula("Sheet1", "I1", "=EXPAND(A1:B2,3,,0)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "I3"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J3"), Value::Number(0.0));

    // Default pad_with is #N/A.
    engine
        .set_cell_formula("Sheet1", "L1", "=EXPAND(A1:B2,3)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "L3"),
        Value::Error(ErrorKind::NA)
    );

    // Cannot shrink arrays.
    engine
        .set_cell_formula("Sheet1", "N1", "=EXPAND(A1:B2,1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "N1"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn map_applies_lambda_elementwise_across_arrays() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=MAP({1;2;3},LAMBDA(x,x*2))")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(6.0));

    engine
        .set_cell_formula("Sheet1", "C1", "=MAP({1;2;3},{10;20;30},LAMBDA(a,b,a+b))")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(11.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(22.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(33.0));

    // Broadcast a scalar as a 1x1 array.
    engine
        .set_cell_formula("Sheet1", "E1", "=MAP({1;2;3},10,LAMBDA(a,b,a+b))")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(11.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(12.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(13.0));

    // Shape mismatch (3x1 vs 1x3) => #VALUE!
    engine
        .set_cell_formula("Sheet1", "G1", "=MAP({1;2;3},{1,2,3},LAMBDA(a,b,a+b))")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "G1"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn map_preserves_lambda_name_for_recursive_let_bindings() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=LET(FACT,LAMBDA(n,IF(n<=1,1,n*FACT(n-1))),MAP({1;2;3},FACT))",
        )
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(6.0));
}

#[test]
fn map_invokes_lambda_even_when_name_collides_with_builtin_function() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(SUM,LAMBDA(x,x+1),MAP({1;2;3},SUM))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(4.0));
}

#[test]
fn makearray_generates_values_from_indices() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(11.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(12.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(13.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(21.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(22.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(23.0));
}

#[test]
fn byrow_and_bycol_apply_lambda_to_vectors() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={1,2,3;4,5,6}")
        .unwrap();
    engine.recalculate_single_threaded();

    engine
        .set_cell_formula("Sheet1", "E1", "=BYROW(A1:C2,LAMBDA(r,SUM(r)))")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(15.0));

    engine
        .set_cell_formula("Sheet1", "F1", "=BYCOL(A1:C2,LAMBDA(c,SUM(c)))")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(9.0));
}

#[test]
fn reduce_and_scan_accumulate_values() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=REDUCE(0,{1,2,3},LAMBDA(a,v,a+v))")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(6.0));

    // Omit initial_value by providing only (array, lambda).
    engine
        .set_cell_formula("Sheet1", "B1", "=REDUCE({1,2,3},LAMBDA(a,v,a+v))")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));

    engine
        .set_cell_formula("Sheet1", "D1", "=SCAN(0,{1,2,3},LAMBDA(a,v,a+v))")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(6.0));
}

#[test]
fn take_and_drop_slice_arrays() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={1,2,3;4,5,6;7,8,9}")
        .unwrap();
    engine.recalculate_single_threaded();

    // Column-only TAKE: omit rows (blank) to take all rows and select columns.
    engine
        .set_cell_formula("Sheet1", "C5", "=TAKE(A1:C3,,2)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C5"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D5"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C6"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D6"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C7"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D7"), Value::Number(8.0));

    engine
        .set_cell_formula("Sheet1", "E1", "=TAKE(A1:C3,2,-2)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(6.0));

    engine
        .set_cell_formula("Sheet1", "G1", "=TAKE(A1:C3,-1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(8.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(9.0));

    engine
        .set_cell_formula("Sheet1", "E4", "=DROP(A1:C3,1,1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E4"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F4"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E5"), Value::Number(8.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F5"), Value::Number(9.0));

    engine
        .set_cell_formula("Sheet1", "G4", "=DROP(A1:C3,-1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "G4"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H4"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I4"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G5"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H5"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I5"), Value::Number(6.0));
}

#[test]
fn choosecols_and_chooserows_support_negative_indices() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={1,2,3;4,5,6}")
        .unwrap();
    engine.recalculate_single_threaded();

    engine
        .set_cell_formula("Sheet1", "E1", "=CHOOSECOLS(A1:C2,3,1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(4.0));

    engine
        .set_cell_formula("Sheet1", "G1", "=CHOOSECOLS(A1:C2,-1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(6.0));

    engine
        .set_cell_formula("Sheet1", "I1", "=CHOOSEROWS(A1:C2,-1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J1"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "K1"), Value::Number(6.0));

    engine
        .set_cell_formula("Sheet1", "L1", "=CHOOSECOLS(A1:C2,0)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "L1"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn hstack_and_vstack_fill_missing_with_na() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=HSTACK({1;2},{3;4;5})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(5.0));

    engine
        .set_cell_formula("Sheet1", "D1", "=VSTACK({1,2},{3})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "E2"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn tocol_and_torow_flatten_arrays_with_ignore_modes() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TOCOL({1,2;3,4})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(4.0));

    engine
        .set_cell_formula("Sheet1", "C1", "=TOCOL({1,,2},1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));

    engine
        .set_cell_formula("Sheet1", "E1", "=TOCOL({1,#VALUE!,2},2)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(2.0));

    engine
        .set_cell_formula("Sheet1", "G1", "=TOROW({1,2;3,4},0,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();
    // Scan by column when the third argument is TRUE.
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J1"), Value::Number(4.0));
}

#[test]
fn wraprows_and_wrapcols_wrap_vectors_with_padding() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=WRAPROWS({1;2;3;4;5},2)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(5.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "B3"),
        Value::Error(ErrorKind::NA)
    );

    engine
        .set_cell_formula("Sheet1", "D1", "=WRAPCOLS({1;2;3;4;5},2)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(4.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "F2"),
        Value::Error(ErrorKind::NA)
    );

    engine
        .set_cell_formula("Sheet1", "H1", "=WRAPROWS({1;2;3},2,0)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "I2"), Value::Number(0.0));
}

#[test]
fn multithreaded_and_singlethreaded_recalc_match_for_dynamic_arrays() {
    fn setup(engine: &mut Engine) {
        // Data for SORT stability: sort by column B ascending, with ties.
        engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
        engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
        engine.set_cell_value("Sheet1", "A4", 4.0).unwrap();

        engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
        engine.set_cell_value("Sheet1", "B2", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "B3", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "B4", 3.0).unwrap();

        engine
            .set_cell_formula("Sheet1", "D1", "=SORT(A1:B4,2,1,FALSE)")
            .unwrap();

        // Data for UNIQUE.
        engine.set_cell_value("Sheet1", "F1", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "F2", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "F3", 2.0).unwrap();
        engine.set_cell_value("Sheet1", "F4", 3.0).unwrap();
        engine.set_cell_value("Sheet1", "F5", 2.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "H1", "=UNIQUE(F1:F5)")
            .unwrap();
    }

    let mut single = Engine::new();
    setup(&mut single);
    single.recalculate_single_threaded();

    let mut multi = Engine::new();
    setup(&mut multi);
    multi.recalculate_multi_threaded();

    // SORT results should match across modes.
    for addr in ["D1", "E1", "D2", "E2", "D3", "E3", "D4", "E4"] {
        assert_eq!(
            multi.get_cell_value("Sheet1", addr),
            single.get_cell_value("Sheet1", addr),
            "mismatch for {addr}"
        );
    }

    // UNIQUE results should match across modes.
    for addr in ["H1", "H2", "H3"] {
        assert_eq!(
            multi.get_cell_value("Sheet1", addr),
            single.get_cell_value("Sheet1", addr),
            "mismatch for {addr}"
        );
    }
}
