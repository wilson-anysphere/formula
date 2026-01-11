use formula_engine::functions::lookup;
use formula_engine::Engine;
use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

#[test]
fn xmatch_finds_case_insensitive_text() {
    let array = vec![Value::from("A"), Value::from("b"), Value::Number(1.0)];
    assert_eq!(lookup::xmatch(&Value::from("B"), &array).unwrap(), 2);
    assert_eq!(lookup::xmatch(&Value::Number(1.0), &array).unwrap(), 3);
    assert_eq!(
        lookup::xmatch(&Value::from("missing"), &array).unwrap_err(),
        ErrorKind::NA
    );
}

#[test]
fn xmatch_and_xlookup_are_case_insensitive_for_unicode_text() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "Straße");
    sheet.set("B1", 123.0);

    // Uses Unicode-aware uppercasing: ß -> SS.
    assert_eq!(sheet.eval("=XMATCH(\"STRASSE\", A1:A1)"), Value::Number(1.0));
    assert_eq!(sheet.eval("=XLOOKUP(\"STRASSE\", A1:A1, B1:B1)"), Value::Number(123.0));

    // Wildcard mode should also use Unicode-aware case folding.
    assert_eq!(sheet.eval("=XMATCH(\"straß*\", A1:A1, 2)"), Value::Number(1.0));
}

#[test]
fn xlookup_returns_if_not_found_when_provided() {
    let lookup_array = vec![Value::from("A"), Value::from("B")];
    let return_array = vec![Value::Number(10.0), Value::Number(20.0)];

    assert_eq!(
        lookup::xlookup(&Value::from("B"), &lookup_array, &return_array, None).unwrap(),
        Value::Number(20.0)
    );

    assert_eq!(
        lookup::xlookup(
            &Value::from("C"),
            &lookup_array,
            &return_array,
            Some(Value::from("not found"))
        )
        .unwrap(),
        Value::from("not found")
    );
}

#[test]
fn xmatch_and_xlookup_work_in_formulas_and_accept_xlfn_prefix() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("A".to_string()));
    sheet.set("A2", Value::Text("b".to_string()));
    sheet.set("A3", Value::Text("C".to_string()));
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("B3", 30.0);

    assert_eq!(sheet.eval("=XMATCH(\"B\", A1:A3)"), Value::Number(2.0));
    assert_eq!(
        sheet.eval("=_xlfn.XMATCH(\"B\", A1:A3)"),
        Value::Number(2.0)
    );

    assert_eq!(
        sheet.eval("=XLOOKUP(\"B\", A1:A3, B1:B3)"),
        Value::Number(20.0)
    );
    assert_eq!(
        sheet.eval("=_xlfn.XLOOKUP(\"B\", A1:A3, B1:B3)"),
        Value::Number(20.0)
    );

    assert_eq!(
        sheet.eval("=XLOOKUP(\"missing\", A1:A3, B1:B3, \"no\")"),
        Value::Text("no".to_string())
    );
    assert_eq!(
        sheet.eval("=XLOOKUP(\"missing\", A1:A3, B1:B3)"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn xmatch_supports_searching_last_to_first() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 2.0);

    assert_eq!(sheet.eval("=XMATCH(2, A1:A3, 0, 1)"), Value::Number(2.0));
    assert_eq!(sheet.eval("=XMATCH(2, A1:A3, 0, -1)"), Value::Number(3.0));
}

#[test]
fn xmatch_supports_wildcards_and_escapes() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "apple");
    sheet.set("A2", "banana");
    sheet.set("A3", "apricot");

    assert_eq!(sheet.eval("=XMATCH(\"a*\", A1:A3, 2)"), Value::Number(1.0));

    sheet.set("B1", "*");
    sheet.set("B2", "?");
    sheet.set("B3", "~a");
    assert_eq!(sheet.eval("=XMATCH(\"~*\", B1:B3, 2)"), Value::Number(1.0));
    assert_eq!(sheet.eval("=XMATCH(\"~?\", B1:B3, 2)"), Value::Number(2.0));
    // `~` only escapes `*`, `?`, or `~`; otherwise it should be treated literally.
    assert_eq!(sheet.eval("=XMATCH(\"~a\", B1:B3, 2)"), Value::Number(3.0));
}

#[test]
fn xmatch_wildcards_coerce_non_text_candidates_to_text() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 10.0);
    sheet.set("A2", 20.0);
    sheet.set("A3", 30.0);

    assert_eq!(sheet.eval("=XMATCH(\"2*\", A1:A3, 2)"), Value::Number(2.0));
}

#[test]
fn xmatch_supports_next_smaller_and_next_larger() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 3.0);
    sheet.set("A3", 5.0);

    assert_eq!(sheet.eval("=XMATCH(4, A1:A3, -1)"), Value::Number(2.0));
    assert_eq!(sheet.eval("=XMATCH(4, A1:A3, 1)"), Value::Number(3.0));
    assert_eq!(sheet.eval("=XMATCH(0, A1:A3, -1)"), Value::Error(ErrorKind::NA));
}

#[test]
fn xmatch_approximate_modes_treat_blanks_like_zero_or_empty_string() {
    let mut sheet = TestSheet::new();

    // Numeric: blank behaves like 0.
    sheet.set("A1", Value::Blank);
    sheet.set("A2", 1.0);
    sheet.set("A3", 2.0);
    assert_eq!(sheet.eval("=XMATCH(0.5, A1:A3, -1)"), Value::Number(1.0));
    assert_eq!(sheet.eval("=XMATCH(0.5, A1:A3, -1, 2)"), Value::Number(1.0));
    assert_eq!(sheet.eval("=XMATCH(0, A1:A3, -1)"), Value::Number(1.0));
    assert_eq!(sheet.eval("=XMATCH(0, A1:A3, 1)"), Value::Number(1.0));

    // Text: blank behaves like empty string.
    sheet.set("B1", Value::Blank);
    sheet.set("B2", "B");
    sheet.set("B3", "C");
    assert_eq!(sheet.eval("=XMATCH(\"A\", B1:B3, -1)"), Value::Number(1.0));
    assert_eq!(sheet.eval("=XMATCH(\"\", B1:B3, -1)"), Value::Number(1.0));
    assert_eq!(sheet.eval("=XMATCH(\"\", B1:B3, 1)"), Value::Number(1.0));
}

#[test]
fn xmatch_approximate_modes_handle_duplicates_like_sorted_insertion_points() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 2.0);
    sheet.set("A4", 2.0);
    sheet.set("A5", 3.0);

    // Next smaller: insertion point for 2.5 is after the last 2.
    assert_eq!(sheet.eval("=XMATCH(2.5, A1:A5, -1)"), Value::Number(4.0));
    assert_eq!(sheet.eval("=XMATCH(2.5, A1:A5, -1, 2)"), Value::Number(4.0));

    // Next larger: insertion point for 1.5 is before the first 2.
    assert_eq!(sheet.eval("=XMATCH(1.5, A1:A5, 1)"), Value::Number(2.0));
    assert_eq!(sheet.eval("=XMATCH(1.5, A1:A5, 1, 2)"), Value::Number(2.0));
}

#[test]
fn xmatch_binary_search_modes() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 3.0);
    sheet.set("A3", 5.0);
    sheet.set("A4", 7.0);

    assert_eq!(sheet.eval("=XMATCH(4, A1:A4, -1, 2)"), Value::Number(2.0));
    assert_eq!(sheet.eval("=XMATCH(4, A1:A4, 1, 2)"), Value::Number(3.0));
    assert_eq!(sheet.eval("=XMATCH(5, A1:A4, 0, 2)"), Value::Number(3.0));

    sheet.set("B1", 7.0);
    sheet.set("B2", 5.0);
    sheet.set("B3", 3.0);
    sheet.set("B4", 1.0);

    assert_eq!(sheet.eval("=XMATCH(4, B1:B4, -1, -2)"), Value::Number(3.0));
    assert_eq!(sheet.eval("=XMATCH(4, B1:B4, 1, -2)"), Value::Number(2.0));
}

#[test]
fn xlookup_supports_binary_search_mode() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 3.0);
    sheet.set("A3", 5.0);
    sheet.set("A4", 7.0);

    sheet.set("B1", 10.0);
    sheet.set("B2", 30.0);
    sheet.set("B3", 50.0);
    sheet.set("B4", 70.0);

    assert_eq!(
        sheet.eval("=XLOOKUP(4, A1:A4, B1:B4, \"no\", -1, 2)"),
        Value::Number(30.0)
    );
    assert_eq!(
        sheet.eval("=XLOOKUP(4, A1:A4, B1:B4, \"no\", 1, 2)"),
        Value::Number(50.0)
    );
}

#[test]
fn xlookup_supports_searching_last_to_first() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 2.0);
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("B3", 30.0);

    assert_eq!(
        sheet.eval("=XLOOKUP(2, A1:A3, B1:B3, \"no\", 0, 1)"),
        Value::Number(20.0)
    );
    assert_eq!(
        sheet.eval("=XLOOKUP(2, A1:A3, B1:B3, \"no\", 0, -1)"),
        Value::Number(30.0)
    );
}

#[test]
fn xlookup_supports_wildcards_and_reverse_search() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "apple");
    sheet.set("A2", "banana");
    sheet.set("A3", "apricot");
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("B3", 30.0);

    assert_eq!(
        sheet.eval("=XLOOKUP(\"a*\", A1:A3, B1:B3, \"no\", 2, 1)"),
        Value::Number(10.0)
    );
    assert_eq!(
        sheet.eval("=XLOOKUP(\"a*\", A1:A3, B1:B3, \"no\", 2, -1)"),
        Value::Number(30.0)
    );
}

#[test]
fn xlookup_supports_binary_search_mode_descending() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 7.0);
    sheet.set("A2", 5.0);
    sheet.set("A3", 3.0);
    sheet.set("A4", 1.0);
    sheet.set("B1", 70.0);
    sheet.set("B2", 50.0);
    sheet.set("B3", 30.0);
    sheet.set("B4", 10.0);

    assert_eq!(
        sheet.eval("=XLOOKUP(4, A1:A4, B1:B4, \"no\", -1, -2)"),
        Value::Number(30.0)
    );
    assert_eq!(
        sheet.eval("=XLOOKUP(4, A1:A4, B1:B4, \"no\", 1, -2)"),
        Value::Number(50.0)
    );
}

#[test]
fn xlookup_spills_rows_and_columns_from_2d_return_arrays() {
    let mut engine = Engine::new();

    // Vertical lookup_array -> spill the matched row horizontally.
    engine.set_cell_value("Sheet1", "A1", "A").unwrap();
    engine.set_cell_value("Sheet1", "A2", "B").unwrap();
    engine.set_cell_value("Sheet1", "A3", "C").unwrap();

    for (row, base) in [(1, 10.0), (2, 20.0), (3, 30.0)] {
        engine.set_cell_value("Sheet1", &format!("B{row}"), base).unwrap();
        engine.set_cell_value("Sheet1", &format!("C{row}"), base + 1.0).unwrap();
        engine.set_cell_value("Sheet1", &format!("D{row}"), base + 2.0).unwrap();
    }

    engine
        .set_cell_formula("Sheet1", "F1", "=XLOOKUP(\"B\", A1:A3, B1:D3)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(21.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(22.0));

    // Horizontal lookup_array -> spill the matched column vertically.
    engine.set_cell_value("Sheet1", "A5", "A").unwrap();
    engine.set_cell_value("Sheet1", "B5", "B").unwrap();
    engine.set_cell_value("Sheet1", "C5", "C").unwrap();

    // 3x3 return array under the headers in row 5.
    for (row_off, base) in [(0, 100.0), (1, 110.0), (2, 120.0)] {
        let row = 6 + row_off;
        engine.set_cell_value("Sheet1", &format!("A{row}"), base).unwrap();
        engine.set_cell_value("Sheet1", &format!("B{row}"), base + 10.0).unwrap();
        engine.set_cell_value("Sheet1", &format!("C{row}"), base + 20.0).unwrap();
    }

    engine
        .set_cell_formula("Sheet1", "E5", "=XLOOKUP(\"B\", A5:C5, A6:C8)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "E5"), Value::Number(110.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E6"), Value::Number(120.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E7"), Value::Number(130.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E8"), Value::Blank);
}

#[test]
fn xlookup_errors_on_mismatched_shapes_and_invalid_modes() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);

    // return_array is too short.
    assert_eq!(sheet.eval("=XLOOKUP(2, A1:A3, B1:B2)"), Value::Error(ErrorKind::Value));

    // Invalid match/search modes.
    assert_eq!(sheet.eval("=XMATCH(2, A1:A3, 3)"), Value::Error(ErrorKind::Value));
    assert_eq!(sheet.eval("=XMATCH(2, A1:A3, 0, 0)"), Value::Error(ErrorKind::Value));
    assert_eq!(
        sheet.eval("=XLOOKUP(2, A1:A3, A1:A3, \"\", 3)"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=XLOOKUP(2, A1:A3, A1:A3, \"\", 0, 0)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn xmatch_and_xlookup_treat_missing_optional_args_as_defaults() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("B3", 30.0);

    // Missing search_mode should default to 1 (first-to-last), not error.
    assert_eq!(sheet.eval("=XMATCH(2, A1:A3, 0,)"), Value::Number(2.0));

    // Missing if_not_found should still default to #N/A when the lookup is absent.
    assert_eq!(
        sheet.eval("=XLOOKUP(4, A1:A3, B1:B3,,0,1)"),
        Value::Error(ErrorKind::NA)
    );

    // Missing search_mode should default to 1 for XLOOKUP as well.
    assert_eq!(sheet.eval("=XLOOKUP(2, A1:A3, B1:B3,,0,)"), Value::Number(20.0));
}

#[test]
fn vlookup_exact_match_and_errors() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("B1", Value::Text("a".to_string()));
    sheet.set("A2", 2.0);
    sheet.set("B2", Value::Text("b".to_string()));
    sheet.set("A3", 3.0);
    sheet.set("B3", Value::Text("c".to_string()));

    assert_eq!(
        sheet.eval("=VLOOKUP(2, A1:B3, 2, FALSE)"),
        Value::Text("b".to_string())
    );
    assert_eq!(
        sheet.eval("=VLOOKUP(4, A1:B3, 2, FALSE)"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        sheet.eval("=VLOOKUP(2, A1:B3, 3, FALSE)"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn vlookup_approximate_match() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("B1", Value::Text("a".to_string()));
    sheet.set("A2", 3.0);
    sheet.set("B2", Value::Text("b".to_string()));
    sheet.set("A3", 5.0);
    sheet.set("B3", Value::Text("c".to_string()));

    assert_eq!(
        sheet.eval("=VLOOKUP(4, A1:B3, 2)"),
        Value::Text("b".to_string())
    );
    assert_eq!(
        sheet.eval("=VLOOKUP(0, A1:B3, 2)"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn hlookup_exact_match() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("B1", 2.0);
    sheet.set("C1", 3.0);
    sheet.set("A2", Value::Text("a".to_string()));
    sheet.set("B2", Value::Text("b".to_string()));
    sheet.set("C2", Value::Text("c".to_string()));

    assert_eq!(
        sheet.eval("=HLOOKUP(2, A1:C2, 2, FALSE)"),
        Value::Text("b".to_string())
    );
}

#[test]
fn index_and_match() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("A".to_string()));
    sheet.set("B1", Value::Text("b".to_string()));
    sheet.set("C1", Value::Text("C".to_string()));

    assert_eq!(
        sheet.eval("=INDEX(A1:C1,1,2)"),
        Value::Text("b".to_string())
    );
    assert_eq!(sheet.eval("=MATCH(\"B\", A1:C1, 0)"), Value::Number(2.0));

    sheet.set("A2", 1.0);
    sheet.set("A3", 3.0);
    sheet.set("A4", 5.0);
    sheet.set("A5", 7.0);
    assert_eq!(sheet.eval("=MATCH(4, A2:A5, 1)"), Value::Number(2.0));

    sheet.set("B2", 7.0);
    sheet.set("B3", 5.0);
    sheet.set("B4", 3.0);
    sheet.set("B5", 1.0);
    assert_eq!(sheet.eval("=MATCH(4, B2:B5, -1)"), Value::Number(2.0));
    assert_eq!(
        sheet.eval("=MATCH(11, B2:B5, -1)"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn choose_selects_index_supports_arrays_and_range_unions() {
    let mut sheet = TestSheet::new();

    assert_eq!(sheet.eval("=CHOOSE(1, 10, 1/0)"), Value::Number(10.0));
    assert_eq!(sheet.eval("=CHOOSE(2, 1/0, 20)"), Value::Number(20.0));
    assert_eq!(
        sheet.eval("=CHOOSE(0, 1, 2)"),
        Value::Error(ErrorKind::Value)
    );

    sheet.set("A1", 10.0);
    sheet.set("B1", 20.0);
    sheet.set_formula("C1", "=CHOOSE({1,2}, A1, B1)");
    sheet.recalculate();
    assert_eq!(sheet.get("C1"), Value::Number(10.0));
    assert_eq!(sheet.get("D1"), Value::Number(20.0));

    sheet.set("A2", 1.0);
    sheet.set("A3", 2.0);
    sheet.set("A4", 3.0);
    sheet.set("B2", 10.0);
    sheet.set("B3", 20.0);
    sheet.set("B4", 30.0);
    assert_eq!(
        sheet.eval("=SUM(CHOOSE({1,2}, A2:A4, B2:B4))"),
        Value::Number(66.0)
    );
}

#[test]
fn getpivotdata_returns_values_from_tabular_pivot_output() {
    let mut sheet = TestSheet::new();

    // Simulated pivot-engine output (tabular layout, 1 row field, 1 value field).
    sheet.set("A1", "Region");
    sheet.set("B1", "Sum of Sales");
    sheet.set("A2", "East");
    sheet.set("B2", 250.0);
    sheet.set("A3", "West");
    sheet.set("B3", 450.0);
    sheet.set("A4", "Grand Total");
    sheet.set("B4", 700.0);

    // pivot_table reference can point anywhere inside the pivot.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", B2, \"Region\", \"East\")"),
        Value::Number(250.0)
    );

    // When no field/item pairs are provided, return the grand total.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1)"),
        Value::Number(700.0)
    );
}

#[test]
fn getpivotdata_supports_multiple_row_fields() {
    let mut sheet = TestSheet::new();

    // Simulated pivot-engine output (tabular layout, 2 row fields, 1 value field).
    sheet.set("A1", "Region");
    sheet.set("B1", "Product");
    sheet.set("C1", "Sum of Sales");

    sheet.set("A2", "East");
    sheet.set("B2", "A");
    sheet.set("C2", 100.0);
    sheet.set("A3", "East");
    sheet.set("B3", "B");
    sheet.set("C3", 150.0);
    sheet.set("A4", "West");
    sheet.set("B4", "A");
    sheet.set("C4", 200.0);
    sheet.set("A5", "West");
    sheet.set("B5", "B");
    sheet.set("C5", 250.0);
    sheet.set("A6", "Grand Total");
    sheet.set("C6", 700.0);

    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A3, \"Region\", \"West\", \"Product\", \"A\")"),
        Value::Number(200.0)
    );
}

#[test]
fn getpivotdata_errors() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", "Region");
    sheet.set("B1", "Sum of Sales");
    sheet.set("A2", "East");
    sheet.set("B2", 250.0);

    // pivot_table must be a reference.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", 1, \"Region\", \"East\")"),
        Value::Error(ErrorKind::Value)
    );

    // Field/item pairs must be complete.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\")"),
        Value::Error(ErrorKind::Value)
    );

    // Unknown field -> #REF!
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Product\", \"A\")"),
        Value::Error(ErrorKind::Ref)
    );

    // Unknown item -> #N/A
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"Missing\")"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn getpivotdata_rejects_column_fields_mvp() {
    let mut sheet = TestSheet::new();

    // Simulated pivot-engine output with a column field (headers are "A - ..." / "B - ...").
    sheet.set("A1", "Region");
    sheet.set("B1", "A - Sum of Sales");
    sheet.set("C1", "B - Sum of Sales");
    sheet.set("A2", "East");
    sheet.set("B2", 100.0);
    sheet.set("C2", 150.0);

    // The MVP only supports pivots with no column fields.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"A - Sum of Sales\", A2, \"Region\", \"East\")"),
        Value::Error(ErrorKind::Ref)
    );
}
