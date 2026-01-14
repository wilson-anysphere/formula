use formula_engine::functions::lookup;
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::value::EntityValue;
use formula_engine::Engine;
use formula_engine::{ErrorKind, Value};
use formula_model::EXCEL_MAX_COLS;

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
fn lookup_matches_rich_values_by_display_string() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Entity(EntityValue::new("Apple")));
    sheet.set("B1", 42.0);
    assert_eq!(
        sheet.eval("=VLOOKUP(\"apple\", A1:B1, 2, FALSE)"),
        Value::Number(42.0)
    );
}

#[test]
fn xmatch_and_xlookup_are_case_insensitive_for_unicode_text() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "Straße");
    sheet.set("B1", 123.0);

    // Uses Unicode-aware uppercasing: ß -> SS.
    assert_eq!(
        sheet.eval("=XMATCH(\"STRASSE\", A1:A1)"),
        Value::Number(1.0)
    );
    assert_eq!(
        sheet.eval("=XLOOKUP(\"STRASSE\", A1:A1, B1:B1)"),
        Value::Number(123.0)
    );

    // Wildcard mode should also use Unicode-aware case folding.
    assert_eq!(
        sheet.eval("=XMATCH(\"straß*\", A1:A1, 2)"),
        Value::Number(1.0)
    );
}

#[test]
fn xmatch_and_xlookup_coerce_numeric_text_via_value() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "1,234.5");
    sheet.set("B1", 42.0);

    assert_eq!(sheet.eval("=XMATCH(1234.5, A1:A1)"), Value::Number(1.0));
    assert_eq!(
        sheet.eval("=XLOOKUP(1234.5, A1:A1, B1:B1)"),
        Value::Number(42.0)
    );
}

#[test]
fn match_and_vlookup_are_case_insensitive_for_unicode_text() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "Straße");
    sheet.set("B1", 42.0);

    assert_eq!(
        sheet.eval("=MATCH(\"STRASSE\", A1:A1, 0)"),
        Value::Number(1.0)
    );
    assert_eq!(
        sheet.eval("=VLOOKUP(\"STRASSE\", A1:B1, 2, FALSE)"),
        Value::Number(42.0)
    );
}

#[test]
fn match_and_vlookup_support_wildcard_exact_matching() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "apple");
    sheet.set("A2", "banana");
    sheet.set("A3", "*");
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("B3", 30.0);

    assert_eq!(sheet.eval("=MATCH(\"b*\", A1:A3, 0)"), Value::Number(2.0));
    assert_eq!(
        sheet.eval("=VLOOKUP(\"b*\", A1:B3, 2, FALSE)"),
        Value::Number(20.0)
    );

    // `~` escapes wildcards in lookup patterns.
    assert_eq!(sheet.eval("=MATCH(\"~*\", A1:A3, 0)"), Value::Number(3.0));
    assert_eq!(
        sheet.eval("=VLOOKUP(\"~*\", A1:B3, 2, FALSE)"),
        Value::Number(30.0)
    );
}

#[test]
fn match_and_vlookup_coerce_numeric_text_via_value() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "1,234.5");
    sheet.set("B1", 42.0);

    assert_eq!(sheet.eval("=MATCH(1234.5, A1:A1, 0)"), Value::Number(1.0));
    assert_eq!(
        sheet.eval("=VLOOKUP(1234.5, A1:B1, 2, FALSE)"),
        Value::Number(42.0)
    );
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
fn lookup_wildcard_numeric_text_coercion_respects_value_locale() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());
    sheet.set("A1", 1.5);
    sheet.set("A2", 2.0);

    // MATCH (exact match_type=0) applies wildcard matching when the lookup value contains
    // wildcards, and coerces non-text candidates to text first.
    assert_eq!(sheet.eval("=MATCH(\"*,5\", A1:A2, 0)"), Value::Number(1.0));

    // XMATCH wildcard mode should also coerce non-text candidates using the workbook value locale.
    assert_eq!(sheet.eval("=XMATCH(\"1,5\", A1:A2, 2)"), Value::Number(1.0));
}

#[test]
fn lookup_numeric_text_parsing_respects_value_locale_for_number_equality() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);
        engine.set_value_locale(ValueLocaleConfig::de_de());

        // MATCH should coerce numeric text to numbers using the workbook value locale.
        engine.set_cell_value("Sheet1", "A1", "1,5").unwrap();
        engine
            .set_cell_formula("Sheet1", "Z1", "=MATCH(1.5, A1:A1, 0)")
            .unwrap();

        // VLOOKUP should do the same for exact matches.
        engine.set_cell_value("Sheet1", "A2", "1.234,5").unwrap();
        engine.set_cell_value("Sheet1", "B2", 42.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "Z2", "=VLOOKUP(1234.5, A2:B2, 2, FALSE)")
            .unwrap();

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(1.0));
        assert_eq!(engine.get_cell_value("Sheet1", "Z2"), Value::Number(42.0));

        if bytecode_enabled {
            assert!(
                engine.bytecode_program_count() > 0,
                "expected lookup formulas to compile to bytecode for this test"
            );
        }
    }

    // XMATCH/XLOOKUP use their own lookup engine; ensure it also respects the workbook value locale
    // when comparing numbers to numeric text.
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());
    sheet.set("A1", "1.234,5");
    sheet.set("B1", 99.0);
    assert_eq!(sheet.eval("=XMATCH(1234.5, A1:A1)"), Value::Number(1.0));
    assert_eq!(
        sheet.eval("=XLOOKUP(1234.5, A1:A1, B1:B1)"),
        Value::Number(99.0)
    );
}

#[test]
fn xmatch_supports_next_smaller_and_next_larger() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 3.0);
    sheet.set("A3", 5.0);

    assert_eq!(sheet.eval("=XMATCH(4, A1:A3, -1)"), Value::Number(2.0));
    assert_eq!(sheet.eval("=XMATCH(4, A1:A3, 1)"), Value::Number(3.0));
    assert_eq!(
        sheet.eval("=XMATCH(0, A1:A3, -1)"),
        Value::Error(ErrorKind::NA)
    );
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

    // Exact match should follow the same insertion-point semantics for duplicates:
    // - match_mode=-1 chooses the last occurrence (after duplicates)
    // - match_mode=1 chooses the first occurrence (before duplicates)
    assert_eq!(sheet.eval("=XMATCH(2, A1:A5, -1)"), Value::Number(4.0));
    assert_eq!(sheet.eval("=XMATCH(2, A1:A5, -1, 2)"), Value::Number(4.0));
    assert_eq!(sheet.eval("=XMATCH(2, A1:A5, 1)"), Value::Number(2.0));
    assert_eq!(sheet.eval("=XMATCH(2, A1:A5, 1, 2)"), Value::Number(2.0));

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
fn xmatch_binary_search_orders_errors_by_numeric_error_code() {
    let array = vec![
        Value::Error(ErrorKind::Null),
        Value::Error(ErrorKind::Div0),
        Value::Error(ErrorKind::Value),
    ];
    assert_eq!(
        lookup::xmatch_with_modes(
            &Value::Error(ErrorKind::Div0),
            &array,
            lookup::MatchMode::Exact,
            lookup::SearchMode::BinaryAscending
        )
        .unwrap(),
        2
    );
}

#[test]
fn xmatch_binary_search_orders_extended_errors_by_numeric_error_code() {
    let array = vec![
        Value::Error(ErrorKind::Field),
        Value::Error(ErrorKind::Connect),
        Value::Error(ErrorKind::Blocked),
    ];
    assert_eq!(
        lookup::xmatch_with_modes(
            &Value::Error(ErrorKind::Connect),
            &array,
            lookup::MatchMode::Exact,
            lookup::SearchMode::BinaryAscending
        )
        .unwrap(),
        2
    );
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
        engine
            .set_cell_value("Sheet1", &format!("B{row}"), base)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("C{row}"), base + 1.0)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("D{row}"), base + 2.0)
            .unwrap();
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
        engine
            .set_cell_value("Sheet1", &format!("A{row}"), base)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("B{row}"), base + 10.0)
            .unwrap();
        engine
            .set_cell_value("Sheet1", &format!("C{row}"), base + 20.0)
            .unwrap();
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
fn xlookup_returns_spill_for_huge_matched_column() {
    // When sheet dimensions grow beyond the array materialization cap, XLOOKUP should fail fast
    // with `#SPILL!` rather than attempting to allocate an enormous output vector.
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 5_000_002, EXCEL_MAX_COLS)
        .unwrap();

    engine.set_cell_value("Sheet1", "A1", "A").unwrap();
    engine.set_cell_value("Sheet1", "B1", "B").unwrap();

    // Horizontal lookup array (A1:B1) with a very tall return array. Looking up "B" returns the
    // matched column (rows x 1), which exceeds the cap.
    engine
        .set_cell_formula("Sheet1", "D1", "=XLOOKUP(\"B\", A1:B1, A2:B5000002)")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Spill)
    );
    assert!(engine.spill_range("Sheet1", "D1").is_none());
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
    assert_eq!(
        sheet.eval("=XLOOKUP(2, A1:A3, B1:B2)"),
        Value::Error(ErrorKind::Value)
    );

    // Invalid match/search modes.
    assert_eq!(
        sheet.eval("=XMATCH(2, A1:A3, 3)"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=XMATCH(2, A1:A3, 0, 0)"),
        Value::Error(ErrorKind::Value)
    );
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
fn xmatch_and_xlookup_reject_2d_lookup_arrays() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("B1", 2.0);
    sheet.set("A2", 3.0);
    sheet.set("B2", 4.0);

    assert_eq!(
        sheet.eval("=XMATCH(2, A1:B2)"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=XLOOKUP(2, A1:B2, A1:A2)"),
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
    assert_eq!(
        sheet.eval("=XLOOKUP(2, A1:A3, B1:B3,,0,)"),
        Value::Number(20.0)
    );
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
fn vlookup_and_hlookup_compile_to_bytecode_backend_for_simple_range_tables() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("B1", "a");
    sheet.set("B2", "b");
    sheet.set("B3", "c");

    sheet.set("D1", 1.0);
    sheet.set("E1", 2.0);
    sheet.set("F1", 3.0);
    sheet.set("D2", "a");
    sheet.set("E2", "b");
    sheet.set("F2", "c");

    sheet.set_formula("C1", "=VLOOKUP(2, A1:B3, 2, FALSE)");
    sheet.set_formula("C2", "=HLOOKUP(2, D1:F2, 2, FALSE)");

    assert_eq!(sheet.bytecode_program_count(), 2);

    sheet.recalculate();

    assert_eq!(sheet.get("C1"), Value::Text("b".to_string()));
    assert_eq!(sheet.get("C2"), Value::Text("b".to_string()));
}

#[test]
fn lookup_functions_compile_to_bytecode_backend_with_let_bound_ranges() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("B1", "a");
    sheet.set("B2", "b");
    sheet.set("B3", "c");

    sheet.set("D1", 1.0);
    sheet.set("E1", 2.0);
    sheet.set("F1", 3.0);
    sheet.set("D2", "a");
    sheet.set("E2", "b");
    sheet.set("F2", "c");

    sheet.set_formula("C1", "=LET(t, A1:B3, VLOOKUP(2, t, 2, FALSE))");
    sheet.set_formula("C2", "=LET(t, D1:F2, HLOOKUP(2, t, 2, FALSE))");
    sheet.set_formula("C3", "=LET(a, A1:A3, MATCH(2, a, 0))");
    sheet.set_formula("C4", "=LET(a, A1:A3, XMATCH(2, a))");
    sheet.set_formula("C5", "=LET(a, A1:A3, b, B1:B3, XLOOKUP(2, a, b))");

    assert_eq!(sheet.bytecode_program_count(), 5);

    sheet.recalculate();

    assert_eq!(sheet.get("C1"), Value::Text("b".to_string()));
    assert_eq!(sheet.get("C2"), Value::Text("b".to_string()));
    assert_eq!(sheet.get("C3"), Value::Number(2.0));
    assert_eq!(sheet.get("C4"), Value::Number(2.0));
    assert_eq!(sheet.get("C5"), Value::Text("b".to_string()));
}

#[test]
fn lookup_functions_compile_to_bytecode_backend_with_spill_range_tables() {
    let mut sheet = TestSheet::new();

    // Build spilled tables with dynamic array formulas; lookups refer to them via the `#` spill
    // operator, which should still be eligible for bytecode lowering.
    sheet.set_formula("A1", "=SEQUENCE(3,2)"); // spills into A1:B3.
    sheet.set_formula("A5", "=SEQUENCE(2,3)"); // spills into A5:C6.
    sheet.set_formula("A10", "=SEQUENCE(3)"); // spills into A10:A12.

    let base_programs = sheet.bytecode_program_count();

    sheet.set_formula("D1", "=VLOOKUP(3, A1#, 2, FALSE)");
    sheet.set_formula("D2", "=HLOOKUP(2, A5#, 2, FALSE)");
    sheet.set_formula("D3", "=MATCH(2, A10#, 0)");
    sheet.set_formula("D4", "=XMATCH(2, A10#)");
    sheet.set_formula("D5", "=XLOOKUP(2, A10#, A10#)");

    assert_eq!(sheet.bytecode_program_count(), base_programs + 5);

    sheet.recalculate();

    assert_eq!(sheet.get("D1"), Value::Number(4.0));
    assert_eq!(sheet.get("D2"), Value::Number(5.0));
    assert_eq!(sheet.get("D3"), Value::Number(2.0));
    assert_eq!(sheet.get("D4"), Value::Number(2.0));
    assert_eq!(sheet.get("D5"), Value::Number(2.0));
}

#[test]
fn vlookup_propagates_spill_range_errors_in_table_array_and_compiles_to_bytecode_backend() {
    let mut sheet = TestSheet::new();

    // A1 is not a spill origin, so `A1#` is a `#REF!` error. VLOOKUP should propagate it from the
    // table_array argument.
    sheet.set("A1", 1.0);

    let base_programs = sheet.bytecode_program_count();
    sheet.set_formula("B1", "=VLOOKUP(1, A1#, 1, FALSE)");

    assert_eq!(sheet.bytecode_program_count(), base_programs + 1);

    sheet.recalculate();
    assert_eq!(sheet.get("B1"), Value::Error(ErrorKind::Ref));
}

#[test]
fn xmatch_and_xlookup_propagate_spill_range_errors_and_compile_to_bytecode_backend() {
    let mut sheet = TestSheet::new();

    // A1 is not a spill origin, so `A1#` is a `#REF!` error. XMATCH/XLOOKUP should propagate it
    // from their lookup_array / return_array arguments.
    sheet.set("A1", 1.0);

    let base_programs = sheet.bytecode_program_count();
    sheet.set_formula("B1", "=XMATCH(1, A1#)");
    sheet.set_formula("B2", "=XLOOKUP(1, A1#, A1#)");

    assert_eq!(sheet.bytecode_program_count(), base_programs + 2);

    sheet.recalculate();
    assert_eq!(sheet.get("B1"), Value::Error(ErrorKind::Ref));
    assert_eq!(sheet.get("B2"), Value::Error(ErrorKind::Ref));
}

#[test]
fn lookup_functions_compile_to_bytecode_backend_for_array_literal_tables() {
    let mut sheet = TestSheet::new();

    sheet.set_formula("A1", "=VLOOKUP(2, {1,\"a\";2,\"b\";3,\"c\"}, 2, FALSE)");
    sheet.set_formula("A2", "=HLOOKUP(2, {1,2,3;10,20,30}, 2, FALSE)");
    sheet.set_formula("A3", "=MATCH(2, {1;2;3}, 0)");

    assert_eq!(sheet.bytecode_program_count(), 3);

    sheet.recalculate();

    assert_eq!(sheet.get("A1"), Value::Text("b".to_string()));
    assert_eq!(sheet.get("A2"), Value::Number(20.0));
    assert_eq!(sheet.get("A3"), Value::Number(2.0));
}

#[test]
fn lookup_functions_compile_to_bytecode_backend_with_let_bound_arrays() {
    let mut sheet = TestSheet::new();

    sheet.set_formula(
        "A1",
        "=LET(t, {1,\"a\";2,\"b\";3,\"c\"}, VLOOKUP(2, t, 2, FALSE))",
    );
    sheet.set_formula("A2", "=LET(t, {1,2,3;10,20,30}, HLOOKUP(2, t, 2, FALSE))");
    sheet.set_formula("A3", "=LET(a, {1;2;3}, MATCH(2, a, 0))");

    assert_eq!(sheet.bytecode_program_count(), 3);

    sheet.recalculate();

    assert_eq!(sheet.get("A1"), Value::Text("b".to_string()));
    assert_eq!(sheet.get("A2"), Value::Number(20.0));
    assert_eq!(sheet.get("A3"), Value::Number(2.0));
}

#[test]
fn lookup_functions_compile_to_bytecode_backend_for_computed_array_tables() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("B3", 30.0);

    sheet.set("D1", 1.0);
    sheet.set("E1", 2.0);
    sheet.set("F1", 3.0);
    sheet.set("D2", 10.0);
    sheet.set("E2", 20.0);
    sheet.set("F2", 30.0);

    sheet.set_formula("C1", "=VLOOKUP(2, A1:B3*1, 2, FALSE)");
    sheet.set_formula("C2", "=HLOOKUP(2, D1:F2*1, 2, FALSE)");
    sheet.set_formula("C3", "=MATCH(20, A1:A3*10, 0)");

    assert_eq!(sheet.bytecode_program_count(), 3);

    sheet.recalculate();

    assert_eq!(sheet.get("C1"), Value::Number(20.0));
    assert_eq!(sheet.get("C2"), Value::Number(20.0));
    assert_eq!(sheet.get("C3"), Value::Number(2.0));
}

#[test]
fn match_and_vlookup_approximate_treat_blanks_like_zero_or_empty_string() {
    let mut sheet = TestSheet::new();

    // Numeric: blank behaves like 0.
    sheet.set("A1", Value::Blank);
    sheet.set("A2", 1.0);
    sheet.set("A3", 2.0);
    sheet.set("B1", "zero");
    sheet.set("B2", "one");
    sheet.set("B3", "two");
    assert_eq!(sheet.eval("=MATCH(0.5, A1:A3, 1)"), Value::Number(1.0));
    assert_eq!(sheet.eval("=MATCH(0, A1:A3, 1)"), Value::Number(1.0));
    assert_eq!(
        sheet.eval("=VLOOKUP(0.5, A1:B3, 2)"),
        Value::Text("zero".to_string())
    );

    // Text: blank behaves like empty string.
    sheet.set("C1", Value::Blank);
    sheet.set("C2", "B");
    sheet.set("C3", "C");
    assert_eq!(sheet.eval("=MATCH(\"A\", C1:C3, 1)"), Value::Number(1.0));
    assert_eq!(sheet.eval("=MATCH(\"\", C1:C3, 1)"), Value::Number(1.0));
}

#[test]
fn match_and_vlookup_approximate_handle_sorted_mixed_type_arrays() {
    let mut sheet = TestSheet::new();
    // Excel sorts numbers before text, so this is a valid ascending order for approximate match.
    sheet.set("A1", 1.0);
    sheet.set("A2", 3.0);
    sheet.set("A3", "A");
    sheet.set("B1", 10.0);
    sheet.set("B2", 30.0);
    sheet.set("B3", 40.0);

    assert_eq!(sheet.eval("=MATCH(2, A1:A3, 1)"), Value::Number(1.0));
    assert_eq!(sheet.eval("=MATCH(4, A1:A3, 1)"), Value::Number(2.0));
    assert_eq!(sheet.eval("=VLOOKUP(2, A1:B3, 2)"), Value::Number(10.0));
    assert_eq!(sheet.eval("=VLOOKUP(4, A1:B3, 2)"), Value::Number(30.0));
}

#[test]
fn match_and_vlookup_approximate_handle_duplicates_like_sorted_insertion_points() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 2.0);
    sheet.set("A4", 2.0);
    sheet.set("A5", 3.0);
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("B3", 30.0);
    sheet.set("B4", 40.0);
    sheet.set("B5", 50.0);

    // Ascending approximate: insertion point for 2.5 is after the last 2.
    assert_eq!(sheet.eval("=MATCH(2.5, A1:A5, 1)"), Value::Number(4.0));
    assert_eq!(sheet.eval("=MATCH(2, A1:A5, 1)"), Value::Number(4.0));
    assert_eq!(sheet.eval("=VLOOKUP(2.5, A1:B5, 2)"), Value::Number(40.0));
    assert_eq!(sheet.eval("=VLOOKUP(2, A1:B5, 2)"), Value::Number(40.0));

    // Descending approximate: insertion point for 2 is after the last 2.
    sheet.set("C1", 3.0);
    sheet.set("C2", 2.0);
    sheet.set("C3", 2.0);
    sheet.set("C4", 2.0);
    sheet.set("C5", 1.0);
    assert_eq!(sheet.eval("=MATCH(2, C1:C5, -1)"), Value::Number(4.0));
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
fn hlookup_supports_wildcard_exact_matching() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "apple");
    sheet.set("B1", "banana");
    sheet.set("C1", "*");
    sheet.set("A2", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("C2", 30.0);

    assert_eq!(
        sheet.eval("=HLOOKUP(\"b*\", A1:C2, 2, FALSE)"),
        Value::Number(20.0)
    );
    assert_eq!(
        sheet.eval("=HLOOKUP(\"~*\", A1:C2, 2, FALSE)"),
        Value::Number(30.0)
    );
}

#[test]
fn lookup_vector_form_is_exact_or_next_smaller_and_returns_last_duplicate() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 2.0);
    sheet.set("A4", 3.0);
    sheet.set("B1", "one");
    sheet.set("B2", "two-a");
    sheet.set("B3", "two-b");
    sheet.set("B4", "three");

    // Exact match returns last duplicate.
    assert_eq!(
        sheet.eval("=LOOKUP(2, A1:A4, B1:B4)"),
        Value::Text("two-b".to_string())
    );

    // Approximate match (next smaller).
    assert_eq!(
        sheet.eval("=LOOKUP(2.5, A1:A4, B1:B4)"),
        Value::Text("two-b".to_string())
    );

    // Out of range low -> #N/A.
    assert_eq!(
        sheet.eval("=LOOKUP(0, A1:A4, B1:B4)"),
        Value::Error(ErrorKind::NA)
    );

    // Out of range high -> last element.
    assert_eq!(
        sheet.eval("=LOOKUP(10, A1:A4, B1:B4)"),
        Value::Text("three".to_string())
    );

    // Missing result_vector defaults to lookup_vector.
    assert_eq!(sheet.eval("=LOOKUP(2.5, A1:A4)"), Value::Number(2.0));
}

#[test]
fn lookup_errors_on_mismatched_vector_lengths() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("B1", 10.0);

    assert_eq!(
        sheet.eval("=LOOKUP(2, A1:A2, B1:B1)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn lookup_array_form_searches_first_column_for_tall_arrays() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("A4", 4.0);
    sheet.set("B1", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("B3", 30.0);
    sheet.set("B4", 40.0);

    // Array is 4x2, so LOOKUP searches A1:A4 and returns from B1:B4.
    assert_eq!(sheet.eval("=LOOKUP(3.5, A1:B4)"), Value::Number(30.0));
}

#[test]
fn lookup_array_form_searches_first_row_for_wide_arrays() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("B1", 2.0);
    sheet.set("C1", 3.0);
    sheet.set("D1", 4.0);
    sheet.set("A2", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("C2", 30.0);
    sheet.set("D2", 40.0);

    // Array is 2x4, so LOOKUP searches A1:D1 and returns from A2:D2.
    assert_eq!(sheet.eval("=LOOKUP(3.5, A1:D2)"), Value::Number(30.0));
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
fn index_returns_references_and_spills_rows_or_columns_when_zero() {
    let mut engine = Engine::new();

    // 3x3 array in A1:C3.
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 200.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 300.0).unwrap();

    // row_num = 0: return the entire column.
    engine
        .set_cell_formula("Sheet1", "E1", "=INDEX(A1:C3, 0, 2)")
        .unwrap();
    // col_num = 0: return the entire row.
    engine
        .set_cell_formula("Sheet1", "E5", "=INDEX(A1:C3, 2, 0)")
        .unwrap();
    // INDEX returns references (so OFFSET can consume them).
    engine
        .set_cell_formula("Sheet1", "H1", "=OFFSET(INDEX(A1:C3, 1, 2), 1, 0)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(200.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E4"), Value::Blank);

    assert_eq!(engine.get_cell_value("Sheet1", "E5"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F5"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G5"), Value::Number(30.0));

    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(20.0));
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
fn getpivotdata_returns_values_from_compact_pivot_output() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotCache, PivotConfig, PivotEngine, PivotField,
        PivotValue, SubtotalPosition, ValueField,
    };

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    fn to_cell_value(v: &PivotValue) -> Value {
        match v {
            PivotValue::Blank => Value::Blank,
            PivotValue::Number(n) => Value::Number(*n),
            PivotValue::Text(s) => Value::Text(s.clone()),
            PivotValue::Bool(b) => Value::Bool(*b),
            PivotValue::Date(d) => Value::Text(d.to_string()),
        }
    }

    // Compute a compact-layout pivot using the engine.
    let data: Vec<Vec<PivotValue>> = vec![
        pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
        pv_row(&["East".into(), "A".into(), 100.into()]),
        pv_row(&["East".into(), "B".into(), 150.into()]),
        pv_row(&["West".into(), "A".into(), 200.into()]),
        pv_row(&["West".into(), "B".into(), 250.into()]),
    ];
    let cache = PivotCache::from_range(&data).unwrap();
    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".into(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Compact,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: false,
        },
    };
    let result = PivotEngine::calculate(&cache, &cfg).unwrap();

    // Write the pivot output into a worksheet.
    let mut sheet = TestSheet::new();
    for (r, row) in result.data.iter().enumerate() {
        for (c, value) in row.iter().enumerate() {
            let addr = formula_engine::eval::CellAddr {
                row: r as u32,
                col: c as u32,
            }
            .to_a1();
            sheet.set(&addr, to_cell_value(value));
        }
    }

    // Row items live in a single "Row Labels" column under Compact layout.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", B2, \"Row Labels\", \"East\")"),
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
fn getpivotdata_is_case_insensitive_for_unicode_fields_items_and_values() {
    let mut sheet = TestSheet::new();

    // Simulated pivot-engine output (tabular layout, 1 row field, 1 value field) with Unicode
    // identifiers that require Unicode-aware case folding (ß -> SS).
    sheet.set("A1", "Straße");
    sheet.set("B1", "Sum of Maß");
    sheet.set("A2", "Maß");
    sheet.set("B2", 10.0);
    sheet.set("A3", "Grand Total");
    sheet.set("B3", 10.0);

    // Ensure we can reference:
    // - the value header "Sum of Maß" as "SUM OF MASS"
    // - the field header "Straße" as "STRASSE"
    // - the item "Maß" as "MASS"
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"SUM OF MASS\", A1, \"STRASSE\", \"MASS\")"),
        Value::Number(10.0)
    );
}

#[test]
fn getpivotdata_supports_column_fields() {
    let mut sheet = TestSheet::new();

    // Simulated pivot-engine output with a column field (headers are "A - ..." / "B - ...")
    // plus a row grand total column.
    sheet.set("A1", "Region");
    sheet.set("B1", "A - Sum of Sales");
    sheet.set("C1", "B - Sum of Sales");
    sheet.set("D1", "Grand Total - Sum of Sales");

    sheet.set("A2", "East");
    sheet.set("B2", 100.0);
    sheet.set("C2", 150.0);
    sheet.set("D2", 250.0);

    sheet.set("A3", "West");
    sheet.set("B3", 200.0);
    sheet.set("C3", 250.0);
    sheet.set("D3", 450.0);

    sheet.set("A4", "Grand Total");
    sheet.set("B4", 300.0);
    sheet.set("C4", 400.0);
    sheet.set("D4", 700.0);

    // `data_field` can refer directly to a rendered header.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"A - Sum of Sales\", A2, \"Region\", \"East\")"),
        Value::Number(100.0)
    );

    // Or a base value field name with a column item criterion.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A2, \"Region\", \"East\", \"Product\", \"A\")"),
        Value::Number(100.0)
    );

    // If no column item is specified, return the row grand total for that value field.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A2, \"Region\", \"East\")"),
        Value::Number(250.0)
    );

    // Column totals can be queried by omitting row criteria.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A2, \"Product\", \"A\")"),
        Value::Number(300.0)
    );

    // When no criteria are provided, return the overall grand total.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A2)"),
        Value::Number(700.0)
    );
}

#[test]
fn getpivotdata_column_item_criteria_is_case_insensitive_for_unicode_text() {
    let mut sheet = TestSheet::new();

    // Simulated pivot-engine output with a column field item containing ß ("Maß").
    sheet.set("A1", "Region");
    sheet.set("B1", "Maß - Sum of Sales");
    sheet.set("C1", "Grand Total - Sum of Sales");

    sheet.set("A2", "East");
    sheet.set("B2", 100.0);
    sheet.set("C2", 100.0);

    sheet.set("A3", "Grand Total");
    sheet.set("B3", 100.0);
    sheet.set("C3", 100.0);

    // Column-field pivots do not render the column field name in the output grid, so the scan-based
    // GETPIVOTDATA implementation treats unknown fields as column criteria. Use "Product" as a
    // placeholder field name (matching other tests); the important part is the item match.
    assert_eq!(
        sheet.eval(
            "=GETPIVOTDATA(\"Sum of Sales\", A2, \"Region\", \"East\", \"Product\", \"MASS\")"
        ),
        Value::Number(100.0)
    );
}

#[test]
fn getpivotdata_supports_multiple_value_fields() {
    let mut sheet = TestSheet::new();

    // Simulated pivot-engine output with 2 value fields.
    sheet.set("A1", "Region");
    sheet.set("B1", "Sum of Sales");
    sheet.set("C1", "Count of Sales");

    sheet.set("A2", "East");
    sheet.set("B2", 250.0);
    sheet.set("C2", 2.0);

    sheet.set("A3", "West");
    sheet.set("B3", 450.0);
    sheet.set("C3", 2.0);

    sheet.set("A4", "Grand Total");
    sheet.set("B4", 700.0);
    sheet.set("C4", 4.0);

    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Number(250.0)
    );
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Count of Sales\", A1, \"Region\", \"East\")"),
        Value::Number(2.0)
    );
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Count of Sales\", A1)"),
        Value::Number(4.0)
    );
}

#[test]
fn getpivotdata_supports_column_fields_and_multiple_value_fields() {
    let mut sheet = TestSheet::new();

    // Simulated pivot-engine output with:
    // - row field: Region
    // - column field: Product (A/B)
    // - value fields: Sum of Sales, Count of Sales
    sheet.set("A1", "Region");
    sheet.set("B1", "A - Sum of Sales");
    sheet.set("C1", "A - Count of Sales");
    sheet.set("D1", "B - Sum of Sales");
    sheet.set("E1", "B - Count of Sales");
    sheet.set("F1", "Grand Total - Sum of Sales");
    sheet.set("G1", "Grand Total - Count of Sales");

    sheet.set("A2", "East");
    sheet.set("B2", 100.0);
    sheet.set("C2", 1.0);
    sheet.set("D2", 150.0);
    sheet.set("E2", 1.0);
    sheet.set("F2", 250.0);
    sheet.set("G2", 2.0);

    sheet.set("A3", "West");
    sheet.set("B3", 200.0);
    sheet.set("C3", 1.0);
    sheet.set("D3", 250.0);
    sheet.set("E3", 1.0);
    sheet.set("F3", 450.0);
    sheet.set("G3", 2.0);

    sheet.set("A4", "Grand Total");
    sheet.set("B4", 300.0);
    sheet.set("C4", 2.0);
    sheet.set("D4", 400.0);
    sheet.set("E4", 2.0);
    sheet.set("F4", 700.0);
    sheet.set("G4", 4.0);

    // Select a specific column item + specific value field.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A2, \"Region\", \"East\", \"Product\", \"A\")"),
        Value::Number(100.0)
    );
    assert_eq!(
        sheet.eval(
            "=GETPIVOTDATA(\"Count of Sales\", A2, \"Region\", \"East\", \"Product\", \"B\")"
        ),
        Value::Number(1.0)
    );

    // Select a row total for a specific value field.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Count of Sales\", A2, \"Region\", \"East\")"),
        Value::Number(2.0)
    );

    // Select a column total (no row criteria).
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A2, \"Product\", \"A\")"),
        Value::Number(300.0)
    );

    // Select an overall grand total for a specific value field.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Count of Sales\", A2)"),
        Value::Number(4.0)
    );
}

#[test]
fn getpivotdata_uses_registry_for_column_fields_and_multiple_values() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotTable, PivotValue,
        SubtotalPosition, ValueField,
    };
    use formula_model::{CellRef, Range};

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    let source = vec![
        pv_row(&[
            "Region".into(),
            "Product".into(),
            "Sales".into(),
            "Units".into(),
        ]),
        pv_row(&["East".into(), "A".into(), 100.into(), 1.into()]),
        pv_row(&["East".into(), "B".into(), 150.into(), 2.into()]),
        pv_row(&["West".into(), "A".into(), 200.into(), 3.into()]),
        pv_row(&["West".into(), "B".into(), 250.into(), 4.into()]),
    ];

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![PivotField::new("Product")],
        value_fields: vec![
            ValueField {
                source_field: "Sales".into(),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            },
            ValueField {
                source_field: "Units".into(),
                name: "Sum of Units".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            },
        ],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: true,
        },
    };

    let pivot = PivotTable::new("PivotTable1", &source, cfg).expect("create pivot");
    let result = pivot.calculate().expect("calculate pivot");

    // Register this pivot as if it were rendered starting at A1.
    let start = CellRef::new(0, 0);
    let end = CellRef::new(
        start.row + result.data.len() as u32 - 1,
        start.col + result.data[0].len() as u32 - 1,
    );
    let destination = Range::new(start, end);

    let mut sheet = TestSheet::new();
    sheet.register_pivot_table(destination, pivot);

    // Column field + value field.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\", \"Product\", \"A\")"),
        Value::Number(100.0)
    );
    // Multiple value fields.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Units\", A1, \"Region\", \"West\", \"Product\", \"B\")"),
        Value::Number(4.0)
    );
    // Partial criteria should return the corresponding subtotal (sum across columns).
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Number(250.0)
    );
    // Grand total.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1)"),
        Value::Number(700.0)
    );
}

#[test]
fn getpivotdata_registry_resolves_data_model_field_refs_against_quoted_cache_headers() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotFieldRef, PivotTable,
        PivotValue, SubtotalPosition, ValueField,
    };
    use formula_model::{CellRef, Range};

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    // The worksheet cache stores a quoted DAX-like column header, which `PivotFieldRef` will parse
    // into a `DataModelColumn` ref. The canonical/display name for that ref is unquoted, so pivot
    // registry matching must be able to map it back to the cache's actual field name.
    let source = vec![
        pv_row(&["'Dim Product'[Category]".into(), "Sales".into()]),
        pv_row(&["East".into(), 100.into()]),
        pv_row(&["West".into(), 200.into()]),
        pv_row(&["East".into(), 150.into()]),
    ];

    let cfg = PivotConfig {
        row_fields: vec![PivotField {
            source_field: PivotFieldRef::from_unstructured("'Dim Product'[Category]"),
            sort_order: Default::default(),
            manual_sort: None,
        }],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".into(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: false,
        },
    };

    let pivot = PivotTable::new("PivotTable1", &source, cfg).expect("create pivot");
    let result = pivot.calculate().expect("calculate pivot");

    // Register this pivot as if it were rendered starting at A1.
    let start = CellRef::new(0, 0);
    let end = CellRef::new(
        start.row + result.data.len() as u32 - 1,
        start.col + result.data[0].len() as u32 - 1,
    );
    let destination = Range::new(start, end);

    let mut sheet = TestSheet::new();
    sheet.register_pivot_table(destination, pivot);

    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Dim Product[Category]\", \"East\")"),
        Value::Number(250.0)
    );
}

#[test]
fn getpivotdata_registry_resolves_data_model_measure_refs_against_unbracketed_cache_headers() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotFieldRef, PivotTable,
        PivotValue, SubtotalPosition, ValueField,
    };
    use formula_model::{CellRef, Range};

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    // The pivot config uses a Data Model measure ref (which has a canonical string form of
    // `[Total Sales]`), but the worksheet/pivot-cache header stores the measure as a plain string
    // (`Total Sales`). Pivot registry matching should still be able to map the measure ref to the
    // cache column.
    let source = vec![
        pv_row(&["Region".into(), "Total Sales".into()]),
        pv_row(&["East".into(), 100.into()]),
        pv_row(&["West".into(), 200.into()]),
        pv_row(&["East".into(), 150.into()]),
    ];

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: PivotFieldRef::DataModelMeasure("Total Sales".to_string()),
            name: "Total Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: false,
        },
    };

    let pivot = PivotTable::new("PivotTable1", &source, cfg).expect("create pivot");
    let result = pivot.calculate().expect("calculate pivot");

    // Register this pivot as if it were rendered starting at A1.
    let start = CellRef::new(0, 0);
    let end = CellRef::new(
        start.row + result.data.len() as u32 - 1,
        start.col + result.data[0].len() as u32 - 1,
    );
    let destination = Range::new(start, end);

    let mut sheet = TestSheet::new();
    sheet.register_pivot_table(destination, pivot);

    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Total Sales\", A1, \"Region\", \"East\")"),
        Value::Number(250.0)
    );
}

#[test]
fn getpivotdata_registry_is_case_insensitive_for_unicode_fields_items_and_values() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotTable, PivotValue,
        SubtotalPosition, ValueField,
    };
    use formula_model::{CellRef, Range};

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    let source = vec![
        pv_row(&["Straße".into(), "Sales".into()]),
        pv_row(&["Maß".into(), 10.into()]),
    ];

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Straße")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".into(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: false,
        },
    };

    let pivot = PivotTable::new("PivotTable1", &source, cfg).expect("create pivot");
    let result = pivot.calculate().expect("calculate pivot");

    // Register this pivot as if it were rendered starting at A1.
    let start = CellRef::new(0, 0);
    let end = CellRef::new(
        start.row + result.data.len() as u32 - 1,
        start.col + result.data[0].len() as u32 - 1,
    );
    let destination = Range::new(start, end);

    let mut sheet = TestSheet::new();
    sheet.register_pivot_table(destination, pivot);

    // Uses Unicode-aware uppercasing: ß -> SS.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"SUM OF SALES\", A1, \"STRASSE\", \"MASS\")"),
        Value::Number(10.0)
    );
}

#[test]
fn getpivotdata_falls_back_to_scan_when_pivot_not_registered() {
    let mut sheet = TestSheet::new();

    // Simulated pivot output (tabular layout, 1 row field, 1 value field).
    sheet.set("A1", "Region");
    sheet.set("B1", "Sum of Sales");
    sheet.set("A2", "East");
    sheet.set("B2", 250.0);
    sheet.set("A3", "West");
    sheet.set("B3", 450.0);
    sheet.set("A4", "Grand Total");
    sheet.set("B4", 700.0);

    // No pivot registry entry was registered for this range, so GETPIVOTDATA should fall back to
    // the legacy grid-scanning heuristics.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", B2, \"Region\", \"East\")"),
        Value::Number(250.0)
    );
}

#[test]
fn getpivotdata_tracks_dynamic_dependency_on_pivot_destination_via_registry() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotTable, PivotValue,
        SubtotalPosition, ValueField,
    };
    use formula_model::{CellRef, Range};

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    let destination = Range::new(CellRef::new(0, 0), CellRef::new(1, 1)); // A1:B2

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".into(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: true,
        },
    };

    let source_v1 = vec![
        pv_row(&["Region".into(), "Sales".into()]),
        pv_row(&["East".into(), 100.into()]),
        pv_row(&["West".into(), 200.into()]),
    ];
    let pivot_v1 = PivotTable::new("PivotTable1", &source_v1, cfg.clone()).expect("create pivot");

    let mut sheet = TestSheet::new();
    sheet.register_pivot_table(destination, pivot_v1);
    sheet.set("A1", "Region"); // stable anchor cell; do not mutate across refreshes.

    // Evaluate once so the dependency graph captures GETPIVOTDATA's dynamic reference to the full
    // pivot destination.
    sheet.set_formula(
        "C1",
        "=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")",
    );
    sheet.recalc();
    assert_eq!(sheet.get("C1"), Value::Number(100.0));

    // Simulate a pivot refresh:
    // - Update the registered pivot metadata (new cache values).
    // - Change a value cell within the pivot destination range (but *not* the pivot_table argument
    //   cell) to trigger dependency propagation.
    let source_v2 = vec![
        pv_row(&["Region".into(), "Sales".into()]),
        pv_row(&["East".into(), 150.into()]),
        pv_row(&["West".into(), 200.into()]),
    ];
    let pivot_v2 = PivotTable::new("PivotTable1", &source_v2, cfg).expect("create pivot");
    sheet.register_pivot_table(destination, pivot_v2);

    // This edit should cause `C1` to be marked dirty only if GETPIVOTDATA recorded a dynamic
    // dependency on the full pivot destination range.
    sheet.set("B2", 150.0);
    sheet.recalc();
    assert_eq!(sheet.get("C1"), Value::Number(150.0));
}

#[test]
fn getpivotdata_registry_replaces_entry_when_destination_changes_for_same_pivot_id() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotTable, PivotValue,
        SubtotalPosition, ValueField,
    };
    use formula_model::{CellRef, Range};

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".into(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: true,
        },
    };

    let source = vec![
        pv_row(&["Region".into(), "Sales".into()]),
        pv_row(&["East".into(), 100.into()]),
        pv_row(&["West".into(), 200.into()]),
    ];

    // Register a pivot with a destination that includes B2.
    let mut pivot_v1 = PivotTable::new("PivotTable1", &source, cfg.clone()).expect("create pivot");
    pivot_v1.id = "pivot-stable-id".to_string();
    let destination_v1 = Range::new(CellRef::new(0, 0), CellRef::new(1, 1)); // A1:B2

    let mut sheet = TestSheet::new();
    sheet.register_pivot_table(destination_v1, pivot_v1);
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", B2, \"Region\", \"East\")"),
        Value::Number(100.0)
    );

    // Refresh the same logical pivot id with a different destination footprint that no longer
    // contains B2.
    let mut pivot_v2 = PivotTable::new("PivotTable1", &source, cfg).expect("create pivot");
    pivot_v2.id = "pivot-stable-id".to_string();
    let destination_v2 = Range::new(CellRef::new(0, 0), CellRef::new(0, 0)); // A1 only
    sheet.register_pivot_table(destination_v2, pivot_v2);

    // A cell that is outside the new destination should no longer resolve to the pivot registry
    // entry (and should fall back to scan-based heuristics, which fail here because the sheet has
    // no rendered pivot output).
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", B2, \"Region\", \"East\")"),
        Value::Error(ErrorKind::Ref)
    );

    // A cell inside the updated destination should still resolve via registry.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Number(100.0)
    );
}

#[test]
fn getpivotdata_registry_shifts_destination_after_insert_rows() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotTable, PivotValue,
        SubtotalPosition, ValueField,
    };
    use formula_engine::EditOp;
    use formula_model::{CellRef, Range};

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".into(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: true,
        },
    };

    let source = vec![
        pv_row(&["Region".into(), "Sales".into()]),
        pv_row(&["East".into(), 100.into()]),
        pv_row(&["West".into(), 200.into()]),
    ];

    let mut pivot = PivotTable::new("PivotTable1", &source, cfg).expect("create pivot");
    pivot.id = "pivot-stable-id".to_string();

    let destination = Range::new(CellRef::new(0, 0), CellRef::new(1, 1)); // A1:B2
    let mut sheet = TestSheet::new();
    sheet.register_pivot_table(destination, pivot);
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Number(100.0)
    );

    // Inserting rows above the pivot output shifts the output down; the pivot registry destination
    // should shift accordingly so GETPIVOTDATA doesn't resolve against the old range.
    sheet.apply_operation(EditOp::InsertRows {
        sheet: "Sheet1".to_string(),
        row: 0,
        count: 1,
    });

    // A1 is now outside the shifted destination (A2:B3). With no rendered pivot grid, the scan
    // fallback fails with #REF!.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Error(ErrorKind::Ref)
    );
    // A2 is inside the shifted destination and should still resolve via registry.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A2, \"Region\", \"East\")"),
        Value::Number(100.0)
    );
}

#[test]
fn getpivotdata_registry_shifts_destination_after_move_range() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotTable, PivotValue,
        SubtotalPosition, ValueField,
    };
    use formula_engine::EditOp;
    use formula_model::{CellRef, Range};

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".into(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: true,
        },
    };

    let source = vec![
        pv_row(&["Region".into(), "Sales".into()]),
        pv_row(&["East".into(), 100.into()]),
        pv_row(&["West".into(), 200.into()]),
    ];

    let mut pivot = PivotTable::new("PivotTable1", &source, cfg).expect("create pivot");
    pivot.id = "pivot-stable-id".to_string();

    let destination = Range::new(CellRef::new(0, 0), CellRef::new(1, 1)); // A1:B2
    let mut sheet = TestSheet::new();
    sheet.register_pivot_table(destination, pivot);
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Number(100.0)
    );

    // Move the pivot destination range to C3:D4.
    sheet.apply_operation(EditOp::MoveRange {
        sheet: "Sheet1".to_string(),
        src: destination,
        dst_top_left: CellRef::new(2, 2), // C3
    });

    // A1 is no longer inside the moved destination (C3:D4).
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", C3, \"Region\", \"East\")"),
        Value::Number(100.0)
    );
}

#[test]
fn getpivotdata_registry_shifts_destination_after_insert_cells_shift_right() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotTable, PivotValue,
        SubtotalPosition, ValueField,
    };
    use formula_engine::EditOp;
    use formula_model::{CellRef, Range};

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".into(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: true,
        },
    };

    let source = vec![
        pv_row(&["Region".into(), "Sales".into()]),
        pv_row(&["East".into(), 100.into()]),
        pv_row(&["West".into(), 200.into()]),
    ];

    let mut pivot = PivotTable::new("PivotTable1", &source, cfg).expect("create pivot");
    pivot.id = "pivot-stable-id".to_string();

    let destination = Range::new(CellRef::new(0, 0), CellRef::new(1, 1)); // A1:B2
    let mut sheet = TestSheet::new();
    sheet.register_pivot_table(destination, pivot);

    // Insert a single column worth of cells at A1:A2, shifting the pivot output right by one column.
    sheet.apply_operation(EditOp::InsertCellsShiftRight {
        sheet: "Sheet1".to_string(),
        range: Range::new(CellRef::new(0, 0), CellRef::new(1, 0)), // A1:A2
    });

    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", B1, \"Region\", \"East\")"),
        Value::Number(100.0)
    );
}

#[test]
fn getpivotdata_registry_shifts_destination_after_delete_cells_shift_up() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotTable, PivotValue,
        SubtotalPosition, ValueField,
    };
    use formula_engine::EditOp;
    use formula_model::{CellRef, Range};

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".into(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: true,
        },
    };

    let source = vec![
        pv_row(&["Region".into(), "Sales".into()]),
        pv_row(&["East".into(), 100.into()]),
        pv_row(&["West".into(), 200.into()]),
    ];

    let mut pivot = PivotTable::new("PivotTable1", &source, cfg).expect("create pivot");
    pivot.id = "pivot-stable-id".to_string();

    // Register a pivot destination starting at A2 so deleting cells above it shifts it upward.
    let destination = Range::new(CellRef::new(1, 0), CellRef::new(2, 1)); // A2:B3
    let mut sheet = TestSheet::new();
    sheet.register_pivot_table(destination, pivot);

    // A1 is outside the destination initially.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Error(ErrorKind::Ref)
    );

    // Delete the row above the pivot in columns A:B; this shifts the pivot up by one row (A1:B2).
    sheet.apply_operation(EditOp::DeleteCellsShiftUp {
        sheet: "Sheet1".to_string(),
        range: Range::new(CellRef::new(0, 0), CellRef::new(0, 1)), // A1:B1
    });

    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Number(100.0)
    );
}

#[test]
fn getpivotdata_registry_shifts_destination_after_delete_cells_shift_left() {
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotTable, PivotValue,
        SubtotalPosition, ValueField,
    };
    use formula_engine::EditOp;
    use formula_model::{CellRef, Range};

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".into(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: true,
        },
    };

    let source = vec![
        pv_row(&["Region".into(), "Sales".into()]),
        pv_row(&["East".into(), 100.into()]),
        pv_row(&["West".into(), 200.into()]),
    ];

    let mut pivot = PivotTable::new("PivotTable1", &source, cfg).expect("create pivot");
    pivot.id = "pivot-stable-id".to_string();

    // Register a pivot destination starting at B1 so deleting cells to its left shifts it left.
    let destination = Range::new(CellRef::new(0, 1), CellRef::new(1, 2)); // B1:C2
    let mut sheet = TestSheet::new();
    sheet.register_pivot_table(destination, pivot);

    // A1 is outside the destination initially.
    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Error(ErrorKind::Ref)
    );

    // Delete the column to the left of the pivot in rows 1..2; this shifts the pivot left by one
    // column (A1:B2).
    sheet.apply_operation(EditOp::DeleteCellsShiftLeft {
        sheet: "Sheet1".to_string(),
        range: Range::new(CellRef::new(0, 0), CellRef::new(1, 0)), // A1:A2
    });

    assert_eq!(
        sheet.eval("=GETPIVOTDATA(\"Sum of Sales\", A1, \"Region\", \"East\")"),
        Value::Number(100.0)
    );
}
