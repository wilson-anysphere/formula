use formula_engine::locale::ValueLocaleConfig;
use formula_engine::sort_filter::{
    apply_autofilter_to_outline, apply_autofilter_to_outline_with_value_locale,
};
use formula_model::{
    CellRef, CellValue, FilterColumn, FilterCriterion, FilterJoin, FilterValue, NumberComparison,
    OpaqueCustomFilter, Outline, Range, SheetAutoFilter, TextMatch, TextMatchKind, Worksheet,
};

#[test]
fn autofilter_updates_outline_filter_hidden_flags_and_can_be_cleared() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_value(CellRef::new(0, 0), CellValue::String("Name".into()));
    sheet.set_value(CellRef::new(1, 0), CellValue::String("Alice".into()));
    sheet.set_value(CellRef::new(2, 0), CellValue::String("Bob".into()));

    let range = Range::from_a1("A1:A3").unwrap();

    let filter = SheetAutoFilter {
        range,
        filter_columns: vec![FilterColumn {
            col_id: 0,
            join: FilterJoin::Any,
            criteria: vec![FilterCriterion::TextMatch(TextMatch {
                kind: TextMatchKind::Contains,
                pattern: "ali".into(),
                case_sensitive: false,
            })],
            values: Vec::new(),
            raw_xml: Vec::new(),
        }],
        sort_state: None,
        raw_xml: Vec::new(),
    };

    let mut outline = Outline::default();
    let result =
        apply_autofilter_to_outline(&sheet, &mut outline, range, Some(&filter)).expect("filter");

    // Row 1 is header (A1) and is never hidden by AutoFilter.
    assert!(!outline.rows.entry(1).hidden.filter);

    // Row 2 ("Alice") remains visible, row 3 ("Bob") is hidden by filter.
    assert_eq!(result.hidden_sheet_rows, vec![2]);
    assert!(!outline.rows.entry(2).hidden.filter);
    assert!(outline.rows.entry(3).hidden.filter);

    // Clearing the filter removes filter hidden flags but preserves the outline map
    // shape (entries should drop back to defaults).
    let cleared = apply_autofilter_to_outline(&sheet, &mut outline, range, None).expect("filter");
    assert_eq!(cleared.hidden_sheet_rows, Vec::<usize>::new());
    assert!(!outline.rows.entry(3).hidden.filter);
}

#[test]
fn autofilter_evaluates_negative_text_ops_from_opaque_custom() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_value(CellRef::new(0, 0), CellValue::String("Name".into()));
    sheet.set_value(CellRef::new(1, 0), CellValue::String("Alice".into()));
    sheet.set_value(CellRef::new(2, 0), CellValue::String("Bob".into()));

    let range = Range::from_a1("A1:A3").unwrap();

    // Negative text operators are preserved as `OpaqueCustom` in the model so they can round-trip
    // through XLSX, but the engine still interprets and evaluates them.
    let filter = SheetAutoFilter {
        range,
        filter_columns: vec![FilterColumn {
            col_id: 0,
            join: FilterJoin::Any,
            criteria: vec![FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                operator: "doesNotContain".into(),
                value: Some("ali".into()),
            })],
            values: Vec::new(),
            raw_xml: Vec::new(),
        }],
        sort_state: None,
        raw_xml: Vec::new(),
    };

    let mut outline = Outline::default();
    let result =
        apply_autofilter_to_outline(&sheet, &mut outline, range, Some(&filter)).expect("filter");

    // Header row is never hidden.
    assert_eq!(result.visible_rows[0], true);

    // `doesNotContain("ali")` hides "Alice" but leaves "Bob" visible.
    assert_eq!(result.visible_rows, vec![true, false, true]);
    assert_eq!(result.hidden_sheet_rows, vec![1]);
    assert!(outline.rows.entry(2).hidden.filter);
    assert!(!outline.rows.entry(3).hidden.filter);
}

#[test]
fn autofilter_preserves_user_hidden_rows() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_value(CellRef::new(0, 0), CellValue::String("Val".into()));
    sheet.set_value(CellRef::new(1, 0), CellValue::String("a".into()));
    sheet.set_value(CellRef::new(2, 0), CellValue::String("b".into()));

    let range = Range::from_a1("A1:A3").unwrap();

    let filter = SheetAutoFilter {
        range,
        filter_columns: vec![FilterColumn {
            col_id: 0,
            join: FilterJoin::Any,
            criteria: vec![FilterCriterion::Equals(FilterValue::Text("a".into()))],
            values: Vec::new(),
            raw_xml: Vec::new(),
        }],
        sort_state: None,
        raw_xml: Vec::new(),
    };

    let mut outline = Outline::default();
    outline.rows.entry_mut(3).hidden.user = true;

    apply_autofilter_to_outline(&sheet, &mut outline, range, Some(&filter)).expect("filter");

    // Row 3 is hidden both by user and by filter.
    assert!(outline.rows.entry(3).hidden.user);
    assert!(outline.rows.entry(3).hidden.filter);

    // Clearing the filter keeps user hidden.
    apply_autofilter_to_outline(&sheet, &mut outline, range, None).expect("filter");
    assert!(outline.rows.entry(3).hidden.user);
    assert!(!outline.rows.entry(3).hidden.filter);
}

#[test]
fn autofilter_blanks_does_not_treat_errors_as_blank() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_value(CellRef::new(0, 0), CellValue::String("Val".into()));
    sheet.set_value(
        CellRef::new(1, 0),
        CellValue::Error(formula_model::ErrorValue::Div0),
    );
    sheet.set_value(CellRef::new(2, 0), CellValue::Empty);

    let range = Range::from_a1("A1:A3").unwrap();
    let filter = SheetAutoFilter {
        range,
        filter_columns: vec![FilterColumn {
            col_id: 0,
            join: FilterJoin::Any,
            criteria: vec![FilterCriterion::Blanks],
            values: Vec::new(),
            raw_xml: Vec::new(),
        }],
        sort_state: None,
        raw_xml: Vec::new(),
    };

    let mut outline = Outline::default();
    let result =
        apply_autofilter_to_outline(&sheet, &mut outline, range, Some(&filter)).expect("filter");

    // Row 2 is an error and should not match the Blanks criterion; row 3 is blank and should
    // remain visible.
    assert_eq!(result.visible_rows, vec![true, false, true]);
    assert_eq!(result.hidden_sheet_rows, vec![1]);
    assert!(outline.rows.entry(2).hidden.filter);
    assert!(!outline.rows.entry(3).hidden.filter);
}

#[test]
fn autofilter_with_value_locale_parses_text_numbers() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_value(CellRef::new(0, 0), CellValue::String("Val".into()));
    sheet.set_value(CellRef::new(1, 0), CellValue::String("1,10".into()));
    sheet.set_value(CellRef::new(2, 0), CellValue::String("1,2".into()));

    let range = Range::from_a1("A1:A3").unwrap();

    let filter = SheetAutoFilter {
        range,
        filter_columns: vec![FilterColumn {
            col_id: 0,
            join: FilterJoin::Any,
            criteria: vec![FilterCriterion::Number(NumberComparison::LessThan(1.15))],
            values: Vec::new(),
            raw_xml: Vec::new(),
        }],
        sort_state: None,
        raw_xml: Vec::new(),
    };

    let mut outline = Outline::default();
    let result = apply_autofilter_to_outline_with_value_locale(
        &sheet,
        &mut outline,
        range,
        Some(&filter),
        ValueLocaleConfig::de_de(),
    )
    .expect("filter");

    // Row 1 is header and should always be visible.
    assert_eq!(result.visible_rows, vec![true, true, false]);
    assert_eq!(result.hidden_sheet_rows, vec![2]);
    assert!(!outline.rows.entry(2).hidden.filter);
    assert!(outline.rows.entry(3).hidden.filter);
}

#[test]
fn autofilter_with_value_locale_parses_legacy_value_list_numbers() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_value(CellRef::new(0, 0), CellValue::String("Val".into()));
    sheet.set_value(CellRef::new(1, 0), CellValue::Number(1.1));
    sheet.set_value(CellRef::new(2, 0), CellValue::Number(1.2));

    let range = Range::from_a1("A1:A3").unwrap();

    // Use the legacy `values` list (no criteria) to ensure locale-aware parsing is applied during
    // model->engine filter conversion.
    let filter = SheetAutoFilter {
        range,
        filter_columns: vec![FilterColumn {
            col_id: 0,
            join: FilterJoin::Any,
            criteria: Vec::new(),
            values: vec!["1,10".into()],
            raw_xml: Vec::new(),
        }],
        sort_state: None,
        raw_xml: Vec::new(),
    };

    let mut outline = Outline::default();
    let result = apply_autofilter_to_outline_with_value_locale(
        &sheet,
        &mut outline,
        range,
        Some(&filter),
        ValueLocaleConfig::de_de(),
    )
    .expect("filter");

    // Header visible, 1.1 row visible, 1.2 row hidden.
    assert_eq!(result.visible_rows, vec![true, true, false]);
    assert_eq!(result.hidden_sheet_rows, vec![2]);
    assert!(!outline.rows.entry(2).hidden.filter);
    assert!(outline.rows.entry(3).hidden.filter);
}
