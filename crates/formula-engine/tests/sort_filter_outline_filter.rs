use formula_engine::sort_filter::{
    apply_autofilter_to_outline, AutoFilter, ColumnFilter, FilterCriterion, FilterJoin, FilterValue,
    TextMatch, TextMatchKind,
};
use formula_model::{CellRef, CellValue, Outline, Range, Worksheet};
use std::collections::BTreeMap;

#[test]
fn autofilter_updates_outline_filter_hidden_flags_and_can_be_cleared() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_value(CellRef::new(0, 0), CellValue::String("Name".into()));
    sheet.set_value(CellRef::new(1, 0), CellValue::String("Alice".into()));
    sheet.set_value(CellRef::new(2, 0), CellValue::String("Bob".into()));

    let range = Range::from_a1("A1:A3").unwrap();

    let filter = AutoFilter {
        range: formula_engine::sort_filter::parse_a1_range("A1:A3").unwrap(),
        columns: BTreeMap::from([(
            0,
            ColumnFilter {
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::TextMatch(TextMatch {
                    kind: TextMatchKind::Contains,
                    pattern: "ali".into(),
                    case_sensitive: false,
                })],
            },
        )]),
    };

    let mut outline = Outline::default();
    let result = apply_autofilter_to_outline(&sheet, &mut outline, range, Some(&filter));

    // Row 1 is header (A1) and is never hidden by AutoFilter.
    assert!(!outline.rows.entry(1).hidden.filter);

    // Row 2 ("Alice") remains visible, row 3 ("Bob") is hidden by filter.
    assert_eq!(result.hidden_sheet_rows, vec![2]);
    assert!(!outline.rows.entry(2).hidden.filter);
    assert!(outline.rows.entry(3).hidden.filter);

    // Clearing the filter removes filter hidden flags but preserves the outline map
    // shape (entries should drop back to defaults).
    let cleared = apply_autofilter_to_outline(&sheet, &mut outline, range, None);
    assert_eq!(cleared.hidden_sheet_rows, Vec::<usize>::new());
    assert!(!outline.rows.entry(3).hidden.filter);
}

#[test]
fn autofilter_preserves_user_hidden_rows() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_value(CellRef::new(0, 0), CellValue::String("Val".into()));
    sheet.set_value(CellRef::new(1, 0), CellValue::String("a".into()));
    sheet.set_value(CellRef::new(2, 0), CellValue::String("b".into()));

    let range = Range::from_a1("A1:A3").unwrap();

    let filter = AutoFilter {
        range: formula_engine::sort_filter::parse_a1_range("A1:A3").unwrap(),
        columns: BTreeMap::from([(
            0,
            ColumnFilter {
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Equals(FilterValue::Text("a".into()))],
            },
        )]),
    };

    let mut outline = Outline::default();
    outline.rows.entry_mut(3).hidden.user = true;

    apply_autofilter_to_outline(&sheet, &mut outline, range, Some(&filter));

    // Row 3 is hidden both by user and by filter.
    assert!(outline.rows.entry(3).hidden.user);
    assert!(outline.rows.entry(3).hidden.filter);

    // Clearing the filter keeps user hidden.
    apply_autofilter_to_outline(&sheet, &mut outline, range, None);
    assert!(outline.rows.entry(3).hidden.user);
    assert!(!outline.rows.entry(3).hidden.filter);
}

